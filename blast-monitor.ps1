<#
.SYNOPSIS
    Standalone Blast Session Performance Monitor
    Collects Horizon Blast metrics from Windows Performance Counters,
    stores historical data, and serves a real-time web dashboard.

.DESCRIPTION
    Single-file, zero-dependency PowerShell script that runs on any
    Windows VDI endpoint with Horizon Agent installed.

    Metrics: Input Lag, RTT, Jitter, Packet Loss, Estimated Bandwidth,
    FPS, Encoder, Encoder CPU, System CPU, Transport Protocol.

    Double-click or run directly - auto-bypasses execution policy and
    self-elevates to admin if needed (HttpListener requires it).

.PARAMETER Port
    HTTP listener port (default: 8888)

.PARAMETER IntervalSeconds
    Collection interval in seconds (default: 5)

.PARAMETER ImportFile
    Optional CSV/JSON file to load at startup for historical viewing

.PARAMETER Stop
    Send stop signal to a running instance on the specified port

.EXAMPLE
    .\blast-monitor.ps1
    .\blast-monitor.ps1 -Port 9090
    .\blast-monitor.ps1 -Stop
    .\blast-monitor.ps1 -ImportFile "session_2026-03-23.json"
#>
param(
    [int]$Port = 8888,
    [int]$IntervalSeconds = 5,
    [string]$ImportFile = "",
    [string]$OutputDir = "",
    [int]$FlushIntervalSeconds = 60,
    [string]$WatchFile = "",
    [switch]$Stop
)

# (No admin elevation needed - TcpListener works without admin)

# ============================================================
# CONSTANTS
# ============================================================
$RetentionHours = 8
$TimeZoneId = "Turkey Standard Time"
$script:TZ = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)

function Get-GMTPlus3 {
    [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $script:TZ)
}

function Format-Timestamp([DateTime]$dt) {
    $dt.ToString("yyyy-MM-ddTHH:mm:ss+03:00")
}

# ============================================================
# ENCODER MAP (from BlastEncoderMap.cs - Horizon Agent 2412)
# ============================================================
$script:EncoderMap = @(
    "none"                          # 0
    "startup"                       # 1
    "static"                        # 2
    "adaptive"                      # 3
    "BlastCodec"                    # 4
    "BlastCodec Lossless"           # 5
    "h264 4:2:0"                    # 6
    "h264 4:4:4"                    # 7
    "hevc 4:2:0"                    # 8
    "hevc 4:4:4"                    # 9
    "NVIDIA NvEnc H264 4:2:0"      # 10
    "NVIDIA NvEnc H264 4:4:4"      # 11
    "NVIDIA NvEnc HEVC 4:2:0"      # 12
    "NVIDIA NvEnc HEVC 4:4:4"      # 13
    "NVIDIA NvEnc HEVC HDR 4:2:0"  # 14
    "NVIDIA NvEnc HEVC HDR 4:4:4"  # 15
    "NVIDIA NvEnc AV1 4:2:0"       # 16
    "NVIDIA NvEnc AV1 4:4:4"       # 17
    "NVIDIA NvEnc AV1 HDR 4:2:0"   # 18
    "NVIDIA NvEnc AV1 HDR 4:4:4"   # 19
    "Switch (Text)"                 # 20
    "Switch (Video)"                # 21
    "Intel H264 SW 4:2:0"          # 22
    "Intel H264 SW 4:4:4"          # 23
    "Intel H264 HW 4:2:0"          # 24
    "Intel H264 HW 4:4:4"          # 25
    "Intel HEVC SW 4:2:0"          # 26
    "Intel HEVC SW 4:4:4"          # 27
    "Intel HEVC HW 4:2:0"          # 28
    "Intel HEVC HW 4:4:4"          # 29
    "Intel AV1 SW 4:2:0"           # 30
    "Intel AV1 SW 4:4:4"           # 31
    "Intel AV1 HW 4:2:0"           # 32
    "Intel AV1 HW 4:4:4"           # 33
    "AMD H264 SW 4:2:0"            # 34
    "AMD H264 SW 4:4:4"            # 35
    "AMD H264 HW 4:2:0"            # 36
    "AMD H264 HW 4:4:4"            # 37
    "AMD HEVC SW 4:2:0"            # 38
    "AMD HEVC SW 4:4:4"            # 39
    "AMD HEVC HW 4:2:0"            # 40
    "AMD HEVC HW 4:4:4"            # 41
    "AMD AV1 SW 4:2:0"             # 42
    "AMD AV1 SW 4:4:4"             # 43
    "AMD AV1 HW 4:2:0"             # 44
    "AMD AV1 HW 4:4:4"             # 45
)

# ============================================================
# DATA STORE
# ============================================================
$script:Samples = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:Running = $true
$script:SessionInfo = @{
    Username     = ""
    ComputerName = $env:COMPUTERNAME
    StartTime    = ""
    EncoderName  = ""
    Transport    = ""
    Tunneled     = $false
}

# ============================================================
# PERFORMANCE COUNTER STATE
# ============================================================
$script:CpuCounter = $null
$script:BlastSessionCategory = $null
$script:BlastImagingCategory = $null
$script:BlastSessionInstance = $null
$script:BlastImagingInstance = $null
$script:InputDelayCategory = $null

# Encoder CPU delta tracking
$script:EncProcId = 0
$script:EncoderPrevCpuTime = [TimeSpan]::Zero
$script:EncoderPrevTimestamp = [DateTime]::MinValue
$script:CachedEncoderCpu = $null
$script:CachedEncoderProcess = $null  # Reuse process handle with Refresh()

# Byte counter delta tracking (cumulative -> KB/s)
$script:PrevBytes = @{
    ImagingTx = $null; AudioTx = $null
    SessionRx = $null; SessionTx = $null
}
$script:PrevBytesTime = [DateTime]::MinValue

# ============================================================
# COUNTER INITIALIZATION
# ============================================================
function Initialize-Counters {
    # System CPU - needs priming
    try {
        $script:CpuCounter = New-Object System.Diagnostics.PerformanceCounter("Processor", "% Processor Time", "_Total")
        $null = $script:CpuCounter.NextValue()
    } catch {
        Write-Host "  [WARN] System CPU counter unavailable: $_" -ForegroundColor Yellow
    }

    # Disk counters - need priming (rate counters return 0 on first read)
    try {
        $script:DiskQueueCounter = New-Object System.Diagnostics.PerformanceCounter("PhysicalDisk", "Avg. Disk Queue Length", "_Total")
        $script:DiskReadCounter = New-Object System.Diagnostics.PerformanceCounter("PhysicalDisk", "Avg. Disk sec/Read", "_Total")
        $script:DiskWriteCounter = New-Object System.Diagnostics.PerformanceCounter("PhysicalDisk", "Avg. Disk sec/Write", "_Total")
        $null = $script:DiskQueueCounter.NextValue()
        $null = $script:DiskReadCounter.NextValue()
        $null = $script:DiskWriteCounter.NextValue()
        Write-Host "  [OK] Disk counters primed" -ForegroundColor Green
    } catch {
        Write-Host "  [WARN] Disk counters unavailable: $_" -ForegroundColor Yellow
    }

    # Blast Session Counters - try Horizon first, then VMware
    foreach ($catName in @("Horizon Blast Session Counters", "VMware Blast Session Counters")) {
        try {
            $cat = New-Object System.Diagnostics.PerformanceCounterCategory($catName)
            $instances = $cat.GetInstanceNames()
            if ($instances.Count -gt 0) {
                $script:BlastSessionCategory = $catName
                $script:BlastSessionInstance = $instances[0]
                Write-Host "  [OK] Blast Session: $catName (instance: $($instances[0]))" -ForegroundColor Green
                break
            }
        } catch { }
    }

    if (-not $script:BlastSessionCategory) {
        Write-Host "  [WARN] No active Blast session detected. Will retry on each collection cycle." -ForegroundColor Yellow
    }

    # Blast Imaging Counters
    foreach ($catName in @("Horizon Blast Imaging Counters", "VMware Blast Imaging Counters")) {
        try {
            $cat = New-Object System.Diagnostics.PerformanceCounterCategory($catName)
            $instances = $cat.GetInstanceNames()
            if ($instances.Count -gt 0) {
                $script:BlastImagingCategory = $catName
                $script:BlastImagingInstance = $instances[0]
                Write-Host "  [OK] Blast Imaging: $catName (instance: $($instances[0]))" -ForegroundColor Green
                break
            }
        } catch { }
    }

    # User Input Delay
    try {
        $cat = New-Object System.Diagnostics.PerformanceCounterCategory("User Input Delay per Session")
        $instances = $cat.GetInstanceNames()
        if ($instances.Count -gt 0) {
            $script:InputDelayCategory = "User Input Delay per Session"
            Write-Host "  [OK] Input Delay available" -ForegroundColor Green
        }
    } catch { }
}

# ============================================================
# SAFE COUNTER READ
# ============================================================
function Read-CounterValue([string]$category, [string]$counterName, [string]$instance) {
    try {
        $c = New-Object System.Diagnostics.PerformanceCounter($category, $counterName, $instance)
        $val = $c.NextValue()
        $c.Dispose()
        return [Math]::Round([double]$val, 2)
    } catch {
        return $null
    }
}

# ============================================================
# REFRESH INSTANCES (for dynamic session attach/detach)
# ============================================================
function Refresh-BlastInstances {
    foreach ($catName in @("Horizon Blast Session Counters", "VMware Blast Session Counters")) {
        try {
            $cat = New-Object System.Diagnostics.PerformanceCounterCategory($catName)
            $instances = $cat.GetInstanceNames()
            if ($instances.Count -gt 0) {
                $script:BlastSessionCategory = $catName
                $script:BlastSessionInstance = $instances[0]
                break
            }
        } catch { }
    }

    foreach ($catName in @("Horizon Blast Imaging Counters", "VMware Blast Imaging Counters")) {
        try {
            $cat = New-Object System.Diagnostics.PerformanceCounterCategory($catName)
            $instances = $cat.GetInstanceNames()
            if ($instances.Count -gt 0) {
                $script:BlastImagingCategory = $catName
                $script:BlastImagingInstance = $instances[0]
                break
            }
        } catch { }
    }

    if ($script:InputDelayCategory) {
        try {
            $cat = New-Object System.Diagnostics.PerformanceCounterCategory($script:InputDelayCategory)
            $instances = $cat.GetInstanceNames()
            # Input Delay instance is the session ID - pick active one
        } catch { }
    }
}

# ============================================================
# ENCODER CPU PROBE (process delta - same logic as PerformanceSampler.cs)
# ============================================================
function Probe-EncoderCpu {
    try {
        # Use Windows Performance Counter directly - much more reliable than Process.TotalProcessorTime
        foreach ($procName in @("VMBlastW", "remotemks")) {
            try {
                $counter = Get-Counter -Counter "\Process($procName)\% Processor Time" -ErrorAction Stop
                $rawValue = $counter.CounterSamples[0].CookedValue
                # CookedValue is per-core %, normalize to system-wide %
                $cpuPercent = [Math]::Round($rawValue / [Environment]::ProcessorCount, 2)
                if ($cpuPercent -lt 0) { $cpuPercent = 0 }
                if ($cpuPercent -gt 100) { $cpuPercent = 100 }
                $script:CachedEncoderCpu = $cpuPercent
                return
            } catch {
                # Process not found in counter, try next name
            }
        }
        $script:CachedEncoderCpu = $null
    } catch {
        Write-Host "  [WARN] EncoderCPU probe error: $_" -ForegroundColor Yellow
        $script:CachedEncoderCpu = $null
    }
}

# ============================================================
# COLLECT SAMPLE
# ============================================================
function Collect-Sample {
    $now = Get-GMTPlus3

    # Refresh instances every cycle (handles session attach/detach)
    Refresh-BlastInstances

    # System CPU
    $systemCpu = $null
    if ($script:CpuCounter) {
        try { $systemCpu = [Math]::Round($script:CpuCounter.NextValue(), 1) } catch { }
    }

    # Blast Session Counters
    $rtt = $null; $jitter = $null; $packetLoss = $null; $bandwidth = $null
    $transport = $null
    if ($script:BlastSessionCategory -and $script:BlastSessionInstance) {
        $cat = $script:BlastSessionCategory
        $inst = $script:BlastSessionInstance

        $rtt        = Read-CounterValue $cat "RTT" $inst
        $jitter     = Read-CounterValue $cat "Jitter (Uplink)" $inst
        $packetLoss = Read-CounterValue $cat "Packet Loss (Uplink)" $inst
        $bandwidth  = Read-CounterValue $cat "Estimated Bandwidth (Uplink)" $inst

        # Session bytes (bidirectional)
        $sessionRxBytes = Read-CounterValue $cat "Received Bytes" $inst
        $sessionTxBytes = Read-CounterValue $cat "Transmitted Bytes" $inst

        # Transport derivation
        $tcpTx = Read-CounterValue $cat "Instantaneous Transmitted Bytes over TCP" $inst
        $udpTx = Read-CounterValue $cat "Instantaneous Transmitted Bytes over UDP" $inst
        $tcpRx = Read-CounterValue $cat "Instantaneous Received Bytes over TCP" $inst
        $udpRx = Read-CounterValue $cat "Instantaneous Received Bytes over UDP" $inst
        $hasTcp = ($tcpTx -gt 0) -or ($tcpRx -gt 0)
        $hasUdp = ($udpTx -gt 0) -or ($udpRx -gt 0)
        if ($hasTcp -and $hasUdp) { $transport = "TCP+UDP" }
        elseif ($hasTcp) { $transport = "TCP" }
        elseif ($hasUdp) { $transport = "UDP" }
    }

    # Blast Imaging Counters
    $fps = $null; $encoderType = $null; $encoderName = $null
    $dirtyFps = $null; $imagingTxBytes = $null
    if ($script:BlastImagingCategory -and $script:BlastImagingInstance) {
        $fps = Read-CounterValue $script:BlastImagingCategory "Frames per second" $script:BlastImagingInstance
        $dirtyFps = Read-CounterValue $script:BlastImagingCategory "Dirty frames per second" $script:BlastImagingInstance
        $imagingTxBytes = Read-CounterValue $script:BlastImagingCategory "Transmitted Bytes" $script:BlastImagingInstance
        $rawEncoder = Read-CounterValue $script:BlastImagingCategory "Encoder Type" $script:BlastImagingInstance
        if ($null -ne $rawEncoder) {
            $encoderType = [int]$rawEncoder
            if ($encoderType -ge 0 -and $encoderType -lt $script:EncoderMap.Count) {
                $encoderName = $script:EncoderMap[$encoderType]
            } else {
                $encoderName = "Unknown ($encoderType)"
            }
        }
    }

    # Blast Audio Counters
    $audioTxBytes = $null
    foreach ($audioCat in @("Horizon Blast Audio Counters", "VMware Blast Audio Counters")) {
        try {
            $cat2 = New-Object System.Diagnostics.PerformanceCounterCategory($audioCat)
            $inst2 = $cat2.GetInstanceNames()
            if ($inst2.Count -gt 0) {
                $audioTxBytes = Read-CounterValue $audioCat "Transmitted Bytes" $inst2[0]
                break
            }
        } catch { }
    }

    # Memory (WMI)
    $memPercent = $null
    try {
        $os = Get-WmiObject Win32_OperatingSystem
        $totalMB = [Math]::Round($os.TotalVisibleMemorySize / 1024, 0)
        $freeMB = [Math]::Round($os.FreePhysicalMemory / 1024, 0)
        $memPercent = [Math]::Round(($totalMB - $freeMB) / $totalMB * 100, 1)
    } catch { }

    # Disk metrics (use primed counters)
    $diskQueue = $null; $diskReadLatency = $null; $diskWriteLatency = $null
    if ($script:DiskQueueCounter) {
        try { $diskQueue = [Math]::Round($script:DiskQueueCounter.NextValue(), 2) } catch { }
    }
    if ($script:DiskReadCounter) {
        try { $diskReadLatency = [Math]::Round($script:DiskReadCounter.NextValue() * 1000, 2) } catch { }
    }
    if ($script:DiskWriteCounter) {
        try { $diskWriteLatency = [Math]::Round($script:DiskWriteCounter.NextValue() * 1000, 2) } catch { }
    }

    # Encoder Max FPS (one-time registry read)
    # Encoder Max FPS (one-time — check policy, then config, multiple vendors)
    if ($null -eq $script:EncoderMaxFps) {
        $script:EncoderMaxFps = 30  # default
        $fpsSource = "default"
        # Priority: User Policy > Computer Policy > User Config > Machine Config (Omnissa then VMware)
        $regPaths = @(
            @{ Path = "HKCU:\SOFTWARE\Policies\Omnissa\Horizon\Blast\Config"; Src = "User Policy (Omnissa)" }
            @{ Path = "HKCU:\SOFTWARE\Policies\VMware, Inc.\VMware Blast\Config"; Src = "User Policy (VMware)" }
            @{ Path = "HKLM:\SOFTWARE\Policies\Omnissa\Horizon\Blast\Config"; Src = "Computer Policy (Omnissa)" }
            @{ Path = "HKLM:\SOFTWARE\Policies\VMware, Inc.\VMware Blast\Config"; Src = "Computer Policy (VMware)" }
            @{ Path = "HKCU:\SOFTWARE\Omnissa\Horizon\Blast\Config"; Src = "User Config (Omnissa)" }
            @{ Path = "HKCU:\SOFTWARE\VMware, Inc.\VMware Blast\Config"; Src = "User Config (VMware)" }
            @{ Path = "HKLM:\SOFTWARE\Omnissa\Horizon\Blast\Config"; Src = "Machine Config (Omnissa)" }
            @{ Path = "HKLM:\SOFTWARE\VMware, Inc.\VMware Blast\Config"; Src = "Machine Config (VMware)" }
        )
        foreach ($rp in $regPaths) {
            try {
                $regVal = Get-ItemProperty -Path $rp.Path -Name "EncoderMaxFPS" -ErrorAction Stop
                if ($regVal.EncoderMaxFPS) {
                    $script:EncoderMaxFps = $regVal.EncoderMaxFPS
                    $fpsSource = $rp.Src
                    break
                }
            } catch { }
        }
        Write-Host ("  [INFO] Encoder Max FPS: {0} (source: {1})" -f $script:EncoderMaxFps, $fpsSource) -ForegroundColor Cyan
    }

    # Input Delay - direct Get-Counter call (works in -File mode, not -Command mode)
    $inputDelay = $null
    try {
        $ilPath = '\User Input Delay per Session(*)\Max Input Delay'
        $ilResult = Get-Counter -Counter $ilPath -ErrorAction Stop
        foreach ($s in $ilResult.CounterSamples) {
            if ($s.InstanceName -eq '1') {
                $inputDelay = [Math]::Round($s.CookedValue, 1)
                break
            }
        }
    } catch {
        if (-not $script:InputDelayWarned) {
            Write-Host ("  [WARN] Input Delay: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
            $script:InputDelayWarned = $true
        }
    }

    # Encoder CPU
    Probe-EncoderCpu

    # CPU Ready (VM Processor - CPU Stolen Time from VMware Tools)
    $cpuReady = $null
    try {
        $stolenCounter = Get-Counter -Counter '\VM Processor(_Total)\CPU stolen time' -ErrorAction Stop
        $cpuReady = [Math]::Round($stolenCounter.CounterSamples[0].CookedValue, 1)
    } catch {
        if (-not $script:CpuReadyWarned) {
            Write-Host "  [WARN] CPU Stolen Time unavailable (VMware Tools needed)" -ForegroundColor Yellow
            $script:CpuReadyWarned = $true
        }
    }

    # Compute byte throughput (cumulative -> KB/s delta)
    $nowUtc = [DateTime]::UtcNow
    $imgKBs = $null; $audioKBs = $null; $sessRxKBs = $null; $sessTxKBs = $null
    if ($script:PrevBytesTime -ne [DateTime]::MinValue) {
        $elapsed = ($nowUtc - $script:PrevBytesTime).TotalSeconds
        if ($elapsed -gt 0) {
            if ($null -ne $imagingTxBytes -and $null -ne $script:PrevBytes.ImagingTx) {
                $imgKBs = [Math]::Round(($imagingTxBytes - $script:PrevBytes.ImagingTx) / 1024 / $elapsed, 1)
            }
            if ($null -ne $audioTxBytes -and $null -ne $script:PrevBytes.AudioTx) {
                $audioKBs = [Math]::Round(($audioTxBytes - $script:PrevBytes.AudioTx) / 1024 / $elapsed, 1)
            }
            if ($null -ne $sessionRxBytes -and $null -ne $script:PrevBytes.SessionRx) {
                $sessRxKBs = [Math]::Round(($sessionRxBytes - $script:PrevBytes.SessionRx) / 1024 / $elapsed, 1)
            }
            if ($null -ne $sessionTxBytes -and $null -ne $script:PrevBytes.SessionTx) {
                $sessTxKBs = [Math]::Round(($sessionTxBytes - $script:PrevBytes.SessionTx) / 1024 / $elapsed, 1)
            }
        }
    }
    $script:PrevBytes.ImagingTx = $imagingTxBytes
    $script:PrevBytes.AudioTx = $audioTxBytes
    $script:PrevBytes.SessionRx = $sessionRxBytes
    $script:PrevBytes.SessionTx = $sessionTxBytes
    $script:PrevBytesTime = $nowUtc

    # Build sample
    $sample = [PSCustomObject]@{
        Timestamp      = Format-Timestamp $now
        InputLag       = $inputDelay
        RTT            = $rtt
        Jitter         = $jitter
        PacketLoss     = $packetLoss
        Bandwidth      = $bandwidth
        FPS            = $fps
        DirtyFPS       = $dirtyFps
        EncoderType    = $encoderType
        EncoderName    = $encoderName
        SystemCPU      = $systemCpu
        EncoderCPU     = $script:CachedEncoderCpu
        CpuReady       = $cpuReady
        MemoryPercent  = $memPercent
        DiskQueue      = $diskQueue
        DiskReadMs     = $diskReadLatency
        DiskWriteMs    = $diskWriteLatency
        ImagingKBs     = $imgKBs
        AudioKBs       = $audioKBs
        SessionRxKBs   = $sessRxKBs
        SessionTxKBs   = $sessTxKBs
        Transport      = $transport
    }

    $script:Samples.Add($sample)

    # Update session info
    if ($encoderName) { $script:SessionInfo.EncoderName = $encoderName }
    if ($transport)   { $script:SessionInfo.Transport = $transport }
    $script:SessionInfo.EncoderMaxFps = $script:EncoderMaxFps

    # Prune old samples (> retention hours)
    $cutoff = (Get-GMTPlus3).AddHours(-$RetentionHours)
    $cutoffStr = Format-Timestamp $cutoff
    while ($script:Samples.Count -gt 0 -and $script:Samples[0].Timestamp -lt $cutoffStr) {
        $script:Samples.RemoveAt(0)
    }
}

# ============================================================
# USERNAME DETECTION
# ============================================================
function Get-SessionUsername {
    try {
        $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs.UserName) { return $cs.UserName }
    } catch { }
    return $env:USERNAME
}

# ============================================================
# TCP HTTP SERVER (no admin required)
# ============================================================
function Send-TcpResponse($client, [int]$statusCode, [string]$statusText, [string]$contentType, [byte[]]$body, [string]$extraHeaders = "") {
    $header = "HTTP/1.1 $statusCode $statusText`r`nContent-Type: $contentType`r`nContent-Length: $($body.Length)`r`nAccess-Control-Allow-Origin: *`r`nConnection: close`r`n$extraHeaders`r`n"
    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
    $stream = $client.GetStream()
    $stream.Write($headerBytes, 0, $headerBytes.Length)
    $stream.Write($body, 0, $body.Length)
    $stream.Flush()
    $client.Close()
}

function Send-JsonResponse($client, $data, [int]$statusCode = 200) {
    $json = $data | ConvertTo-Json -Depth 10 -Compress
    $body = [System.Text.Encoding]::UTF8.GetBytes($json)
    Send-TcpResponse $client $statusCode "OK" "application/json; charset=utf-8" $body
}

function Send-HtmlResponse($client, [string]$html) {
    $body = [System.Text.Encoding]::UTF8.GetBytes($html)
    Send-TcpResponse $client 200 "OK" "text/html; charset=utf-8" $body
}

function Send-FileResponse($client, [string]$content, [string]$contentType, [string]$filename) {
    $body = [System.Text.Encoding]::UTF8.GetBytes($content)
    $extra = "Content-Disposition: attachment; filename=`"$filename`"`r`n"
    Send-TcpResponse $client 200 "OK" "$contentType; charset=utf-8" $body $extra
}

function Send-404($client) {
    $body = [System.Text.Encoding]::UTF8.GetBytes("Not Found")
    Send-TcpResponse $client 404 "Not Found" "text/plain" $body
}

function Parse-HttpRequest($client) {
    $stream = $client.GetStream()
    $stream.ReadTimeout = 2000
    $buffer = New-Object byte[] 65536
    try {
        $bytesRead = $stream.Read($buffer, 0, $buffer.Length)
    } catch {
        return $null
    }
    if ($bytesRead -eq 0) { return $null }
    $raw = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $bytesRead)

    # Parse first line: "GET /path HTTP/1.1"
    $lines = $raw -split "`r`n"
    if ($lines.Count -eq 0) { return $null }
    $parts = $lines[0] -split " "
    if ($parts.Count -lt 2) { return $null }

    # Find body (after empty line)
    $bodyStart = $raw.IndexOf("`r`n`r`n")
    $body = ""
    if ($bodyStart -ge 0) { $body = $raw.Substring($bodyStart + 4) }

    return @{
        Method = $parts[0]
        Path   = ($parts[1] -split "\?")[0]
        Query  = if ($parts[1] -match "\?(.+)") { $Matches[1] } else { "" }
        Body   = $body
        Raw    = $raw
    }
}

# ============================================================
# API HANDLERS
# ============================================================
function Handle-ApiCurrent($client) {
    $latest = if ($script:Samples.Count -gt 0) { $script:Samples[$script:Samples.Count - 1] } else { $null }
    $data = @{
        sample      = $latest
        sessionInfo = $script:SessionInfo
        sampleCount = $script:Samples.Count
    }
    Send-JsonResponse $client $data
}

function Handle-ApiHistory($client, [string]$query) {
    $minutes = 0
    if ($query -match "minutes=(\d+)") { $minutes = [int]$Matches[1] }

    if ($minutes -gt 0) {
        $cutoff = Format-Timestamp ((Get-GMTPlus3).AddMinutes(-$minutes))
        $filtered = $script:Samples | Where-Object { $_.Timestamp -ge $cutoff }
        Send-JsonResponse $client @($filtered)
    } else {
        Send-JsonResponse $client @($script:Samples)
    }
}

function Handle-ApiExportCsv($client) {
    $now = (Get-GMTPlus3).ToString("yyyy-MM-dd_HHmm")
    $csv = ($script:Samples | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
    Send-FileResponse $client $csv "text/csv" "blast_session_$now.csv"
}

function Handle-ApiExportJson($client) {
    $now = (Get-GMTPlus3).ToString("yyyy-MM-dd_HHmm")
    $data = @{
        sessionInfo = $script:SessionInfo
        exportTime  = Format-Timestamp (Get-GMTPlus3)
        samples     = @($script:Samples)
    }
    $json = $data | ConvertTo-Json -Depth 10
    Send-FileResponse $client $json "application/json" "blast_session_$now.json"
}

function Handle-ApiImport($client, [string]$body) {
    try {
        # Try JSON first
        $parsed = $body | ConvertFrom-Json -ErrorAction Stop
        if ($parsed.samples) {
            Send-JsonResponse $client @{
                ok          = $true
                format      = "json"
                sessionInfo = $parsed.sessionInfo
                samples     = @($parsed.samples)
                count       = $parsed.samples.Count
            }
        } else {
            throw "No samples array in JSON"
        }
    } catch {
        # Try CSV
        try {
            $lines = $body -split "`n" | Where-Object { $_.Trim() }
            $csvData = $lines -join "`n" | ConvertFrom-Csv -ErrorAction Stop
            Send-JsonResponse $client @{
                ok          = $true
                format      = "csv"
                sessionInfo = $null
                samples     = @($csvData)
                count       = $csvData.Count
            }
        } catch {
            Send-JsonResponse $client @{ ok = $false; error = "Parse failed" } 400
        }
    }
}

function Handle-ApiStop($client) {
    Send-JsonResponse $client @{ ok = $true; message = "Stopping..." }
    $script:Running = $false
}

# ============================================================
# DASHBOARD HTML
# ============================================================
function Get-DashboardHtml {
return @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Blast Session Monitor</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-annotation@3.1.0/dist/chartjs-plugin-annotation.min.js"></script>
<style>
:root{--bg:#0f0f1a;--card:#1a1a2e;--card2:#16213e;--accent:#e94560;--blue:#4a9eff;--green:#00d084;--yellow:#f0b429;--red:#e94560;--text:#e0e0e0;--text2:#888;--border:#2a2a4a}
:root.light{--bg:#f0f2f5;--card:#ffffff;--card2:#f8f9fa;--accent:#e94560;--blue:#2563eb;--green:#16a34a;--yellow:#d97706;--red:#dc2626;--text:#1a1a2e;--text2:#6b7280;--border:#d1d5db}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:var(--bg);color:var(--text);overflow-x:hidden}
.header{background:var(--card);border-bottom:1px solid var(--border);padding:12px 24px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.header h1{font-size:16px;font-weight:600;color:#f47920;letter-spacing:1px}
.header-info{display:flex;gap:16px;font-size:12px;color:var(--text2)}
.header-info .pill{background:var(--card2);padding:3px 10px;border-radius:12px;border:1px solid var(--border)}
.header-info .pill b{color:var(--text)}
.no-session{background:var(--card);border:1px solid var(--yellow);border-radius:8px;padding:16px 24px;margin:16px 24px;text-align:center;color:var(--yellow);font-size:14px}
.main{padding:16px 24px;display:flex;flex-direction:column;gap:16px}
.gauges{display:grid;grid-template-columns:repeat(5,1fr);gap:12px}
.gauge-card{background:var(--card);border-radius:8px;padding:16px 8px 8px;text-align:center;border:1px solid var(--border)}
.gauge-card .label{font-size:11px;color:var(--text2);margin-bottom:4px;text-transform:uppercase;letter-spacing:0.5px}
.gauge-card .value{font-size:28px;font-weight:700;font-variant-numeric:tabular-nums}
.gauge-card .unit{font-size:11px;color:var(--text2)}
.gauge-canvas{width:100%;max-width:160px;margin:0 auto}
.cpu-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
.cpu-card{background:var(--card);border-radius:8px;padding:14px 18px;border:1px solid var(--border)}
.cpu-card .label{font-size:11px;color:var(--text2);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px}
.cpu-bar-bg{height:20px;background:var(--card2);border-radius:4px;overflow:hidden;position:relative}
.cpu-bar{height:100%;border-radius:4px;transition:width 0.5s ease}
.cpu-bar-text{position:absolute;right:8px;top:50%;transform:translateY(-50%);font-size:12px;font-weight:600}
.charts-section{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.chart-card{background:var(--card);border-radius:8px;padding:14px;border:1px solid var(--border)}
.chart-card .title{font-size:12px;color:var(--text2);margin-bottom:8px;text-transform:uppercase;letter-spacing:0.5px;display:flex;justify-content:space-between}
.chart-card .peak{font-size:10px;color:var(--accent)}
.chart-container{height:180px;position:relative}
.conn-row{display:flex;gap:12px;flex-wrap:wrap}
.conn-pill{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:10px 18px;font-size:13px}
.conn-pill .k{color:var(--text2);font-size:11px;text-transform:uppercase;margin-right:6px}
.conn-pill .v{font-weight:600}
.toolbar{background:var(--card);border-top:1px solid var(--border);padding:12px 24px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;position:sticky;bottom:0}
.btn{background:var(--card2);color:var(--text);border:1px solid var(--border);padding:8px 16px;border-radius:6px;font-size:12px;cursor:pointer;transition:all 0.15s}
.btn:hover{border-color:var(--accent);color:var(--accent)}
.btn-danger{border-color:var(--red)}
.btn-danger:hover{background:var(--red);color:#fff}
.import-mode-banner{background:var(--card2);border:1px solid var(--blue);border-radius:8px;padding:10px 20px;margin:0;display:flex;justify-content:space-between;align-items:center;color:var(--blue);font-size:13px}
.time-range{margin-left:auto;display:flex;gap:4px}
.time-range button{background:none;border:1px solid var(--border);color:var(--text2);padding:4px 10px;border-radius:4px;font-size:11px;cursor:pointer}
.time-range button.active{border-color:var(--accent);color:var(--accent)}
.status-dot{width:8px;height:8px;border-radius:50%;display:inline-block;margin-right:6px}
.status-dot.live{background:var(--green);box-shadow:0 0 6px var(--green)}
.status-dot.stale{background:var(--yellow)}
.status-dot.off{background:var(--red)}
@media(max-width:900px){.gauges{grid-template-columns:repeat(3,1fr)}.charts-section{grid-template-columns:1fr}}
@media(max-width:600px){.gauges{grid-template-columns:repeat(2,1fr)}.cpu-row{grid-template-columns:1fr 1fr}}
</style>
</head>
<body>
<div class="header">
    <div style="display:flex;align-items:flex-end"><img alt="APRO" style="height:28px;margin-right:16px" src="data:image/png;base64,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"><h1>BLAST SESSION MONITOR</h1></div>
    <div class="header-info">
        <span class="pill"><span class="status-dot live" id="statusDot"></span><b id="statusText">Connecting...</b></span>
        <span class="pill">User: <b id="hdrUser">-</b></span>
        <span class="pill">Machine: <b id="hdrMachine">-</b></span>
        <span class="pill">Samples: <b id="hdrCount">0</b></span>
    </div>
</div>

<div id="noSession" class="no-session" style="display:none">No active Blast session detected. Waiting for connection...</div>

<div id="importBanner" class="import-mode-banner" style="display:none">
    <span>IMPORTED DATA - viewing historical session</span>
    <button class="btn" onclick="exitImportMode()">Back to Live</button>
</div>

<div class="main">
    <!-- Gauges -->
    <div class="gauges">
        <div class="gauge-card">
            <div class="label">Input Lag</div>
            <canvas id="gaugeInputLag" class="gauge-canvas" width="160" height="100"></canvas>
            <div class="value" id="valInputLag">-</div>
            <div class="unit">ms</div>
        </div>
        <div class="gauge-card">
            <div class="label">RTT</div>
            <canvas id="gaugeRTT" class="gauge-canvas" width="160" height="100"></canvas>
            <div class="value" id="valRTT">-</div>
            <div class="unit">ms</div>
        </div>
        <div class="gauge-card">
            <div class="label">Jitter</div>
            <canvas id="gaugeJitter" class="gauge-canvas" width="160" height="100"></canvas>
            <div class="value" id="valJitter">-</div>
            <div class="unit">ms</div>
        </div>
        <div class="gauge-card">
            <div class="label">Packet Loss</div>
            <canvas id="gaugePacketLoss" class="gauge-canvas" width="160" height="100"></canvas>
            <div class="value" id="valPacketLoss">-</div>
            <div class="unit">%</div>
        </div>
        <div class="gauge-card">
            <div class="label">Bandwidth</div>
            <canvas id="gaugeBandwidth" class="gauge-canvas" width="160" height="100"></canvas>
            <div class="value" id="valBandwidth">-</div>
            <div class="unit">Mbps</div>
        </div>
    </div>

    <!-- CPU Bars -->
    <div class="cpu-row">
        <div class="cpu-card">
            <div class="label">System CPU</div>
            <div class="cpu-bar-bg">
                <div class="cpu-bar" id="barSystemCPU" style="width:0%;background:var(--blue)"></div>
                <div class="cpu-bar-text" id="txtSystemCPU">-</div>
            </div>
        </div>
        <div class="cpu-card">
            <div class="label">Encoder CPU</div>
            <div class="cpu-bar-bg">
                <div class="cpu-bar" id="barEncoderCPU" style="width:0%;background:var(--accent)"></div>
                <div class="cpu-bar-text" id="txtEncoderCPU">-</div>
            </div>
        </div>
        <div class="cpu-card">
            <div class="label">Memory</div>
            <div class="cpu-bar-bg">
                <div class="cpu-bar" id="barMemory" style="width:0%;background:#9b59b6"></div>
                <div class="cpu-bar-text" id="txtMemory">-</div>
            </div>
        </div>
    </div>

    <!-- Time Series Charts -->
    <div style="display:flex;justify-content:flex-end">
        <div class="time-range" id="timeRange">
            <button data-minutes="5">5m</button>
            <button data-minutes="15">15m</button>
            <button data-minutes="30" class="active">30m</button>
            <button data-minutes="60">1h</button>
            <button data-minutes="240">4h</button>
            <button data-minutes="0">All</button>
        </div>
    </div>

    <div class="charts-section">
        <div class="chart-card">
            <div class="title"><span>Input Lag & RTT (ms)</span><span class="peak" id="peakLatency"></span></div>
            <div class="chart-container"><canvas id="chartLatency"></canvas></div>
        </div>
        <div class="chart-card">
            <div class="title"><span>Jitter (ms) & Packet Loss (%)</span><span class="peak" id="peakJitter"></span></div>
            <div class="chart-container"><canvas id="chartJitter"></canvas></div>
        </div>
        <div class="chart-card">
            <div class="title"><span>Bandwidth (Mbps) & FPS</span><span class="peak" id="peakBw"></span></div>
            <div class="chart-container"><canvas id="chartBw"></canvas></div>
        </div>
        <div class="chart-card">
            <div class="title"><span>CPU Usage (%)</span><span class="peak" id="peakCpu"></span></div>
            <div class="chart-container"><canvas id="chartCpu"></canvas></div>
        </div>
        <div class="chart-card">
            <div class="title"><span>FPS vs Dirty FPS</span><span class="peak" id="peakFps"></span></div>
            <div class="chart-container"><canvas id="chartFps"></canvas></div>
        </div>
        <div class="chart-card">
            <div class="title"><span>Memory (%)</span><span class="peak" id="peakMem"></span></div>
            <div class="chart-container"><canvas id="chartMem"></canvas></div>
        </div>
        <div class="chart-card">
            <div class="title"><span>Disk: Queue, Read & Write Latency</span><span class="peak" id="peakDisk"></span></div>
            <div class="chart-container"><canvas id="chartDisk"></canvas></div>
        </div>
        <div class="chart-card" style="grid-column:span 2">
            <div class="title"><span>Channel Bytes (Imaging, Audio, Session)</span><span class="peak" id="peakBytes"></span></div>
            <div class="chart-container"><canvas id="chartBytes"></canvas></div>
        </div>
    </div>

    <!-- Connection Info -->
    <div class="conn-row">
        <div class="conn-pill"><span class="k">Encoder</span><span class="v" id="connEncoder">-</span></div>
        <div class="conn-pill"><span class="k">Max FPS</span><span class="v" id="connMaxFps">-</span></div>
        <div class="conn-pill"><span class="k">Transport</span><span class="v" id="connTransport">-</span></div>
        <div class="conn-pill"><span class="k">FPS</span><span class="v" id="connFps">-</span></div>
        <div class="conn-pill"><span class="k">Dirty FPS</span><span class="v" id="connDirtyFps">-</span></div>
        <div class="conn-pill"><span class="k">CPU Ready</span><span class="v" id="connCpuReady" style="color:#ff6b00">-</span></div>
    </div>
</div>

<div class="toolbar">
    <button class="btn" onclick="exportCSV()">Export CSV</button>
    <button class="btn" onclick="exportJSON()">Export JSON</button>
    <input type="file" id="importInput" accept=".csv,.json" style="display:none" onchange="handleImport(this)">
    <button class="btn" onclick="document.getElementById('importInput').click()">Import File</button>
    <button class="btn btn-danger" onclick="stopCollector()">Stop Collector</button>
    <span style="flex:1"></span>
    <button class="btn" id="themeToggle" onclick="toggleTheme()" title="Toggle light/dark theme">☀️ Light</button>
</div>

<script>
// ============================================================
// STATE
// ============================================================
let liveMode = true;
let importedSamples = null;
let importedSessionInfo = null;
let selectedMinutes = 30;
let charts = {};
let gaugeCtxs = {};
let lastData = null;

const GAUGE_CONFIG = {
    InputLag:   { min:0, max:500, thresholds:[50,150],  unit:'ms',   invert:false },
    RTT:        { min:0, max:300, thresholds:[50,150],  unit:'ms',   invert:false },
    Jitter:     { min:0, max:100, thresholds:[10,30],   unit:'ms',   invert:false },
    PacketLoss: { min:0, max:10,  thresholds:[1,5],     unit:'%',    invert:false },
    Bandwidth:  { min:0, max:200, thresholds:[50,100],  unit:'Mbps', invert:true  }
};

let COLORS = {
    accent:'#e94560', blue:'#4a9eff', green:'#00d084',
    yellow:'#f0b429', text2:'#888', grid:'#2a2a4a', gaugebg:'#2a2a4a'
};
const DARK_COLORS = {
    accent:'#e94560', blue:'#4a9eff', green:'#00d084',
    yellow:'#f0b429', text2:'#888', grid:'#2a2a4a', gaugebg:'#2a2a4a'
};
const LIGHT_COLORS = {
    accent:'#dc2626', blue:'#2563eb', green:'#16a34a',
    yellow:'#d97706', text2:'#6b7280', grid:'#e5e7eb', gaugebg:'#e5e7eb'
};

function toggleTheme() {
    const root = document.documentElement;
    const isLight = root.classList.toggle('light');
    COLORS = isLight ? {...LIGHT_COLORS} : {...DARK_COLORS};
    const btn = document.getElementById('themeToggle');
    btn.textContent = isLight ? '\uD83C\uDF19 Dark' : '\u2600\uFE0F Light';
    localStorage.setItem('hat-theme', isLight ? 'light' : 'dark');
    // Re-render charts with new colors
    Object.values(charts).forEach(c => {
        if (c && c.options) {
            c.options.scales.x.ticks.color = COLORS.text2;
            c.options.scales.x.grid.color = COLORS.grid;
            Object.keys(c.options.scales).forEach(k => {
                if (k !== 'x') {
                    c.options.scales[k].ticks.color = COLORS.text2;
                    c.options.scales[k].grid.color = COLORS.grid;
                }
            });
            c.update('none');
        }
    });
    // Re-draw gauges
    updateGauges(lastData);
}

// Restore saved theme
(function() {
    if (localStorage.getItem('hat-theme') === 'light') {
        document.documentElement.classList.add('light');
        COLORS = {...LIGHT_COLORS};
        const btn = document.getElementById('themeToggle');
        if (btn) btn.textContent = '\uD83C\uDF19 Dark';
    }
})();

// ============================================================
// GAUGE DRAWING (canvas-based semicircle)
// ============================================================
function drawGauge(canvasId, value, config) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const w = canvas.width, h = canvas.height;
    const cx = w/2, cy = h - 5, r = Math.min(w/2, h) - 10;

    ctx.clearRect(0,0,w,h);

    // Background arc
    ctx.beginPath();
    ctx.arc(cx, cy, r, Math.PI, 0, false);
    ctx.lineWidth = 12;
    ctx.strokeStyle = COLORS.gaugebg;
    ctx.stroke();

    if (value === null || value === undefined) return;

    // Clamp value
    let pct = Math.max(0, Math.min(1, (value - config.min) / (config.max - config.min)));

    // Color based on thresholds
    let color;
    if (config.invert) {
        if (value >= config.thresholds[1]) color = COLORS.green;
        else if (value >= config.thresholds[0]) color = COLORS.yellow;
        else color = COLORS.accent;
    } else {
        if (value <= config.thresholds[0]) color = COLORS.green;
        else if (value <= config.thresholds[1]) color = COLORS.yellow;
        else color = COLORS.accent;
    }

    // Value arc
    const startAngle = Math.PI;
    const endAngle = Math.PI + (Math.PI * pct);
    ctx.beginPath();
    ctx.arc(cx, cy, r, startAngle, endAngle, false);
    ctx.lineWidth = 12;
    ctx.strokeStyle = color;
    ctx.lineCap = 'round';
    ctx.stroke();
}

// ============================================================
// CHART SETUP
// ============================================================
function createChart(canvasId, datasets, yAxes) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    const scales = { x: { type:'category', ticks:{color:COLORS.text2, maxTicksLimit:8, font:{size:10}}, grid:{color:COLORS.grid} } };
    yAxes.forEach((a,i) => {
        scales[a.id] = {
            position: i===0?'left':'right',
            ticks:{color:a.color||COLORS.text2, font:{size:10}},
            grid:{color: i===0?COLORS.grid:'transparent'},
            title:{display:!!a.label, text:a.label||'', color:COLORS.text2, font:{size:10}}
        };
    });
    return new Chart(ctx, {
        type:'line',
        data:{ labels:[], datasets },
        options:{
            responsive:true, maintainAspectRatio:false,
            animation:{duration:300},
            interaction:{mode:'index', intersect:false},
            plugins:{
                legend:{display:true, position:'top', labels:{color:COLORS.text2, boxWidth:10, font:{size:10}}},
                tooltip:{
                    backgroundColor:'#1a1a2eee', titleColor:'#e0e0e0', bodyColor:'#e0e0e0',
                    borderColor:'#2a2a4a', borderWidth:1, padding:10,
                    callbacks:{
                        title: function(items){ return items[0]?.label || ''; }
                    }
                }
            },
            scales
        }
    });
}

function initCharts() {
    charts.latency = createChart('chartLatency', [
        {label:'Input Lag', data:[], borderColor:COLORS.accent, backgroundColor:COLORS.accent+'22', fill:true, tension:0.3, pointRadius:0, yAxisID:'y1'},
        {label:'RTT', data:[], borderColor:COLORS.blue, backgroundColor:COLORS.blue+'22', fill:true, tension:0.3, pointRadius:0, yAxisID:'y1'}
    ], [{id:'y1', color:COLORS.text2, label:'ms'}]);

    charts.jitter = createChart('chartJitter', [
        {label:'Jitter', data:[], borderColor:COLORS.yellow, backgroundColor:COLORS.yellow+'22', fill:true, tension:0.3, pointRadius:0, yAxisID:'y1'},
        {label:'Pkt Loss', data:[], borderColor:COLORS.accent, backgroundColor:COLORS.accent+'22', fill:true, tension:0.3, pointRadius:0, yAxisID:'y2'}
    ], [{id:'y1', color:COLORS.yellow, label:'ms'}, {id:'y2', color:COLORS.accent, label:'%'}]);

    charts.bw = createChart('chartBw', [
        {label:'Bandwidth', data:[], borderColor:COLORS.blue, backgroundColor:COLORS.blue+'22', fill:true, tension:0.3, pointRadius:0, yAxisID:'y1'},
        {label:'FPS', data:[], borderColor:COLORS.green, backgroundColor:COLORS.green+'22', fill:true, tension:0.3, pointRadius:0, yAxisID:'y2'}
    ], [{id:'y1', color:COLORS.blue, label:'Mbps'}, {id:'y2', color:COLORS.green, label:'fps'}]);

    charts.cpu = createChart('chartCpu', [
        {label:'System CPU', data:[], borderColor:COLORS.blue, backgroundColor:COLORS.blue+'22', fill:true, tension:0.3, pointRadius:0, yAxisID:'y1'},
        {label:'Encoder CPU', data:[], borderColor:COLORS.accent, backgroundColor:COLORS.accent+'22', fill:true, tension:0.3, pointRadius:0, yAxisID:'y1'},
        {label:'CPU Ready', data:[], borderColor:'#ff6b00', backgroundColor:'#ff6b0022', fill:false, tension:0.3, pointRadius:0, borderDash:[5,3], yAxisID:'y2'}
    ], [{id:'y1', color:COLORS.text2, label:'%'}, {id:'y2', color:'#ff6b00', label:'ms (stolen)'}]);

    charts.fps = createChart('chartFps', [
        {label:'FPS', data:[], borderColor:COLORS.green, backgroundColor:COLORS.green+'22', fill:true, tension:0.3, pointRadius:0, yAxisID:'y1'},
        {label:'Dirty FPS', data:[], borderColor:'#e67e22', backgroundColor:'#e67e2222', fill:false, tension:0.3, pointRadius:0, borderDash:[4,2], yAxisID:'y1'}
    ], [{id:'y1', color:COLORS.text2, label:'fps'}]);

    charts.mem = createChart('chartMem', [
        {label:'Memory %', data:[], borderColor:'#9b59b6', backgroundColor:'#9b59b622', fill:true, tension:0.3, pointRadius:0, yAxisID:'y1'}
    ], [{id:'y1', color:'#9b59b6', label:'%'}]);

    charts.disk = createChart('chartDisk', [
        {label:'Queue', data:[], borderColor:'#1abc9c', backgroundColor:'#1abc9c22', fill:true, tension:0.3, pointRadius:0, yAxisID:'y1'},
        {label:'Read ms', data:[], borderColor:'#3498db', backgroundColor:'#3498db22', fill:false, tension:0.3, pointRadius:0, yAxisID:'y2'},
        {label:'Write ms', data:[], borderColor:'#e74c3c', backgroundColor:'#e74c3c22', fill:false, tension:0.3, pointRadius:0, yAxisID:'y2'}
    ], [{id:'y1', color:'#1abc9c', label:'queue'}, {id:'y2', color:'#e74c3c', label:'ms'}]);

    charts.bytes = createChart('chartBytes', [
        {label:'Imaging TX', data:[], borderColor:'#e94560', backgroundColor:'#e9456022', fill:true, tension:0.3, pointRadius:0, yAxisID:'y1'},
        {label:'Audio TX', data:[], borderColor:'#f39c12', backgroundColor:'#f39c1222', fill:false, tension:0.3, pointRadius:0, yAxisID:'y1'},
        {label:'Session RX', data:[], borderColor:'#2ecc71', backgroundColor:'#2ecc7122', fill:false, tension:0.3, pointRadius:0, borderDash:[4,2], yAxisID:'y1'},
        {label:'Session TX', data:[], borderColor:'#3498db', backgroundColor:'#3498db22', fill:false, tension:0.3, pointRadius:0, borderDash:[4,2], yAxisID:'y1'}
    ], [{id:'y1', color:COLORS.text2, label:'KB/s'}]);
}

// ============================================================
// UPDATE UI
// ============================================================
function formatTime(ts) {
    if (!ts) return '';
    const d = new Date(ts);
    return d.toLocaleTimeString('tr-TR', {hour:'2-digit',minute:'2-digit',second:'2-digit'});
}

function val(v, dec=1) {
    if (v === null || v === undefined || v === '') return null;
    return Math.round(parseFloat(v) * Math.pow(10,dec)) / Math.pow(10,dec);
}

function updateGauges(sample) {
    if (!sample) return;
    const metrics = {
        InputLag: val(sample.InputLag),
        RTT: val(sample.RTT),
        Jitter: val(sample.Jitter),
        PacketLoss: val(sample.PacketLoss, 2),
        Bandwidth: val(sample.Bandwidth) ? val(sample.Bandwidth) / 1000 : null
    };

    for (const [key, cfg] of Object.entries(GAUGE_CONFIG)) {
        const v = metrics[key];
        drawGauge('gauge'+key, v, cfg);
        const el = document.getElementById('val'+key);
        if (el) el.textContent = v !== null ? v : '-';
    }

    // CPU bars
    const sysCpu = val(sample.SystemCPU);
    const encCpu = val(sample.EncoderCPU);
    document.getElementById('barSystemCPU').style.width = (sysCpu||0)+'%';
    document.getElementById('txtSystemCPU').textContent = sysCpu !== null ? sysCpu+'%' : '-';
    document.getElementById('barEncoderCPU').style.width = (encCpu||0)+'%';
    document.getElementById('txtEncoderCPU').textContent = encCpu !== null ? encCpu+'%' : '-';

    // CPU bar color
    document.getElementById('barSystemCPU').style.background = (sysCpu||0)>80?COLORS.accent:(sysCpu||0)>60?COLORS.yellow:COLORS.blue;
    document.getElementById('barEncoderCPU').style.background = (encCpu||0)>80?COLORS.accent:(encCpu||0)>60?COLORS.yellow:'#e94560';

    // Connection info
    document.getElementById('connEncoder').textContent = sample.EncoderName || '-';
    document.getElementById('connTransport').textContent = sample.Transport || '-';
    document.getElementById('connFps').textContent = val(sample.FPS) !== null ? val(sample.FPS) : '-';
    document.getElementById('connDirtyFps').textContent = val(sample.DirtyFPS) !== null ? val(sample.DirtyFPS) : '-';
    document.getElementById('connCpuReady').textContent = val(sample.CpuReady) !== null ? val(sample.CpuReady) + 'ms' : '-';

    // Memory & Disk bars
    const mem = val(sample.MemoryPercent);
    const dq = val(sample.DiskQueue);
    document.getElementById('barMemory').style.width = (mem||0)+'%';
    document.getElementById('txtMemory').textContent = mem !== null ? mem+'%' : '-';
    document.getElementById('barMemory').style.background = (mem||0)>85?'#e94560':(mem||0)>70?'#f0b429':'#9b59b6';
}

function findPeak(samples, key) {
    let max = -Infinity, maxTs = '';
    for (const s of samples) {
        const v = parseFloat(s[key]);
        if (!isNaN(v) && v > max) { max = v; maxTs = s.Timestamp; }
    }
    return max > -Infinity ? { value: Math.round(max*100)/100, time: formatTime(maxTs) } : null;
}

function updateCharts(samples) {
    if (!samples || !samples.length) return;

    const labels = samples.map(s => formatTime(s.Timestamp));
    const get = (key, div=1) => samples.map(s => { const v=val(s[key]); return v!==null?v/div:null; });

    charts.latency.data.labels = labels;
    charts.latency.data.datasets[0].data = get('InputLag');
    charts.latency.data.datasets[1].data = get('RTT');
    charts.latency.update('none');

    charts.jitter.data.labels = labels;
    charts.jitter.data.datasets[0].data = get('Jitter');
    charts.jitter.data.datasets[1].data = get('PacketLoss');
    charts.jitter.update('none');

    charts.bw.data.labels = labels;
    charts.bw.data.datasets[0].data = get('Bandwidth', 1000);
    charts.bw.data.datasets[1].data = get('FPS');
    charts.bw.update('none');

    charts.cpu.data.labels = labels;
    charts.cpu.data.datasets[0].data = get('SystemCPU');
    charts.cpu.data.datasets[1].data = get('EncoderCPU');
    charts.cpu.data.datasets[2].data = get('CpuReady');
    charts.cpu.update('none');

    charts.fps.data.labels = labels;
    charts.fps.data.datasets[0].data = get('FPS');
    charts.fps.data.datasets[1].data = get('DirtyFPS');
    charts.fps.update('none');

    charts.mem.data.labels = labels;
    charts.mem.data.datasets[0].data = get('MemoryPercent');
    charts.mem.update('none');

    charts.disk.data.labels = labels;
    charts.disk.data.datasets[0].data = get('DiskQueue');
    charts.disk.data.datasets[1].data = get('DiskReadMs');
    charts.disk.data.datasets[2].data = get('DiskWriteMs');
    charts.disk.update('none');

    charts.bytes.data.labels = labels;
    charts.bytes.data.datasets[0].data = get('ImagingKBs');
    charts.bytes.data.datasets[1].data = get('AudioKBs');
    charts.bytes.data.datasets[2].data = get('SessionRxKBs');
    charts.bytes.data.datasets[3].data = get('SessionTxKBs');
    charts.bytes.update('none');

    // Peaks
    const peakIL = findPeak(samples,'InputLag');
    const peakRTT = findPeak(samples,'RTT');
    document.getElementById('peakLatency').textContent = peakIL ? `Peak IL: ${peakIL.value}ms @ ${peakIL.time}` : '';

    const peakJ = findPeak(samples,'Jitter');
    document.getElementById('peakJitter').textContent = peakJ ? `Peak: ${peakJ.value}ms @ ${peakJ.time}` : '';

    const peakB = findPeak(samples,'Bandwidth');
    document.getElementById('peakBw').textContent = peakB ? `Peak: ${Math.round(peakB.value/1000)}Mbps @ ${peakB.time}` : '';

    const peakC = findPeak(samples,'SystemCPU');
    document.getElementById('peakCpu').textContent = peakC ? `Peak: ${peakC.value}% @ ${peakC.time}` : '';
}

function updateSessionInfo(info) {
    if (!info) return;
    document.getElementById('hdrUser').textContent = info.Username || '-';
    document.getElementById('hdrMachine').textContent = info.ComputerName || '-';
    if (info.EncoderMaxFps) document.getElementById('connMaxFps').textContent = info.EncoderMaxFps;
}

// ============================================================
// POLLING
// ============================================================
async function poll() {
    if (!liveMode) return;
    try {
        const resp = await fetch('/api/current');
        const data = await resp.json();

        const hasSample = data.sample && data.sample.RTT !== null;
        document.getElementById('noSession').style.display = hasSample ? 'none' : 'block';
        document.getElementById('statusDot').className = 'status-dot ' + (hasSample ? 'live' : 'stale');
        document.getElementById('statusText').textContent = hasSample ? 'Live' : 'No Session';
        document.getElementById('hdrCount').textContent = data.sampleCount || 0;

        lastData = data.sample;
        updateGauges(data.sample);
        updateSessionInfo(data.sessionInfo);

        // History
        const histUrl = selectedMinutes > 0 ? `/api/history?minutes=${selectedMinutes}` : '/api/history';
        const histResp = await fetch(histUrl);
        const history = await histResp.json();
        updateCharts(history);
    } catch(e) {
        document.getElementById('statusDot').className = 'status-dot off';
        document.getElementById('statusText').textContent = 'Disconnected';
    }
}

// ============================================================
// IMPORT / EXPORT
// ============================================================
function exportCSV() { window.location.href = '/api/export/csv'; }
function exportJSON() { window.location.href = '/api/export/json'; }

async function handleImport(input) {
    const file = input.files[0];
    if (!file) return;

    const text = await file.text();
    try {
        const resp = await fetch('/api/import', { method:'POST', body: text });
        const data = await resp.json();
        if (data.ok) {
            importedSamples = data.samples;
            importedSessionInfo = data.sessionInfo;
            liveMode = false;
            document.getElementById('importBanner').style.display = 'flex';
            document.getElementById('statusDot').className = 'status-dot off';
            document.getElementById('statusText').textContent = 'Import Mode';
            if (importedSessionInfo) updateSessionInfo(importedSessionInfo);
            document.getElementById('hdrCount').textContent = data.count;

            // Show last sample in gauges
            if (importedSamples.length > 0) {
                updateGauges(importedSamples[importedSamples.length - 1]);
            }
            updateCharts(importedSamples);
        } else {
            alert('Import failed: ' + (data.error || 'Unknown error'));
        }
    } catch(e) {
        alert('Import error: ' + e.message);
    }
    input.value = '';
}

function exitImportMode() {
    liveMode = true;
    importedSamples = null;
    importedSessionInfo = null;
    document.getElementById('importBanner').style.display = 'none';
    poll();
}

async function stopCollector() {
    if (!confirm('Stop the collector?')) return;
    try {
        await fetch('/api/stop', {method:'POST'});
        document.getElementById('statusDot').className = 'status-dot off';
        document.getElementById('statusText').textContent = 'Stopped';
    } catch(e) {}
}

// ============================================================
// TIME RANGE BUTTONS
// ============================================================
document.getElementById('timeRange').addEventListener('click', e => {
    if (e.target.tagName !== 'BUTTON') return;
    document.querySelectorAll('#timeRange button').forEach(b => b.classList.remove('active'));
    e.target.classList.add('active');
    selectedMinutes = parseInt(e.target.dataset.minutes);
    if (liveMode) poll();
    else if (importedSamples) {
        let filtered = importedSamples;
        if (selectedMinutes > 0 && importedSamples.length > 0) {
            const last = new Date(importedSamples[importedSamples.length-1].Timestamp);
            const cutoff = new Date(last.getTime() - selectedMinutes*60000);
            filtered = importedSamples.filter(s => new Date(s.Timestamp) >= cutoff);
        }
        updateCharts(filtered);
    }
});

// ============================================================
// INIT
// ============================================================
initCharts();
poll();
setInterval(poll, 5000);
</script>
</body>
</html>
'@
}

# ============================================================
# REQUEST ROUTER
# ============================================================
function Handle-Request($client, $req) {
    $route = "$($req.Method) $($req.Path)"
    try {
        switch ($route) {
            "GET /"                { Send-HtmlResponse $client (Get-DashboardHtml) }
"GET /api/current"    { Handle-ApiCurrent $client }
            "GET /api/history"    { Handle-ApiHistory $client $req.Query }
            "GET /api/export/csv" { Handle-ApiExportCsv $client }
            "GET /api/export/json"{ Handle-ApiExportJson $client }
            "POST /api/import"    { Handle-ApiImport $client $req.Body }
            "POST /api/stop"      { Handle-ApiStop $client }
            default               { Send-404 $client }
        }
    } catch {
        try {
            $errBody = [System.Text.Encoding]::UTF8.GetBytes("Internal Error: $_")
            Send-TcpResponse $client 500 "Internal Server Error" "text/plain" $errBody
        } catch { }
    }
}

# ============================================================
# STOP COMMAND
# ============================================================
if ($Stop) {
    try {
        $null = Invoke-RestMethod -Uri "http://localhost:$Port/api/stop" -Method POST -TimeoutSec 5
        Write-Host "Stop signal sent to instance on port $Port" -ForegroundColor Green
    } catch {
        Write-Host "No running instance found on port $Port" -ForegroundColor Yellow
    }
    return
}

# ============================================================
# IMPORT AT STARTUP
# ============================================================
if ($ImportFile -and (Test-Path $ImportFile)) {
    Write-Host "Loading import file: $ImportFile"
    $content = Get-Content $ImportFile -Raw
    $ext = [System.IO.Path]::GetExtension($ImportFile).ToLower()
    try {
        if ($ext -eq ".json") {
            $parsed = $content | ConvertFrom-Json
            if ($parsed.samples) {
                foreach ($s in $parsed.samples) {
                    $script:Samples.Add([PSCustomObject]@{
                        Timestamp   = $s.Timestamp
                        InputLag    = $s.InputLag
                        RTT         = $s.RTT
                        Jitter      = $s.Jitter
                        PacketLoss  = $s.PacketLoss
                        Bandwidth   = $s.Bandwidth
                        FPS         = $s.FPS
                        EncoderType = $s.EncoderType
                        EncoderName = $s.EncoderName
                        SystemCPU   = $s.SystemCPU
                        EncoderCPU  = $s.EncoderCPU
                        Transport   = $s.Transport
                    })
                }
                if ($parsed.sessionInfo) {
                    $script:SessionInfo.Username = $parsed.sessionInfo.Username
                    $script:SessionInfo.ComputerName = $parsed.sessionInfo.ComputerName
                }
                Write-Host "  Loaded $($script:Samples.Count) samples from JSON" -ForegroundColor Green
            }
        } elseif ($ext -eq ".csv") {
            $csvData = $content | ConvertFrom-Csv
            foreach ($row in $csvData) {
                $script:Samples.Add([PSCustomObject]@{
                    Timestamp   = $row.Timestamp
                    InputLag    = if ($row.InputLag) { [double]$row.InputLag } else { $null }
                    RTT         = if ($row.RTT) { [double]$row.RTT } else { $null }
                    Jitter      = if ($row.Jitter) { [double]$row.Jitter } else { $null }
                    PacketLoss  = if ($row.PacketLoss) { [double]$row.PacketLoss } else { $null }
                    Bandwidth   = if ($row.Bandwidth) { [double]$row.Bandwidth } else { $null }
                    FPS         = if ($row.FPS) { [double]$row.FPS } else { $null }
                    EncoderType = if ($row.EncoderType) { [int]$row.EncoderType } else { $null }
                    EncoderName = $row.EncoderName
                    SystemCPU   = if ($row.SystemCPU) { [double]$row.SystemCPU } else { $null }
                    EncoderCPU  = if ($row.EncoderCPU) { [double]$row.EncoderCPU } else { $null }
                    Transport   = $row.Transport
                })
            }
            Write-Host "  Loaded $($script:Samples.Count) samples from CSV" -ForegroundColor Green
        }
    } catch {
        Write-Host "  [ERROR] Failed to load import file: $_" -ForegroundColor Red
    }
}

# ============================================================
# OUTPUT DIR FLUSH FUNCTION
# ============================================================
function Flush-ToOutputDir {
    if (-not $OutputDir) { return }
    try {
        if (-not (Test-Path $OutputDir)) {
            New-Item -ItemType Directory -Path $OutputDir -Force -ErrorAction Stop | Out-Null
        }
        $filePath = Join-Path $OutputDir "blast_live.json"
        $data = @{
            sessionInfo = $script:SessionInfo
            exportTime  = Format-Timestamp (Get-GMTPlus3)
            samples     = @($script:Samples)
        }
        $data | ConvertTo-Json -Depth 10 | Set-Content $filePath -Encoding UTF8 -ErrorAction Stop
    } catch {
        # Silent after first warning — don't spam console
        if (-not $script:FlushWarnShown) {
            Write-Host "  [WARN] Output dir not writable: $OutputDir — data kept in memory only" -ForegroundColor Yellow
            $script:FlushWarnShown = $true
        }
    }
}

# ============================================================
# WATCH FILE FUNCTION (admin mode - read remote JSON)
# ============================================================
function Read-WatchFile {
    if (-not $WatchFile) { return }
    try {
        if (-not (Test-Path $WatchFile)) { return }
        $content = Get-Content $WatchFile -Raw -ErrorAction Stop
        $parsed = $content | ConvertFrom-Json -ErrorAction Stop
        if ($parsed.samples) {
            $script:Samples.Clear()
            foreach ($s in $parsed.samples) {
                $script:Samples.Add($s)
            }
            if ($parsed.sessionInfo) {
                $script:SessionInfo.Username = $parsed.sessionInfo.Username
                $script:SessionInfo.ComputerName = $parsed.sessionInfo.ComputerName
                $script:SessionInfo.EncoderName = $parsed.sessionInfo.EncoderName
                $script:SessionInfo.Transport = $parsed.sessionInfo.Transport
                $script:SessionInfo.EncoderMaxFps = $parsed.sessionInfo.EncoderMaxFps
            }
        }
    } catch { }
}

# ============================================================
# MAIN
# ============================================================
$isWatchMode = [bool]$WatchFile

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
if ($isWatchMode) {
    Write-Host "  BLAST SESSION MONITOR (ADMIN VIEW)" -ForegroundColor Cyan
} else {
    Write-Host "  BLAST SESSION MONITOR" -ForegroundColor Cyan
}
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Port:     $Port"

if ($isWatchMode) {
    Write-Host "  Mode:     Watch (read-only)" -ForegroundColor Yellow
    Write-Host "  Source:   $WatchFile"
} else {
    Write-Host "  Interval: ${IntervalSeconds}s"
    Write-Host "  Timezone: GMT+3 (Turkey)"
    Write-Host "  Retention: ${RetentionHours}h"
    if ($OutputDir) {
        Write-Host "  Output:   $OutputDir (every ${FlushIntervalSeconds}s)" -ForegroundColor Green
    }
}
Write-Host ""

if (-not $isWatchMode) {
    Write-Host "Initializing counters..." -ForegroundColor Gray
    Initialize-Counters
    $script:SessionInfo.Username = Get-SessionUsername
    $script:SessionInfo.StartTime = Format-Timestamp (Get-GMTPlus3)
    Write-Host "  User: $($script:SessionInfo.Username)" -ForegroundColor Gray
} else {
    Write-Host "Waiting for data from: $WatchFile" -ForegroundColor Gray
    Read-WatchFile
    if ($script:Samples.Count -gt 0) {
        Write-Host ("  Loaded {0} samples" -f $script:Samples.Count) -ForegroundColor Green
    }
}

$tcpListener = $null
try {
    $tcpListener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
    $tcpListener.Start()

    Write-Host ""
    Write-Host "  Dashboard: http://localhost:$Port" -ForegroundColor Green
    Write-Host "  Press Ctrl+C to stop" -ForegroundColor Gray
    Write-Host ""

    $lastCollect = [DateTime]::MinValue
    $lastFlush = [DateTime]::MinValue
    $lastWatch = [DateTime]::MinValue

    while ($script:Running) {
        # Handle pending HTTP requests (non-blocking)
        while ($tcpListener.Pending()) {
            try {
                $client = $tcpListener.AcceptTcpClient()
                $client.ReceiveTimeout = 2000
                $client.SendTimeout = 5000
                $req = Parse-HttpRequest $client
                if ($req) {
                    Handle-Request $client $req
                } else {
                    $client.Close()
                }
            } catch {
                Write-Host "  [HTTP] Error: $_" -ForegroundColor DarkYellow
                try { $client.Close() } catch { }
            }
        }

        # Small sleep to avoid busy-wait
        Start-Sleep -Milliseconds 50

        $now = [DateTime]::UtcNow

        if ($isWatchMode) {
            # WATCH MODE: periodically re-read the JSON file
            if (($now - $lastWatch).TotalSeconds -ge 5) {
                $prevCount = $script:Samples.Count
                Read-WatchFile
                $newCount = $script:Samples.Count
                if ($newCount -ne $prevCount) {
                    Write-Host ("  [{0}] Refreshed: {1} samples from {2}" -f (Format-Timestamp (Get-GMTPlus3)), $newCount, $script:SessionInfo.ComputerName) -ForegroundColor DarkGray
                }
                $lastWatch = $now
            }
        } else {
            # COLLECT MODE: gather metrics
            if (($now - $lastCollect).TotalSeconds -ge $IntervalSeconds) {
                try {
                    Collect-Sample
                    $count = $script:Samples.Count
                    $latest = if ($count -gt 0) { $script:Samples[$count-1] } else { $null }
                    if ($latest -and $latest.RTT -ne $null) {
                        if ($latest.Bandwidth) { $bwMbps = [Math]::Round($latest.Bandwidth / 1000, 1) } else { $bwMbps = "-" }
                        if ($null -ne $latest.EncoderCPU) { $encCpu = [string]$latest.EncoderCPU + "%" } else { $encCpu = "-" }
                        if ($latest.EncoderName) { $enc = $latest.EncoderName } else { $enc = "-" }
                        if ($latest.Transport) { $tp = $latest.Transport } else { $tp = "-" }
                        if ($null -ne $latest.InputLag) { $il = [string]$latest.InputLag + "ms" } else { $il = "-" }
                        if ($null -ne $latest.PacketLoss) { $pl = [string]$latest.PacketLoss + "%" } else { $pl = "-" }
                        if ($null -ne $latest.CpuReady) { $rdy = [string]$latest.CpuReady + "ms" } else { $rdy = "-" }
                        Write-Host ("  [{0}] RTT={1}ms Jitter={2}ms IL={3} PL={4} BW={5}Mbps FPS={6} CPU={7}% EncCPU={8} RDY={9} Enc={10} TP={11} [{12} samples]" -f $latest.Timestamp, $latest.RTT, $latest.Jitter, $il, $pl, $bwMbps, $latest.FPS, $latest.SystemCPU, $encCpu, $rdy, $enc, $tp, $count) -ForegroundColor DarkGray
                    } else {
                        Write-Host ("  [{0}] No active Blast session [{1} samples]" -f (Format-Timestamp (Get-GMTPlus3)), $count) -ForegroundColor DarkYellow
                    }
                } catch {
                    Write-Host "  [ERROR] Collection failed: $_" -ForegroundColor Red
                }
                $lastCollect = $now
            }

            # Flush to OutputDir periodically
            if ($OutputDir -and ($now - $lastFlush).TotalSeconds -ge $FlushIntervalSeconds) {
                Flush-ToOutputDir
                $lastFlush = $now
            }
        }
    }
} catch {
    if ($script:Running) {
        Write-Host "ERROR: $_" -ForegroundColor Red
        if ($_.Exception.Message -match "address already in use|access denied") {
            Write-Host "Port $Port might be in use. Try: -Port 9090" -ForegroundColor Yellow
        }
    }
} finally {
    if ($tcpListener) {
        $tcpListener.Stop()
    }
    Write-Host ""
    # Final flush on exit
    if ($OutputDir -and $script:Samples.Count -gt 0) {
        Write-Host "  Final flush to $OutputDir..." -ForegroundColor Gray
        Flush-ToOutputDir
    }
    Write-Host "Monitor stopped. $($script:Samples.Count) samples collected." -ForegroundColor Cyan
}
