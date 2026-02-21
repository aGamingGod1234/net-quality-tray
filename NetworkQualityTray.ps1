[CmdletBinding()]
param(
    [switch]$InstallStartup,
    [switch]$RemoveStartup,
    [switch]$InstallShortcut,
    [switch]$RemoveShortcut,
    [switch]$NoTray,
    [switch]$VerboseLogging,
    [int]$RunForSeconds = 0,
    [switch]$BypassSingleton
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:AppDisplayName = "Network Quality Accesser"
$script:AppTitle = "Network Quality Accesser"
$script:QualityThresholds = [pscustomobject]@{
    HighMinScore = 70
    PoorMinScore = 45
    VeryPoorMinScore = 25
}
$script:TierColorMap = @{}

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Net.Http
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
} catch {
    throw "Windows Forms and Drawing assemblies are required. Run this on Windows PowerShell/PowerShell with desktop support."
}

if (-not ("NativeIconMethods" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class NativeIconMethods
{
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool DestroyIcon(IntPtr handle);
}
"@
}

function Clamp-Value {
    param(
        [double]$Value,
        [double]$Min = 0.0,
        [double]$Max = 1.0
    )
    if ($Value -lt $Min) { return $Min }
    if ($Value -gt $Max) { return $Max }
    return $Value
}

function Get-Mean {
    param([object[]]$Values)
    if (-not $Values -or $Values.Count -eq 0) { return 0.0 }
    $sum = 0.0
    foreach ($value in $Values) {
        $sum += [double]$value
    }
    return $sum / $Values.Count
}

function Get-StdDev {
    param(
        [object[]]$Values,
        [double]$Mean
    )
    if (-not $Values -or $Values.Count -lt 2) { return 0.0 }
    $sum = 0.0
    foreach ($value in $Values) {
        $delta = ([double]$value) - $Mean
        $sum += ($delta * $delta)
    }
    return [math]::Sqrt($sum / ($Values.Count - 1))
}

function Add-WindowSample {
    param(
        [System.Collections.Generic.Queue[double]]$Queue,
        [double]$Value,
        [int]$MaxCount
    )
    if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) { return }
    $Queue.Enqueue([double]$Value)
    while ($Queue.Count -gt $MaxCount) {
        [void]$Queue.Dequeue()
    }
}

function Get-PrimaryNetworkInterface {
    $all = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | Where-Object {
        $_.OperationalStatus -eq [System.Net.NetworkInformation.OperationalStatus]::Up -and
        $_.NetworkInterfaceType -ne [System.Net.NetworkInformation.NetworkInterfaceType]::Loopback -and
        $_.NetworkInterfaceType -ne [System.Net.NetworkInformation.NetworkInterfaceType]::Tunnel -and
        $_.NetworkInterfaceType -ne [System.Net.NetworkInformation.NetworkInterfaceType]::Unknown
    }

    if (-not $all -or $all.Count -eq 0) { return $null }

    $withGateway = @()
    foreach ($iface in $all) {
        try {
            if ($iface.GetIPProperties().GatewayAddresses.Count -gt 0) {
                $withGateway += $iface
            }
        } catch {
            continue
        }
    }

    $candidates = if ($withGateway.Count -gt 0) { $withGateway } else { $all }
    return $candidates | Sort-Object -Property Speed -Descending | Select-Object -First 1
}

function Measure-LatencyMetrics {
    param(
        [string[]]$Hosts,
        [string]$PreferredHost,
        [int]$Samples = 3,
        [int]$TimeoutMs = 900
    )

    $orderedHosts = @()
    if ($PreferredHost) { $orderedHosts += $PreferredHost }
    foreach ($targetHost in $Hosts) {
        if ($targetHost -and ($orderedHosts -notcontains $targetHost)) {
            $orderedHosts += $targetHost
        }
    }

    $ping = [System.Net.NetworkInformation.Ping]::new()
    try {
        $selectedHost = $null
        $firstRtt = $null
        foreach ($targetHost in $orderedHosts) {
            try {
                $reply = $ping.Send($targetHost, $TimeoutMs)
                if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                    $selectedHost = $targetHost
                    $firstRtt = [double]$reply.RoundtripTime
                    break
                }
            } catch {
                continue
            }
        }

        if (-not $selectedHost) {
            return [pscustomobject]@{
                Success  = $false
                Host     = $null
                AvgMs    = [double]::NaN
                JitterMs = [double]::NaN
                LossPct  = 100.0
                Samples  = 0
                Failures = $Samples
            }
        }

        $rtts = [System.Collections.Generic.List[double]]::new()
        $rtts.Add($firstRtt)
        $failures = 0

        for ($i = 1; $i -lt $Samples; $i++) {
            try {
                $reply = $ping.Send($selectedHost, $TimeoutMs)
                if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                    $rtts.Add([double]$reply.RoundtripTime)
                } else {
                    $failures++
                }
            } catch {
                $failures++
            }
        }

        if ($rtts.Count -eq 0) {
            return [pscustomobject]@{
                Success  = $false
                Host     = $selectedHost
                AvgMs    = [double]::NaN
                JitterMs = [double]::NaN
                LossPct  = 100.0
                Samples  = 0
                Failures = $Samples
            }
        }

        $avg = Get-Mean -Values $rtts.ToArray()
        $std = Get-StdDev -Values $rtts.ToArray() -Mean $avg
        $totalSent = $rtts.Count + $failures
        $lossPct = if ($totalSent -gt 0) { ($failures / $totalSent) * 100.0 } else { 100.0 }

        return [pscustomobject]@{
            Success  = $true
            Host     = $selectedHost
            AvgMs    = [math]::Round($avg, 1)
            JitterMs = [math]::Round($std, 1)
            LossPct  = [math]::Round($lossPct, 1)
            Samples  = $totalSent
            Failures = $failures
        }
    } finally {
        $ping.Dispose()
    }
}

function Measure-DownloadProbe {
    param(
        [System.Net.Http.HttpClient]$Client,
        [string]$Endpoint,
        [int]$BytesToRead,
        [int]$MaxDurationMs
    )

    $requestUri = "{0}?bytes={1}" -f $Endpoint, $BytesToRead
    $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $requestUri)
    $response = $null
    $stream = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $bytesRead = 0L

    try {
        $response = $Client.SendAsync($request, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
        if (-not $response.IsSuccessStatusCode) {
            throw "Download probe failed with HTTP $([int]$response.StatusCode)."
        }

        $stream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $buffer = New-Object byte[] 32768
        while ($bytesRead -lt $BytesToRead -and $sw.ElapsedMilliseconds -lt $MaxDurationMs) {
            $read = $stream.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
            $bytesRead += $read
        }

        $sw.Stop()
        if ($bytesRead -lt 20480) {
            throw "Download probe returned too little data ($bytesRead bytes)."
        }

        $seconds = [math]::Max($sw.Elapsed.TotalSeconds, 0.001)
        $mbps = ($bytesRead * 8.0 / 1000000.0) / $seconds

        return [pscustomobject]@{
            Success    = $true
            Mbps       = [math]::Round($mbps, 2)
            Bytes      = $bytesRead
            DurationMs = [int]$sw.ElapsedMilliseconds
            Error      = $null
        }
    } catch {
        $sw.Stop()
        return [pscustomobject]@{
            Success    = $false
            Mbps       = 0.0
            Bytes      = $bytesRead
            DurationMs = [int]$sw.ElapsedMilliseconds
            Error      = $_.Exception.Message
        }
    } finally {
        if ($stream) { $stream.Dispose() }
        if ($response) { $response.Dispose() }
        $request.Dispose()
    }
}

function Measure-UploadProbe {
    param(
        [System.Net.Http.HttpClient]$Client,
        [string]$Endpoint,
        [byte[]]$Buffer,
        [int]$BytesToSend,
        [int]$MaxDurationMs
    )

    $content = [System.Net.Http.ByteArrayContent]::new($Buffer, 0, $BytesToSend)
    $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/octet-stream")
    $response = $null
    $cts = [System.Threading.CancellationTokenSource]::new()
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        $cts.CancelAfter($MaxDurationMs + 2000)
        $response = $Client.PostAsync($Endpoint, $content, $cts.Token).GetAwaiter().GetResult()
        if (-not $response.IsSuccessStatusCode) {
            throw "Upload probe failed with HTTP $([int]$response.StatusCode)."
        }

        [void]$response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
        $sw.Stop()

        $seconds = [math]::Max($sw.Elapsed.TotalSeconds, 0.001)
        $mbps = ($BytesToSend * 8.0 / 1000000.0) / $seconds

        return [pscustomobject]@{
            Success    = $true
            Mbps       = [math]::Round($mbps, 2)
            Bytes      = $BytesToSend
            DurationMs = [int]$sw.ElapsedMilliseconds
            Error      = $null
        }
    } catch {
        $sw.Stop()
        return [pscustomobject]@{
            Success    = $false
            Mbps       = 0.0
            Bytes      = $BytesToSend
            DurationMs = [int]$sw.ElapsedMilliseconds
            Error      = $_.Exception.Message
        }
    } finally {
        if ($response) { $response.Dispose() }
        $content.Dispose()
        $cts.Dispose()
    }
}

function New-EndpointFailoverState {
    param([string[]]$Endpoints)
    if (-not $Endpoints -or $Endpoints.Count -eq 0) {
        throw "At least one endpoint is required for failover state."
    }

    $state = [pscustomobject]@{
        Endpoints     = $Endpoints
        Cursor        = 0
        FailureCount  = @{}
        NextAllowedAt = @{}
    }

    foreach ($endpoint in $Endpoints) {
        $state.FailureCount[$endpoint] = 0
        $state.NextAllowedAt[$endpoint] = [datetime]::MinValue
    }

    return $state
}

function Get-EndpointProbeCandidate {
    param(
        [object]$State,
        [datetime]$Now
    )

    $count = $State.Endpoints.Count
    if ($count -le 0) {
        return [pscustomobject]@{
            Available = $false
            Endpoint  = $null
            Index     = -1
            DelaySec  = 2
        }
    }

    for ($offset = 0; $offset -lt $count; $offset++) {
        $index = ($State.Cursor + $offset) % $count
        $endpoint = $State.Endpoints[$index]
        $nextAt = [datetime]$State.NextAllowedAt[$endpoint]
        if ($Now -ge $nextAt) {
            return [pscustomobject]@{
                Available = $true
                Endpoint  = $endpoint
                Index     = $index
                DelaySec  = 0
            }
        }
    }

    $soonestEndpoint = $State.Endpoints[0]
    $soonestIndex = 0
    $soonestAt = [datetime]$State.NextAllowedAt[$soonestEndpoint]
    for ($i = 1; $i -lt $count; $i++) {
        $endpoint = $State.Endpoints[$i]
        $candidate = [datetime]$State.NextAllowedAt[$endpoint]
        if ($candidate -lt $soonestAt) {
            $soonestAt = $candidate
            $soonestEndpoint = $endpoint
            $soonestIndex = $i
        }
    }

    $delay = [int][math]::Ceiling([math]::Max(1.0, ($soonestAt - $Now).TotalSeconds))
    return [pscustomobject]@{
        Available = $false
        Endpoint  = $soonestEndpoint
        Index     = $soonestIndex
        DelaySec  = $delay
    }
}

function Register-EndpointProbeFailure {
    param(
        [object]$State,
        [int]$Index,
        [datetime]$Now,
        [int]$MaxBackoffSec = 180
    )

    if ($Index -lt 0 -or $Index -ge $State.Endpoints.Count) { return }
    $endpoint = $State.Endpoints[$Index]
    $newFailures = [int]$State.FailureCount[$endpoint] + 1
    $State.FailureCount[$endpoint] = $newFailures

    $exp = [int][math]::Min(7, $newFailures)
    $backoff = [int][math]::Min($MaxBackoffSec, [math]::Pow(2, $exp))
    $jitter = Get-Random -Minimum 0 -Maximum 3
    $State.NextAllowedAt[$endpoint] = $Now.AddSeconds($backoff + $jitter)
    $State.Cursor = ($Index + 1) % $State.Endpoints.Count
}

function Register-EndpointProbeSuccess {
    param(
        [object]$State,
        [int]$Index
    )

    if ($Index -lt 0 -or $Index -ge $State.Endpoints.Count) { return }
    $endpoint = $State.Endpoints[$Index]
    $State.FailureCount[$endpoint] = 0
    $State.NextAllowedAt[$endpoint] = [datetime]::MinValue
    $State.Cursor = $Index
}

function Resolve-DownloadProbeUri {
    param(
        [string]$Endpoint,
        [int]$BytesToRead
    )

    if ($Endpoint -match "\{bytes\}") {
        return $Endpoint -replace "\{bytes\}", [string]$BytesToRead
    }

    if ($Endpoint -match "\?") {
        return "{0}&bytes={1}" -f $Endpoint, $BytesToRead
    }
    return "{0}?bytes={1}" -f $Endpoint, $BytesToRead
}

function Start-DownloadProbeAsync {
    param(
        [System.Net.Http.HttpClient]$Client,
        [object]$EndpointState,
        [int]$BytesToRead,
        [int]$MaxDurationMs,
        [datetime]$Now
    )

    $candidate = Get-EndpointProbeCandidate -State $EndpointState -Now $Now
    if (-not $candidate.Available) {
        return [pscustomobject]@{
            Started  = $false
            DelaySec = $candidate.DelaySec
            State    = $null
        }
    }

    $uri = Resolve-DownloadProbeUri -Endpoint $candidate.Endpoint -BytesToRead $BytesToRead
    $cts = [System.Threading.CancellationTokenSource]::new()
    $cts.CancelAfter($MaxDurationMs + 2000)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $task = $Client.GetAsync($uri, [System.Net.Http.HttpCompletionOption]::ResponseContentRead, $cts.Token)

    return [pscustomobject]@{
        Started  = $true
        DelaySec = 0
        State    = [pscustomobject]@{
            Task          = $task
            Cts           = $cts
            Stopwatch     = $sw
            Endpoint      = $candidate.Endpoint
            EndpointIndex = $candidate.Index
            BytesToRead   = $BytesToRead
        }
    }
}

function Complete-DownloadProbeAsync {
    param([object]$ProbeState)

    if (-not $ProbeState) { return $null }
    $task = $ProbeState.Task
    if (-not $task.IsCompleted) { return $null }

    $response = $null
    try {
        if ($task.IsCanceled) {
            return [pscustomobject]@{
                Completed    = $true
                Success      = $false
                Mbps         = 0.0
                Bytes        = 0
                DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
                Error        = "Download probe canceled."
                Endpoint     = $ProbeState.Endpoint
                EndpointIndex= $ProbeState.EndpointIndex
            }
        }

        if ($task.IsFaulted) {
            $message = if ($task.Exception -and $task.Exception.InnerException) {
                $task.Exception.InnerException.Message
            } elseif ($task.Exception) {
                $task.Exception.Message
            } else {
                "Download probe failed."
            }

            return [pscustomobject]@{
                Completed    = $true
                Success      = $false
                Mbps         = 0.0
                Bytes        = 0
                DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
                Error        = $message
                Endpoint     = $ProbeState.Endpoint
                EndpointIndex= $ProbeState.EndpointIndex
            }
        }

        $response = $task.GetAwaiter().GetResult()
        if (-not $response.IsSuccessStatusCode) {
            return [pscustomobject]@{
                Completed    = $true
                Success      = $false
                Mbps         = 0.0
                Bytes        = 0
                DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
                Error        = "Download probe failed with HTTP $([int]$response.StatusCode)."
                Endpoint     = $ProbeState.Endpoint
                EndpointIndex= $ProbeState.EndpointIndex
            }
        }

        $bytes = $response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
        $bytesRead = [math]::Min([long]$bytes.LongLength, [long]$ProbeState.BytesToRead)
        if ($bytesRead -lt 20480) {
            return [pscustomobject]@{
                Completed    = $true
                Success      = $false
                Mbps         = 0.0
                Bytes        = $bytesRead
                DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
                Error        = "Download probe returned too little data ($bytesRead bytes)."
                Endpoint     = $ProbeState.Endpoint
                EndpointIndex= $ProbeState.EndpointIndex
            }
        }

        $seconds = [math]::Max($ProbeState.Stopwatch.Elapsed.TotalSeconds, 0.001)
        $mbps = ($bytesRead * 8.0 / 1000000.0) / $seconds
        return [pscustomobject]@{
            Completed    = $true
            Success      = $true
            Mbps         = [math]::Round($mbps, 2)
            Bytes        = $bytesRead
            DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
            Error        = $null
            Endpoint     = $ProbeState.Endpoint
            EndpointIndex= $ProbeState.EndpointIndex
        }
    } catch {
        return [pscustomobject]@{
            Completed    = $true
            Success      = $false
            Mbps         = 0.0
            Bytes        = 0
            DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
            Error        = $_.Exception.Message
            Endpoint     = $ProbeState.Endpoint
            EndpointIndex= $ProbeState.EndpointIndex
        }
    } finally {
        if ($ProbeState.Stopwatch) { $ProbeState.Stopwatch.Stop() }
        if ($response) { $response.Dispose() }
        if ($ProbeState.Cts) { $ProbeState.Cts.Dispose() }
    }
}

function Start-UploadProbeAsync {
    param(
        [System.Net.Http.HttpClient]$Client,
        [object]$EndpointState,
        [byte[]]$Buffer,
        [int]$BytesToSend,
        [int]$MaxDurationMs,
        [datetime]$Now
    )

    $candidate = Get-EndpointProbeCandidate -State $EndpointState -Now $Now
    if (-not $candidate.Available) {
        return [pscustomobject]@{
            Started  = $false
            DelaySec = $candidate.DelaySec
            State    = $null
        }
    }

    $content = [System.Net.Http.ByteArrayContent]::new($Buffer, 0, $BytesToSend)
    $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/octet-stream")
    $cts = [System.Threading.CancellationTokenSource]::new()
    $cts.CancelAfter($MaxDurationMs + 2000)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $task = $Client.PostAsync($candidate.Endpoint, $content, $cts.Token)

    return [pscustomobject]@{
        Started  = $true
        DelaySec = 0
        State    = [pscustomobject]@{
            Task          = $task
            Cts           = $cts
            Stopwatch     = $sw
            Endpoint      = $candidate.Endpoint
            EndpointIndex = $candidate.Index
            BytesToSend   = $BytesToSend
            Content       = $content
        }
    }
}

function Complete-UploadProbeAsync {
    param([object]$ProbeState)

    if (-not $ProbeState) { return $null }
    $task = $ProbeState.Task
    if (-not $task.IsCompleted) { return $null }

    $response = $null
    try {
        if ($task.IsCanceled) {
            return [pscustomobject]@{
                Completed    = $true
                Success      = $false
                Mbps         = 0.0
                Bytes        = $ProbeState.BytesToSend
                DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
                Error        = "Upload probe canceled."
                Endpoint     = $ProbeState.Endpoint
                EndpointIndex= $ProbeState.EndpointIndex
            }
        }

        if ($task.IsFaulted) {
            $message = if ($task.Exception -and $task.Exception.InnerException) {
                $task.Exception.InnerException.Message
            } elseif ($task.Exception) {
                $task.Exception.Message
            } else {
                "Upload probe failed."
            }

            return [pscustomobject]@{
                Completed    = $true
                Success      = $false
                Mbps         = 0.0
                Bytes        = $ProbeState.BytesToSend
                DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
                Error        = $message
                Endpoint     = $ProbeState.Endpoint
                EndpointIndex= $ProbeState.EndpointIndex
            }
        }

        $response = $task.GetAwaiter().GetResult()
        if (-not $response.IsSuccessStatusCode) {
            return [pscustomobject]@{
                Completed    = $true
                Success      = $false
                Mbps         = 0.0
                Bytes        = $ProbeState.BytesToSend
                DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
                Error        = "Upload probe failed with HTTP $([int]$response.StatusCode)."
                Endpoint     = $ProbeState.Endpoint
                EndpointIndex= $ProbeState.EndpointIndex
            }
        }

        [void]$response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
        $seconds = [math]::Max($ProbeState.Stopwatch.Elapsed.TotalSeconds, 0.001)
        $mbps = ($ProbeState.BytesToSend * 8.0 / 1000000.0) / $seconds
        return [pscustomobject]@{
            Completed    = $true
            Success      = $true
            Mbps         = [math]::Round($mbps, 2)
            Bytes        = $ProbeState.BytesToSend
            DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
            Error        = $null
            Endpoint     = $ProbeState.Endpoint
            EndpointIndex= $ProbeState.EndpointIndex
        }
    } catch {
        return [pscustomobject]@{
            Completed    = $true
            Success      = $false
            Mbps         = 0.0
            Bytes        = $ProbeState.BytesToSend
            DurationMs   = [int]$ProbeState.Stopwatch.ElapsedMilliseconds
            Error        = $_.Exception.Message
            Endpoint     = $ProbeState.Endpoint
            EndpointIndex= $ProbeState.EndpointIndex
        }
    } finally {
        if ($ProbeState.Stopwatch) { $ProbeState.Stopwatch.Stop() }
        if ($response) { $response.Dispose() }
        if ($ProbeState.Content) { $ProbeState.Content.Dispose() }
        if ($ProbeState.Cts) { $ProbeState.Cts.Dispose() }
    }
}

function Get-IconLevel {
    param([int]$Score)
    if ($Score -ge 80) { return 4 }
    if ($Score -ge 55) { return 3 }
    if ($Score -ge 30) { return 2 }
    if ($Score -ge 10) { return 1 }
    return 0
}

function Get-IconStateKey {
    param(
        [int]$Score,
        [string]$Tier,
        [bool]$Paused = $false
    )
    if ($Paused) { return "paused" }
    if ($Tier -eq "Offline") { return "offline" }
    $level = Get-IconLevel -Score $Score
    return "{0}:{1}" -f $Tier, $level
}

function Get-ConsistencyScore {
    param(
        [System.Collections.Generic.Queue[double]]$DownHistory,
        [System.Collections.Generic.Queue[double]]$UpHistory,
        [System.Collections.Generic.Queue[double]]$SuccessHistory
    )

    $stabilityScores = [System.Collections.Generic.List[double]]::new()

    foreach ($window in @($DownHistory.ToArray(), $UpHistory.ToArray())) {
        if ($window.Count -ge 3) {
            $mean = Get-Mean -Values $window
            if ($mean -gt 0.0) {
                $std = Get-StdDev -Values $window -Mean $mean
                $cv = $std / $mean
                $stabilityScores.Add(1.0 - (Clamp-Value -Value ($cv / 0.9)))
            }
        }
    }

    $base = if ($stabilityScores.Count -gt 0) {
        Get-Mean -Values $stabilityScores.ToArray()
    } else {
        0.65
    }

    $successWindow = $SuccessHistory.ToArray()
    $successRate = if ($successWindow.Count -gt 0) {
        Get-Mean -Values $successWindow
    } else {
        1.0
    }

    return Clamp-Value -Value (($base * 0.7) + ($successRate * 0.3))
}

function Get-LinearScoreHigherBetter {
    param(
        [double]$Value,
        [double]$BadThreshold,
        [double]$GoodThreshold
    )
    if ($Value -le $BadThreshold) { return 0.0 }
    if ($Value -ge $GoodThreshold) { return 1.0 }
    return Clamp-Value -Value (($Value - $BadThreshold) / ($GoodThreshold - $BadThreshold))
}

function Get-LinearScoreLowerBetter {
    param(
        [double]$Value,
        [double]$GoodThreshold,
        [double]$BadThreshold
    )
    if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) { return 0.0 }
    if ($Value -le $GoodThreshold) { return 1.0 }
    if ($Value -ge $BadThreshold) { return 0.0 }
    return Clamp-Value -Value (1.0 - (($Value - $GoodThreshold) / ($BadThreshold - $GoodThreshold)))
}

function Get-QualityTier {
    param(
        [int]$Score,
        [bool]$Offline
    )
    if ($Offline) { return "Offline" }

    $highMin = 70
    $poorMin = 45
    $veryPoorMin = 25
    if ($script:QualityThresholds) {
        $highMin = [int]$script:QualityThresholds.HighMinScore
        $poorMin = [int]$script:QualityThresholds.PoorMinScore
        $veryPoorMin = [int]$script:QualityThresholds.VeryPoorMinScore
    }

    if ($Score -ge $highMin) { return "High" }
    if ($Score -ge $poorMin) { return "Poor" }
    if ($Score -ge $veryPoorMin) { return "VeryPoor" }
    return "Bad"
}

function Get-TierColor {
    param([string]$Tier)
    if ($script:TierColorMap -and $script:TierColorMap.ContainsKey($Tier)) {
        return $script:TierColorMap[$Tier]
    }
    switch ($Tier) {
        "High" { return [System.Drawing.Color]::FromArgb(46, 204, 113) }
        "Poor" { return [System.Drawing.Color]::FromArgb(241, 196, 15) }
        "VeryPoor" { return [System.Drawing.Color]::FromArgb(230, 126, 34) }
        "Bad" { return [System.Drawing.Color]::FromArgb(231, 76, 60) }
        "Offline" { return [System.Drawing.Color]::FromArgb(149, 17, 17) }
        "Paused" { return [System.Drawing.Color]::FromArgb(160, 160, 160) }
        default { return [System.Drawing.Color]::FromArgb(120, 120, 120) }
    }
}

function Compute-QualitySnapshot {
    param(
        [object]$PrimaryInterface,
        [object]$LatencyMetrics,
        [object]$DownProbe,
        [object]$UpProbe,
        [double]$LastGoodDownMbps,
        [datetime]$LastGoodDownAt,
        [double]$LastGoodUpMbps,
        [datetime]$LastGoodUpAt,
        [double]$Consistency,
        [datetime]$Now
    )

    $downAgeSec = [math]::Max(0.0, ($Now - $LastGoodDownAt).TotalSeconds)
    $upAgeSec = [math]::Max(0.0, ($Now - $LastGoodUpAt).TotalSeconds)
    $downDecay = [math]::Max(0.35, 1.0 - ($downAgeSec / 240.0))
    $upDecay = [math]::Max(0.35, 1.0 - ($upAgeSec / 240.0))

    $effectiveDown = if ($DownProbe.Success) {
        [double]$DownProbe.Mbps
    } elseif ($downAgeSec -lt 180.0) {
        $LastGoodDownMbps * $downDecay
    } else {
        0.0
    }

    $effectiveUp = if ($UpProbe.Success) {
        [double]$UpProbe.Mbps
    } elseif ($upAgeSec -lt 180.0) {
        $LastGoodUpMbps * $upDecay
    } else {
        0.0
    }

    $hasRecentThroughput = ($downAgeSec -lt 180.0) -or ($upAgeSec -lt 180.0)
    $offline = (-not $PrimaryInterface) -or (
        (-not $LatencyMetrics.Success) -and
        (-not $hasRecentThroughput) -and
        (-not $DownProbe.Success) -and
        (-not $UpProbe.Success)
    )

    $downScore = Get-LinearScoreHigherBetter -Value $effectiveDown -BadThreshold 3.0 -GoodThreshold 120.0
    $upScore = Get-LinearScoreHigherBetter -Value $effectiveUp -BadThreshold 1.0 -GoodThreshold 40.0
    $latencyScore = Get-LinearScoreLowerBetter -Value $LatencyMetrics.AvgMs -GoodThreshold 25.0 -BadThreshold 250.0
    $jitterScore = Get-LinearScoreLowerBetter -Value $LatencyMetrics.JitterMs -GoodThreshold 6.0 -BadThreshold 90.0
    $lossScore = Get-LinearScoreLowerBetter -Value $LatencyMetrics.LossPct -GoodThreshold 0.0 -BadThreshold 20.0
    $consistencyScore = Clamp-Value -Value $Consistency

    if ($offline) {
        $quality = 0
    } else {
        $quality = (
            ($downScore * 0.32) +
            ($upScore * 0.18) +
            ($latencyScore * 0.20) +
            ($jitterScore * 0.10) +
            ($lossScore * 0.10) +
            ($consistencyScore * 0.10)
        ) * 100.0

        if (-not $DownProbe.Success -or -not $UpProbe.Success) {
            $quality *= 0.92
        }
    }

    $quality = [int][math]::Round([math]::Max(0.0, [math]::Min(100.0, $quality)))
    $tier = Get-QualityTier -Score $quality -Offline $offline
    $linkMbps = if ($PrimaryInterface) { [math]::Round(($PrimaryInterface.Speed / 1000000.0), 1) } else { 0.0 }

    return [pscustomobject]@{
        Timestamp          = $Now
        InterfaceName      = if ($PrimaryInterface) { $PrimaryInterface.Name } else { "Disconnected" }
        InterfaceType      = if ($PrimaryInterface) { $PrimaryInterface.NetworkInterfaceType.ToString() } else { "None" }
        LinkMbps           = $linkMbps
        QualityScore       = $quality
        Tier               = $tier
        Offline            = $offline
        DownloadMbps       = [math]::Round($effectiveDown, 2)
        UploadMbps         = [math]::Round($effectiveUp, 2)
        LatencyMs          = if ($LatencyMetrics.Success) { $LatencyMetrics.AvgMs } else { [double]::NaN }
        JitterMs           = if ($LatencyMetrics.Success) { $LatencyMetrics.JitterMs } else { [double]::NaN }
        LossPct            = if ($LatencyMetrics.Success) { $LatencyMetrics.LossPct } else { 100.0 }
        ConsistencyScore   = [math]::Round($consistencyScore, 2)
        LatencyHost        = if ($LatencyMetrics.Success) { $LatencyMetrics.Host } else { "None" }
        LastDownloadError  = if ($DownProbe.Success) { "" } else { $DownProbe.Error }
        LastUploadError    = if ($UpProbe.Success) { "" } else { $UpProbe.Error }
    }
}

function New-QualityIcon {
    param(
        [int]$Score,
        [string]$Tier,
        [bool]$Paused = $false
    )

    $bitmap = New-Object System.Drawing.Bitmap 16, 16
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $activeBrush = $null
    $inactiveBrush = $null
    $outlinePen = $null
    $markPen = $null

    try {
        $graphics.Clear([System.Drawing.Color]::Transparent)
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

        $activeColor = if ($Paused) { Get-TierColor -Tier "Paused" } else { Get-TierColor -Tier $Tier }
        $activeBrush = [System.Drawing.SolidBrush]::new($activeColor)
        $inactiveBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(75, 85, 95))
        $outlinePen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(140, 90, 90, 90), 1.0)
        $graphics.DrawRectangle($outlinePen, 1, 1, 13, 13)

        if ($Paused) {
            $graphics.FillRectangle($activeBrush, 4, 4, 2, 8)
            $graphics.FillRectangle($activeBrush, 8, 4, 2, 8)
        } elseif ($Tier -eq "Offline") {
            $markPen = [System.Drawing.Pen]::new($activeColor, 1.8)
            $graphics.DrawLine($markPen, 4, 4, 10, 10)
            $graphics.DrawLine($markPen, 10, 4, 4, 10)
        } else {
            $barHeights = @(3, 6, 9, 12)
            $bars = Get-IconLevel -Score $Score
            for ($i = 0; $i -lt 4; $i++) {
                $height = $barHeights[$i]
                $x = 2 + ($i * 3)
                $y = 14 - $height
                $rect = [System.Drawing.Rectangle]::new($x, $y, 2, $height)
                if ($i -lt $bars) {
                    $graphics.FillRectangle($activeBrush, $rect)
                } else {
                    $graphics.FillRectangle($inactiveBrush, $rect)
                }
            }
        }

        $hIcon = $bitmap.GetHicon()
        try {
            return [System.Drawing.Icon]::FromHandle($hIcon).Clone()
        } finally {
            [void][NativeIconMethods]::DestroyIcon($hIcon)
        }
    } finally {
        if ($markPen) { $markPen.Dispose() }
        if ($outlinePen) { $outlinePen.Dispose() }
        if ($activeBrush) { $activeBrush.Dispose() }
        if ($inactiveBrush) { $inactiveBrush.Dispose() }
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Set-NotifyIconTextSafe {
    param(
        [System.Windows.Forms.NotifyIcon]$NotifyIcon,
        [string]$Text
    )
    $safe = if ($Text.Length -gt 63) { $Text.Substring(0, 63) } else { $Text }
    try {
        $NotifyIcon.Text = $safe
    } catch {
        $NotifyIcon.Text = $script:AppDisplayName
    }
}

function Format-SnapshotLine {
    param([object]$Snapshot)
    if (-not $Snapshot) { return "[No Data] Waiting for probes..." }

    $latency = if ([double]::IsNaN([double]$Snapshot.LatencyMs)) { "n/a" } else { "$($Snapshot.LatencyMs) ms" }
    return "[{0:HH:mm:ss}] {1} | Q:{2}% | D:{3} Mbps U:{4} Mbps | L:{5} J:{6} Loss:{7}% | C:{8}" -f `
        $Snapshot.Timestamp,
        $Snapshot.Tier,
        $Snapshot.QualityScore,
        $Snapshot.DownloadMbps,
        $Snapshot.UploadMbps,
        $latency,
        $Snapshot.JitterMs,
        $Snapshot.LossPct,
        $Snapshot.ConsistencyScore
}

function Show-SnapshotDialog {
    param([object]$Snapshot)
    if (-not $Snapshot) {
        [System.Windows.Forms.MessageBox]::Show("No samples collected yet.", $script:AppTitle, "OK", "Information") | Out-Null
        return
    }

    $latencyText = if ([double]::IsNaN([double]$Snapshot.LatencyMs)) { "n/a" } else { "$($Snapshot.LatencyMs) ms" }
    $message = @"
Interface: $($Snapshot.InterfaceName) [$($Snapshot.InterfaceType)]
Link speed: $($Snapshot.LinkMbps) Mbps

Quality: $($Snapshot.QualityScore)% ($($Snapshot.Tier))
Download: $($Snapshot.DownloadMbps) Mbps
Upload: $($Snapshot.UploadMbps) Mbps
Latency: $latencyText
Jitter: $($Snapshot.JitterMs) ms
Packet Loss: $($Snapshot.LossPct)%
Consistency: $([math]::Round($Snapshot.ConsistencyScore * 100, 1))%
Latency Host: $($Snapshot.LatencyHost)

Last download probe error: $($Snapshot.LastDownloadError)
Last upload probe error: $($Snapshot.LastUploadError)
"@
    [System.Windows.Forms.MessageBox]::Show($message, $script:AppTitle, "OK", "Information") | Out-Null
}

function Set-StartupRegistration {
    param([bool]$Enable)
    $runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    $valueName = "NetQualityTray"
    $scriptPath = Get-ScriptFilePath

    if ($Enable) {
        $scriptDir = Split-Path -Parent $scriptPath
        $nativeExe = Join-Path $scriptDir "NetQualitySentinel.exe"
        if (Test-Path $nativeExe) {
            $launch = "`"$nativeExe`""
        } else {
            $launcherPath = Ensure-LauncherScript -ScriptPath $scriptPath
            $wscriptPath = Join-Path $env:SystemRoot "System32\wscript.exe"
            $launch = "`"$wscriptPath`" //B //Nologo `"$launcherPath`""
        }
        New-ItemProperty -Path $runKey -Name $valueName -Value $launch -PropertyType String -Force | Out-Null
    } else {
        Remove-ItemProperty -Path $runKey -Name $valueName -ErrorAction SilentlyContinue
    }
}

function Get-DefaultColorHexMap {
    return @{
        High     = "#2ECC71"
        Poor     = "#F1C40F"
        VeryPoor = "#E67E22"
        Bad      = "#E74C3C"
        Offline  = "#951111"
        Paused   = "#A0A0A0"
    }
}

function Normalize-ColorHex {
    param(
        [string]$Value,
        [string]$Fallback
    )

    $fallbackHex = if ([string]::IsNullOrWhiteSpace($Fallback)) { "#808080" } else { $Fallback }
    $raw = if ([string]::IsNullOrWhiteSpace($Value)) { $fallbackHex } else { $Value.Trim() }
    if ($raw.StartsWith("#")) { $raw = $raw.Substring(1) }
    if ($raw -notmatch "^[0-9A-Fa-f]{6}$") {
        $raw = $fallbackHex.TrimStart("#")
    }
    return ("#" + $raw.ToUpperInvariant())
}

function Convert-HexToColor {
    param(
        [string]$Hex,
        [System.Drawing.Color]$Fallback = [System.Drawing.Color]::Gray
    )

    if ([string]::IsNullOrWhiteSpace($Hex)) { return $Fallback }
    $value = $Hex.Trim()
    if ($value.StartsWith("#")) { $value = $value.Substring(1) }
    if ($value -notmatch "^[0-9A-Fa-f]{6}$") { return $Fallback }

    try {
        $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
        $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
        $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
        return [System.Drawing.Color]::FromArgb($r, $g, $b)
    } catch {
        return $Fallback
    }
}

function Convert-ColorToHex {
    param([System.Drawing.Color]$Color)
    return "#{0:X2}{1:X2}{2:X2}" -f $Color.R, $Color.G, $Color.B
}

function Get-ReadableTextColor {
    param([System.Drawing.Color]$Background)
    $luma = (0.299 * $Background.R) + (0.587 * $Background.G) + (0.114 * $Background.B)
    if ($luma -ge 140) {
        return [System.Drawing.Color]::FromArgb(20, 20, 20)
    }
    return [System.Drawing.Color]::FromArgb(245, 245, 245)
}

function Get-DeviceThemePalette {
    $isLight = $true
    try {
        $themePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $value = Get-ItemPropertyValue -Path $themePath -Name "AppsUseLightTheme" -ErrorAction Stop
        $isLight = ([int]$value -ne 0)
    } catch {
        $isLight = $true
    }

    if ($isLight) {
        return [pscustomobject]@{
            IsLight    = $true
            Background = [System.Drawing.Color]::FromArgb(248, 249, 251)
            Surface    = [System.Drawing.Color]::FromArgb(255, 255, 255)
            Input      = [System.Drawing.Color]::FromArgb(255, 255, 255)
            Foreground = [System.Drawing.Color]::FromArgb(26, 30, 36)
            Muted      = [System.Drawing.Color]::FromArgb(92, 101, 112)
            Border     = [System.Drawing.Color]::FromArgb(214, 220, 228)
            ButtonBg   = [System.Drawing.Color]::FromArgb(236, 239, 244)
            ButtonFg   = [System.Drawing.Color]::FromArgb(26, 30, 36)
        }
    }

    return [pscustomobject]@{
        IsLight    = $false
        Background = [System.Drawing.Color]::FromArgb(30, 33, 40)
        Surface    = [System.Drawing.Color]::FromArgb(39, 44, 54)
        Input      = [System.Drawing.Color]::FromArgb(51, 57, 69)
        Foreground = [System.Drawing.Color]::FromArgb(236, 240, 246)
        Muted      = [System.Drawing.Color]::FromArgb(166, 176, 188)
        Border     = [System.Drawing.Color]::FromArgb(70, 77, 90)
        ButtonBg   = [System.Drawing.Color]::FromArgb(58, 66, 80)
        ButtonFg   = [System.Drawing.Color]::FromArgb(236, 240, 246)
    }
}

function Apply-ThemeToControl {
    param(
        [System.Windows.Forms.Control]$Control,
        [object]$Palette
    )

    if (-not $Control) { return }

    if ($Control -is [System.Windows.Forms.Form]) {
        $Control.BackColor = $Palette.Background
        $Control.ForeColor = $Palette.Foreground
    } elseif ($Control -is [System.Windows.Forms.GroupBox]) {
        $Control.BackColor = $Palette.Surface
        $Control.ForeColor = $Palette.Foreground
    } elseif ($Control -is [System.Windows.Forms.Label]) {
        $Control.BackColor = [System.Drawing.Color]::Transparent
        $Control.ForeColor = $Palette.Foreground
    } elseif ($Control -is [System.Windows.Forms.NumericUpDown]) {
        $Control.BackColor = $Palette.Input
        $Control.ForeColor = $Palette.Foreground
        $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    } elseif ($Control -is [System.Windows.Forms.CheckBox]) {
        $Control.BackColor = [System.Drawing.Color]::Transparent
        $Control.ForeColor = $Palette.Foreground
    } elseif ($Control -is [System.Windows.Forms.Panel] -or $Control -is [System.Windows.Forms.FlowLayoutPanel] -or $Control -is [System.Windows.Forms.TableLayoutPanel]) {
        $Control.BackColor = $Palette.Background
        $Control.ForeColor = $Palette.Foreground
    } elseif ($Control -is [System.Windows.Forms.Button]) {
        if ($Control.Tag -eq "color-picker") {
            $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $Control.FlatAppearance.BorderSize = 1
            $Control.FlatAppearance.BorderColor = $Palette.Border
        } else {
            $Control.BackColor = $Palette.ButtonBg
            $Control.ForeColor = $Palette.ButtonFg
            $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $Control.FlatAppearance.BorderSize = 1
            $Control.FlatAppearance.BorderColor = $Palette.Border
        }
    }

    foreach ($child in $Control.Controls) {
        Apply-ThemeToControl -Control $child -Palette $Palette
    }
}

function Refresh-VisualConfiguration {
    param([pscustomobject]$Config)

    $defaults = Get-DefaultColorHexMap

    $Config.ColorHighHex = Normalize-ColorHex -Value $Config.ColorHighHex -Fallback $defaults.High
    $Config.ColorPoorHex = Normalize-ColorHex -Value $Config.ColorPoorHex -Fallback $defaults.Poor
    $Config.ColorVeryPoorHex = Normalize-ColorHex -Value $Config.ColorVeryPoorHex -Fallback $defaults.VeryPoor
    $Config.ColorBadHex = Normalize-ColorHex -Value $Config.ColorBadHex -Fallback $defaults.Bad
    $Config.ColorOfflineHex = Normalize-ColorHex -Value $Config.ColorOfflineHex -Fallback $defaults.Offline
    $Config.ColorPausedHex = Normalize-ColorHex -Value $Config.ColorPausedHex -Fallback $defaults.Paused

    $Config.QualityHighMinScore = [int][math]::Min(100, [math]::Max(1, [int]$Config.QualityHighMinScore))
    $Config.QualityPoorMinScore = [int][math]::Min(99, [math]::Max(0, [int]$Config.QualityPoorMinScore))
    $Config.QualityVeryPoorMinScore = [int][math]::Min(98, [math]::Max(0, [int]$Config.QualityVeryPoorMinScore))

    if ($Config.QualityPoorMinScore -ge $Config.QualityHighMinScore) {
        $Config.QualityPoorMinScore = [math]::Max(0, $Config.QualityHighMinScore - 1)
    }
    if ($Config.QualityVeryPoorMinScore -ge $Config.QualityPoorMinScore) {
        $Config.QualityVeryPoorMinScore = [math]::Max(0, $Config.QualityPoorMinScore - 1)
    }

    $script:QualityThresholds = [pscustomobject]@{
        HighMinScore     = [int]$Config.QualityHighMinScore
        PoorMinScore     = [int]$Config.QualityPoorMinScore
        VeryPoorMinScore = [int]$Config.QualityVeryPoorMinScore
    }

    $script:TierColorMap = @{
        High     = Convert-HexToColor -Hex $Config.ColorHighHex -Fallback ([System.Drawing.Color]::FromArgb(46, 204, 113))
        Poor     = Convert-HexToColor -Hex $Config.ColorPoorHex -Fallback ([System.Drawing.Color]::FromArgb(241, 196, 15))
        VeryPoor = Convert-HexToColor -Hex $Config.ColorVeryPoorHex -Fallback ([System.Drawing.Color]::FromArgb(230, 126, 34))
        Bad      = Convert-HexToColor -Hex $Config.ColorBadHex -Fallback ([System.Drawing.Color]::FromArgb(231, 76, 60))
        Offline  = Convert-HexToColor -Hex $Config.ColorOfflineHex -Fallback ([System.Drawing.Color]::FromArgb(149, 17, 17))
        Paused   = Convert-HexToColor -Hex $Config.ColorPausedHex -Fallback ([System.Drawing.Color]::FromArgb(160, 160, 160))
    }
}

function Get-ScriptFilePath {
    if ($PSCommandPath) { return $PSCommandPath }
    throw "Cannot determine script path. Save this script to disk before using this action."
}

function Get-LauncherScriptPath {
    $scriptPath = Get-ScriptFilePath
    $scriptDir = Split-Path -Parent $scriptPath
    return Join-Path $scriptDir "Start-NetworkQualityTray.vbs"
}

function Ensure-LauncherScript {
    param([string]$ScriptPath = (Get-ScriptFilePath))

    $launcherPath = Get-LauncherScriptPath
    $content = @(
        'Set shell = CreateObject("WScript.Shell")'
        ('shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""{0}""", 0, False' -f $ScriptPath)
    )

    Set-Content -Path $launcherPath -Value $content -Encoding ASCII -Force
    return $launcherPath
}

function Set-StartMenuShortcut {
    param([bool]$Enable)
    $scriptPath = Get-ScriptFilePath
    $programsPath = [Environment]::GetFolderPath("Programs")
    $shortcutPath = Join-Path $programsPath "$($script:AppDisplayName).lnk"

    if ($Enable) {
        $scriptDir = Split-Path -Parent $scriptPath
        $nativeExe = Join-Path $scriptDir "NetQualitySentinel.exe"
        if (Test-Path $nativeExe) {
            $targetPath = $nativeExe
            $arguments = ""
        } else {
            $launcherPath = Ensure-LauncherScript -ScriptPath $scriptPath
            $targetPath = Join-Path $env:SystemRoot "System32\wscript.exe"
            $arguments = "//B //Nologo `"$launcherPath`""
        }
        $workingDir = Split-Path -Parent $scriptPath
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetPath
        $shortcut.Arguments = $arguments
        $shortcut.WorkingDirectory = $workingDir
        $shortcut.WindowStyle = 1
        $shortcut.Description = "Live network quality monitor"
        $shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,220"
        $shortcut.Save()
    } else {
        Remove-Item -Path $shortcutPath -ErrorAction SilentlyContinue
    }

    return $shortcutPath
}

function Get-StartupRegistrationEnabled {
    $runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    try {
        $item = Get-ItemProperty -Path $runKey -Name "NetQualityTray" -ErrorAction Stop
        return -not [string]::IsNullOrWhiteSpace([string]$item.NetQualityTray)
    } catch {
        return $false
    }
}

function Normalize-Config {
    param([pscustomobject]$Config)

    $Config.LatencySamples = [int][math]::Min(8, [math]::Max(1, [int]$Config.LatencySamples))
    $Config.LatencyTimeoutMs = [int][math]::Min(5000, [math]::Max(300, [int]$Config.LatencyTimeoutMs))
    $Config.LatencyIntervalSec = [int][math]::Min(30, [math]::Max(1, [int]$Config.LatencyIntervalSec))

    $Config.DownloadSmallBytes = [int][math]::Min(5000000, [math]::Max(50000, [int]$Config.DownloadSmallBytes))
    $Config.DownloadFullBytes = [int][math]::Min(15000000, [math]::Max(100000, [int]$Config.DownloadFullBytes))
    $Config.UploadSmallBytes = [int][math]::Min(3000000, [math]::Max(30000, [int]$Config.UploadSmallBytes))
    $Config.UploadFullBytes = [int][math]::Min(8000000, [math]::Max(80000, [int]$Config.UploadFullBytes))
    $Config.DownloadProbeMaxMs = [int][math]::Min(15000, [math]::Max(800, [int]$Config.DownloadProbeMaxMs))
    $Config.UploadProbeMaxMs = [int][math]::Min(15000, [math]::Max(800, [int]$Config.UploadProbeMaxMs))
    $Config.MaxEndpointBackoffSec = [int][math]::Min(900, [math]::Max(10, [int]$Config.MaxEndpointBackoffSec))
    $Config.HttpTimeoutSec = [int][math]::Min(60, [math]::Max(3, [int]$Config.HttpTimeoutSec))

    $Config.SpeedIntervalPoorSec = [int][math]::Min(120, [math]::Max(2, [int]$Config.SpeedIntervalPoorSec))
    $Config.SpeedIntervalNormalSec = [int][math]::Min(180, [math]::Max(3, [int]$Config.SpeedIntervalNormalSec))
    $Config.SpeedIntervalGoodSec = [int][math]::Min(300, [math]::Max(5, [int]$Config.SpeedIntervalGoodSec))
    $Config.FullProbeIntervalSec = [int][math]::Min(3600, [math]::Max(30, [int]$Config.FullProbeIntervalSec))
    $Config.HistorySize = [int][math]::Min(120, [math]::Max(4, [int]$Config.HistorySize))

    if ($Config.DownloadFullBytes -lt $Config.DownloadSmallBytes) {
        $Config.DownloadFullBytes = $Config.DownloadSmallBytes
    }
    if ($Config.UploadFullBytes -lt $Config.UploadSmallBytes) {
        $Config.UploadFullBytes = $Config.UploadSmallBytes
    }

    if ($Config.SpeedIntervalNormalSec -lt $Config.SpeedIntervalPoorSec) {
        $Config.SpeedIntervalNormalSec = $Config.SpeedIntervalPoorSec
    }
    if ($Config.SpeedIntervalGoodSec -lt $Config.SpeedIntervalNormalSec) {
        $Config.SpeedIntervalGoodSec = $Config.SpeedIntervalNormalSec
    }

    if (-not $Config.DownloadEndpoints -or @($Config.DownloadEndpoints).Count -eq 0) {
        $Config.DownloadEndpoints = @("https://speed.cloudflare.com/__down?bytes={bytes}")
    } else {
        $Config.DownloadEndpoints = @($Config.DownloadEndpoints | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    if (-not $Config.UploadEndpoints -or @($Config.UploadEndpoints).Count -eq 0) {
        $Config.UploadEndpoints = @("https://speed.cloudflare.com/__up")
    } else {
        $Config.UploadEndpoints = @($Config.UploadEndpoints | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    $Config.QualityHighMinScore = [int][math]::Min(100, [math]::Max(1, [int]$Config.QualityHighMinScore))
    $Config.QualityPoorMinScore = [int][math]::Min(99, [math]::Max(0, [int]$Config.QualityPoorMinScore))
    $Config.QualityVeryPoorMinScore = [int][math]::Min(98, [math]::Max(0, [int]$Config.QualityVeryPoorMinScore))
    if ($Config.QualityPoorMinScore -ge $Config.QualityHighMinScore) {
        $Config.QualityPoorMinScore = [math]::Max(0, $Config.QualityHighMinScore - 1)
    }
    if ($Config.QualityVeryPoorMinScore -ge $Config.QualityPoorMinScore) {
        $Config.QualityVeryPoorMinScore = [math]::Max(0, $Config.QualityPoorMinScore - 1)
    }

    $defaults = Get-DefaultColorHexMap
    $Config.ColorHighHex = Normalize-ColorHex -Value $Config.ColorHighHex -Fallback $defaults.High
    $Config.ColorPoorHex = Normalize-ColorHex -Value $Config.ColorPoorHex -Fallback $defaults.Poor
    $Config.ColorVeryPoorHex = Normalize-ColorHex -Value $Config.ColorVeryPoorHex -Fallback $defaults.VeryPoor
    $Config.ColorBadHex = Normalize-ColorHex -Value $Config.ColorBadHex -Fallback $defaults.Bad
    $Config.ColorOfflineHex = Normalize-ColorHex -Value $Config.ColorOfflineHex -Fallback $defaults.Offline
    $Config.ColorPausedHex = Normalize-ColorHex -Value $Config.ColorPausedHex -Fallback $defaults.Paused
}

function Load-ConfigFromFile {
    param(
        [pscustomobject]$Config,
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) { return }
    try {
        $raw = Get-Content -Path $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return }
        $loaded = $raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        return
    }

    foreach ($prop in $loaded.PSObject.Properties) {
        $name = [string]$prop.Name
        if (-not ($Config.PSObject.Properties.Name -contains $name)) { continue }

        $current = $Config.$name
        $incoming = $prop.Value

        if ($current -is [System.Array]) {
            $Config.$name = @($incoming | ForEach-Object { [string]$_ })
            continue
        }

        if ($current -is [int] -or $current -is [long] -or $current -is [double]) {
            try {
                $Config.$name = [double]$incoming
            } catch {
                continue
            }
            continue
        }

        if ($current -is [bool]) {
            try {
                $Config.$name = [bool]$incoming
            } catch {
                continue
            }
            continue
        }

        $Config.$name = $incoming
    }
}

function Save-ConfigToFile {
    param(
        [pscustomobject]$Config,
        [string]$Path
    )

    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
    $json = $Config | ConvertTo-Json -Depth 6
    Set-Content -Path $Path -Value $json -Encoding UTF8
}

function Reset-SpeedCycleState {
    param(
        [object]$SpeedCycle,
        [switch]$CancelRunning
    )

    if (-not $SpeedCycle) { return }

    if ($CancelRunning) {
        if ($SpeedCycle.DownloadTask -and $SpeedCycle.DownloadTask.Cts) {
            try { $SpeedCycle.DownloadTask.Cts.Cancel() } catch { }
            try { $SpeedCycle.DownloadTask.Cts.Dispose() } catch { }
        }
        if ($SpeedCycle.UploadTask) {
            if ($SpeedCycle.UploadTask.Cts) {
                try { $SpeedCycle.UploadTask.Cts.Cancel() } catch { }
                try { $SpeedCycle.UploadTask.Cts.Dispose() } catch { }
            }
            if ($SpeedCycle.UploadTask.Content) {
                try { $SpeedCycle.UploadTask.Content.Dispose() } catch { }
            }
        }
    }

    $SpeedCycle.Active = $false
    $SpeedCycle.RunFullProbe = $false
    $SpeedCycle.DownloadTask = $null
    $SpeedCycle.UploadTask = $null
    $SpeedCycle.DownResult = $null
    $SpeedCycle.UpResult = $null
}

function Set-WpfColorSwatch {
    param(
        [System.Windows.Shapes.Rectangle]$Swatch,
        [string]$Hex
    )

    $safeHex = Normalize-ColorHex -Value $Hex -Fallback "#808080"
    $Swatch.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($safeHex)
}

function Show-SettingsDialog {
    param(
        [pscustomobject]$Config,
        [string]$ConfigPath,
        [object]$Snapshot,
        [bool]$Paused
    )

    $result = [pscustomobject]@{
        Saved         = $false
        ForceProbe    = $false
        ExitRequested = $false
        Paused        = $Paused
    }

    if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
        [System.Windows.Forms.MessageBox]::Show(
            "Settings UI requires STA thread mode.",
            $script:AppTitle,
            "OK",
            "Warning"
        ) | Out-Null
        return $result
    }

    $palette = Get-DeviceThemePalette
    $defaults = Get-DefaultColorHexMap

    $bgHex = Convert-ColorToHex -Color $palette.Background
    $surfaceHex = Convert-ColorToHex -Color $palette.Surface
    $surfaceAltHex = if ($palette.IsLight) { "#F5F7FA" } else { "#343A45" }
    $fgHex = Convert-ColorToHex -Color $palette.Foreground
    $mutedHex = Convert-ColorToHex -Color $palette.Muted
    $borderHex = Convert-ColorToHex -Color $palette.Border
    $inputHex = Convert-ColorToHex -Color $palette.Input
    $accentHex = if ($palette.IsLight) { "#005FB8" } else { "#4CC2FF" }
    $accentHoverHex = if ($palette.IsLight) { "#0B6AC6" } else { "#70D4FF" }
    $accentBorderHex = if ($palette.IsLight) { "#0056AA" } else { "#46B5E6" }
    $accentTextHex = if ($palette.IsLight) { "#FFFFFF" } else { "#002336" }
    $headerTagHex = if ($palette.IsLight) { "#EDF3FB" } else { "#1F2A36" }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$($script:AppTitle) Settings"
        Width="1080"
        Height="720"
        MinWidth="860"
        MinHeight="620"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        WindowStyle="SingleBorderWindow"
        FontFamily="Segoe UI Variable Text"
        FontSize="13"
        SnapsToDevicePixels="True"
        UseLayoutRounding="True"
        Background="$bgHex"
        Foreground="$fgHex">
  <Window.Resources>
    <SolidColorBrush x:Key="CardBrush" Color="$surfaceHex"/>
    <SolidColorBrush x:Key="CardAltBrush" Color="$surfaceAltHex"/>
    <SolidColorBrush x:Key="BorderBrush" Color="$borderHex"/>
    <SolidColorBrush x:Key="InputBrush" Color="$inputHex"/>
    <SolidColorBrush x:Key="MutedBrush" Color="$mutedHex"/>
    <SolidColorBrush x:Key="AccentBrush" Color="$accentHex"/>
    <SolidColorBrush x:Key="AccentHoverBrush" Color="$accentHoverHex"/>
    <SolidColorBrush x:Key="AccentBorderBrush" Color="$accentBorderHex"/>
    <SolidColorBrush x:Key="AccentTextBrush" Color="$accentTextHex"/>
    <SolidColorBrush x:Key="HeaderTagBrush" Color="$headerTagHex"/>

    <Style x:Key="CardStyle" TargetType="Border">
      <Setter Property="CornerRadius" Value="12"/>
      <Setter Property="Background" Value="{StaticResource CardBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>

    <Style x:Key="SectionTitleStyle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="17"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,4"/>
    </Style>

    <Style x:Key="MutedTextStyle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>

    <Style x:Key="FieldTextBoxStyle" TargetType="TextBox">
      <Setter Property="Height" Value="32"/>
      <Setter Property="Padding" Value="8,5"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="Background" Value="{StaticResource InputBrush}"/>
      <Setter Property="Foreground" Value="$fgHex"/>
    </Style>

    <Style x:Key="SectionComboStyle" TargetType="ComboBox">
      <Setter Property="Height" Value="32"/>
      <Setter Property="Padding" Value="8,3"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="Background" Value="{StaticResource InputBrush}"/>
      <Setter Property="Foreground" Value="$fgHex"/>
    </Style>

    <Style x:Key="SecondaryButtonStyle" TargetType="Button">
      <Setter Property="Height" Value="32"/>
      <Setter Property="Padding" Value="14,0"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="Background" Value="{StaticResource CardAltBrush}"/>
      <Setter Property="Foreground" Value="$fgHex"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <Style x:Key="PrimaryButtonStyle" TargetType="Button" BasedOn="{StaticResource SecondaryButtonStyle}">
      <Setter Property="Margin" Value="0"/>
      <Setter Property="BorderBrush" Value="{StaticResource AccentBorderBrush}"/>
      <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource AccentTextBrush}"/>
    </Style>
  </Window.Resources>

  <Grid Margin="20">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Style="{StaticResource CardStyle}" Padding="16">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel>
          <TextBlock Text="$($script:AppDisplayName)" FontSize="24" FontWeight="SemiBold"/>
          <TextBlock Text="Network quality settings" Margin="0,4,0,0" Style="{StaticResource MutedTextStyle}"/>
        </StackPanel>
        <Border Grid.Column="1" Background="{StaticResource HeaderTagBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" CornerRadius="8" Padding="12,6" Margin="12,0,0,0" VerticalAlignment="Center">
          <TextBlock Text="Windows-style settings" Foreground="{StaticResource MutedBrush}" FontSize="12"/>
        </Border>
      </Grid>
    </Border>

    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,14,0,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="330"/>
          <ColumnDefinition Width="14"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Column="0">
          <Border Style="{StaticResource CardStyle}" Padding="14" Margin="0,0,0,12">
            <StackPanel>
              <TextBlock Text="Live status" Style="{StaticResource SectionTitleStyle}"/>
              <TextBlock x:Name="txtLive" Margin="0,4,0,0" Style="{StaticResource MutedTextStyle}"/>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource CardStyle}" Padding="14" Margin="0,0,0,12">
            <StackPanel>
              <TextBlock Text="General" Style="{StaticResource SectionTitleStyle}"/>
              <CheckBox x:Name="chkStartup" Content="Start with Windows" Margin="0,10,0,6"/>
              <CheckBox x:Name="chkPaused" Content="Pause monitoring"/>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource CardStyle}" Padding="14">
            <StackPanel>
              <TextBlock Text="Actions" Style="{StaticResource SectionTitleStyle}"/>
              <TextBlock Text="Open status details without leaving settings." Style="{StaticResource MutedTextStyle}" Margin="0,2,0,10"/>
              <Button x:Name="btnStatus" Content="Status details" Style="{StaticResource SecondaryButtonStyle}" Width="132" Margin="0"/>
            </StackPanel>
          </Border>
        </StackPanel>

        <StackPanel Grid.Column="2">
          <Border Style="{StaticResource CardStyle}" Padding="14" Margin="0,0,0,12">
          <StackPanel>
            <TextBlock Text="Quality thresholds" Style="{StaticResource SectionTitleStyle}"/>
            <TextBlock Text="Define what score range maps to each quality tier." Margin="0,0,0,12" Style="{StaticResource MutedTextStyle}"/>

            <Grid Margin="0,0,0,10">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="180"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="88"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="High starts at" VerticalAlignment="Center"/>
              <Slider x:Name="sldHigh" Grid.Column="1" Minimum="1" Maximum="100" TickFrequency="1" IsSnapToTickEnabled="True" Margin="12,0"/>
              <TextBox x:Name="txtHigh" Grid.Column="2" Style="{StaticResource FieldTextBoxStyle}" HorizontalContentAlignment="Center"/>
            </Grid>

            <Grid Margin="0,0,0,10">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="180"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="88"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Poor starts at" VerticalAlignment="Center"/>
              <Slider x:Name="sldPoor" Grid.Column="1" Minimum="0" Maximum="99" TickFrequency="1" IsSnapToTickEnabled="True" Margin="12,0"/>
              <TextBox x:Name="txtPoor" Grid.Column="2" Style="{StaticResource FieldTextBoxStyle}" HorizontalContentAlignment="Center"/>
            </Grid>

            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="180"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="88"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Very Poor starts at" VerticalAlignment="Center"/>
              <Slider x:Name="sldVeryPoor" Grid.Column="1" Minimum="0" Maximum="98" TickFrequency="1" IsSnapToTickEnabled="True" Margin="12,0"/>
              <TextBox x:Name="txtVeryPoor" Grid.Column="2" Style="{StaticResource FieldTextBoxStyle}" HorizontalContentAlignment="Center"/>
            </Grid>

            <TextBlock x:Name="txtRanges" Margin="0,12,0,0" Style="{StaticResource MutedTextStyle}"/>
          </StackPanel>
        </Border>

        <Border Style="{StaticResource CardStyle}" Padding="14">
          <StackPanel>
            <TextBlock Text="Quality colors" Style="{StaticResource SectionTitleStyle}"/>
            <TextBlock Text="Select a tier, choose a palette chip, or enter a hex value directly." Margin="0,0,0,12" Style="{StaticResource MutedTextStyle}"/>

            <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
              <TextBlock Text="Apply palette to" VerticalAlignment="Center" Margin="0,0,8,0"/>
              <ComboBox x:Name="cmbTargetTier" Width="180" Style="{StaticResource SectionComboStyle}">
                <ComboBoxItem Content="High"/>
                <ComboBoxItem Content="Poor"/>
                <ComboBoxItem Content="VeryPoor"/>
                <ComboBoxItem Content="Bad"/>
                <ComboBoxItem Content="Offline"/>
                <ComboBoxItem Content="Paused"/>
              </ComboBox>
            </StackPanel>

            <Border CornerRadius="8" Background="{StaticResource CardAltBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Padding="8" Margin="0,0,0,12">
              <WrapPanel x:Name="chipPanel"/>
            </Border>

            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="150"/>
                <ColumnDefinition Width="116"/>
                <ColumnDefinition Width="140"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Grid.RowDefinitions>
                <RowDefinition Height="36"/>
                <RowDefinition Height="36"/>
                <RowDefinition Height="36"/>
                <RowDefinition Height="36"/>
                <RowDefinition Height="36"/>
                <RowDefinition Height="36"/>
              </Grid.RowDefinitions>

              <TextBlock Grid.Row="0" Text="High" VerticalAlignment="Center"/>
              <Rectangle x:Name="swatchHigh" Grid.Row="0" Grid.Column="1" Width="96" Height="24" RadiusX="5" RadiusY="5" Stroke="$borderHex" StrokeThickness="1" Margin="6,5"/>
              <TextBox x:Name="txtColorHigh" Grid.Row="0" Grid.Column="2" Margin="6,2" Style="{StaticResource FieldTextBoxStyle}"/>

              <TextBlock Grid.Row="1" Text="Poor" VerticalAlignment="Center"/>
              <Rectangle x:Name="swatchPoor" Grid.Row="1" Grid.Column="1" Width="96" Height="24" RadiusX="5" RadiusY="5" Stroke="$borderHex" StrokeThickness="1" Margin="6,5"/>
              <TextBox x:Name="txtColorPoor" Grid.Row="1" Grid.Column="2" Margin="6,2" Style="{StaticResource FieldTextBoxStyle}"/>

              <TextBlock Grid.Row="2" Text="VeryPoor" VerticalAlignment="Center"/>
              <Rectangle x:Name="swatchVeryPoor" Grid.Row="2" Grid.Column="1" Width="96" Height="24" RadiusX="5" RadiusY="5" Stroke="$borderHex" StrokeThickness="1" Margin="6,5"/>
              <TextBox x:Name="txtColorVeryPoor" Grid.Row="2" Grid.Column="2" Margin="6,2" Style="{StaticResource FieldTextBoxStyle}"/>

              <TextBlock Grid.Row="3" Text="Bad" VerticalAlignment="Center"/>
              <Rectangle x:Name="swatchBad" Grid.Row="3" Grid.Column="1" Width="96" Height="24" RadiusX="5" RadiusY="5" Stroke="$borderHex" StrokeThickness="1" Margin="6,5"/>
              <TextBox x:Name="txtColorBad" Grid.Row="3" Grid.Column="2" Margin="6,2" Style="{StaticResource FieldTextBoxStyle}"/>

              <TextBlock Grid.Row="4" Text="Offline" VerticalAlignment="Center"/>
              <Rectangle x:Name="swatchOffline" Grid.Row="4" Grid.Column="1" Width="96" Height="24" RadiusX="5" RadiusY="5" Stroke="$borderHex" StrokeThickness="1" Margin="6,5"/>
              <TextBox x:Name="txtColorOffline" Grid.Row="4" Grid.Column="2" Margin="6,2" Style="{StaticResource FieldTextBoxStyle}"/>

              <TextBlock Grid.Row="5" Text="Paused" VerticalAlignment="Center"/>
              <Rectangle x:Name="swatchPaused" Grid.Row="5" Grid.Column="1" Width="96" Height="24" RadiusX="5" RadiusY="5" Stroke="$borderHex" StrokeThickness="1" Margin="6,5"/>
              <TextBox x:Name="txtColorPaused" Grid.Row="5" Grid.Column="2" Margin="6,2" Style="{StaticResource FieldTextBoxStyle}"/>
            </Grid>
          </StackPanel>
        </Border>
        </StackPanel>
      </Grid>
    </ScrollViewer>

    <Border Grid.Row="2" Style="{StaticResource CardStyle}" Padding="12">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Text="Settings are saved to settings.json and applied immediately." VerticalAlignment="Center" Style="{StaticResource MutedTextStyle}"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="btnProbeNow" Content="Probe Now" Width="102" Style="{StaticResource SecondaryButtonStyle}"/>
          <Button x:Name="btnExitApp" Content="Exit App" Width="96" Style="{StaticResource SecondaryButtonStyle}"/>
          <Button x:Name="btnClose" Content="Close" Width="96" Style="{StaticResource SecondaryButtonStyle}"/>
          <Button x:Name="btnSave" Content="Save" Width="102" Style="{StaticResource PrimaryButtonStyle}"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
"@

    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $find = {
        param([string]$Name)
        $window.FindName($Name)
    }

    $txtLive = & $find "txtLive"
    if ($Snapshot) {
        $txtLive.Text = "Live: $($Snapshot.Tier) | Score $($Snapshot.QualityScore)% | D $($Snapshot.DownloadMbps) Mbps | U $($Snapshot.UploadMbps) Mbps"
    } else {
        $txtLive.Text = "Live: waiting for first probe..."
    }

    $sldHigh = & $find "sldHigh"
    $sldPoor = & $find "sldPoor"
    $sldVeryPoor = & $find "sldVeryPoor"
    $txtHigh = & $find "txtHigh"
    $txtPoor = & $find "txtPoor"
    $txtVeryPoor = & $find "txtVeryPoor"
    $txtRanges = & $find "txtRanges"

    $sldHigh.Value = [double][int]$Config.QualityHighMinScore
    $sldPoor.Value = [double][int]$Config.QualityPoorMinScore
    $sldVeryPoor.Value = [double][int]$Config.QualityVeryPoorMinScore

    $txtColorHigh = & $find "txtColorHigh"
    $txtColorPoor = & $find "txtColorPoor"
    $txtColorVeryPoor = & $find "txtColorVeryPoor"
    $txtColorBad = & $find "txtColorBad"
    $txtColorOffline = & $find "txtColorOffline"
    $txtColorPaused = & $find "txtColorPaused"

    $swatchHigh = & $find "swatchHigh"
    $swatchPoor = & $find "swatchPoor"
    $swatchVeryPoor = & $find "swatchVeryPoor"
    $swatchBad = & $find "swatchBad"
    $swatchOffline = & $find "swatchOffline"
    $swatchPaused = & $find "swatchPaused"

    $txtColorHigh.Text = Normalize-ColorHex -Value $Config.ColorHighHex -Fallback $defaults.High
    $txtColorPoor.Text = Normalize-ColorHex -Value $Config.ColorPoorHex -Fallback $defaults.Poor
    $txtColorVeryPoor.Text = Normalize-ColorHex -Value $Config.ColorVeryPoorHex -Fallback $defaults.VeryPoor
    $txtColorBad.Text = Normalize-ColorHex -Value $Config.ColorBadHex -Fallback $defaults.Bad
    $txtColorOffline.Text = Normalize-ColorHex -Value $Config.ColorOfflineHex -Fallback $defaults.Offline
    $txtColorPaused.Text = Normalize-ColorHex -Value $Config.ColorPausedHex -Fallback $defaults.Paused

    Set-WpfColorSwatch -Swatch $swatchHigh -Hex $txtColorHigh.Text
    Set-WpfColorSwatch -Swatch $swatchPoor -Hex $txtColorPoor.Text
    Set-WpfColorSwatch -Swatch $swatchVeryPoor -Hex $txtColorVeryPoor.Text
    Set-WpfColorSwatch -Swatch $swatchBad -Hex $txtColorBad.Text
    Set-WpfColorSwatch -Swatch $swatchOffline -Hex $txtColorOffline.Text
    Set-WpfColorSwatch -Swatch $swatchPaused -Hex $txtColorPaused.Text

    $txtHigh.Text = [string][int]$sldHigh.Value
    $txtPoor.Text = [string][int]$sldPoor.Value
    $txtVeryPoor.Text = [string][int]$sldVeryPoor.Value

    $syncThresholds = {
        $high = [int][math]::Round($sldHigh.Value)
        $poor = [int][math]::Round($sldPoor.Value)
        $very = [int][math]::Round($sldVeryPoor.Value)

        if ($poor -ge $high) {
            $poor = [math]::Max(0, $high - 1)
            $sldPoor.Value = $poor
        }
        if ($very -ge $poor) {
            $very = [math]::Max(0, $poor - 1)
            $sldVeryPoor.Value = $very
        }

        $txtHigh.Text = [string]$high
        $txtPoor.Text = [string]$poor
        $txtVeryPoor.Text = [string]$very
        $txtRanges.Text = "High: $high-100 | Poor: $poor-$([math]::Max($high - 1, $poor)) | Very Poor: $very-$([math]::Max($poor - 1, $very)) | Bad: 0-$([math]::Max($very - 1, 0))"
    }

    $setSliderFromText = {
        param($textbox, $slider, [int]$min, [int]$max)
        $v = 0
        if (-not [int]::TryParse($textbox.Text, [ref]$v)) {
            $v = [int][math]::Round($slider.Value)
        }
        if ($v -lt $min) { $v = $min }
        if ($v -gt $max) { $v = $max }
        $slider.Value = $v
        & $syncThresholds
    }

    $sldHigh.add_ValueChanged($syncThresholds)
    $sldPoor.add_ValueChanged($syncThresholds)
    $sldVeryPoor.add_ValueChanged($syncThresholds)
    $txtHigh.add_LostFocus({ & $setSliderFromText $txtHigh $sldHigh 1 100 })
    $txtPoor.add_LostFocus({ & $setSliderFromText $txtPoor $sldPoor 0 99 })
    $txtVeryPoor.add_LostFocus({ & $setSliderFromText $txtVeryPoor $sldVeryPoor 0 98 })
    & $syncThresholds

    $syncColor = {
        param($textbox, $swatch)
        $textbox.Text = Normalize-ColorHex -Value $textbox.Text -Fallback "#808080"
        Set-WpfColorSwatch -Swatch $swatch -Hex $textbox.Text
    }

    $txtColorHigh.add_LostFocus({ & $syncColor $txtColorHigh $swatchHigh })
    $txtColorPoor.add_LostFocus({ & $syncColor $txtColorPoor $swatchPoor })
    $txtColorVeryPoor.add_LostFocus({ & $syncColor $txtColorVeryPoor $swatchVeryPoor })
    $txtColorBad.add_LostFocus({ & $syncColor $txtColorBad $swatchBad })
    $txtColorOffline.add_LostFocus({ & $syncColor $txtColorOffline $swatchOffline })
    $txtColorPaused.add_LostFocus({ & $syncColor $txtColorPaused $swatchPaused })

    $cmbTargetTier = & $find "cmbTargetTier"
    $cmbTargetTier.SelectedIndex = 0
    $chipPanel = & $find "chipPanel"

    $tierMap = @{
        High     = @{ Box = $txtColorHigh; Swatch = $swatchHigh }
        Poor     = @{ Box = $txtColorPoor; Swatch = $swatchPoor }
        VeryPoor = @{ Box = $txtColorVeryPoor; Swatch = $swatchVeryPoor }
        Bad      = @{ Box = $txtColorBad; Swatch = $swatchBad }
        Offline  = @{ Box = $txtColorOffline; Swatch = $swatchOffline }
        Paused   = @{ Box = $txtColorPaused; Swatch = $swatchPaused }
    }

    $chipHexes = @(
        "#2ECC71", "#27AE60", "#009688", "#00BCD4", "#4CAF50", "#8BC34A",
        "#F1C40F", "#FFC107", "#FF9800", "#FFB300", "#FFE082", "#FFD54F",
        "#E67E22", "#FF7043", "#FF5722", "#F4511E", "#FB8C00", "#EF6C00",
        "#E74C3C", "#EF5350", "#D32F2F", "#C62828", "#B71C1C", "#8E0000",
        "#607D8B", "#9E9E9E", "#616161", "#424242", "#212121", "#111111"
    )

    foreach ($hex in $chipHexes) {
        $chipColor = $hex
        $chip = [System.Windows.Controls.Button]::new()
        $chip.Width = 34
        $chip.Height = 34
        $chip.Margin = [System.Windows.Thickness]::new(3)
        $chip.BorderThickness = [System.Windows.Thickness]::new(1)
        $chip.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($borderHex)
        $chip.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($chipColor)
        $chip.ToolTip = $chipColor
        $chipHandler = {
            $selected = [System.Windows.Controls.ComboBoxItem]$cmbTargetTier.SelectedItem
            if (-not $selected) { return }
            $tierName = [string]$selected.Content
            if (-not $tierMap.ContainsKey($tierName)) { return }

            $entry = $tierMap[$tierName]
            $entry.Box.Text = $chipColor
            & $syncColor $entry.Box $entry.Swatch
        }.GetNewClosure()
        $chip.add_Click($chipHandler)
        [void]$chipPanel.Children.Add($chip)
    }

    $chkStartup = & $find "chkStartup"
    $chkPaused = & $find "chkPaused"
    $chkStartup.IsChecked = [bool](Get-StartupRegistrationEnabled)
    $chkPaused.IsChecked = [bool]$Paused

    $btnStatus = & $find "btnStatus"
    $btnProbeNow = & $find "btnProbeNow"
    $btnExitApp = & $find "btnExitApp"
    $btnClose = & $find "btnClose"
    $btnSave = & $find "btnSave"

    $btnStatus.add_Click({
        Show-SnapshotDialog -Snapshot $Snapshot
    })

    $btnProbeNow.add_Click({
        $result.ForceProbe = $true
        $window.Close()
    })

    $btnExitApp.add_Click({
        $result.ExitRequested = $true
        $window.Close()
    })

    $btnClose.add_Click({
        $window.Close()
    })

    $btnSave.add_Click({
        try {
            & $syncThresholds
            foreach ($tierName in $tierMap.Keys) {
                $entry = $tierMap[$tierName]
                & $syncColor $entry.Box $entry.Swatch
            }

            $Config.QualityHighMinScore = [int][math]::Round($sldHigh.Value)
            $Config.QualityPoorMinScore = [int][math]::Round($sldPoor.Value)
            $Config.QualityVeryPoorMinScore = [int][math]::Round($sldVeryPoor.Value)

            $Config.ColorHighHex = [string]$txtColorHigh.Text
            $Config.ColorPoorHex = [string]$txtColorPoor.Text
            $Config.ColorVeryPoorHex = [string]$txtColorVeryPoor.Text
            $Config.ColorBadHex = [string]$txtColorBad.Text
            $Config.ColorOfflineHex = [string]$txtColorOffline.Text
            $Config.ColorPausedHex = [string]$txtColorPaused.Text

            Normalize-Config -Config $Config
            Refresh-VisualConfiguration -Config $Config
            Save-ConfigToFile -Config $Config -Path $ConfigPath

            if ([bool]$chkStartup.IsChecked) {
                Set-StartupRegistration -Enable $true
            } else {
                Set-StartupRegistration -Enable $false
            }

            $result.Paused = [bool]$chkPaused.IsChecked
            $result.Saved = $true
            $result.ForceProbe = $true
            $window.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to save settings: $($_.Exception.Message)",
                "$($script:AppTitle) Settings",
                "OK",
                "Error"
            ) | Out-Null
        }
    })

    [void]$window.ShowDialog()
    return $result
}
$config = [pscustomobject]@{
    LatencyHosts            = @("1.1.1.1", "8.8.8.8", "9.9.9.9")
    LatencySamples          = 3
    LatencyTimeoutMs        = 900
    LatencyIntervalSec      = 3

    DownloadEndpoints       = @(
        "https://speed.cloudflare.com/__down?bytes={bytes}",
        "https://speed.hetzner.de/1MB.bin",
        "https://proof.ovh.net/files/1Mb.dat"
    )
    UploadEndpoints         = @(
        "https://speed.cloudflare.com/__up",
        "https://postman-echo.com/post",
        "https://httpbin.org/post"
    )
    DownloadSmallBytes      = 150000
    DownloadFullBytes       = 1200000
    UploadSmallBytes        = 70000
    UploadFullBytes         = 300000
    DownloadProbeMaxMs      = 3500
    UploadProbeMaxMs        = 3500
    MaxEndpointBackoffSec   = 180
    HttpTimeoutSec          = 8

    SpeedIntervalGoodSec    = 30
    SpeedIntervalNormalSec  = 18
    SpeedIntervalPoorSec    = 10
    FullProbeIntervalSec    = 300
    HistorySize             = 12

    QualityHighMinScore     = 70
    QualityPoorMinScore     = 45
    QualityVeryPoorMinScore = 25
    ColorHighHex            = "#2ECC71"
    ColorPoorHex            = "#F1C40F"
    ColorVeryPoorHex        = "#E67E22"
    ColorBadHex             = "#E74C3C"
    ColorOfflineHex         = "#951111"
    ColorPausedHex          = "#A0A0A0"
}

$scriptPath = Get-ScriptFilePath
$scriptDir = Split-Path -Parent $scriptPath
$settingsFilePath = Join-Path $scriptDir "settings.json"
Load-ConfigFromFile -Config $config -Path $settingsFilePath
Normalize-Config -Config $config
Refresh-VisualConfiguration -Config $config

if ($InstallShortcut) {
    $path = Set-StartMenuShortcut -Enable $true
    Write-Host "Start menu shortcut installed: $path"
    exit 0
}

if ($RemoveShortcut) {
    $path = Set-StartMenuShortcut -Enable $false
    Write-Host "Start menu shortcut removed: $path"
    exit 0
}

if ($InstallStartup) {
    Set-StartupRegistration -Enable $true
    Write-Host "Startup registration installed."
    exit 0
}

if ($RemoveStartup) {
    Set-StartupRegistration -Enable $false
    Write-Host "Startup registration removed."
    exit 0
}

try {
    [void](Set-StartMenuShortcut -Enable $true)
} catch {
    if ($VerboseLogging -or $NoTray) {
        Write-Host "Start menu shortcut update failed: $($_.Exception.Message)"
    }
}

$singleInstanceMutex = $null
$mutexAcquired = $false
if (-not $BypassSingleton) {
    $mutexCreated = $false
    $singleInstanceMutex = [System.Threading.Mutex]::new($false, "Local\NetQualityTray.Singleton", [ref]$mutexCreated)
    try {
        $mutexAcquired = $singleInstanceMutex.WaitOne(0, $false)
    } catch [System.Threading.AbandonedMutexException] {
        $mutexAcquired = $true
    }

    if (-not $mutexAcquired) {
        $message = "$($script:AppTitle) is already running."
        if (-not $NoTray) {
            [System.Windows.Forms.MessageBox]::Show($message, $script:AppTitle, "OK", "Information") | Out-Null
        }
        Write-Host $message
        exit 0
    }
}

$handler = [System.Net.Http.HttpClientHandler]::new()
$handler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
$client = [System.Net.Http.HttpClient]::new($handler)
$client.Timeout = [TimeSpan]::FromSeconds($config.HttpTimeoutSec)
$client.DefaultRequestHeaders.UserAgent.ParseAdd("NetworkQualityTray/1.0")

$downloadEndpointState = New-EndpointFailoverState -Endpoints $config.DownloadEndpoints
$uploadEndpointState = New-EndpointFailoverState -Endpoints $config.UploadEndpoints

$maxUploadBuffer = [math]::Max($config.UploadSmallBytes, $config.UploadFullBytes)
$uploadBuffer = New-Object byte[] $maxUploadBuffer
[System.Random]::new().NextBytes($uploadBuffer)

$downHistory = [System.Collections.Generic.Queue[double]]::new()
$upHistory = [System.Collections.Generic.Queue[double]]::new()
$successHistory = [System.Collections.Generic.Queue[double]]::new()

$lastGoodDownMbps = 0.0
$lastGoodUpMbps = 0.0
$lastGoodDownAt = (Get-Date).AddYears(-10)
$lastGoodUpAt = (Get-Date).AddYears(-10)

$lastDownProbe = [pscustomobject]@{
    Success = $false; Mbps = 0.0; Bytes = 0; DurationMs = 0; Error = "No probe yet."
}
$lastUpProbe = [pscustomobject]@{
    Success = $false; Mbps = 0.0; Bytes = 0; DurationMs = 0; Error = "No probe yet."
}
$lastLatency = [pscustomobject]@{
    Success = $false; Host = $null; AvgMs = [double]::NaN; JitterMs = [double]::NaN; LossPct = 100.0; Samples = 0; Failures = 0
}

$script:ShouldExit = $false
$script:Paused = $false
$script:ForceProbe = $false
$script:LastSnapshot = $null
$script:CurrentIcon = $null
$script:LastIconKey = $null
$script:LastTooltip = ""
$script:SettingsDialogOpen = $false
$script:SettingsChanged = $false

$speedCycle = [pscustomobject]@{
    Active       = $false
    RunFullProbe = $false
    DownloadTask = $null
    UploadTask   = $null
    DownResult   = $null
    UpResult     = $null
}

$preferredHost = $null
$notifyIcon = $null
$logEvery = 4
$nextLogAt = Get-Date

if (-not $NoTray) {
    $notifyIcon = [System.Windows.Forms.NotifyIcon]::new()
    $notifyIcon.Visible = $true
    $initialIcon = New-QualityIcon -Score 0 -Tier "Offline"
    $notifyIcon.Icon = $initialIcon
    $script:CurrentIcon = $initialIcon
    $script:LastIconKey = "offline"
    $notifyIcon.Text = "$($script:AppDisplayName) (right-click for settings)"
    $notifyIcon.add_DoubleClick({
        Show-SnapshotDialog -Snapshot $script:LastSnapshot
    })
    $notifyIcon.add_MouseUp({
        param($sender, $eventArgs)
        if ($eventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
        if ($script:SettingsDialogOpen) { return }

        $script:SettingsDialogOpen = $true
        try {
            $dialogResult = Show-SettingsDialog `
                -Config $config `
                -ConfigPath $settingsFilePath `
                -Snapshot $script:LastSnapshot `
                -Paused $script:Paused

            if ($dialogResult) {
                $script:Paused = [bool]$dialogResult.Paused
                if ([bool]$dialogResult.Saved) {
                    $script:SettingsChanged = $true
                }
                if ([bool]$dialogResult.ForceProbe) {
                    $script:ForceProbe = $true
                }
                if ([bool]$dialogResult.ExitRequested) {
                    $script:ShouldExit = $true
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Unable to open settings: $($_.Exception.Message)",
                $script:AppTitle,
                "OK",
                "Error"
            ) | Out-Null
        } finally {
            $script:SettingsDialogOpen = $false
        }
    })
}

$startAt = Get-Date
$nextLatencyProbeAt = Get-Date
$nextSpeedProbeAt = Get-Date
$nextFullProbeAt = (Get-Date).AddSeconds($config.FullProbeIntervalSec)

try {
    while (-not $script:ShouldExit) {
        $now = Get-Date

        if ($RunForSeconds -gt 0 -and ($now - $startAt).TotalSeconds -ge $RunForSeconds) {
            $script:ShouldExit = $true
            break
        }

        try {
            if ($script:SettingsChanged) {
                $script:SettingsChanged = $false
                Normalize-Config -Config $config
                Refresh-VisualConfiguration -Config $config
                Save-ConfigToFile -Config $config -Path $settingsFilePath

                $client.Timeout = [TimeSpan]::FromSeconds($config.HttpTimeoutSec)
                $downloadEndpointState = New-EndpointFailoverState -Endpoints $config.DownloadEndpoints
                $uploadEndpointState = New-EndpointFailoverState -Endpoints $config.UploadEndpoints
                Reset-SpeedCycleState -SpeedCycle $speedCycle -CancelRunning

                $requiredUploadBuffer = [math]::Max($config.UploadSmallBytes, $config.UploadFullBytes)
                if ((-not $uploadBuffer) -or $uploadBuffer.Length -lt $requiredUploadBuffer) {
                    $uploadBuffer = New-Object byte[] $requiredUploadBuffer
                    [System.Random]::new().NextBytes($uploadBuffer)
                }

                $nextLatencyProbeAt = $now
                $nextSpeedProbeAt = $now
                $nextFullProbeAt = $now.AddSeconds($config.FullProbeIntervalSec)
                $script:LastIconKey = $null
                $script:LastTooltip = ""
                $script:ForceProbe = $true
            }

            if ($script:ForceProbe) {
                $nextLatencyProbeAt = $now
                $nextSpeedProbeAt = $now
                $script:ForceProbe = $false
            }

            if (-not $script:Paused) {
            $primaryInterface = Get-PrimaryNetworkInterface

            if ($now -ge $nextLatencyProbeAt) {
                $lastLatency = Measure-LatencyMetrics `
                    -Hosts $config.LatencyHosts `
                    -PreferredHost $preferredHost `
                    -Samples $config.LatencySamples `
                    -TimeoutMs $config.LatencyTimeoutMs

                if ($lastLatency.Success) {
                    $preferredHost = $lastLatency.Host
                }

                $nextLatencyProbeAt = $now.AddSeconds($config.LatencyIntervalSec)
            }

            if (($now -ge $nextSpeedProbeAt) -and (-not $speedCycle.Active)) {
                $speedCycle.Active = $true
                $speedCycle.RunFullProbe = $now -ge $nextFullProbeAt
                $speedCycle.DownResult = $null
                $speedCycle.UpResult = $null

                $downloadBytes = if ($speedCycle.RunFullProbe) { $config.DownloadFullBytes } else { $config.DownloadSmallBytes }
                $uploadBytes = if ($speedCycle.RunFullProbe) { $config.UploadFullBytes } else { $config.UploadSmallBytes }

                $downStart = Start-DownloadProbeAsync `
                    -Client $client `
                    -EndpointState $downloadEndpointState `
                    -BytesToRead $downloadBytes `
                    -MaxDurationMs $config.DownloadProbeMaxMs `
                    -Now $now

                $upStart = Start-UploadProbeAsync `
                    -Client $client `
                    -EndpointState $uploadEndpointState `
                    -Buffer $uploadBuffer `
                    -BytesToSend $uploadBytes `
                    -MaxDurationMs $config.UploadProbeMaxMs `
                    -Now $now

                $speedCycle.DownloadTask = $downStart.State
                $speedCycle.UploadTask = $upStart.State

                $delaySec = [int][math]::Max($downStart.DelaySec, $upStart.DelaySec)
                if (-not $downStart.Started) {
                    $speedCycle.DownResult = [pscustomobject]@{
                        Completed     = $true
                        Success       = $false
                        Mbps          = 0.0
                        Bytes         = 0
                        DurationMs    = 0
                        Error         = "Download endpoints temporarily in backoff."
                        Endpoint      = "none"
                        EndpointIndex = -1
                    }
                }
                if (-not $upStart.Started) {
                    $speedCycle.UpResult = [pscustomobject]@{
                        Completed     = $true
                        Success       = $false
                        Mbps          = 0.0
                        Bytes         = 0
                        DurationMs    = 0
                        Error         = "Upload endpoints temporarily in backoff."
                        Endpoint      = "none"
                        EndpointIndex = -1
                    }
                }

                if (-not $downStart.Started -and -not $upStart.Started) {
                    $speedCycle.Active = $false
                    $nextSpeedProbeAt = $now.AddSeconds([math]::Max(2, $delaySec))
                }
            }

            if ($speedCycle.Active) {
                if ($speedCycle.DownloadTask) {
                    $done = Complete-DownloadProbeAsync -ProbeState $speedCycle.DownloadTask
                    if ($done) {
                        $speedCycle.DownResult = $done
                        $speedCycle.DownloadTask = $null
                        if ($done.Success) {
                            Register-EndpointProbeSuccess -State $downloadEndpointState -Index $done.EndpointIndex
                        } else {
                            Register-EndpointProbeFailure `
                                -State $downloadEndpointState `
                                -Index $done.EndpointIndex `
                                -Now $now `
                                -MaxBackoffSec $config.MaxEndpointBackoffSec
                        }
                    }
                }

                if ($speedCycle.UploadTask) {
                    $done = Complete-UploadProbeAsync -ProbeState $speedCycle.UploadTask
                    if ($done) {
                        $speedCycle.UpResult = $done
                        $speedCycle.UploadTask = $null
                        if ($done.Success) {
                            Register-EndpointProbeSuccess -State $uploadEndpointState -Index $done.EndpointIndex
                        } else {
                            Register-EndpointProbeFailure `
                                -State $uploadEndpointState `
                                -Index $done.EndpointIndex `
                                -Now $now `
                                -MaxBackoffSec $config.MaxEndpointBackoffSec
                        }
                    }
                }

                if ((-not $speedCycle.DownloadTask) -and (-not $speedCycle.UploadTask)) {
                    if ($speedCycle.DownResult) {
                        $lastDownProbe = [pscustomobject]@{
                            Success    = [bool]$speedCycle.DownResult.Success
                            Mbps       = [double]$speedCycle.DownResult.Mbps
                            Bytes      = [int64]$speedCycle.DownResult.Bytes
                            DurationMs = [int]$speedCycle.DownResult.DurationMs
                            Error      = [string]$speedCycle.DownResult.Error
                        }
                        if ($lastDownProbe.Success) {
                            $lastGoodDownMbps = [double]$lastDownProbe.Mbps
                            $lastGoodDownAt = $now
                            Add-WindowSample -Queue $downHistory -Value $lastGoodDownMbps -MaxCount $config.HistorySize
                        }
                    }

                    if ($speedCycle.UpResult) {
                        $lastUpProbe = [pscustomobject]@{
                            Success    = [bool]$speedCycle.UpResult.Success
                            Mbps       = [double]$speedCycle.UpResult.Mbps
                            Bytes      = [int64]$speedCycle.UpResult.Bytes
                            DurationMs = [int]$speedCycle.UpResult.DurationMs
                            Error      = [string]$speedCycle.UpResult.Error
                        }
                        if ($lastUpProbe.Success) {
                            $lastGoodUpMbps = [double]$lastUpProbe.Mbps
                            $lastGoodUpAt = $now
                            Add-WindowSample -Queue $upHistory -Value $lastGoodUpMbps -MaxCount $config.HistorySize
                        }
                    }

                    $probeSuccess = if (
                        ($speedCycle.DownResult -and $speedCycle.DownResult.Success) -or
                        ($speedCycle.UpResult -and $speedCycle.UpResult.Success)
                    ) { 1.0 } else { 0.0 }
                    Add-WindowSample -Queue $successHistory -Value $probeSuccess -MaxCount $config.HistorySize

                    if ($speedCycle.RunFullProbe) {
                        $nextFullProbeAt = $now.AddSeconds($config.FullProbeIntervalSec)
                    }

                    $speedCycle.Active = $false
                    $nextSpeedProbeAt = $now
                }
            }

            $consistency = Get-ConsistencyScore `
                -DownHistory $downHistory `
                -UpHistory $upHistory `
                -SuccessHistory $successHistory

            $snapshot = Compute-QualitySnapshot `
                -PrimaryInterface $primaryInterface `
                -LatencyMetrics $lastLatency `
                -DownProbe $lastDownProbe `
                -UpProbe $lastUpProbe `
                -LastGoodDownMbps $lastGoodDownMbps `
                -LastGoodDownAt $lastGoodDownAt `
                -LastGoodUpMbps $lastGoodUpMbps `
                -LastGoodUpAt $lastGoodUpAt `
                -Consistency $consistency `
                -Now $now

            $script:LastSnapshot = $snapshot

            if (($now -ge $nextSpeedProbeAt) -and (-not $speedCycle.Active)) {
                if ($snapshot.Offline -or $snapshot.QualityScore -lt 40) {
                    $nextIntervalSec = $config.SpeedIntervalPoorSec
                } elseif ($snapshot.QualityScore -lt 70) {
                    $nextIntervalSec = $config.SpeedIntervalNormalSec
                } else {
                    $nextIntervalSec = $config.SpeedIntervalGoodSec
                }
                $nextSpeedProbeAt = $now.AddSeconds($nextIntervalSec)
            }

            if ($notifyIcon) {
                $iconKey = Get-IconStateKey -Score $snapshot.QualityScore -Tier $snapshot.Tier
                if ($iconKey -ne $script:LastIconKey) {
                    $newIcon = New-QualityIcon -Score $snapshot.QualityScore -Tier $snapshot.Tier
                    $oldIcon = $script:CurrentIcon
                    $notifyIcon.Icon = $newIcon
                    $script:CurrentIcon = $newIcon
                    $script:LastIconKey = $iconKey
                    if ($oldIcon) { $oldIcon.Dispose() }
                }

                $latencyText = if ([double]::IsNaN([double]$snapshot.LatencyMs)) { "n/a" } else { "$($snapshot.LatencyMs)ms" }
                $tooltip = "{0} Q:{1}% D:{2} U:{3} L:{4}" -f `
                    $snapshot.Tier,
                    $snapshot.QualityScore,
                    $snapshot.DownloadMbps,
                    $snapshot.UploadMbps,
                    $latencyText

                if ($tooltip -ne $script:LastTooltip) {
                    Set-NotifyIconTextSafe -NotifyIcon $notifyIcon -Text $tooltip
                    $script:LastTooltip = $tooltip
                }
            }

            if ($NoTray -or $VerboseLogging) {
                if ($now -ge $nextLogAt) {
                    Write-Host (Format-SnapshotLine -Snapshot $snapshot)
                    $nextLogAt = $now.AddSeconds($logEvery)
                }
            }
            } else {
                if ($notifyIcon) {
                    $pausedKey = "paused"
                    if ($script:LastIconKey -ne $pausedKey) {
                        $pausedIcon = New-QualityIcon -Score 0 -Tier "Paused" -Paused $true
                        $oldIcon = $script:CurrentIcon
                        $notifyIcon.Icon = $pausedIcon
                        $script:CurrentIcon = $pausedIcon
                        $script:LastIconKey = $pausedKey
                        if ($oldIcon) { $oldIcon.Dispose() }
                    }

                    $pausedTooltip = "$($script:AppDisplayName) (paused)"
                    if ($script:LastTooltip -ne $pausedTooltip) {
                        Set-NotifyIconTextSafe -NotifyIcon $notifyIcon -Text $pausedTooltip
                        $script:LastTooltip = $pausedTooltip
                    }
                }
            }
        } catch {
            if ($NoTray -or $VerboseLogging) {
                Write-Host "Runtime recovery: $($_.Exception.Message)"
            }
            Reset-SpeedCycleState -SpeedCycle $speedCycle -CancelRunning
            $nextLatencyProbeAt = $now.AddSeconds(2)
            $nextSpeedProbeAt = $now.AddSeconds(2)
            $script:LastIconKey = $null
            $script:LastTooltip = ""
        }

        if ($notifyIcon) {
            [System.Windows.Forms.Application]::DoEvents()
        }
        Start-Sleep -Milliseconds 250
    }
} finally {
    Reset-SpeedCycleState -SpeedCycle $speedCycle -CancelRunning

    if ($client) { $client.Dispose() }
    if ($handler) { $handler.Dispose() }

    if ($notifyIcon) {
        $notifyIcon.Visible = $false
        if ($script:CurrentIcon) {
            $script:CurrentIcon.Dispose()
            $script:CurrentIcon = $null
        }
        $notifyIcon.Dispose()
    }

    if ($singleInstanceMutex) {
        if ($mutexAcquired) {
            try { $singleInstanceMutex.ReleaseMutex() } catch { }
        }
        $singleInstanceMutex.Dispose()
    }
}




