[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [Uri]$Url,
    [ValidateSet(1, 2, 3)]
    [int]$ProtocolMask = 2,
    [switch]$Required,
    [ValidateRange(1000, 120000)]
    [int]$TimeoutMilliseconds = 15000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($Url.Scheme -ne "https" -or [string]::IsNullOrWhiteSpace($Url.Host) -or
    -not [string]::IsNullOrWhiteSpace($Url.UserInfo) -or -not [string]::IsNullOrWhiteSpace($Url.Fragment)) {
    throw "Probe requires an absolute HTTPS URL without user-info or fragments."
}

if ($null -eq ("VBAHttpWinHttpProbe" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class VBAHttpWinHttpProbe
{
    public const uint WinHttpAccessTypeDefaultProxy = 0;
    public const uint WinHttpFlagSecure = 0x00800000;
    public const uint WinHttpOptionEnableHttpProtocol = 133;
    public const uint WinHttpOptionHttpProtocolUsed = 134;
    public const uint WinHttpOptionHttpProtocolRequired = 145;

    [DllImport("winhttp.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr WinHttpOpen(string userAgent, uint accessType, string proxyName, string proxyBypass, uint flags);

    [DllImport("winhttp.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr WinHttpConnect(IntPtr session, string serverName, ushort serverPort, uint reserved);

    [DllImport("winhttp.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr WinHttpOpenRequest(IntPtr connection, string verb, string objectName, IntPtr version, IntPtr referrer, IntPtr acceptTypes, uint flags);

    [DllImport("winhttp.dll", SetLastError = true)]
    public static extern bool WinHttpSetTimeouts(IntPtr handle, int resolve, int connect, int send, int receive);

    [DllImport("winhttp.dll", SetLastError = true)]
    public static extern bool WinHttpSetOption(IntPtr handle, uint option, ref uint buffer, uint bufferLength);

    [DllImport("winhttp.dll", SetLastError = true)]
    public static extern bool WinHttpSendRequest(IntPtr request, IntPtr headers, uint headersLength, IntPtr optional, uint optionalLength, uint totalLength, IntPtr context);

    [DllImport("winhttp.dll", SetLastError = true)]
    public static extern bool WinHttpReceiveResponse(IntPtr request, IntPtr reserved);

    [DllImport("winhttp.dll", SetLastError = true)]
    public static extern bool WinHttpQueryOption(IntPtr handle, uint option, ref uint buffer, ref uint bufferLength);

    [DllImport("winhttp.dll", SetLastError = true)]
    public static extern bool WinHttpCloseHandle(IntPtr handle);

    public static int LastError()
    {
        return Marshal.GetLastWin32Error();
    }
}
"@
}

$session = [IntPtr]::Zero
$connection = [IntPtr]::Zero
$request = [IntPtr]::Zero
$stage = "none"
$mask = [uint32]$ProtocolMask
$requiredValue = if ($Required) { [uint32]1 } else { [uint32]0 }
$protocolUsed = [uint32]0
$protocolLength = [uint32]4

try {
    $stage = "WinHttpOpen"
    $session = [VBAHttpWinHttpProbe]::WinHttpOpen("VBA-HTTP protocol probe", [VBAHttpWinHttpProbe]::WinHttpAccessTypeDefaultProxy, $null, $null, 0)
    if ($session -eq [IntPtr]::Zero) { throw "WinHttpOpen failed: $([VBAHttpWinHttpProbe]::LastError())" }

    $port = if ($Url.IsDefaultPort) { [uint16]443 } else { [uint16]$Url.Port }
    $stage = "WinHttpConnect"
    $connection = [VBAHttpWinHttpProbe]::WinHttpConnect($session, $Url.Host, $port, 0)
    if ($connection -eq [IntPtr]::Zero) { throw "WinHttpConnect failed: $([VBAHttpWinHttpProbe]::LastError())" }

    $target = $Url.AbsolutePath
    if ([string]::IsNullOrEmpty($target)) { $target = "/" }
    if (-not [string]::IsNullOrEmpty($Url.Query)) { $target += $Url.Query }
    $stage = "WinHttpOpenRequest"
    $request = [VBAHttpWinHttpProbe]::WinHttpOpenRequest($connection, "GET", $target, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [VBAHttpWinHttpProbe]::WinHttpFlagSecure)
    if ($request -eq [IntPtr]::Zero) { throw "WinHttpOpenRequest failed: $([VBAHttpWinHttpProbe]::LastError())" }

    $stage = "WinHttpSetTimeouts"
    if (-not [VBAHttpWinHttpProbe]::WinHttpSetTimeouts($request, $TimeoutMilliseconds, $TimeoutMilliseconds, $TimeoutMilliseconds, $TimeoutMilliseconds)) {
        throw "WinHttpSetTimeouts failed: $([VBAHttpWinHttpProbe]::LastError())"
    }

    $stage = "WinHttpSetOption(133)"
    if (-not [VBAHttpWinHttpProbe]::WinHttpSetOption($request, [VBAHttpWinHttpProbe]::WinHttpOptionEnableHttpProtocol, [ref]$mask, 4)) {
        throw "WinHttpSetOption(133) failed: $([VBAHttpWinHttpProbe]::LastError())"
    }

    if ($Required) {
        $stage = "WinHttpSetOption(145)"
        if (-not [VBAHttpWinHttpProbe]::WinHttpSetOption($request, [VBAHttpWinHttpProbe]::WinHttpOptionHttpProtocolRequired, [ref]$requiredValue, 4)) {
            throw "WinHttpSetOption(145) failed: $([VBAHttpWinHttpProbe]::LastError())"
        }
    }

    $stage = "WinHttpSendRequest"
    if (-not [VBAHttpWinHttpProbe]::WinHttpSendRequest($request, [IntPtr]::Zero, 0, [IntPtr]::Zero, 0, 0, [IntPtr]::Zero)) {
        throw "WinHttpSendRequest failed: $([VBAHttpWinHttpProbe]::LastError())"
    }

    $stage = "WinHttpReceiveResponse"
    if (-not [VBAHttpWinHttpProbe]::WinHttpReceiveResponse($request, [IntPtr]::Zero)) {
        throw "WinHttpReceiveResponse failed: $([VBAHttpWinHttpProbe]::LastError())"
    }

    $stage = "WinHttpQueryOption(134)"
    if (-not [VBAHttpWinHttpProbe]::WinHttpQueryOption($request, [VBAHttpWinHttpProbe]::WinHttpOptionHttpProtocolUsed, [ref]$protocolUsed, [ref]$protocolLength)) {
        throw "WinHttpQueryOption(134) failed: $([VBAHttpWinHttpProbe]::LastError())"
    }

    [pscustomobject]@{
        status = "passed"
        url = $Url.AbsoluteUri
        requested_mask = $ProtocolMask
        required = [bool]$Required
        protocol_used_flag = $protocolUsed
        stage = $stage
    } | ConvertTo-Json -Compress
}
catch {
    [pscustomobject]@{
        status = "failed"
        url = $Url.AbsoluteUri
        requested_mask = $ProtocolMask
        required = [bool]$Required
        stage = $stage
        error = $_.Exception.Message
    } | ConvertTo-Json -Compress
    exit 1
}
finally {
    if ($request -ne [IntPtr]::Zero) { [void][VBAHttpWinHttpProbe]::WinHttpCloseHandle($request) }
    if ($connection -ne [IntPtr]::Zero) { [void][VBAHttpWinHttpProbe]::WinHttpCloseHandle($connection) }
    if ($session -ne [IntPtr]::Zero) { [void][VBAHttpWinHttpProbe]::WinHttpCloseHandle($session) }
}
