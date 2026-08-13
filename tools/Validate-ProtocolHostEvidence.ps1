[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [ValidateSet("HTTP/2", "HTTP/3")]
    [string]$ExpectedProtocol = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    throw "Protocol host evidence does not exist: $Path"
}

$result = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
$protocols = @("HTTP/2", "HTTP/3")

function Require-Property($Object, [string]$Name) {
    if ($null -eq $Object -or $null -eq $Object.PSObject.Properties[$Name]) {
        throw "Protocol host evidence is missing '$Name'."
    }
    return $Object.PSObject.Properties[$Name].Value
}

function Assert-Sha256([string]$Value, [string]$Label) {
    if ($Value -notmatch '^[0-9a-fA-F]{64}$') {
        throw "$Label must be a 64-character SHA-256 value."
    }
}

function Assert-NoSensitiveField($Object, [string]$Location = "root") {
    if ($null -eq $Object) { return }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in $Object.Keys) {
            if ([string]$key -match '(?i)(authorization|password|secret|token|cookie|body|header|query|user.?info|file.?path|local.?path|url)') {
                throw "Protocol host evidence contains a forbidden field '$Location.$key'."
            }
            Assert-NoSensitiveField $Object[$key] "$Location.$key"
        }
        return
    }
    if ($Object -is [System.Management.Automation.PSCustomObject]) {
        foreach ($property in $Object.PSObject.Properties) {
            if ($property.Name -match '(?i)(authorization|password|secret|token|cookie|body|header|query|user.?info|file.?path|local.?path|url)') {
                throw "Protocol host evidence contains a forbidden field '$Location.$($property.Name)'."
            }
            Assert-NoSensitiveField $property.Value "$Location.$($property.Name)"
        }
        return
    }
    if ($Object -is [System.Collections.IEnumerable] -and -not ($Object -is [string])) {
        $index = 0
        foreach ($item in $Object) {
            Assert-NoSensitiveField $item "$Location[$index]"
            $index++
        }
    }
}

if ([int](Require-Property $result "schema_version") -ne 1 -or
    [string](Require-Property $result "benchmark") -ne "protocol-host-validation") {
    throw "Protocol host evidence schema identity is invalid."
}
if ([string](Require-Property $result "status") -ne "passed" -or
    [string](Require-Property $result "mode") -ne "required" -or
    [bool](Require-Property $result "external_network") -ne $true) {
    throw "Protocol host evidence must be a required-mode external-network pass."
}

$requested = [string](Require-Property $result "requested_protocol")
$observed = [string](Require-Property $result "observed_protocol")
if ($requested -notin $protocols -or $observed -notin $protocols -or $requested -ne $observed) {
    throw "Requested and observed protocols are not an exact HTTP/2 or HTTP/3 match."
}
if (-not [string]::IsNullOrWhiteSpace($ExpectedProtocol) -and $requested -ne $ExpectedProtocol) {
    throw "Evidence protocol '$requested' does not match expected '$ExpectedProtocol'."
}

$runUtc = [DateTime]::MinValue
if (-not [DateTime]::TryParse([string](Require-Property $result "run_utc"), [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$runUtc) -or $runUtc.Kind -ne [DateTimeKind]::Utc) {
    throw "run_utc must be a valid UTC ISO-8601 timestamp."
}
if ([string](Require-Property $result "source_revision") -notmatch '^[0-9a-fA-F]{40,64}$') {
    throw "source_revision must be a full hexadecimal Git revision."
}

$target = Require-Property $result "target"
if ([string](Require-Property $target "scheme") -ne "https" -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $target "host"))) {
    throw "Protocol host target must contain an HTTPS scheme and host."
}
$port = [int](Require-Property $target "port")
if ($port -lt 1 -or $port -gt 65535) { throw "Protocol host target port is invalid." }

$bridge = Require-Property $result "bridge"
if ([string](Require-Property $bridge "architecture") -notin @("X86", "X64") -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $bridge "name")) -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $bridge "version")) -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $bridge "runtime"))) {
    throw "Protocol host bridge metadata is incomplete."
}

$artifact = Require-Property $result "artifact"
if ([string]::IsNullOrWhiteSpace([string](Require-Property $artifact "name"))) {
    throw "Protocol host artifact name is missing."
}
Assert-Sha256 ([string](Require-Property $artifact "sha256")) "artifact.sha256"
Assert-Sha256 ([string](Require-Property $artifact "manifest_sha256")) "artifact.manifest_sha256"

$build = Require-Property $result "build"
if ([string](Require-Property $build "vbe_compile") -ne "passed" -or
    [bool](Require-Property $build "source_applied") -ne $true -or
    [bool](Require-Property $build "workbook_saved") -ne $true -or
    [bool](Require-Property $build "workbook_closed") -ne $true -or
    [string](Require-Property $build "excel_cleanup") -ne "clean") {
    throw "Protocol host evidence does not prove a clean release build."
}

Assert-NoSensitiveField $result
Write-Output "Protocol host evidence valid: $requested on $($target.host):$port ($($bridge.architecture))."
