[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [ValidateSet("HTTP/2")]
    [string]$ExpectedProtocol = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    throw "Protocol host evidence does not exist: $Path"
}

$result = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
$protocols = @("HTTP/2")

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

function Assert-OnlyProperties($Object, [string[]]$Allowed, [string]$Location) {
    if ($null -eq $Object) { throw "$Location is missing." }
    foreach ($property in $Object.PSObject.Properties) {
        if ($Allowed -notcontains $property.Name) {
            throw "$Location contains an unsupported property '$($property.Name)'."
        }
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
Assert-OnlyProperties $result @("schema_version", "benchmark", "status", "run_utc", "source_revision", "requested_protocol", "observed_protocol", "mode", "external_network", "target", "bridge", "office", "environment", "artifact", "build") "root"
if ([string](Require-Property $result "status") -ne "passed" -or
    [string](Require-Property $result "mode") -ne "required" -or
    [bool](Require-Property $result "external_network") -ne $true) {
    throw "Protocol host evidence must be a required-mode external-network pass."
}

$requested = [string](Require-Property $result "requested_protocol")
$observed = [string](Require-Property $result "observed_protocol")
if ($requested -notin $protocols -or $observed -notin $protocols -or $requested -ne $observed) {
    throw "Requested and observed protocol must be an exact HTTP/2 match."
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
Assert-OnlyProperties $target @("scheme", "host", "port") "target"
if ([string](Require-Property $target "scheme") -ne "https" -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $target "host"))) {
    throw "Protocol host target must contain an HTTPS scheme and host."
}
$port = [int](Require-Property $target "port")
if ($port -lt 1 -or $port -gt 65535) { throw "Protocol host target port is invalid." }

$bridge = Require-Property $result "bridge"
Assert-OnlyProperties $bridge @("name", "version", "runtime", "architecture") "bridge"
if ([string](Require-Property $bridge "architecture") -ne "X64" -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $bridge "name")) -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $bridge "version")) -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $bridge "runtime"))) {
    throw "Protocol host bridge metadata is incomplete."
}

$artifact = Require-Property $result "artifact"
Assert-OnlyProperties $artifact @("name", "sha256", "manifest_sha256") "artifact"
if ([string]::IsNullOrWhiteSpace([string](Require-Property $artifact "name"))) {
    throw "Protocol host artifact name is missing."
}
Assert-Sha256 ([string](Require-Property $artifact "sha256")) "artifact.sha256"
Assert-Sha256 ([string](Require-Property $artifact "manifest_sha256")) "artifact.manifest_sha256"

$build = Require-Property $result "build"
Assert-OnlyProperties $build @("vbe_compile", "source_applied", "workbook_saved", "workbook_closed", "excel_cleanup") "build"
if ([string](Require-Property $build "vbe_compile") -ne "passed" -or
    [bool](Require-Property $build "source_applied") -ne $true -or
    [bool](Require-Property $build "workbook_saved") -ne $true -or
    [bool](Require-Property $build "workbook_closed") -ne $true -or
    [string](Require-Property $build "excel_cleanup") -ne "clean") {
    throw "Protocol host evidence does not prove a clean release build."
}

$office = Require-Property $result "office"
Assert-OnlyProperties $office @("version", "build", "operating_system") "office"
if ([string]::IsNullOrWhiteSpace([string](Require-Property $office "version")) -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $office "build")) -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $office "operating_system"))) {
    throw "Office metadata is incomplete."
}

$environment = Require-Property $result "environment"
Assert-OnlyProperties $environment @("windows", "winhttp") "environment"
if ([string]::IsNullOrWhiteSpace([string](Require-Property $environment "windows"))) {
    throw "Windows metadata is missing."
}
$winhttp = Require-Property $environment "winhttp"
$preflight = Require-Property $winhttp "preflight"
Assert-OnlyProperties $winhttp @("enabled_protocol_option", "used_protocol_option", "required_protocol_option", "observed_protocol", "preflight") "environment.winhttp"
if ([int](Require-Property $winhttp "enabled_protocol_option") -ne 133 -or
    [int](Require-Property $winhttp "used_protocol_option") -ne 134 -or
    [int](Require-Property $winhttp "required_protocol_option") -ne 145 -or
    [string](Require-Property $winhttp "observed_protocol") -ne $observed) {
    throw "WinHTTP protocol option metadata is inconsistent."
}
Assert-OnlyProperties $preflight @("status", "requested_mask", "protocol_used_flag", "required", "stage") "environment.winhttp.preflight"
$expectedMask = 1
if ([string](Require-Property $preflight "status") -ne "passed" -or
    [int](Require-Property $preflight "requested_mask") -ne $expectedMask -or
    [int](Require-Property $preflight "protocol_used_flag") -ne $expectedMask -or
    [bool](Require-Property $preflight "required") -ne $true -or
    [string]::IsNullOrWhiteSpace([string](Require-Property $preflight "stage"))) {
    throw "WinHTTP capability preflight metadata is incomplete or inconsistent."
}

Assert-NoSensitiveField $result
Write-Output "Protocol host evidence valid: $requested on $($target.host):$port ($($bridge.architecture))."
