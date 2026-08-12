[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$requiredDirectories = @(
    (Join-Path $projectRoot "src\classes"),
    (Join-Path $projectRoot "src\forms")
)

foreach ($directory in $requiredDirectories) {
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $directory)
    }
}
