[CmdletBinding()]
param(
    [string]$OutputDirectory,
    [string]$InnoCompiler
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$version = (Get-Content -Raw -LiteralPath (Join-Path $projectRoot 'VERSION')).Trim()
if ($version -notmatch '\A(0|[1-9][0-9]{0,4})\.(0|[1-9][0-9]{0,4})\.(0|[1-9][0-9]{0,4})\z' -or
    ($version.Split('.') | Where-Object { [int]$_ -gt 65535 })) {
    throw 'VERSION must contain a three-part Windows product version.'
}

if (-not $InnoCompiler) {
    $command = Get-Command ISCC.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) { $InnoCompiler = $command.Source }
    else { $InnoCompiler = Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe' }
}
if (-not (Test-Path -LiteralPath $InnoCompiler -PathType Leaf)) {
    throw 'Inno Setup 6.3 or later is required. Specify -InnoCompiler if it is installed in a custom location.'
}
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $projectRoot 'dist' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
$payload = Join-Path $projectRoot 'bin\installer-release'

& (Join-Path $PSScriptRoot 'Build.ps1') -Configuration Release -OutputDirectory $payload
$assembly = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $payload 'PptGcs.AddIn.dll'))
$installerName = "PPT-GCS-Setup-$version-x64"
$setup = Join-Path $projectRoot 'installer\Setup.iss'
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$innoArguments = @(
    '/Qp', "/O$OutputDirectory", "/F$installerName", "/DProductVersion=$version", "/DPayloadDir=$payload",
    "/DManagedIdentity=$($assembly.FullName)", "/DManagedVersion=$($assembly.Version)", $setup
)
& $InnoCompiler @innoArguments
if ($LASTEXITCODE -ne 0) { throw "Inno Setup failed with exit code $LASTEXITCODE." }

$installer = Join-Path $OutputDirectory "$installerName.exe"
if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) { throw "The installer was not produced: $installer" }
Write-Output "Built $installer"
