[CmdletBinding()]
param([ValidateSet('Debug', 'Release')][string]$Configuration = 'Debug')

$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$vswherePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswherePath)) { throw 'Visual Studio Build Tools (MSBuild and Roslyn) are required.' }
$msbuildPath = & $vswherePath -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
if (-not $msbuildPath) { throw 'MSBuild was not found. Install the Visual Studio .NET desktop build tools.' }
$compilerPath = Join-Path (Split-Path -Parent $msbuildPath) 'Roslyn\csc.exe'
if (-not (Test-Path -LiteralPath $compilerPath)) { throw "C# compiler not found: $compilerPath" }
$outputDirectory = Join-Path $projectRoot 'bin'
$packageDirectory = Join-Path $projectRoot '.packages'
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$projectPath = Join-Path $projectRoot 'src\PptGcs.AddIn\PptGcs.AddIn.csproj'
& $msbuildPath $projectPath /restore /m /nologo /verbosity:minimal "/p:Configuration=$Configuration" "/p:OutDir=$outputDirectory\" "/p:RestorePackagesPath=$packageDirectory"
if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE." }

foreach ($testName in @('CoreTests', 'IdentityTests', 'NativeConstraintStoreTests', 'PaneTests', 'RadialTests', 'ShapeSupportTests', 'OfficeIconTests', 'TextPointerTests')) {
    $testPath = Join-Path $projectRoot "tests\$testName.cs"
    $referenceDirectory = Join-Path $packageDirectory 'microsoft.netframework.referenceassemblies.net48\1.0.3\build\.NETFramework\v4.8'
    $referenceNames = @('mscorlib', 'System', 'System.Core', 'System.Xml', 'System.Xml.Linq', 'System.Drawing', 'System.Windows.Forms', 'System.Web.Extensions')
    $compilerArguments = @('/nologo', '/noconfig', '/nostdlib+', '/langversion:7.3', '/target:exe', "/out:$outputDirectory\$testName.exe")
    foreach ($referenceName in $referenceNames) {
        $compilerArguments += '/reference:' + (Join-Path $referenceDirectory "$referenceName.dll")
    }
    $compilerArguments += "/reference:$outputDirectory\PptGcs.Core.dll"
    $compilerArguments += "/reference:$outputDirectory\MathNet.Numerics.dll"
    if ($testName -ne 'CoreTests') { $compilerArguments += "/reference:$outputDirectory\PptGcs.AddIn.dll" }
    $compilerArguments += $testPath
    & $compilerPath @compilerArguments
    if ($LASTEXITCODE -ne 0) { throw "$testName compilation failed with exit code $LASTEXITCODE." }
    & (Join-Path $outputDirectory "$testName.exe")
    if ($LASTEXITCODE -ne 0) { throw "$testName failed with exit code $LASTEXITCODE." }
}
& (Join-Path $PSScriptRoot 'BuildPaneHost.ps1')
if ($LASTEXITCODE -ne 0) { throw 'Native task pane host build failed.' }
Write-Output "Built $outputDirectory\PptGcs.AddIn.dll"
