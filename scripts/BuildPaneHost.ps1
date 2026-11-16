[CmdletBinding()]
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$projectRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if(-not $OutputDirectory){$OutputDirectory=Join-Path $projectRoot 'bin'}
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$vsRoot=& $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath | Select-Object -First 1
if(-not $vsRoot){throw 'Visual Studio C++ x64 tools and ATL are required for the native task pane host.'}
$vcvars=Join-Path $vsRoot 'VC\Auxiliary\Build\vcvars64.bat'
$source=Join-Path $projectRoot 'src\PptGcs.PaneHost\ConstraintPaneHost.cpp'
$exports=Join-Path $projectRoot 'src\PptGcs.PaneHost\PptGcs.PaneHost.def'
$intermediate=Join-Path $projectRoot 'src\PptGcs.PaneHost\obj'
New-Item -ItemType Directory -Force -Path $OutputDirectory,$intermediate | Out-Null
$batchPath=Join-Path $intermediate 'build-host.cmd'
# Only known local paths are inserted, each in double quotes. This script does
# not enumerate or delete files through cmd; cmd only sets the MSVC environment.
$batch=@"
@echo off
call "$vcvars" >nul
if errorlevel 1 exit /b %errorlevel%
cl /nologo /LD /MT /O2 /EHsc /W4 /DUNICODE /D_UNICODE /D_WINDLL /D_USRDLL /D_ATL_STATIC_REGISTRY /Fo"$intermediate\ConstraintPaneHost.obj" "$source" /link /NOLOGO /MACHINE:X64 /DEF:"$exports" /OUT:"$OutputDirectory\PptGcs.PaneHost.dll" /IMPLIB:"$intermediate\PptGcs.PaneHost.lib" ole32.lib oleaut32.lib uuid.lib user32.lib gdi32.lib comctl32.lib
exit /b %errorlevel%
"@
[IO.File]::WriteAllText($batchPath,$batch,[Text.Encoding]::ASCII)
& $env:ComSpec /d /c $batchPath
if($LASTEXITCODE -ne 0){throw "Native task pane build failed ($LASTEXITCODE)."}
Write-Output "Built $OutputDirectory\PptGcs.PaneHost.dll"
