param([ValidateSet('undo','redo')][string]$Direction,[int]$Step)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$name='{0}-{1:00}.json' -f $Direction,$Step
$path=Join-Path $root ('artifacts\layout-v2-validation\'+$name)
$null=& (Join-Path $PSScriptRoot 'TestLayoutDocument.ps1') -Action Snapshot -OutputPath $path
python (Join-Path $PSScriptRoot 'CheckLayoutHistoryStep.py') $Direction $Step
if ($LASTEXITCODE -ne 0) { throw "Native $Direction step $Step did not match" }
