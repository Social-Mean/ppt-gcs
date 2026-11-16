param([ValidateSet('undo','redo')][string]$Direction,[int]$StartStep=1)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$evidence=Join-Path $root 'artifacts\layout-v2-validation'
$before=Get-Content -LiteralPath (Join-Path $evidence 'before.json') -Raw | ConvertFrom-Json
$plan=Get-Content -LiteralPath (Join-Path $evidence 'history-plan.json') -Raw | ConvertFrom-Json
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$command=if($Direction -eq 'undo'){'Undo'}else{'Redo'}
for($step=$StartStep; $step -le $plan.gestureCount; $step++) {
    if($ppt.ActivePresentation.FullName -ne $before.presentation -or $ppt.ActivePresentation.Slides.Count -ne 6) {
        throw 'Active presentation changed; no native command was sent.'
    }
    if(-not $ppt.CommandBars.GetEnabledMso($command)) { throw "$command is disabled before step $step" }
    # Execute PowerPoint's own built-in history command, never set geometry here.
    $ppt.CommandBars.ExecuteMso($command)
    & (Join-Path $PSScriptRoot 'CaptureLayoutHistory.ps1') -Direction $Direction -Step $step
}
