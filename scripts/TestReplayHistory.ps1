[CmdletBinding()]
param([string]$EvidencePath)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $EvidencePath) { $EvidencePath=Join-Path $root 'artifacts\layout-v2-replay' }
New-Item -ItemType Directory -Force -Path $EvidencePath | Out-Null
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$fixture=[IO.Path]::GetFullPath((Join-Path $root 'artifacts\layout-v2.pptx'))
if ($ppt.ActivePresentation.FullName -ne $fixture) { throw 'Open the pristine layout-v2 fixture first.' }
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
function Capture([string]$name) {
    $null=& (Join-Path $PSScriptRoot 'TestLayoutDocument.ps1') -Action Snapshot -OutputPath (Join-Path $EvidencePath ($name+'.json'))
}
function NativeHistory([string]$command) {
    if ($ppt.ActivePresentation.FullName -ne $fixture) { throw 'The active presentation changed.' }
    if (-not $ppt.CommandBars.GetEnabledMso($command)) { throw "$command is disabled" }
    $ppt.CommandBars.ExecuteMso($command)
}
function AssertSameSlides([string]$actualName,[string]$expectedName) {
    $actual=Get-Content -LiteralPath (Join-Path $EvidencePath ($actualName+'.json')) -Raw | ConvertFrom-Json
    $expected=Get-Content -LiteralPath (Join-Path $EvidencePath ($expectedName+'.json')) -Raw | ConvertFrom-Json
    if (($actual.slides | ConvertTo-Json -Depth 25 -Compress) -ne ($expected.slides | ConvertTo-Json -Depth 25 -Compress)) {
        throw "$actualName differs from $expectedName"
    }
    if (($actual.ConstraintXml | ConvertTo-Json -Compress) -ne ($expected.ConstraintXml | ConvertTo-Json -Compress)) {
        throw "$actualName XML differs from $expectedName"
    }
}
Capture 'before'
$case=0
for ($slide=1; $slide -le 6; $slide++) {
    $ppt.ActiveWindow.View.GotoSlide($slide)
    foreach ($member in @(0,1)) {
        $case++
        $prefix='case-{0:00}' -f $case
        $snapshot=$addin.GetSnapshot() | ConvertFrom-Json
        $shapes=@($snapshot.slides[$slide-1].geometry)
        $driver=if ($slide -ge 5 -and $member -eq 1) {$shapes[2]} else {$shapes[$member]}
        $mode=if($slide -eq 3){'ResizeSE'}elseif($slide -eq 4){'ResizeS'}else{'Move'}
        # Deliberately leave a significant delta for the release-only sample.
        $sign=if($member -eq 0){1}else{-1}
        if ($slide -eq 1 -or $slide -eq 4 -or $slide -eq 6) {
            $offsets='[[0,{0}],[0,{1}],[0,{2}]]' -f (10*$sign),(20*$sign),(35*$sign)
        } elseif ($slide -eq 3) {
            $offsets='[[{0},{0}],[{1},{1}],[{2},{3}]]' -f (10*$sign),(20*$sign),(35*$sign),(32*$sign)
        } else {
            $offsets='[[{0},0],[{1},0],[{2},0]]' -f (10*$sign),(20*$sign),(35*$sign)
        }
        $null=$addin.ReplayGestureForTest([string]$driver.Id,$mode,$offsets)
        Capture ($prefix+'-after')
        NativeHistory 'Undo'
        Capture ($prefix+'-undo')
        $previous=if($case -eq 1){'before'}else{'case-{0:00}-after' -f ($case-1)}
        AssertSameSlides ($prefix+'-undo') $previous
        NativeHistory 'Redo'
        Capture ($prefix+'-redo')
        AssertSameSlides ($prefix+'-redo') ($prefix+'-after')
        Write-Output "PASS replay $case/12: one native Undo and Redo"
    }
}
foreach($step in @(1,2)) {
    NativeHistory 'Undo'
    Capture ('chain-undo-{0:00}' -f $step)
    AssertSameSlides ('chain-undo-{0:00}' -f $step) ('case-{0:00}-after' -f (12-$step))
}
foreach($step in @(1,2)) {
    NativeHistory 'Redo'
    Capture ('chain-redo-{0:00}' -f $step)
    AssertSameSlides ('chain-redo-{0:00}' -f $step) ('case-{0:00}-after' -f (10+$step))
}
Copy-Item -LiteralPath (Join-Path $env:LOCALAPPDATA 'PptGcs\layout-v2.log') -Destination (Join-Path $EvidencePath 'gestures.log')
python (Join-Path $PSScriptRoot 'VerifyReplayEvidence.py') --artifacts $EvidencePath
if ($LASTEXITCODE -ne 0) { throw 'Programmatic replay evidence failed.' }
