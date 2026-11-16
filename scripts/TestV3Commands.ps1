[CmdletBinding()]
param([switch]$CheckPane)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$evidenceRoot=Join-Path $root 'artifacts\v3-command-validation'
$directory=Join-Path $evidenceRoot ('run-'+[DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfff'))
New-Item -ItemType Directory -Force -Path $directory | Out-Null
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$returnWindow=$ppt.ActiveWindow
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
$fixture=Join-Path $directory 'commands.pptx'
$p=$null
$checks=@()
$passed=$false
$failure=$null
function Assert-TestDocument {
    if($null -eq $p -or $ppt.ActivePresentation.FullName -ne $fixture -or $p.Tags.Item('PPTGCS_TEST') -ne 'V3') { throw 'The active document is not the owned V3 command fixture.' }
}
function Tags($object) {
    @(for($i=1;$i -le $object.Tags.Count;$i++) {[pscustomobject]@{Name=$object.Tags.Name($i);Value=$object.Tags.Value($i)}}) | Sort-Object Name
}
function Read-State {
    Assert-TestDocument
    [pscustomobject][ordered]@{
        Path=$p.FullName
        PresentationTags=@(Tags $p)
        ConstraintXml=$addin.GetConstraintXml()
        Slides=@(foreach($slide in $p.Slides) {
            [pscustomobject][ordered]@{Id=$slide.SlideID;Name=$slide.Name;Tags=@(Tags $slide);Shapes=@(foreach($shape in $slide.Shapes) {
                [pscustomobject][ordered]@{NativeId=$shape.Id;Name=$shape.Name;Tags=@(Tags $shape);Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height;Rotation=$shape.Rotation}
            })}
        })
    }
}
function Snapshot([string]$name) {
    $state=Read-State
    [pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('O');ProcessId=@(Get-Process POWERPNT | ForEach-Object Id);State=$state} | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath (Join-Path $directory ($name+'.json')) -Encoding UTF8
    return $state
}
function Same-State($actual,$expected) {
    ($actual | ConvertTo-Json -Depth 30 -Compress) -ceq ($expected | ConvertTo-Json -Depth 30 -Compress)
}
function Assert-State($actual,$expected,[string]$label) {
    if(-not (Same-State $actual $expected)) {throw "${label}: full geometry, tags or XML differs"}
    $script:checks += $label
    Write-Output "PASS $label"
}
function Invoke-NativeConstraintHistory([string]$command,[string]$name,$expected) {
    Assert-TestDocument
    if(-not $ppt.CommandBars.GetEnabledMso($command)) {throw "$command is disabled before $name"}
    $ppt.CommandBars.ExecuteMso($command)
    # Allow selection events to run before judging metadata history.
    Start-Sleep -Milliseconds 200
    Assert-State (Snapshot $name) $expected "native $command $name"
}
function Assert-Close([double]$actual,[double]$expected,[string]$label) {
    if([Math]::Abs($actual-$expected) -gt 0.01) {throw "$label expected $expected, got $actual"}
}
function Write-Pane([string]$name) {
    $model=$addin.GetPaneState() | ConvertFrom-Json
    $model | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $directory ($name+'.json')) -Encoding UTF8
    return $model
}
try {
    $p=$ppt.Presentations.Add()
    $p.Tags.Add('PPTGCS_TEST','V3')
    $p.PageSetup.SlideWidth=960
    $p.PageSetup.SlideHeight=540
    $slide=$p.Slides.Add(1,12)
    $slide.Name='CommandHistory'
    $a=$slide.Shapes.AddShape(1,100,140,100,80)
    $a.Name='History_A';$a.LockAspectRatio=0;$a.TextFrame.AutoSize=0
    $b=$slide.Shapes.AddShape(1,450,240,150,110)
    $b.Name='History_B';$b.LockAspectRatio=0;$b.TextFrame.AutoSize=0
    $p.SaveAs([string]$fixture,24)
    Start-Sleep -Milliseconds 250
    $p.Close()
    $p=$ppt.Presentations.Open([string]$fixture,0,0,-1)
    $slide=$p.Slides.Item(1);$a=$slide.Shapes.Item('History_A');$b=$slide.Shapes.Item('History_B')
    $a.Select()
    $states=@((Snapshot 'state-00-unconstrained'))
    $fixedId=$addin.AddConstraintCommand('FixedWidth','{"value":180}')
    if(-not $fixedId){throw 'AddConstraintCommand returned no constraint ID'}
    Assert-Close $a.Width 180 'FixedWidth add'
    $states+=Snapshot 'state-01-added'
    $addin.UpdateConstraintCommand([string]$fixedId,'{"value":220}')
    Assert-Close $a.Width 220 'FixedWidth update'
    $states+=Snapshot 'state-02-updated'
    $addin.SetConstraintEnabled([string]$fixedId,$false)
    $states+=Snapshot 'state-03-disabled'
    $disabled=Write-Pane 'pane-disabled'
    if(@($disabled.Relations | Where-Object {$_.Id -eq $fixedId -and $_.Enabled -eq $false -and $_.Status -eq '已禁用'}).Count -ne 1){throw 'Pane did not show disabled relation'}
    $addin.SetConstraintEnabled([string]$fixedId,$true)
    $states+=Snapshot 'state-04-enabled'
    $addin.DeleteConstraintCommand([string]$fixedId)
    $states+=Snapshot 'state-05-deleted'
    if(@(($addin.GetSnapshot() | ConvertFrom-Json).slides[0].constraints).Count -ne 0){throw 'Delete did not remove the selected relation'}
    for($step=1;$step -le 5;$step++){Invoke-NativeConstraintHistory 'Undo' ('undo-{0:00}' -f $step) $states[5-$step]}
    for($step=1;$step -le 5;$step++){Invoke-NativeConstraintHistory 'Redo' ('redo-{0:00}' -f $step) $states[$step]}
    $checks+='five atomic add/update/disable/enable/delete transactions'

    # Real solver conflict: two different fixed widths cannot satisfy SameWidth.
    $slide.Shapes.Range(@('History_A','History_B')).Select()
    $sameId=$addin.AddConstraintCommand('SameWidth','{}')
    $a.Select()
    $lockedId=$addin.AddConstraintCommand('FixedWidth','{"value":180}')
    $b.Select()
    $beforeConflict=Snapshot 'before-conflicting-add'
    $conflictThrown=$false
    try {$null=$addin.AddConstraintCommand('FixedWidth','{"value":120}')} catch {$conflictThrown=$true;$_.Exception.Message | Set-Content -LiteralPath (Join-Path $directory 'conflict-message.txt') -Encoding UTF8}
    if(-not $conflictThrown){throw 'Conflicting required widths were accepted'}
    Assert-State (Snapshot 'after-conflicting-add') $beforeConflict 'conflicting add is atomic'

    $aId=$a.Tags.Item('PPTGCS_ID');$bId=$b.Tags.Item('PPTGCS_ID')
    $addin.SelectConstraintObjects((@($aId,$bId) | ConvertTo-Json -Compress))
    $selected=Write-Pane 'pane-selected-pair'
    if(@($selected.Objects | Where-Object Selected).Count -ne 2){throw 'Pane model does not reflect selected pair'}
    if($ppt.ActiveWindow.Selection.ShapeRange.Count -ne 2){throw 'Relationship selection did not select both native shapes'}
    $checks+='relation members select native shapes and pane model follows'
    if($CheckPane) {
        $null=$addin.OpenConstraintPane()
        $diagnostics=$addin.GetPaneDiagnostics() | ConvertFrom-Json
        $diagnostics | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'pane-diagnostics.json') -Encoding UTF8
        if(-not $diagnostics.factory -or -not $diagnostics.created -or -not $diagnostics.visible -or $diagnostics.host -eq 0 -or $diagnostics.control -eq 0){throw 'The native task pane did not activate'}
        $checks+='native task pane visible with active hosted control'
    }

    $b.Delete()
    Start-Sleep -Milliseconds 200
    $missing=Write-Pane 'pane-missing-member'
    if(@($missing.Relations | Where-Object {$_.Id -eq $sameId -and $_.Status -like '*缺失*'}).Count -ne 1){throw 'Pane did not expose the missing object relation'}
    $beforeCleanup=Snapshot 'before-missing-cleanup'
    $addin.DeleteConstraintCommand([string]$sameId)
    $afterCleanup=Snapshot 'after-missing-cleanup'
    $clean=Write-Pane 'pane-missing-cleaned'
    if(@($clean.Relations | Where-Object Id -eq $sameId).Count -ne 0 -or @($clean.Relations | Where-Object Id -eq $lockedId).Count -ne 1){throw 'Single relation cleanup removed wrong constraints'}
    Invoke-NativeConstraintHistory 'Undo' 'undo-missing-cleanup' $beforeCleanup
    Invoke-NativeConstraintHistory 'Redo' 'redo-missing-cleanup' $afterCleanup
    $checks+='missing relation remains visible and can be individually removed'
    $p.Save()
    $null=Snapshot 'saved'
    $passed=$true
} catch {
    $failure=$_.Exception.ToString()
    $failure | Set-Content -LiteralPath (Join-Path $directory 'failure.txt') -Encoding UTF8
    throw
} finally {
    if($null -ne $p) {
        try {if($p.FullName -eq $fixture){$p.Save();$p.Close()}} catch { }
    }
    if($null -ne $returnWindow){try{$returnWindow.Activate()}catch{}}
    $report=[pscustomobject]@{Passed=$passed;PhysicalMouseInput=$false;CheckPane=[bool]$CheckPane;Checks=$checks;Failure=$failure;Directory=$directory;Utc=[DateTime]::UtcNow.ToString('O')}
    $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'acceptance.json') -Encoding UTF8
    $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $evidenceRoot 'latest.json') -Encoding UTF8
    $report | ConvertTo-Json -Depth 10
}
