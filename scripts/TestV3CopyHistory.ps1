[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$evidenceRoot=Join-Path $root 'artifacts\v3-copy-validation'
$directory=Join-Path $evidenceRoot ('run-'+[DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfff'))
New-Item -ItemType Directory -Force -Path $directory | Out-Null
$fixture=Join-Path $directory 'copy-history.pptx'
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$returnWindow=$ppt.ActiveWindow
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
$p=$null;$passed=$false;$failure=$null;$checks=@();$copyIds=@()
function Assert-CopyFixture {
    if($null -eq $p -or $ppt.ActivePresentation.FullName -ne $fixture -or $p.Tags.Item('PPTGCS_TEST') -ne 'V3'){throw 'Copy history test lost its owned fixture; no command was sent.'}
}
function Read-NativeTags($object) {
    @(for($i=1;$i -le $object.Tags.Count;$i++){[pscustomobject]@{Name=$object.Tags.Name($i);Value=$object.Tags.Value($i)}}) | Sort-Object Name
}
function Read-CopyState {
    Assert-CopyFixture
    [pscustomobject][ordered]@{
        Path=$p.FullName
        PresentationTags=@(Read-NativeTags $p)
        ConstraintXml=$addin.GetConstraintXml()
        Slides=@(foreach($slide in $p.Slides){
            [pscustomobject][ordered]@{Id=$slide.SlideID;Name=$slide.Name;Tags=@(Read-NativeTags $slide);Shapes=@(foreach($shape in $slide.Shapes){
                [pscustomobject][ordered]@{NativeId=$shape.Id;Name=$shape.Name;Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height;Rotation=$shape.Rotation;Tags=@(Read-NativeTags $shape)}
            })}
        })
        Snapshot=($addin.GetSnapshot() | ConvertFrom-Json)
    }
}
function Save-CopySnapshot([string]$name) {
    $state=Read-CopyState
    [pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('O');ProcessId=@(Get-Process POWERPNT | ForEach-Object Id);State=$state} | ConvertTo-Json -Depth 35 | Set-Content -LiteralPath (Join-Path $directory ($name+'.json')) -Encoding UTF8
    return $state
}
function Equal-CopyState($actual,$expected) {($actual | ConvertTo-Json -Depth 35 -Compress) -ceq ($expected | ConvertTo-Json -Depth 35 -Compress)}
function Invoke-CopyNativeHistory([string]$command,[string]$name,$expected) {
    Assert-CopyFixture
    if(-not $ppt.CommandBars.GetEnabledMso($command)){throw "$command unavailable before $name"}
    $ppt.CommandBars.ExecuteMso($command)
    Start-Sleep -Milliseconds 200
    $actual=Save-CopySnapshot $name
    if(-not (Equal-CopyState $actual $expected)){throw "$name did not restore all native geometry, tags and XML"}
    $script:checks += $name
    Write-Output "PASS native $command $name"
}
try {
    $priorCount=$ppt.Presentations.Count
    $addin.CreateDemo($null)
    if($ppt.Presentations.Count -ne $priorCount+1){throw 'CreateDemo did not create a separate fixture.'}
    $p=$ppt.ActivePresentation
    if($p.Slides.Count -ne 6 -or $p.Tags.Item('PPTGCS_TEST') -ne 'V3'){throw 'Created presentation is not the six-case V3 demo.'}
    $p.SaveAs([string]$fixture,24)
    Start-Sleep -Milliseconds 250
    $p.Close()
    $p=$ppt.Presentations.Open([string]$fixture,0,0,-1)
    $ppt.ActiveWindow.View.GotoSlide(1)
    $slide=$p.Slides.Item(1)
    $originalA=$slide.Shapes.Item('GCS_SameY_A')
    $originalB=$slide.Shapes.Item('GCS_SameY_B')
    $originalIds=@($originalA.Tags.Item('PPTGCS_ID'),$originalB.Tags.Item('PPTGCS_ID'))
    $before=Save-CopySnapshot 'before'
    $ppt.StartNewUndoEntry()
    $copies=$slide.Shapes.Range(@('GCS_SameY_A','GCS_SameY_B')).Duplicate()
    $copies.Select()
    $deadline=[DateTime]::UtcNow.AddSeconds(3)
    do {
        $copyIds=@($copies.Item(1).Tags.Item('PPTGCS_ID'),$copies.Item(2).Tags.Item('PPTGCS_ID'))
        $ready=($copyIds[0] -and $copyIds[1] -and $copyIds[0] -ne $copyIds[1] -and @($copyIds | Where-Object {$_ -in $originalIds}).Count -eq 0)
        if(-not $ready){Start-Sleep -Milliseconds 50}
    } while(-not $ready -and [DateTime]::UtcNow -lt $deadline)
    if(-not $ready){throw 'Copies did not receive independent identities after selection.'}
    $ppt.StartNewUndoEntry()
    $copied=Save-CopySnapshot 'copied'
    if($copied.ConstraintXml -cne $before.ConstraintXml){throw 'Copying changed the existing relation document.'}
    foreach($oldSlide in $before.Slides){
        $newSlide=@($copied.Slides | Where-Object Id -eq $oldSlide.Id)[0]
        foreach($shape in $oldSlide.Shapes){
            $current=@($newSlide.Shapes | Where-Object NativeId -eq $shape.NativeId)[0]
            if(($shape | ConvertTo-Json -Depth 15 -Compress) -cne ($current | ConvertTo-Json -Depth 15 -Compress)){throw 'Copying changed an original native shape.'}
        }
    }
    $allIds=@($copied.Snapshot.slides | ForEach-Object {$_.geometry} | ForEach-Object Id)
    if(@($allIds | Select-Object -Unique).Count -ne $allIds.Count){throw 'Duplicate identity remains in copied document.'}
    $checks+='copies get new GUIDs; original geometry, tags and relations unchanged'
    Write-Output 'PASS copying preserves originals and assigns distinct IDs'

    # The original pair already has equal heights, so this is a metadata-only
    # relationship addition: it must still create exactly one native history item.
    $copies.Select()
    $relationId=$addin.AddConstraintCommand('SameHeight','{}')
    if(-not $relationId){throw 'No new relation ID was returned.'}
    $linked=Save-CopySnapshot 'linked'
    $newRelation=@($linked.Snapshot.slides[0].constraints | Where-Object Id -eq $relationId)
    if($newRelation.Count -ne 1 -or $newRelation[0].Kind -ne 'SameHeight' -or $newRelation[0].ShapeIds.Count -ne 2 -or @($newRelation[0].ShapeIds | Where-Object {$_ -notin $copyIds}).Count -ne 0){throw 'New relation did not reference just the two copies.'}
    if(($linked.Slides | ConvertTo-Json -Depth 25 -Compress) -cne ($copied.Slides | ConvertTo-Json -Depth 25 -Compress)){throw 'Equal-height metadata creation unexpectedly changed native shapes.'}
    $checks+='metadata-only relation references only copies'
    Invoke-CopyNativeHistory 'Undo' 'undo-link' $copied
    Invoke-CopyNativeHistory 'Undo' 'undo-copy' $before
    Invoke-CopyNativeHistory 'Redo' 'redo-copy' $copied
    Invoke-CopyNativeHistory 'Redo' 'redo-link' $linked
    $p.Save()
    $null=Save-CopySnapshot 'saved'
    $passed=$true
} catch {
    $failure=$_.Exception.ToString()
    $failure | Set-Content -LiteralPath (Join-Path $directory 'failure.txt') -Encoding UTF8
    throw
} finally {
    if($null -ne $p){try{if($p.FullName -eq $fixture){$p.Save();$p.Close()}}catch{}}
    if($null -ne $returnWindow){try{$returnWindow.Activate()}catch{}}
    $report=[pscustomobject]@{Passed=$passed;PhysicalMouseInput=$false;Checks=$checks;CopyIds=$copyIds;NativeUndoSteps=2;NativeRedoSteps=2;MetadataOnlyLink=$true;Directory=$directory;Failure=$failure;Utc=[DateTime]::UtcNow.ToString('O')}
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $directory 'acceptance.json') -Encoding UTF8
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $evidenceRoot 'latest.json') -Encoding UTF8
    $report | ConvertTo-Json -Depth 12
}
