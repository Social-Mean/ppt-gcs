[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root 'artifacts\copy-identity-validation'
$fixture=Join-Path $directory 'copy-identity.pptx'
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
if ($ppt.ActivePresentation.FullName -ne $fixture) { throw 'Unexpected active document.' }
$p=$ppt.ActivePresentation
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
function Snapshot([string]$name) {
    $slides=@(foreach($slide in $p.Slides) {
        [pscustomobject]@{Id=$slide.SlideID;Shapes=@(foreach($shape in $slide.Shapes) {
            [pscustomobject]@{NativeId=$shape.Id;Name=$shape.Name;Guid=$shape.Tags.Item('PPTGCS_ID');Owner=$shape.Tags.Item('PPTGCS_OWNER');Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height}
        })}
    })
    $snapshot=[pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('O');ProcessId=(Get-Process POWERPNT).Id;DocumentToken=$p.Tags.Item('PPTGCS_DOCUMENT');Xml=$p.CustomXMLParts.SelectByNamespace('urn:ppt-gcs:v2').Item(1).XML;Slides=$slides}
    $snapshot | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $directory ($name+'.json')) -Encoding UTF8
    return $snapshot
}
function AssertState($actual,$expected,[string]$label) {
    if (($actual.Slides | ConvertTo-Json -Depth 15 -Compress) -ne ($expected.Slides | ConvertTo-Json -Depth 15 -Compress) -or $actual.Xml -ne $expected.Xml -or $actual.DocumentToken -ne $expected.DocumentToken) { throw "${label}: complete shape/tag/XML state differs" }
    Write-Output "PASS $label"
}
function AssertUnique($snapshot) {
    $ids=@($snapshot.Slides | ForEach-Object {$_.Shapes} | Where-Object Guid | ForEach-Object Guid)
    if (@($ids | Select-Object -Unique).Count -ne $ids.Count) { throw 'Duplicate GUID remains' }
}
$ppt.ActiveWindow.View.GotoSlide(1)
$slide=$p.Slides.Item(1)
$before=Snapshot 'before'
$ppt.StartNewUndoEntry()
$copies=$slide.Shapes.Range(@('GCS_SameY_A','GCS_SameY_B')).Duplicate()
$copies.Select()
# PowerPoint delivers its selection event asynchronously after external COM calls.
$deadline=[DateTime]::UtcNow.AddSeconds(3)
while($copies.Item(1).Tags.Item('PPTGCS_ID') -eq $slide.Shapes.Item('GCS_SameY_A').Tags.Item('PPTGCS_ID') -or $copies.Item(2).Tags.Item('PPTGCS_ID') -eq $slide.Shapes.Item('GCS_SameY_B').Tags.Item('PPTGCS_ID')) {
    if([DateTime]::UtcNow -ge $deadline){throw 'PowerPoint did not deliver copy identity repair'}
    Start-Sleep -Milliseconds 50
}
$ppt.StartNewUndoEntry()
$copied=Snapshot 'copied'
AssertUnique $copied
if ($copied.Xml -ne $before.Xml) { throw 'Copy changed original constraints' }
foreach($original in $before.Slides[0].Shapes) {
    $current=$copied.Slides[0].Shapes | Where-Object NativeId -eq $original.NativeId
    if (($current | ConvertTo-Json -Compress) -ne ($original | ConvertTo-Json -Compress)) { throw 'Copy changed an original shape' }
}
$copyIds=@($copies.Item(1).Tags.Item('PPTGCS_ID'),$copies.Item(2).Tags.Item('PPTGCS_ID'))
if($copyIds[0] -eq $slide.Shapes.Item('GCS_SameY_A').Tags.Item('PPTGCS_ID') -or $copyIds[1] -eq $slide.Shapes.Item('GCS_SameY_B').Tags.Item('PPTGCS_ID')){throw 'Copy retained source GUID'}
Write-Output 'PASS copies receive unique IDs; original geometry, IDs and XML unchanged'
$ppt.CommandBars.ExecuteMso('Undo')
Start-Sleep -Milliseconds 250
AssertState (Snapshot 'undo-copy') $before 'native undo copy'
$ppt.CommandBars.ExecuteMso('Redo')
Start-Sleep -Milliseconds 250
AssertState (Snapshot 'redo-copy') $copied 'native redo copy retains assigned GUIDs'
# PowerPoint ShapeRange accepts collection indices, which happen to differ from
# native shape IDs when text boxes or deletions are present. Select by GUID.
$selectionIndices=@(for($index=1;$index -le $slide.Shapes.Count;$index++) {if($slide.Shapes.Item($index).Tags.Item('PPTGCS_ID') -in $copyIds){$index}})
$slide.Shapes.Range($selectionIndices).Select()
$addin.LinkSameWidth($null)
$linked=Snapshot 'linked'
$linkedData=$addin.GetSnapshot() | ConvertFrom-Json
$newRelation=@($linkedData.slides[0].constraints | Where-Object Kind -eq 'SameWidth')
if($newRelation.Count -ne 1 -or @($newRelation[0].ShapeIds | Where-Object {$_ -notin $copyIds}).Count -ne 0){throw 'New constraint does not reference only copies'}
Write-Output 'PASS new constraint references only the two copies'
$null=$addin.ReplayGestureForTest([string]$copyIds[0],'ResizeE','[[12,0],[25,0],[40,0]]')
$afterReplay=Snapshot 'after-replay'
$result=$addin.GetSnapshot() | ConvertFrom-Json
$first=$result.slides[0].geometry | Where-Object Id -eq $copyIds[0]
$second=$result.slides[0].geometry | Where-Object Id -eq $copyIds[1]
if([Math]::Abs($first.Width-$second.Width) -gt 0.01){throw 'Copy constraint failed to propagate'}
foreach($original in $before.Slides[0].Shapes) {
    $current=$afterReplay.Slides[0].Shapes | Where-Object NativeId -eq $original.NativeId
    if (($current | ConvertTo-Json -Compress) -ne ($original | ConvertTo-Json -Compress)) { throw 'Editing copies changed an original shape' }
}
$ppt.CommandBars.ExecuteMso('Undo')
Start-Sleep -Milliseconds 250
AssertState (Snapshot 'undo-replay') $linked 'native undo copy resize'
$ppt.CommandBars.ExecuteMso('Redo')
Start-Sleep -Milliseconds 250
AssertState (Snapshot 'redo-replay') $afterReplay 'native redo copy resize'
$p.Save()
$null=Snapshot 'saved'
Copy-Item -LiteralPath (Join-Path $env:LOCALAPPDATA 'PptGcs\layout-v2.log') -Destination (Join-Path $directory 'events.log')
[pscustomobject]@{Passed=$true;Cases=@('copy unique IDs','original constraints and geometry retained','native copy Undo/Redo preserves IDs','new constraint on copies','copy resize propagation and native history');PhysicalMouseInput=$false;ConstraintCreationUndoSupported=$false;CopyIds=$copyIds} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'acceptance.json') -Encoding UTF8
