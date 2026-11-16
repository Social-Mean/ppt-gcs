[CmdletBinding()]
param([switch]$VerifyReopen)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root 'artifacts\copy-identity-validation\variants'
$fixture=Join-Path $root 'artifacts\copy-identity-validation\copy-identity.pptx'
New-Item -ItemType Directory -Force -Path $directory | Out-Null
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
if ($ppt.ActivePresentation.FullName -ne $fixture) { throw 'Unexpected active document; open the copy-identity test fixture.' }
$p=$ppt.ActivePresentation
if ($p.Slides.Count -ne 6 -or $p.Slides.Item(1).Tags.Item('PPTGCS_CASE') -ne 'SameY') { throw 'Fixture markers do not match.' }
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
$previous=Get-Content -LiteralPath (Join-Path $root 'artifacts\copy-identity-validation\acceptance.json') -Raw | ConvertFrom-Json
if (-not $previous.Passed -or $previous.CopyIds.Count -ne 2) { throw 'Run TestCopyIdentity.ps1 successfully first.' }
$copyIds=@($previous.CopyIds)

function Read-State($document) {
    $slides=@(foreach($slide in $document.Slides) {
        [pscustomobject]@{Id=$slide.SlideID;Shapes=@(foreach($shape in $slide.Shapes) {
            [pscustomobject]@{NativeId=$shape.Id;Name=$shape.Name;Guid=$shape.Tags.Item('PPTGCS_ID');Owner=$shape.Tags.Item('PPTGCS_OWNER');Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height}
        })}
    })
    $parts=$document.CustomXMLParts.SelectByNamespace('urn:ppt-gcs:v2')
    $xml=@(for($part=1;$part -le $parts.Count;$part++) {$parts.Item($part).XML})
    return [pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('O');ProcessId=(Get-Process POWERPNT).Id;Path=$document.FullName;DocumentToken=$document.Tags.Item('PPTGCS_DOCUMENT');Xml=$xml;Slides=$slides}
}
function Snapshot([string]$name,$document=$p) {
    $state=Read-State $document
    $state | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $directory ($name+'.json')) -Encoding UTF8
    return $state
}
function Same-State($actual,$expected) {
    return (($actual.Slides | ConvertTo-Json -Depth 15 -Compress) -eq ($expected.Slides | ConvertTo-Json -Depth 15 -Compress) -and
        ($actual.Xml | ConvertTo-Json -Compress) -eq ($expected.Xml | ConvertTo-Json -Compress) -and $actual.DocumentToken -eq $expected.DocumentToken)
}
function Wait-Until([scriptblock]$condition,[string]$label) {
    $deadline=[DateTime]::UtcNow.AddSeconds(3)
    while (-not (& $condition)) {
        if ([DateTime]::UtcNow -ge $deadline) { throw "Timed out waiting for $label" }
        Start-Sleep -Milliseconds 50
    }
}
function Assert-Unique($state) {
    $ids=@($state.Slides | ForEach-Object {$_.Shapes} | Where-Object Guid | ForEach-Object Guid)
    if (@($ids | Select-Object -Unique).Count -ne $ids.Count) { throw 'Duplicate shape GUID remains.' }
}
function Assert-ExistingShapes($actual,$expected,[string]$label) {
    foreach($expectedSlide in $expected.Slides) {
        $actualSlide=@($actual.Slides | Where-Object Id -eq $expectedSlide.Id)
        if ($actualSlide.Count -ne 1) { throw "${label}: slide identity changed" }
        foreach($expectedShape in $expectedSlide.Shapes) {
            $actualShape=@($actualSlide[0].Shapes | Where-Object NativeId -eq $expectedShape.NativeId)
            if ($actualShape.Count -ne 1 -or ($actualShape[0] | ConvertTo-Json -Compress) -ne ($expectedShape | ConvertTo-Json -Compress)) {
                throw "${label}: an existing shape or its tags changed"
            }
        }
    }
}
function Assert-Shapes($actualSlide,$expectedShapes,[string]$label) {
    foreach($shape in $expectedShapes) {
        $actual=@($actualSlide.Shapes | Where-Object NativeId -eq $shape.NativeId)
        if ($actual.Count -ne 1 -or ($actual[0] | ConvertTo-Json -Compress) -ne ($shape | ConvertTo-Json -Compress)) { throw "${label}: shape changed unexpectedly" }
    }
}
function Native-History([string]$command,$expected,[string]$name) {
    if ($ppt.ActivePresentation.FullName -ne $fixture -or -not $ppt.CommandBars.GetEnabledMso($command)) { throw "Unsafe or disabled native $command" }
    $ppt.CommandBars.ExecuteMso($command)
    Wait-Until { Same-State (Read-State $p) $expected } "native $command state"
    $null=Snapshot $name
    Write-Output "PASS native $command $name"
}
function Assert-Fixture {
    if ($ppt.ActivePresentation.FullName -ne $fixture) { throw 'Active presentation changed; no add-in command was sent.' }
}
function Find-Shape($slide,[string]$guid) {
    $matches=@(foreach($shape in $slide.Shapes) {if($shape.Tags.Item('PPTGCS_ID') -eq $guid){$shape}})
    if ($matches.Count -ne 1) { throw "Expected exactly one shape with GUID $guid" }
    return $matches[0]
}
function Wait-NewIds($range,$document,$slide,$sourceIds) {
    Wait-Until {
        $token=$document.Tags.Item('PPTGCS_DOCUMENT')
        if (-not $token) { return $false }
        for($i=1;$i -le $range.Count;$i++) {
            $shape=$range.Item($i)
            $id=$shape.Tags.Item('PPTGCS_ID')
            $owner=$shape.Tags.Item('PPTGCS_OWNER')
            if (-not $id -or $id -in $sourceIds -or $owner -ne ($token+'/'+$slide.SlideID+'/'+$shape.Id)) { return $false }
        }
        return $true
    } 'selection event to assign copy identities'
}
function Constraint-Data { return ($addin.GetSnapshot() | ConvertFrom-Json).slides | ForEach-Object {$_.constraints} }
function Assert-RelationsRetained($current,$expected) {
    foreach($relation in $expected) {
        $same=@($current | Where-Object Id -eq $relation.Id)
        if ($same.Count -ne 1 -or ($same[0] | ConvertTo-Json -Depth 8 -Compress) -ne ($relation | ConvertTo-Json -Depth 8 -Compress)) { throw 'An existing constraint was changed or removed.' }
    }
}

if ($VerifyReopen) {
    # Read-only verification after the parent runner restarts PowerPoint.
    # Never rerun the copy/link/resize sequence in this mode.
    $savedState=Get-Content -LiteralPath (Join-Path $directory 'saved.json') -Raw | ConvertFrom-Json
    $reopened=Snapshot 'after-reopen'
    $beforePids=@($savedState.ProcessId)
    $afterPids=@($reopened.ProcessId)
    $differentProcess=($beforePids.Count -eq 1 -and $afterPids.Count -eq 1 -and $beforePids[0] -ne $afterPids[0])
    $sameDocument=($savedState.Path -eq $fixture -and $reopened.Path -eq $fixture)
    $sameState=Same-State $reopened $savedState
    $result=[pscustomobject]@{
        Passed=($sameDocument -and $differentProcess -and $sameState)
        PhysicalMouseInput=$false
        DocumentPath=$fixture
        BeforeProcessId=$beforePids
        AfterProcessId=$afterPids
        DifferentProcess=$differentProcess
        StateUnchanged=$sameState
        Checks=@('same saved test document','different PowerPoint process','all slides and native IDs','GUIDs and ownership tags','geometry','constraint XML','document identity token')
    }
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'reopen-check.json') -Encoding UTF8
    if (-not $result.Passed) { throw 'Reopen verification failed; inspect variants/reopen-check.json and after-reopen.json.' }
    Write-Output 'PASS new-process reopen preserves all slides, geometry, identities, ownership and constraint XML'
    return
}

$ppt.ActiveWindow.View.GotoSlide(1)
$slide=$p.Slides.Item(1)
$before=Snapshot 'before'
$originalIds=@($slide.Shapes.Item('GCS_SameY_A').Tags.Item('PPTGCS_ID'),$slide.Shapes.Item('GCS_SameY_B').Tags.Item('PPTGCS_ID'))
$originalShapes=@($before.Slides[0].Shapes | Where-Object {$_.Guid -in $originalIds})
if ($originalShapes.Count -ne 2) { throw 'Could not resolve the two originals.' }

# Exercise the second copy as the driver, including the opposite edge anchor.
Assert-Fixture
$null=$addin.ReplayGestureForTest([string]$copyIds[1],'ResizeW','[[12,0],[24,0],[35,0]]')
$copyResized=Snapshot 'second-copy-resized'
$copyGeometry=@($copyResized.Slides[0].Shapes | Where-Object {$_.Guid -in $copyIds})
$beforeSecond=@($before.Slides[0].Shapes | Where-Object Guid -eq $copyIds[1])[0]
if ($copyGeometry.Count -ne 2 -or [Math]::Abs($copyGeometry[0].Width-$copyGeometry[1].Width) -gt 0.01 -or [Math]::Abs($copyGeometry[0].Width-$beforeSecond.Width) -lt 0.01) { throw 'Second copy did not drive the width constraint.' }
Assert-Shapes $copyResized.Slides[0] $originalShapes 'Resizing the second copy'
if (($copyResized.Xml | ConvertTo-Json -Compress) -ne ($before.Xml | ConvertTo-Json -Compress)) { throw 'Copy resize changed constraints.' }
Native-History 'Undo' $before 'undo-second-copy-resize'
Native-History 'Redo' $copyResized 'redo-second-copy-resize'
Write-Output 'PASS second copy drives its constraint without moving original shapes'

# The original relation must still propagate only between the original pair.
Assert-Fixture
$null=$addin.ReplayGestureForTest([string]$originalIds[0],'Move','[[0,12],[0,24],[0,40]]')
$originalMoved=Snapshot 'original-moved'
$originalGeometry=@($originalMoved.Slides[0].Shapes | Where-Object {$_.Guid -in $originalIds})
if ([Math]::Abs($originalGeometry[0].Top-$originalGeometry[1].Top) -gt 0.01 -or [Math]::Abs($originalGeometry[0].Top-$originalShapes[0].Top) -lt 0.01) { throw 'Original same-Y relation failed.' }
Assert-Shapes $originalMoved.Slides[0] @($copyResized.Slides[0].Shapes | Where-Object {$_.Guid -notin $originalIds}) 'Moving the original'
Native-History 'Undo' $copyResized 'undo-original-move'
Write-Output 'PASS original pair remains linked and does not pull its copies'

# A copy of a constrained copy is independent, including the remove command.
$sourceCopy=Find-Shape $slide $copyIds[0]
$extra=$sourceCopy.Duplicate()
$extra.Select()
Wait-NewIds $extra $p $slide @($copyIds+$originalIds)
$extraId=$extra.Item(1).Tags.Item('PPTGCS_ID')
$extraCopied=Snapshot 'copy-of-copy'
Assert-Unique $extraCopied
if (($extraCopied.Xml | ConvertTo-Json -Compress) -ne ($copyResized.Xml | ConvertTo-Json -Compress)) { throw 'Copy of copy changed constraint XML.' }
Assert-Fixture
$addin.UnlinkSelected($null)
$extraUnlinked=Snapshot 'unlinked-independent-copy'
if (-not (Same-State $extraUnlinked $extraCopied)) { throw 'Unlinking an unbound copy changed existing shapes or XML.' }
Write-Output 'PASS copy-of-copy gets a fresh GUID; unlinking it preserves all existing constraints'

# Use the native clipboard object-model path, rather than Duplicate, across slides.
$relationsBefore=@(Constraint-Data)
$slide.Shapes.Range(@('GCS_SameY_A','GCS_SameY_B')).Copy()
$ppt.ActiveWindow.View.GotoSlide(2)
$targetSlide=$p.Slides.Item(2)
$pasted=$targetSlide.Shapes.Paste()
$pasted.Select()
Wait-NewIds $pasted $p $targetSlide $originalIds
$crossSlideIds=@(for($i=1;$i -le $pasted.Count;$i++) {$pasted.Item($i).Tags.Item('PPTGCS_ID')})
$crossSlide=Snapshot 'cross-slide-pasted'
Assert-Unique $crossSlide
Assert-ExistingShapes $crossSlide $extraUnlinked 'Cross-slide paste'
if (($crossSlide.Xml | ConvertTo-Json -Compress) -ne ($extraUnlinked.Xml | ConvertTo-Json -Compress)) { throw 'Cross-slide paste changed original XML.' }
$pasted.Select()
Assert-Fixture
$addin.LinkSameHeight($null)
$heightLinked=Snapshot 'cross-slide-height-linked'
$relationsAfter=@(Constraint-Data)
Assert-RelationsRetained $relationsAfter $relationsBefore
$added=@($relationsAfter | Where-Object {$_.Id -notin @($relationsBefore | ForEach-Object Id)})
if ($added.Count -ne 1 -or $added[0].Kind -ne 'SameHeight' -or $added[0].SlideId -ne $targetSlide.SlideID -or @($added[0].ShapeIds | Where-Object {$_ -notin $crossSlideIds}).Count -ne 0 -or $added[0].ShapeIds.Count -ne 2) { throw 'New same-height relation does not reference just the pasted copies.' }
Assert-ExistingShapes $heightLinked $extraUnlinked 'Creating a relation on pasted shapes'
Write-Output 'PASS cross-slide Copy/Paste identities and new same-height relation; original relations retained'

# Cross-document paste must allocate its own document namespace as well as GUIDs.
$crossPresentation=$null
$crossDocumentPath=Join-Path $directory 'cross-document.pptx'
$crossDocumentIds=@()
$crossDocumentToken=$null
try {
    $ppt.ActiveWindow.View.GotoSlide(1)
    $slide.Shapes.Range(@('GCS_SameY_A','GCS_SameY_B')).Copy()
    $crossPresentation=$ppt.Presentations.Add()
    $newSlide=$crossPresentation.Slides.Add(1,12)
    $documentCopies=$newSlide.Shapes.Paste()
    $documentCopies.Select()
    Wait-NewIds $documentCopies $crossPresentation $newSlide $originalIds
    $crossDocumentToken=$crossPresentation.Tags.Item('PPTGCS_DOCUMENT')
    if ($crossDocumentToken -eq $p.Tags.Item('PPTGCS_DOCUMENT')) { throw 'Pasted document reused the source document namespace.' }
    $crossDocumentIds=@(for($i=1;$i -le $documentCopies.Count;$i++) {$documentCopies.Item($i).Tags.Item('PPTGCS_ID')})
    $crossPresentation.SaveAs([string]$crossDocumentPath,24)
    $crossState=Snapshot 'cross-document-saved' $crossPresentation
    Assert-Unique $crossState
    if ($crossState.Xml.Count -ne 0) { throw 'Cross-document paste unexpectedly imported source constraints.' }
    if (-not (Same-State (Read-State $p) $heightLinked)) { throw 'Cross-document paste changed the source fixture.' }
    Write-Output 'PASS cross-document Copy/Paste assigns independent document token and GUIDs'
}
finally {
    if ($null -ne $crossPresentation) {
        # Preserve failure evidence too; close only this temporary test document.
        if ($crossPresentation.Saved -eq 0) { $crossPresentation.SaveAs([string]$crossDocumentPath,24) }
        $crossPresentation.Close()
    }
    $p.Windows.Item(1).Activate()
    $ppt.ActiveWindow.View.GotoSlide(1)
}
$p.Save()
$saved=Snapshot 'saved'
Assert-Unique $saved
Copy-Item -LiteralPath (Join-Path $env:LOCALAPPDATA 'PptGcs\layout-v2.log') -Destination (Join-Path $directory 'events.log')
[pscustomobject]@{
    Passed=$true;PhysicalMouseInput=$false;ConstraintCreationUndoSupported=$false
    Cases=@('second copy drives width and native Undo/Redo','original same-Y relation stays independent','copy of copy and safe unlink','cross-slide Copy/Paste and new same-height relation','cross-document independent token and GUIDs')
    CopyIds=$copyIds;CopyOfCopyId=$extraId;CrossSlideIds=$crossSlideIds;CrossDocumentIds=$crossDocumentIds;CrossDocumentToken=$crossDocumentToken
} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'acceptance.json') -Encoding UTF8
