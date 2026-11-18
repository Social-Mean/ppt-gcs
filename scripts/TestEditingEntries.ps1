#requires -Version 5.1
[CmdletBinding()]
param([ValidateRange(100, 5000)][int]$HistoryDelayMilliseconds = 250)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OfficeComSession.ps1')
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$token = [guid]::NewGuid().ToString('D').ToUpperInvariant()
$artifactDirectory = Join-Path $repository ('artifacts/editing-entries-' + $token.ToLowerInvariant())
$fixturePath = [IO.Path]::GetFullPath((Join-Path $artifactDirectory 'EditingEntries.pptx'))
$testTag = 'PPTGCS_EDITING_ENTRIES_TEST'
$session = $null; $ppt = $null; $document = $null; $ownedDocument = $null; $window = $null; $addin = $null
$fixtureCreated = $false; $fixtureSaved = $false; $failure = $null
$filterRestored = $false
$results = New-Object 'Collections.Generic.List[object]'
$cleanupErrors = New-Object 'Collections.Generic.List[string]'
$note = 'Programmatic production API replay; this does not validate physical mouse or keyboard interaction.'

function Write-Evidence($Name, $Value) {
    $Value | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath (Join-Path $artifactDirectory ($Name + '.json')) -Encoding UTF8
}
function Assert-OwnedReference {
    if ($null -eq $document -or -not $fixtureCreated -or -not (Test-OfficeComIdentity $document $ownedDocument)) { throw 'The explicitly created test presentation reference is required.' }
    foreach ($original in $session.OriginalDocuments) {
        if (Test-OfficeComIdentity $document $original.Document) { throw 'An original user document can never be a test fixture.' }
    }
}
function Assert-OwnedDocument([switch]$Active) {
    Assert-OwnedReference
    if ([string]$document.Tags.Item('PPTGCS_TEST') -cne 'V3' -or [string]$document.Tags.Item($testTag) -cne $token -or
        ($fixtureSaved -and [IO.Path]::GetFullPath([string]$document.FullName) -ine $fixturePath)) { throw 'The exact owned V3 editing-entry fixture is required.' }
    if ($Active -and (-not (Test-OfficeComIdentity $ppt.ActivePresentation $document) -or
        -not (Test-OfficeComIdentity $ppt.ActiveWindow $window) -or [int]$window.ViewType -ne 9)) {
        throw 'The owned fixture and its original normal-view window must remain active.'
    }
}
function Get-Tags($Owner) {
    $values = @{}
    for ($tagIndex = 1; $tagIndex -le $Owner.Tags.Count; $tagIndex++) {
        $values[[string]$Owner.Tags.Name($tagIndex)] = [string]$Owner.Tags.Value($tagIndex)
    }
    $sorted = [ordered]@{}
    foreach ($tagName in @($values.Keys | Sort-Object)) { $sorted[$tagName] = $values[$tagName] }
    return $sorted
}
function Get-Selection {
    $selection = $window.Selection
    $type = [int]$selection.Type
    $nativeIds = @(); $childIds = @()
    if ($type -eq 2 -or $type -eq 3) {
        $nativeIds = @(for ($index = 1; $index -le $selection.ShapeRange.Count; $index++) { [int]$selection.ShapeRange.Item($index).Id })
    }
    $hasChild = [bool]$selection.HasChildShapeRange
    if ($hasChild) { $childIds = @(for ($index = 1; $index -le $selection.ChildShapeRange.Count; $index++) { [int]$selection.ChildShapeRange.Item($index).Id }) }
    return [pscustomobject]@{ Type = $type; HasChild = $hasChild; NativeIds = $nativeIds; ChildNativeIds = $childIds; ActivePane = [int]$window.ActivePane.ViewType }
}
function Get-Snapshot([string]$Name) {
    Assert-OwnedDocument -Active
    $rows = New-Object 'Collections.Generic.List[object]'
    $slides = New-Object 'Collections.Generic.List[object]'
    function Read-Shape($Shape, [string]$Key, [int]$ParentId) {
        $text = $null
        if ([int]$Shape.HasTextFrame -ne 0) { $text = [string]$Shape.TextFrame.TextRange.Text }
        $rows.Add([pscustomobject]@{
            Key = $Key; NativeId = [int]$Shape.Id; ParentNativeId = $ParentId; Name = [string]$Shape.Name
            Type = [int]$Shape.Type; Child = [int]$Shape.Child; Tags = Get-Tags $Shape; Text = $text
            Left = [double]$Shape.Left; Top = [double]$Shape.Top; Width = [double]$Shape.Width; Height = [double]$Shape.Height
            Rotation = [double]$Shape.Rotation; LockAspectRatio = [int]$Shape.LockAspectRatio
        })
        if ([int]$Shape.Type -eq 6) {
            for ($childIndex = 1; $childIndex -le $Shape.GroupItems.Count; $childIndex++) {
                $child = $Shape.GroupItems.Item($childIndex)
                Read-Shape $child ($Key + '/' + [string]$child.Name) ([int]$Shape.Id)
            }
        }
    }
    for ($slideIndex = 1; $slideIndex -le $document.Slides.Count; $slideIndex++) {
        $slide = $document.Slides.Item($slideIndex)
        $slides.Add([pscustomobject]@{ Index = $slideIndex; Id = [int]$slide.SlideID; Name = [string]$slide.Name; Tags = Get-Tags $slide })
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shape = $slide.Shapes.Item($shapeIndex)
            Read-Shape $shape ($slideIndex.ToString() + '/' + [string]$shape.Name) 0
        }
    }
    Assert-OwnedDocument -Active
    $snapshot = [pscustomobject]@{
        Utc = [datetime]::UtcNow.ToString('O'); ConstraintXml = [string]$addin.GetConstraintXml()
        DocumentTags = Get-Tags $document; SlideWidth = [double]$document.PageSetup.SlideWidth; SlideHeight = [double]$document.PageSetup.SlideHeight
        SnapToGrid = [int]$document.SnapToGrid; GridDistance = [double]$document.GridDistance
        Slides = @($slides.ToArray()); Rows = @($rows.ToArray()); Selection = Get-Selection
        ActiveSlideId = [int]$window.View.Slide.SlideID; ActiveSlideIndex = [int]$window.View.Slide.SlideIndex; Note = $note
    }
    Write-Evidence $Name $snapshot
    return $snapshot
}
function Assert-SameSnapshot($Actual, $Expected, [string]$Name) {
    foreach ($property in @('ConstraintXml', 'SlideWidth', 'SlideHeight')) {
        if ($Actual.$property -cne $Expected.$property) { throw "$Name changed document metadata: $property" }
    }
    foreach ($property in @('DocumentTags', 'Slides')) {
        if (($Actual.$property | ConvertTo-Json -Depth 15 -Compress) -cne ($Expected.$property | ConvertTo-Json -Depth 15 -Compress)) { throw "$Name changed $property" }
    }
    if ($Actual.Rows.Count -ne $Expected.Rows.Count) { throw "$Name changed the recursive shape count" }
    $byKey = @{}
    foreach ($row in $Expected.Rows) { if ($byKey.ContainsKey($row.Key)) { throw "Duplicate expected shape path: $($row.Key)" }; $byKey[$row.Key] = $row }
    foreach ($row in $Actual.Rows) {
        $want = $byKey[$row.Key]
        if ($null -eq $want) { throw "$Name introduced shape path $($row.Key)" }
        foreach ($property in @('NativeId', 'ParentNativeId', 'Name', 'Type', 'Child', 'Text', 'LockAspectRatio')) {
            if ($row.$property -cne $want.$property) { throw "$Name changed $($row.Key).$property" }
        }
        if (($row.Tags | ConvertTo-Json -Depth 5 -Compress) -cne ($want.Tags | ConvertTo-Json -Depth 5 -Compress)) { throw "$Name changed shape tags: $($row.Key)" }
        foreach ($axis in @('Left', 'Top', 'Width', 'Height', 'Rotation')) {
            if ([math]::Abs($row.$axis - $want.$axis) -gt .01) { throw "$Name $($row.Key).$axis expected $($want.$axis), got $($row.$axis)" }
        }
    }
    # Selection is observed in every snapshot, but is neither an Undo contract
    # nor persisted by SaveAs. Grid settings are observed; their persistence is
    # outside this constraint/geometry test. Exact membership is checked below.
}
function Assert-Relations($Snapshot, [string]$Name) {
    [xml]$xml = $Snapshot.ConstraintXml
    $byId = @{}
    foreach ($row in $Snapshot.Rows) {
        $id = [string]$row.Tags['PPTGCS_ID']
        if ($id) { $byId[$id.ToUpperInvariant()] = $row }
    }
    foreach ($relation in $xml.DocumentElement.ChildNodes) {
        if ($relation.LocalName -ne 'constraint' -or $relation.GetAttribute('enabled') -eq 'false') { continue }
        $members = @(foreach ($shape in $relation.ChildNodes) {
            if ($shape.LocalName -ne 'shape') { continue }
            $id = $shape.GetAttribute('id').ToUpperInvariant()
            if (-not $byId.ContainsKey($id)) { throw "$Name has a missing constraint member: $id" }
            $byId[$id]
        })
        $axis = $null
        switch ($relation.GetAttribute('kind')) {
            'SameY' { $axis = 'Top' }
            'SameWidth' { $axis = 'Width' }
            'SameHeight' { $axis = 'Height' }
            'FixedPosition' {
                $left = [double]::Parse($relation.GetAttribute('value'), [Globalization.CultureInfo]::InvariantCulture)
                $top = [double]::Parse($relation.GetAttribute('secondaryValue'), [Globalization.CultureInfo]::InvariantCulture)
                if ([math]::Abs($members[0].Left - $left) -gt .01 -or [math]::Abs($members[0].Top - $top) -gt .01) { throw "$Name violates FixedPosition" }
            }
            default { throw "$Name contains a constraint kind outside this fixture: $($relation.GetAttribute('kind'))" }
        }
        if ($null -eq $axis) { continue }
        foreach ($member in $members) {
            if ([math]::Abs($member.$axis - $members[0].$axis) -gt .01) { throw "$Name violates $($relation.GetAttribute('kind')) on $($member.Key)" }
        }
    }
}
function Assert-EditedMembers($After, $Before, [int[]]$ExpectedIds, [bool]$Children) {
    $prefix = $Before.ActiveSlideIndex.ToString() + '/'
    $beforeRows = @($Before.Rows | Where-Object { $_.Key.StartsWith($prefix, [StringComparison]::Ordinal) })
    $afterRows = @{}
    foreach ($row in $After.Rows) { $afterRows[$row.Key] = $row }
    if ($ExpectedIds.Count -gt 1) {
        $selected = @($beforeRows | Where-Object { $ExpectedIds -contains $_.NativeId -and ([bool]($_.Child -ne 0)) -eq $Children })
        if ($selected.Count -ne $ExpectedIds.Count) { throw 'The frozen selected members could not be resolved uniquely.' }
        $dx = $afterRows[$selected[0].Key].Left - $selected[0].Left
        $dy = $afterRows[$selected[0].Key].Top - $selected[0].Top
        foreach ($member in $selected) {
            $actual = $afterRows[$member.Key]
            if ([math]::Abs(($actual.Left - $member.Left) - $dx) -gt .01 -or [math]::Abs(($actual.Top - $member.Top) - $dy) -gt .01 -or
                [math]::Abs($actual.Width - $member.Width) -gt .01 -or [math]::Abs($actual.Height - $member.Height) -gt .01) {
                throw "Selected members did not translate rigidly: $($member.Key)"
            }
        }
    }
    if ($Children) {
        foreach ($member in @($beforeRows | Where-Object { $_.Child -ne 0 -and $ExpectedIds -notcontains $_.NativeId })) {
            foreach ($axis in @('Left', 'Top', 'Width', 'Height', 'Rotation')) {
                if ([math]::Abs($afterRows[$member.Key].$axis - $member.$axis) -gt .01) { throw "An unselected native sibling changed: $($member.Key).$axis" }
            }
        }
    }
}
function Assert-SelectedIds([int[]]$Expected, [switch]$Children) {
    $state = Get-Selection
    if ($state.Type -ne 2 -or $state.HasChild -ne [bool]$Children) { throw 'Expected an exact shape selection without a text caret.' }
    $actual = if ($Children) { @($state.ChildNativeIds) } else { @($state.NativeIds) }
    if ((($actual | Sort-Object) -join ',') -cne (($Expected | Sort-Object) -join ',')) { throw 'The exact native selected object IDs were not preserved.' }
}
function Select-TopShapes([int]$Page, [string[]]$Names) {
    Assert-OwnedDocument -Active
    $window.View.GotoSlide($Page)
    for ($paneIndex = 1; $paneIndex -le $window.Panes.Count; $paneIndex++) {
        if ([int]$window.Panes.Item($paneIndex).ViewType -eq 1) { $window.Panes.Item($paneIndex).Activate(); break }
    }
    Assert-OwnedDocument -Active
    $window.Selection.Unselect()
    $document.Slides.Item($Page).Shapes.Range([object[]]$Names).Select()
    Assert-SelectedIds @(foreach ($name in $Names) { [int]$document.Slides.Item($Page).Shapes.Item($name).Id })
}
function Select-GroupChild([string]$Name) {
    Select-TopShapes 3 @('AB')
    $document.Slides.Item(3).Shapes.Item('AB').GroupItems.Item($Name).Select()
    Assert-SelectedIds @([int]$document.Slides.Item(3).Shapes.Item('AB').GroupItems.Item($Name).Id) -Children
}
function Invoke-NativeHistory([ValidateSet('Undo', 'Redo')][string]$Command, [string]$Name) {
    Assert-OwnedDocument -Active
    if ([int]$window.Selection.Type -ne 2 -or [bool]$window.Selection.HasChildShapeRange -or [int]$window.ActivePane.ViewType -ne 1) {
        throw 'Native history requires a top-level shape selection in the slide pane.'
    }
    if (-not $ppt.CommandBars.GetEnabledMso($Command)) { throw "$Command is unexpectedly disabled." }
    $ppt.CommandBars.ExecuteMso($Command)
    # This is a separate automation call after the production replay returned.
    # A delayed read catches late native writes rather than fixing them.
    Start-Sleep -Milliseconds $HistoryDelayMilliseconds
    return Get-Snapshot $Name
}
function Test-Replay([string]$Name, [scriptblock]$Action, [int[]]$ExpectedIds, [switch]$Children, [switch]$Rejected) {
    $before = Get-Snapshot ($Name + '-before')
    Assert-OwnedDocument -Active
    $null = & $Action
    $after = Get-Snapshot ($Name + '-after')
    Assert-Relations $after $Name
    if ($ExpectedIds.Count -gt 0) { Assert-SelectedIds $ExpectedIds -Children:$Children }
    if ($Rejected) {
        Assert-SameSnapshot $after $before $Name
        $results.Add([pscustomobject]@{ Name = $Name; Kind = 'AtomicRejection'; Passed = $true })
        Write-Output "PASS $Name atomic rejection"
        return
    }
    Assert-EditedMembers $after $before $ExpectedIds ([bool]$Children)
    $moved = $false; $expected = @{}
    foreach ($row in $before.Rows) { $expected[$row.Key] = $row }
    foreach ($row in $after.Rows) { foreach ($axis in @('Left', 'Top', 'Width', 'Height')) { if ([math]::Abs($row.$axis - $expected[$row.Key].$axis) -gt .01) { $moved = $true } } }
    if (-not $moved) { throw "$Name did not change geometry." }
    $page = [int]$window.View.Slide.SlideIndex
    Select-TopShapes $page $(if ($page -eq 3) { @('AB') } else { @('A') })
    $undo = Invoke-NativeHistory 'Undo' ($Name + '-undo')
    Assert-SameSnapshot $undo $before ($Name + ' one Undo')
    Assert-Relations $undo ($Name + ' one Undo')
    $redo = Invoke-NativeHistory 'Redo' ($Name + '-redo')
    Assert-SameSnapshot $redo $after ($Name + ' one Redo')
    Assert-Relations $redo ($Name + ' one Redo')
    $results.Add([pscustomobject]@{ Name = $Name; Kind = 'OneNativeUndoRedo'; Passed = $true })
    Write-Output "PASS $Name one native Undo/Redo"
}
function Add-Rectangle($Slide, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height) {
    Assert-OwnedDocument -Active
    $shape = $Slide.Shapes.AddShape(1, $Left, $Top, $Width, $Height)
    $shape.Name = $Name; $shape.LockAspectRatio = 0; $shape.TextFrame.AutoSize = 0; $shape.TextFrame2.AutoSize = 0
    $shape.TextFrame.TextRange.Text = $Name; $shape.TextFrame.TextRange.Font.Size = 24
    return $shape
}
function Add-Hint($Slide, [string]$Text) {
    Assert-OwnedDocument -Active
    $shape = $Slide.Shapes.AddTextbox(1, 25, 20, 910, 64)
    $shape.Name = 'Instructions'; $shape.TextFrame.AutoSize = 0; $shape.TextFrame.TextRange.Text = $Text; $shape.TextFrame.TextRange.Font.Size = 18
}
function Add-Relation([int]$Page, [string[]]$Names, [string]$Kind, [string]$Parameters = '{}') {
    Select-TopShapes $Page $Names
    Assert-OwnedDocument -Active
    $id = [string]$addin.AddConstraintCommand($Kind, $Parameters)
    if ([string]::IsNullOrEmpty($id)) { throw "Failed to create $Kind in the owned fixture." }
    return $id
}

try {
    New-Item -ItemType Directory -Path $artifactDirectory | Out-Null
    $session = Start-OfficeComSession
    $ppt = $session.Application
    Write-Evidence 'original-documents' @($session.OriginalDocuments | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path; Saved = $_.Saved } })
    $addin = $ppt.COMAddIns.Item('PptGcs.Connect').Object
    if ($null -eq $addin) { throw 'The installed production PptGcs.Connect add-in must be loaded.' }
    $document = $ppt.Presentations.Add(-1); $ownedDocument = $document; $fixtureCreated = $true; $window = $ppt.ActiveWindow
    $document.Tags.Add('PPTGCS_TEST', 'V3'); $document.Tags.Add($testTag, $token)
    $document.PageSetup.SlideWidth = 960; $document.PageSetup.SlideHeight = 540
    $slide = $document.Slides.Add(1, 12)
    Assert-OwnedDocument -Active
    if (Test-Path -LiteralPath $fixturePath) { throw 'The unique fixture path must not already exist.' }
    $document.SaveAs($fixturePath, 24); $fixtureSaved = $true
    Assert-OwnedDocument -Active
    $slide.Name = 'SelectionMove'
    Add-Hint $slide 'Programmatic replay: select A or B with free D and a zero-width line; the other constrained shape follows.'
    $null = Add-Rectangle $slide 'A' 140 150 140 90
    $null = Add-Rectangle $slide 'B' 540 150 140 90
    $null = Add-Rectangle $slide 'D' 300 320 100 70
    $line = $slide.Shapes.AddLine(460, 320, 460, 430); $line.Name = 'ZeroWidthLine'
    if ([math]::Abs([double]$line.Width) -gt .0001) { throw 'The free native line must have zero width.' }
    $null = Add-Relation 1 @('A', 'B') 'SameY'
    $slide = $document.Slides.Add(2, 12); $slide.Name = 'Conflict'
    Add-Hint $slide 'Programmatic replay: fixed A plus free D must reject a selection move atomically.'
    $null = Add-Rectangle $slide 'A' 140 150 140 90
    $null = Add-Rectangle $slide 'B' 540 150 140 90
    $null = Add-Rectangle $slide 'D' 300 320 100 70
    $null = Add-Relation 2 @('A', 'B') 'SameY'
    $null = Add-Relation 2 @('A') 'FixedPosition' '{"value":140,"second":150}'
    $slide = $document.Slides.Add(3, 12); $slide.Name = 'GroupChild'
    Add-Hint $slide 'Programmatic replay: edit actual A/B inside AB; external C follows group top, width and height.'
    $null = Add-Rectangle $slide 'A' 140 150 140 90
    $null = Add-Rectangle $slide 'B' 300 270 120 70
    $group = $slide.Shapes.Range([object[]]@('A', 'B')).Group(); $group.Name = 'AB'
    $null = Add-Rectangle $slide 'C' 580 $group.Top $group.Width $group.Height
    $null = Add-Relation 3 @('AB', 'C') 'SameY'
    $null = Add-Relation 3 @('AB', 'C') 'SameWidth'
    $null = Add-Relation 3 @('AB', 'C') 'SameHeight'
    $slide = $document.Slides.Add(4, 12); $slide.Name = 'NativeFields'
    Add-Hint $slide 'Reserved for manual native field interaction. This script does not simulate field, mouse or keyboard input.'
    $null = Add-Rectangle $slide 'A' 140 170 150 100
    $null = Add-Rectangle $slide 'B' 540 170 150 100
    $null = Add-Relation 4 @('A', 'B') 'SameWidth'
    $null = Add-Relation 4 @('A', 'B') 'SameHeight'
    Select-TopShapes 1 @('A', 'D', 'ZeroWidthLine')
    $document.Save()
    Write-Evidence 'fixture' ([pscustomobject]@{ Path = $fixturePath; Token = $token; TestTag = $testTag; SlideCount = [int]$document.Slides.Count; Note = $note })
    Write-Output $note

    $ids = @(foreach ($name in @('A', 'D', 'ZeroWidthLine')) { [int]$document.Slides.Item(1).Shapes.Item($name).Id })
    Test-Replay 'multi-a-extras' { $addin.ReplaySelectionMoveForTest('["A","D","ZeroWidthLine"]', '[[12,15],[28,32],[43,51]]') } $ids
    Select-TopShapes 1 @('B', 'D', 'ZeroWidthLine')
    $ids = @(foreach ($name in @('B', 'D', 'ZeroWidthLine')) { [int]$document.Slides.Item(1).Shapes.Item($name).Id })
    Test-Replay 'multi-b-extras' { $addin.ReplaySelectionMoveForTest('["B","D","ZeroWidthLine"]', '[[-10,-15],[-27,-34],[-38,-49]]') } $ids
    Select-TopShapes 1 @('A', 'B', 'D')
    $ids = @(foreach ($name in @('A', 'B', 'D')) { [int]$document.Slides.Item(1).Shapes.Item($name).Id })
    Test-Replay 'multi-two-drivers' { $addin.ReplaySelectionMoveForTest('["A","B","D"]', '[[10,12],[23,26],[38,41]]') } $ids
    Select-TopShapes 1 @('A'); $document.SnapToGrid = 0
    Test-Replay 'keyboard-down-repeats' { $addin.ReplayKeyboardNudgeForTest(40, 8, $false) } @([int]$document.Slides.Item(1).Shapes.Item('A').Id)
    Select-TopShapes 1 @('B', 'D', 'ZeroWidthLine')
    $ids = @(foreach ($name in @('B', 'D', 'ZeroWidthLine')) { [int]$document.Slides.Item(1).Shapes.Item($name).Id })
    Test-Replay 'keyboard-multi-up' { $addin.ReplayKeyboardNudgeForTest(38, 5, $true) } $ids
    Select-TopShapes 1 @('A'); $document.SnapToGrid = -1
    Test-Replay 'keyboard-grid-right' { $addin.ReplayKeyboardNudgeForTest(39, 3, $false) } @([int]$document.Slides.Item(1).Shapes.Item('A').Id)
    $document.SnapToGrid = 0
    Select-TopShapes 2 @('A', 'D')
    $ids = @(foreach ($name in @('A', 'D')) { [int]$document.Slides.Item(2).Shapes.Item($name).Id })
    Test-Replay 'multi-fixed-conflict' { $addin.ReplaySelectionMoveForTest('["A","D"]', '[[10,10],[30,40]]') } $ids -Rejected
    Select-TopShapes 3 @('AB')
    $childA = [int]$document.Slides.Item(3).Shapes.Item('AB').GroupItems.Item('A').Id
    $childB = [int]$document.Slides.Item(3).Shapes.Item('AB').GroupItems.Item('B').Id
    Test-Replay 'child-a-move' { $addin.ReplayGroupChildGestureForTest('["AB","A"]', 'Move', '[[0,-12],[0,-25],[0,-42]]') } @($childA) -Children
    Select-TopShapes 3 @('AB')
    Test-Replay 'child-b-width' { $addin.ReplayGroupChildGestureForTest('["AB","B"]', 'ResizeE', '[[12,0],[28,0],[44,0]]') } @($childB) -Children
    Select-TopShapes 3 @('AB')
    Test-Replay 'child-a-height' { $addin.ReplayGroupChildGestureForTest('["AB","A"]', 'ResizeN', '[[0,-9],[0,-19],[0,-32]]') } @($childA) -Children
    Select-TopShapes 3 @('AB')
    Test-Replay 'child-a-b-move' { $addin.ReplayGroupChildrenMoveForTest('[["AB","A"],["AB","B"]]', '[[8,10],[17,23],[31,38]]') } @($childA, $childB) -Children
    Select-GroupChild 'A'
    Test-Replay 'keyboard-child-down' { $addin.ReplayKeyboardNudgeForTest(40, 6, $true) } @($childA) -Children

    Select-TopShapes 3 @('AB')
    $group = $document.Slides.Item(3).Shapes.Item('AB')
    $parameters = @{ value = [double]$group.Left; second = [double]$group.Top } | ConvertTo-Json -Compress
    $fixed = Add-Relation 3 @('AB') 'FixedPosition' $parameters
    try {
        $beforeMarker = Get-Snapshot 'child-reject-marker-before'
        Assert-OwnedDocument -Active
        $ppt.StartNewUndoEntry()
        $free = $document.Slides.Item(1).Shapes.Item('D'); $free.Left = [single]([double]$free.Left + 7)
        $ppt.StartNewUndoEntry()
        $beforeReject = Get-Snapshot 'child-reject-before'
        Assert-OwnedDocument -Active
        $null = $addin.ReplayGroupChildGestureForTest('["AB","A"]', 'Move', '[[0,-13],[0,-28]]')
        $afterReject = Get-Snapshot 'child-reject-after'
        Assert-SameSnapshot $afterReject $beforeReject 'Fixed-root child atomic rejection'
        Assert-SelectedIds @($childA) -Children
        Select-TopShapes 3 @('AB')
        $undo = Invoke-NativeHistory 'Undo' 'child-reject-undo'
        Assert-SameSnapshot $undo $beforeMarker 'Rejected child must add no empty native Undo item'
        $redo = Invoke-NativeHistory 'Redo' 'child-reject-redo'
        Assert-SameSnapshot $redo $beforeReject 'Sentinel one Redo'
        $results.Add([pscustomobject]@{ Name = 'child-reject-history'; Kind = 'NoEmptyUndoSentinel'; Passed = $true })
        Write-Output 'PASS fixed-root outward child rejection preserves the preceding native Undo/Redo item'
    }
    finally { Assert-OwnedDocument -Active; $addin.DeleteConstraintCommand($fixed) }

    Select-TopShapes 1 @('A', 'D', 'ZeroWidthLine')
    Assert-OwnedDocument -Active
    $document.Save()
    $saved = Get-Snapshot 'saved-before-close'
    $document.Close(); $document = $null; $window = $null
    $document = $ppt.Presentations.Open($fixturePath, 0, 0, -1); $ownedDocument = $document; $window = $ppt.ActiveWindow
    Assert-OwnedDocument -Active
    $reopened = Get-Snapshot 'reopened'
    Assert-SameSnapshot $reopened $saved 'Save and reopen'
    Assert-Relations $reopened 'Save and reopen'
    Write-Output 'PASS save/reopen retained all constraint XML, tags and recursive geometry; selection was recorded but not assumed persistent'
    $results.Add([pscustomobject]@{ Name = 'save-reopen'; Kind = 'Persistence'; Passed = $true })
    Select-TopShapes 1 @('A', 'D', 'ZeroWidthLine')
    $ids = @(foreach ($name in @('A', 'D', 'ZeroWidthLine')) { [int]$document.Slides.Item(1).Shapes.Item($name).Id })
    Test-Replay 'reopened-multi-follow' { $addin.ReplaySelectionMoveForTest('["A","D","ZeroWidthLine"]', '[[8,10],[19,24],[29,37]]') } $ids
    Assert-OwnedDocument -Active
    $document.Save()
    $null = Get-Snapshot 'final'
}
catch { $failure = $_ }
finally {
    try {
        if ($null -ne $document -and $fixtureCreated) {
            if ($fixtureSaved) { Assert-OwnedDocument } else { Assert-OwnedReference }
            # Discard only transient changes in this newly created test deck.
            # User documents are never saved, marked Saved, or closed.
            $document.Saved = -1
            $document.Close(); $document = $null
        }
    }
    catch { $cleanupErrors.Add('Owned fixture cleanup: ' + $_.Exception.Message) }
    finally {
        if ($null -ne $session) {
            try { if ($null -ne $session.OriginalWindow) { $session.OriginalWindow.Activate() } }
            catch { $cleanupErrors.Add('Original window activation: ' + $_.Exception.Message) }
            try {
                $originalAfter = @(foreach ($original in $session.OriginalDocuments) {
                    $state = [pscustomobject]@{ Name = [string]$original.Document.Name; Path = [string]$original.Document.FullName; Saved = [int]$original.Document.Saved; OriginalSaved = $original.Saved }
                    if ($state.Path -cne $original.Path -or $state.Saved -ne $original.Saved) { $cleanupErrors.Add('An original document changed state: ' + $original.Name) }
                    $state
                })
                Write-Evidence 'original-documents-after' $originalAfter
            }
            catch { $cleanupErrors.Add('Original document state read: ' + $_.Exception.Message) }
            finally {
                try { Stop-OfficeComSession $session; $filterRestored = $true }
                catch { $cleanupErrors.Add('OLE message filter restoration: ' + $_.Exception.Message) }
            }
        }
    }
    if (Test-Path -LiteralPath $artifactDirectory) {
        Write-Evidence 'result' ([pscustomobject]@{
            Passed = ($null -eq $failure -and $cleanupErrors.Count -eq 0); Fixture = $fixturePath; TestTag = $testTag; Token = $token
            Cases = @($results.ToArray()); NativeUndoRedoCases = @($results | Where-Object Kind -eq 'OneNativeUndoRedo').Count
            Error = $(if ($null -ne $failure) { $failure.Exception.Message } else { $null }); CleanupErrors = @($cleanupErrors.ToArray()); Note = $note
            OleMessageFilterRestored = $filterRestored
        })
    }
}
if ($null -ne $failure) { throw $failure }
if ($cleanupErrors.Count -gt 0) { throw ('Test cleanup failed: ' + ($cleanupErrors -join '; ')) }
Write-Output ('PASS editing entries. Evidence: ' + $artifactDirectory)
Write-Output $note
