[CmdletBinding()]
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\artifacts\component-validation'),
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OfficeComSession.ps1')
$session = Start-OfficeComSession
$token = [guid]::NewGuid().ToString('D').ToUpperInvariant()
$run = [IO.Path]::GetFullPath((Join-Path $OutputDirectory $token))
$path = Join-Path $run 'ComponentCommands.pptx'
$document = $null
$passed = $false

function Assert-Owned {
    if ($document.FullName -ine $path -or $document.Tags.Item('PPTGCS_COMPONENT_TEST') -cne $token -or
        -not (Test-OfficeComIdentity $ppt.ActivePresentation $document)) { throw 'The active presentation must be the owned fixture.' }
}

function Select-NamedShapes([int]$page, [string[]]$names) {
    Assert-Owned
    $window = $ppt.ActiveWindow
    $window.View.GotoSlide($page)
    foreach ($pane in $window.Panes) { if ($pane.ViewType -eq 1) { $pane.Activate(); break } }
    $document.Slides.Item($page).Shapes.Range([object[]]$names).Select()
    Assert-Owned
}

function Read-Stage([string]$label) {
    Assert-Owned
    $state = [pscustomobject]@{
        Snapshot = $addin.GetSnapshot() | ConvertFrom-Json
        Undo = [bool]$ppt.CommandBars.GetEnabledMso('Undo')
        Redo = [bool]$ppt.CommandBars.GetEnabledMso('Redo')
        Page = [int]$ppt.ActiveWindow.View.Slide.SlideIndex
        Shapes = @(foreach ($slide in $document.Slides) {
            foreach ($shape in $slide.Shapes) {
                [pscustomobject]@{
                    Page = [int]$slide.SlideIndex; NativeId = [int]$shape.Id
                    Id = [string]$shape.Tags.Item('PPTGCS_ID'); Owner = [string]$shape.Tags.Item('PPTGCS_OWNER')
                    Name = [string]$shape.Name; Text = [string]$shape.TextFrame.TextRange.Text
                    X = [double]$shape.Left; Y = [double]$shape.Top; W = [double]$shape.Width; H = [double]$shape.Height
                }
            }
        })
    }
    $state | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath (Join-Path $run "$label.json") -Encoding UTF8
    return $state
}

function Assert-Readback($actual, $expected, [string]$label) {
    if ($actual.Snapshot.constraintXml -cne $expected.Snapshot.constraintXml -or $actual.Shapes.Count -ne $expected.Shapes.Count) {
        throw "$label changed document data or shape count."
    }
    for ($index = 0; $index -lt $actual.Shapes.Count; $index++) {
        $shape = $actual.Shapes[$index]; $old = $expected.Shapes[$index]
        foreach ($field in @('Page', 'NativeId', 'Id', 'Owner', 'Name', 'Text')) {
            if ($shape.$field -cne $old.$field) { throw "$label changed $field for $($shape.Name)." }
        }
        foreach ($field in @('X', 'Y', 'W', 'H')) {
            if ([Math]::Abs([double]$shape.$field - [double]$old.$field) -gt .005) { throw "$label changed $field for $($shape.Name)." }
        }
    }
}

function Assert-History([string]$command, [string]$label, $expected) {
    Assert-Owned
    if (-not $ppt.CommandBars.GetEnabledMso($command)) { throw "$command is unavailable." }
    $ppt.CommandBars.ExecuteMso($command)
    Start-Sleep -Milliseconds 150
    $actual = Read-Stage $label
    Assert-Readback $actual $expected $label
    return $actual
}

function Assert-ParameterGeometry($state, [int]$page, [double]$value) {
    $rules = $state.Snapshot.slides[$page - 1].constraints
    $width = $rules | Where-Object Kind -eq FixedWidth | Select-Object -First 1
    $gap = $rules | Where-Object Kind -eq FixedHorizontalGap | Select-Object -First 1
    $a = $state.Shapes | Where-Object { $_.Page -eq $page -and $_.Id -ceq $width.ShapeIds[0] }
    $b = $state.Shapes | Where-Object { $_.Page -eq $page -and $_.Id -ceq $gap.ShapeIds[1] }
    if ($null -eq $a -or $null -eq $b -or $width.Value -ne $value -or $gap.Value -ne $value -or
        [Math]::Abs($a.W - $value) -gt .005 -or [Math]::Abs($b.X - $a.X - $a.W - $value) -gt .005) {
        throw "Incorrect named size or gap on page $page."
    }
}

function Add-Card($slide, [string]$name, [int]$type, [double]$x, [double]$y, [double]$w, [double]$h) {
    $shape = $slide.Shapes.AddShape($type, $x, $y, $w, $h)
    $shape.Name = $name
    $shape.TextFrame.AutoSize = 0
    $shape.TextFrame.TextRange.Text = $name
}

try {
    $ppt = $session.Application
    $addin = $ppt.COMAddIns.Item('PptGcs.Connect').Object
    if ($null -eq $addin) { throw 'PPT GCS must be installed and loaded.' }
    New-Item -ItemType Directory -Path $run | Out-Null
    $document = $ppt.Presentations.Add(-1)
    foreach ($original in $session.OriginalDocuments) {
        if (Test-OfficeComIdentity $document $original.Document) { throw 'Cannot modify a pre-existing presentation.' }
    }
    $document.Tags.Add('PPTGCS_TEST', 'V3')
    $document.Tags.Add('PPTGCS_COMPONENT_TEST', $token)
    $document.PageSetup.SlideWidth = 960
    $document.PageSetup.SlideHeight = 540
    foreach ($index in 1..4) { [void]$document.Slides.Add($index, 12) }
    $source = $document.Slides.Item(1)
    $source.Name = 'Original'
    Add-Card $source CardA 1 80 110 84 48
    Add-Card $source CardB 9 250 225 100 48
    Add-Card $source External 1 495 220 115 62
    Add-Card $source Unconstrained 3 655 340 40 40
    $document.Slides.Item(2).Name = 'Independent'
    $document.Slides.Item(3).Name = 'Shared'
    $document.Slides.Item(4).Name = 'UI target'
    $document.SaveAs($path, 24)
    Assert-Owned

    $parameter = $addin.AddNamedParameterCommand('Slot', 84)
    $binding = @{ parameterId = $parameter } | ConvertTo-Json -Compress
    Select-NamedShapes 1 @('CardA')
    $null = $addin.AddConstraintCommand('FixedWidth', $binding)
    Select-NamedShapes 1 @('CardA', 'CardB')
    $null = $addin.AddConstraintCommand('FixedHorizontalGap', $binding)
    Select-NamedShapes 1 @('CardB', 'External')
    $null = $addin.AddConstraintCommand('SameY', '{}')
    Select-NamedShapes 1 @('CardA')
    $disabled = $addin.AddConstraintCommand('FixedHeight', '{"value":48}')
    $addin.SetConstraintEnabled($disabled, $false)
    $document.Save(); $document.Close(); $document = $null
    $document = $ppt.Presentations.Open($path, 0, 0, -1)
    Select-NamedShapes 1 @('CardA', 'CardB', 'Unconstrained')
    $before = Read-Stage 'before'

    $ids = $addin.CopyConstrainedComponent([int]$document.Slides.Item(2).SlideID, $false) | ConvertFrom-Json
    $independent = Read-Stage 'independent'
    $copiedRules = $independent.Snapshot.slides[1].constraints
    if ($ids.Count -ne 3 -or $independent.Page -ne 2 -or $independent.Snapshot.parameters.Count -ne 2 -or
        $copiedRules.Count -ne 3 -or ($copiedRules | Where-Object Kind -eq SameY) -or
        ($copiedRules | Where-Object { $_.Kind -eq 'FixedHeight' -and $_.Enabled })) { throw 'The independent copy did not preserve its internal rules.' }
    $independentParameter = ($copiedRules | Where-Object Kind -eq FixedWidth).ParameterId
    if ($independentParameter -ceq $parameter -or @($copiedRules | Where-Object ParameterId -ceq $independentParameter).Count -ne 2) {
        throw 'A private parameter must be cloned once and reused inside the copy.'
    }
    Assert-ParameterGeometry $independent 2 84
    $undo = Assert-History Undo 'undo-independent' $before
    if ($undo.Undo) { throw 'The copy added extra Undo history.' }
    [void](Assert-History Redo 'redo-independent' $independent)

    Select-NamedShapes 2 @('CardA', 'CardB')
    $samePageIds = $addin.CopyConstrainedComponent([int]$document.Slides.Item(2).SlideID, $false) | ConvertFrom-Json
    $onPage = Read-Stage 'same-page'
    if ($samePageIds.Count -ne 2 -or $onPage.Page -ne 2 -or $onPage.Snapshot.slides[1].constraints.Count -ne 6) {
        throw 'The same-page duplicate did not carry its internal relationships.'
    }
    [void](Assert-History Undo 'undo-same-page' $independent)
    [void](Assert-History Redo 'redo-same-page' $onPage)
    [void](Assert-History Undo 'restore-cross-page-baseline' $independent)

    Select-NamedShapes 1 @('CardA', 'CardB', 'Unconstrained')
    $sharedIds = $addin.CopyConstrainedComponent([int]$document.Slides.Item(3).SlideID, $true) | ConvertFrom-Json
    $shared = Read-Stage 'shared'
    if ($sharedIds.Count -ne 3 -or $shared.Snapshot.parameters.Count -ne 2 -or
        @($shared.Snapshot.slides[2].constraints | Where-Object ParameterId -ceq $parameter).Count -ne 2) { throw 'Explicit sharing was lost.' }
    if (@($shared.Shapes | Where-Object Id | Group-Object Id | Where-Object Count -gt 1).Count) { throw 'A copy reused a shape identity.' }
    $copies = @($shared.Shapes | Where-Object Page -ne 1)
    if ($copies.Count -ne 6 -or @($copies | Where-Object { -not $_.Id -or $_.Text -cne $_.Name }).Count) {
        throw 'The copies did not keep all editable text and new identities.'
    }
    [void](Assert-History Undo 'undo-shared' $independent)
    [void](Assert-History Redo 'redo-shared' $shared)

    Select-NamedShapes 3 @('CardB')
    $null = $addin.AddConstraintCommand('FixedWidth', '{"value":84}')
    $fixed = Read-Stage 'fixed-before-conflict'
    Select-NamedShapes 3 @('CardA', 'CardB')
    $null = $addin.AddConstraintCommand('SameWidth', '{}')
    $conflict = Read-Stage 'before-conflicted-update'
    $rejected = $false
    try { $addin.UpdateNamedParameterCommand($parameter, 'Rejected', 110) }
    catch { $rejected = $true; $_.Exception.ToString() | Set-Content -LiteralPath (Join-Path $run 'rejection.txt') -Encoding UTF8 }
    if (-not $rejected) { throw 'An update conflicting on a different slide was accepted.' }
    Assert-Readback (Read-Stage 'rejected-update') $conflict 'Conflict rollback across pages'
    [void](Assert-History Undo 'undo-same-width' $fixed)
    [void](Assert-History Undo 'undo-locked-width' $shared)

    $addin.UpdateNamedParameterCommand($parameter, 'Slot width and gap', 110)
    $updated = Read-Stage 'updated'
    Assert-ParameterGeometry $updated 1 110
    Assert-ParameterGeometry $updated 2 84
    Assert-ParameterGeometry $updated 3 110
    [void](Assert-History Undo 'undo-parameter' $shared)
    [void](Assert-History Redo 'redo-parameter' $updated)

    $rejected = $false
    try { $addin.DeleteNamedParameterCommand($parameter) } catch { $rejected = $true }
    if (-not $rejected) { throw 'The deletion of a referenced parameter was accepted.' }
    Assert-Readback (Read-Stage 'rejected-deletion') $updated 'Rejected parameter deletion'
    $document.Save(); $document.Close(); $document = $null
    $document = $ppt.Presentations.Open($path, 0, 0, -1)
    Assert-Readback (Read-Stage 'reopened') $updated 'Save and reopen'
    Select-NamedShapes 1 @('CardA', 'CardB', 'Unconstrained')
    [pscustomobject]@{ Passed = $true; Path = $path; Token = $token; Shared = $parameter; Independent = $independentParameter;
        NativeSingleUndoRedo = $true; Reopen = $true; ConflictedUpdate = 'all pages unchanged'; BlockedDeletion = $true } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'acceptance.json') -Encoding UTF8
    $passed = $true
    Write-Output "Passed same-slide and cross-slide copy, parameter sharing, atomic conflicts, native history and persistence. $run"
}
finally {
    if ($null -ne $document -and (-not $KeepOpen -or -not $passed) -and $document.Tags.Item('PPTGCS_COMPONENT_TEST') -ceq $token) {
        $document.Saved = -1
        $document.Close()
    }
    if ((-not $KeepOpen -or -not $passed) -and $session.OriginalWindow) { $session.OriginalWindow.Activate() }
    Stop-OfficeComSession $session
}
