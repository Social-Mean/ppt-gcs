[CmdletBinding()]
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\artifacts\reference-constraints'),
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OfficeComSession.ps1')
$session = Start-OfficeComSession
$token = [Guid]::NewGuid().ToString('D').ToUpperInvariant()
$runDirectory = [IO.Path]::GetFullPath((Join-Path $OutputDirectory $token))
$path = Join-Path $runDirectory 'ReferenceConstraints.pptx'
$document = $null
$passed = $false

function Assert-Owned {
    if ($document.FullName -ine $path -or $document.Tags.Item('PPTGCS_REFERENCE_TEST') -cne $token -or
        -not (Test-OfficeComIdentity $ppt.ActivePresentation $document)) { throw 'The active presentation is not the test fixture.' }
}

function Add-Rectangle($slide, [string]$name, [double]$x, [double]$y, [double]$width, [double]$height) {
    $shape = $slide.Shapes.AddShape(1, $x, $y, $width, $height)
    $shape.Name = $name
    if ($name -eq 'Frame') { $shape.Fill.Visible = 0; $shape.Line.Weight = 2 }
    else { $shape.TextFrame.AutoSize = 0; $shape.TextFrame.TextRange.Text = $name }
}

function Add-Rule([int]$page, [string[]]$members, [string]$kind, [double]$value, [string]$reference) {
    Assert-Owned
    $window.View.GotoSlide($page)
    foreach ($pane in $window.Panes) { if ($pane.ViewType -eq 1) { $pane.Activate(); break } }
    $slide = $document.Slides.Item($page)
    $slide.Shapes.Range([object[]]$members).Select()
    Assert-Owned
    if ($reference) {
        $referenceKey = 'shape:{0}/{1}' -f $slide.SlideID,$slide.Shapes.Item($reference).Id
        return $addin.AddContainerConstraintCommand($kind, $value, $referenceKey)
    }
    return $addin.AddConstraintCommand($kind, (@{ value = $value } | ConvertTo-Json -Compress))
}

function Read-State {
    Assert-Owned
    $pages = @(foreach ($slide in $document.Slides) {
        [pscustomobject]@{
            Id = [int]$slide.SlideID
            Name = [string]$slide.Name
            Shapes = @(foreach ($shape in $slide.Shapes) {
                [pscustomobject]@{
                    Id = [int]$shape.Id; Name = [string]$shape.Name; Tag = [string]$shape.Tags.Item('PPTGCS_ID')
                    X = [double]$shape.Left; Y = [double]$shape.Top; W = [double]$shape.Width; H = [double]$shape.Height
                }
            })
        }
    })
    [pscustomobject]@{ Width = [double]$document.PageSetup.SlideWidth; Height = [double]$document.PageSetup.SlideHeight;
        Pages = $pages; Xml = [string]$addin.GetConstraintXml() }
}

function Assert-Near([double]$actual, [double]$expected, [string]$label) {
    if ([Math]::Abs($actual - $expected) -gt 0.005) { throw "$label expected $expected, observed $actual" }
}

function Assert-Same($actual, $expected, [string]$label) {
    if ($actual.Xml -cne $expected.Xml -or $actual.Pages.Count -ne $expected.Pages.Count) { throw "$label changed the relationship document." }
    Assert-Near $actual.Width $expected.Width "$label page width"
    Assert-Near $actual.Height $expected.Height "$label page height"
    for ($i = 0; $i -lt $actual.Pages.Count; $i++) {
        $page = $actual.Pages[$i]; $old = $expected.Pages[$i]
        if ($page.Id -ne $old.Id -or $page.Shapes.Count -ne $old.Shapes.Count) { throw "$label changed slide membership." }
        for ($j = 0; $j -lt $page.Shapes.Count; $j++) {
            $shape = $page.Shapes[$j]; $prior = $old.Shapes[$j]
            if ($shape.Id -ne $prior.Id -or $shape.Tag -cne $prior.Tag -or $shape.Name -cne $prior.Name) {
                throw "$label changed shape identity."
            }
            foreach ($prop in @('X','Y','W','H')) { Assert-Near $shape.$prop $prior.$prop "$label $($shape.Name) $prop" }
        }
    }
}

function Assert-ReferenceGeometry($state) {
    $p = @{}; foreach ($shape in $state.Pages[0].Shapes) { $p[$shape.Name] = $shape }
    foreach ($shape in @($p.A,$p.B)) {
        Assert-Near $shape.X 45 "$($shape.Name) page left"
        Assert-Near ($state.Width-$shape.X-$shape.W) 45 "$($shape.Name) page right"
    }
    Assert-Near $p.A.Y 20 'page top'
    Assert-Near ($state.Height-$p.B.Y-$p.B.H) 22 'page bottom'
    Assert-Near ($p.C.X+$p.C.W/2-$state.Width/2) 7 'page horizontal center offset'
    Assert-Near ($p.C.Y+$p.C.H/2-$state.Height/2) -12 'page vertical center offset'
    $s = @{}; foreach ($shape in $state.Pages[1].Shapes) { $s[$shape.Name] = $shape }
    $frame = $s.Frame
    foreach ($shape in @($s.A,$s.B)) {
        Assert-Near ($shape.X-$frame.X) 22 "$($shape.Name) container left"
        Assert-Near ($frame.X+$frame.W-$shape.X-$shape.W) 34 "$($shape.Name) container right"
    }
    Assert-Near ($s.A.Y-$frame.Y) 24 'container top'
    Assert-Near ($frame.Y+$frame.H-$s.B.Y-$s.B.H) 26 'container bottom'
    Assert-Near ($s.C.X+$s.C.W/2-$frame.X-$frame.W/2) 6 'container horizontal center offset'
    Assert-Near ($s.C.Y+$s.C.H/2-$frame.Y-$frame.H/2) -9 'container vertical center offset'
    $fixed = $state.Pages[2].Shapes[0]
    Assert-Near $fixed.X 70 'absolute left on a third slide'
    Assert-Near $fixed.Y 95 'absolute top on a third slide'
}

try {
    $ppt = $session.Application
    $addin = $ppt.COMAddIns.Item('PptGcs.Connect').Object
    if ($null -eq $addin) { throw 'PPT GCS must be installed and loaded.' }
    New-Item -ItemType Directory -Path $runDirectory -Force | Out-Null
    $document = $ppt.Presentations.Add(-1)
    foreach ($previous in $session.OriginalDocuments) {
        if (Test-OfficeComIdentity $previous.Document $document) { throw 'The test must own its presentation.' }
    }
    $document.Tags.Add('PPTGCS_REFERENCE_TEST', $token)
    $document.Tags.Add('PPTGCS_TEST', 'V3')
    $document.PageSetup.SlideWidth = 960
    $document.PageSetup.SlideHeight = 540
    $document.SnapToGrid = 0
    $window = $document.Windows.Item(1)
    $page = $document.Slides.Add(1, 12); $page.Name = 'Page margins'
    Add-Rectangle $page A 180 60 100 65
    Add-Rectangle $page B 260 400 140 65
    Add-Rectangle $page C 440 210 110 60
    Add-Rectangle $page Free 720 355 92 50
    $page = $document.Slides.Add(2, 12); $page.Name = 'Container insets'
    Add-Rectangle $page A 260 165 110 48
    Add-Rectangle $page B 280 318 135 54
    Add-Rectangle $page C 370 230 100 55
    Add-Rectangle $page Frame 210 130 430 270
    $page = $document.Slides.Add(3, 12); $page.Name = 'Absolute geometry'
    Add-Rectangle $page Fixed 70 95 140 65
    Add-Rectangle $page Free 410 220 90 50
    $document.SaveAs($path, 24)

    $ids = [ordered]@{}
    $ids.PageLeft = Add-Rule 1 @('A','B') PageLeft 45
    $ids.PageRight = Add-Rule 1 @('A','B') PageRight 45
    $ids.PageTop = Add-Rule 1 @('A') PageTop 20
    $ids.PageBottom = Add-Rule 1 @('B') PageBottom 22
    $ids.PageCenterX = Add-Rule 1 @('C') PageCenterX 7
    $ids.PageCenterY = Add-Rule 1 @('C') PageCenterY -12
    $ids.ContainerLeft = Add-Rule 2 @('A','B','Frame') ContainerLeft 22 Frame
    $ids.ContainerRight = Add-Rule 2 @('A','B','Frame') ContainerRight 34 Frame
    $ids.ContainerTop = Add-Rule 2 @('A','Frame') ContainerTop 24 Frame
    $ids.ContainerBottom = Add-Rule 2 @('B','Frame') ContainerBottom 26 Frame
    $ids.ContainerCenterX = Add-Rule 2 @('C','Frame') ContainerCenterX 6 Frame
    $ids.ContainerCenterY = Add-Rule 2 @('C','Frame') ContainerCenterY -9 Frame
    $window.View.GotoSlide(3)
    foreach ($pane in $window.Panes) { if ($pane.ViewType -eq 1) { $pane.Activate(); break } }
    $document.Slides.Item(3).Shapes.Item('Fixed').Select()
    Assert-Owned
    $ids.FixedPosition = $addin.AddConstraintCommand('FixedPosition', '{"value":70,"second":95}')

    $document.Save()
    $document.Close()
    $document = $null
    $document = $ppt.Presentations.Open($path, 0, 0, -1)
    $window = $ppt.ActiveWindow
    $initial = Read-State
    Assert-ReferenceGeometry $initial
    if ($ppt.CommandBars.GetEnabledMso('Undo')) { throw 'The reopened fixture started with Undo enabled.' }
    $rejection = $null
    try { $addin.ResizePageCommand(72, 610) } catch { $rejection = $_.Exception.ToString() }
    if (-not $rejection -or $rejection -notmatch 'PageLeft|页面左边距') { throw 'An impossible page width must report the conflicting margin.' }
    Assert-Same (Read-State) $initial 'preflight rejection'
    if ($ppt.CommandBars.GetEnabledMso('Undo')) { throw 'A rejected resize added an Undo entry.' }

    $addin.ResizePageCommand(1040, 610)
    $resized = Read-State
    Assert-ReferenceGeometry $resized
    if ($resized.Xml -cne $initial.Xml) { throw 'Page resize altered persistent constraint parameters.' }
    Assert-Near $resized.Width 1040 'native resized slide width'
    Assert-Near $resized.Height 610 'native resized slide height'
    foreach ($index in @(0,2)) {
        $free = $resized.Pages[$index].Shapes | Where-Object Name -eq Free
        $old = $initial.Pages[$index].Shapes | Where-Object Name -eq Free
        Assert-Near $free.X ($old.X+40) 'native free-shape recenter X'
        Assert-Near $free.Y ($old.Y+35) 'native free-shape recenter Y'
        Assert-Near $free.W $old.W 'native free-shape width'
        Assert-Near $free.H $old.H 'native free-shape height'
    }
    $ppt.CommandBars.ExecuteMso('Undo')
    Start-Sleep -Milliseconds 200
    Assert-Same (Read-State) $initial 'one native Undo'
    if ($ppt.CommandBars.GetEnabledMso('Undo')) { throw 'Page resizing created more than one Undo entry.' }
    $ppt.CommandBars.ExecuteMso('Redo')
    Start-Sleep -Milliseconds 200
    Assert-Same (Read-State) $resized 'one native Redo'
    if ($ppt.CommandBars.GetEnabledMso('Redo')) { throw 'Page resizing created more than one Redo entry.' }

    $document.Save()
    $document.Close()
    $document = $null
    $document = $ppt.Presentations.Open($path, 0, 0, -1)
    Assert-Same (Read-State) $resized 'save and reopen'
    Assert-ReferenceGeometry (Read-State)
    $window = $ppt.ActiveWindow
    $window.View.GotoSlide(2)
    $result = [pscustomobject]@{ Passed = $true; Path = $path; Token = $token; Relations = $ids;
        Initial = $initial; Resized = $resized; Rejection = $rejection; History = 'one Undo, one Redo'; Reopened = $true }
    $result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $runDirectory 'acceptance.json') -Encoding UTF8
    $passed = $true
    Write-Output "Passed all 12 reference kinds, rejected size, multi-slide resize, native history and reopen. Report: $runDirectory\acceptance.json"
}
finally {
    if ($null -ne $document -and (-not $KeepOpen -or -not $passed) -and $document.Tags.Item('PPTGCS_REFERENCE_TEST') -ceq $token) {
        $document.Saved = -1
        $document.Close()
    }
    if ((-not $KeepOpen -or -not $passed) -and $session.OriginalWindow) { $session.OriginalWindow.Activate() }
    Stop-OfficeComSession $session
}
