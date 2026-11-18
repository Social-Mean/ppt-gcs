[CmdletBinding()]
param([string]$OutputDirectory = (Join-Path $PSScriptRoot '..\artifacts\suggestion-validation'))

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OfficeComSession.ps1')
$session = Start-OfficeComSession
$token = [Guid]::NewGuid().ToString('D').ToUpperInvariant()
$run = [IO.Path]::GetFullPath((Join-Path $OutputDirectory $token))
$path = Join-Path $run 'SuggestionValidation.pptx'
$document = $null
$ready = $false

function Add-Box($slide, [string]$name, [double]$x, [double]$y, [double]$w, [double]$h) {
    $shape = $slide.Shapes.AddShape(1, $x, $y, $w, $h)
    $shape.Name = $name
    $shape.TextFrame.AutoSize = 0
    $shape.TextFrame.TextRange.Text = $name
}

function Select-Box([int]$page, [string[]]$names) {
    if (-not (Test-OfficeComIdentity $ppt.ActivePresentation $document) -or $document.Tags.Item('PPTGCS_SUGGESTION_TEST') -cne $token) {
        throw 'The active presentation must be the test fixture.'
    }
    $window = $ppt.ActiveWindow
    $window.View.GotoSlide($page)
    foreach ($pane in $window.Panes) { if ($pane.ViewType -eq 1) { $pane.Activate(); break } }
    $document.Slides.Item($page).Shapes.Range([object[]]$names).Select()
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
    $document.Tags.Add('PPTGCS_SUGGESTION_TEST', $token)
    $document.PageSetup.SlideWidth = 960
    $document.PageSetup.SlideHeight = 540
    $document.SnapToGrid = 0
    $slide = $document.Slides.Add(1, 12); $slide.Name = 'Top alignment'
    Add-Box $slide A 111 175.6 70 42
    Add-Box $slide B 280 180 110 65
    $slide = $document.Slides.Add(2, 12); $slide.Name = 'Same width'
    Add-Box $slide A 120 120 80 50
    Add-Box $slide B 300 200 130 92
    Select-Box 2 @('A')
    $null = $addin.AddConstraintCommand('FixedPosition', '{"value":120,"second":120}')
    $slide = $document.Slides.Add(3, 12); $slide.Name = 'Equal horizontal gaps'
    Add-Box $slide A 80 170 70 54
    Add-Box $slide B 222 177 58 43
    Add-Box $slide C 396 172 95 67
    Select-Box 3 @('C')
    $null = $addin.AddConstraintCommand('FixedHeight', '{"value":67}')
    $slide = $document.Slides.Add(4, 12); $slide.Name = 'Already aligned'
    Add-Box $slide A 131 181 63 42
    Add-Box $slide B 280 181 110 65
    Select-Box 4 @('A','B')
    $null = $addin.AddConstraintCommand('SameY', '{}')
    $slide = $document.Slides.Add(5, 12); $slide.Name = 'Conflicting candidates'
    Add-Box $slide A 120 180 80 42
    Add-Box $slide B 300 182 130 100
    Select-Box 5 @('A')
    $null = $addin.AddConstraintCommand('FixedPosition', '{"value":120,"second":180}')
    Select-Box 5 @('B')
    $null = $addin.AddConstraintCommand('FixedPosition', '{"value":300,"second":182}')
    $slide = $document.Slides.Add(6, 12); $slide.Name = 'Cancellation and focus'
    Add-Box $slide A 111 130 70 42
    Add-Box $slide B 280 180 110 65
    Select-Box 6 @('A')
    $null = $addin.AddConstraintCommand('FixedWidth', '{"value":70}')
    $document.SaveAs($path, 24)
    $document.Close(); $document = $null
    $document = $ppt.Presentations.Open($path, 0, 0, -1)
    Select-Box 1 @('A')
    $addin.SetCanvasDisplay('Selection')
    [pscustomobject]@{ Path = $path; Token = $token; FixtureReady = $true } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'fixture.json') -Encoding UTF8
    $ready = $true
    Write-Output "Prepared the native suggestion fixture at $run."
}
finally {
    if ($null -ne $document -and -not $ready -and $document.Tags.Item('PPTGCS_SUGGESTION_TEST') -ceq $token) {
        $document.Saved = -1
        $document.Close()
    }
    if (-not $ready -and $session.OriginalWindow) { $session.OriginalWindow.Activate() }
    Stop-OfficeComSession $session
}
