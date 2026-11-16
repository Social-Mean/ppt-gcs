$ErrorActionPreference='Stop'
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$document=$ppt.ActivePresentation
if ($document.Slides.Count -ne 6) { throw 'Expected the six-slide layout test' }
$results=@()
for ($index=1; $index -le 6; $index++) {
    $slide=$document.Slides.Item($index)
    if (-not $slide.Tags.Item('PPTGCS_CASE')) { throw 'Missing layout test marker' }
    $ppt.ActiveWindow.View.GotoSlide($index)
    $selected=$slide.Shapes.Item('GCS_'+$slide.Tags.Item('PPTGCS_CASE')+'_A')
    $selected.Select()
    $diagnostics=$ppt.COMAddIns.Item('PptGcs.Connect').Object.GetDiagnostics() | ConvertFrom-Json
    if ($diagnostics.error) { throw $diagnostics.error }
    foreach($shape in $diagnostics.shapes) {
        if (-not $shape.canvas -or $shape.rangeId -ne $shape.nativeId -or $shape.handles.Count -ne 8) {
            throw ('Preflight failed for '+$shape.name)
        }
        if (@($shape.handles | Where-Object { -not $_.canvas }).Count) { throw ('Handle surface preflight failed for '+$shape.name) }
    }
    $results += $diagnostics
}
$ppt.ActiveWindow.View.GotoSlide(1)
$folder=Join-Path $PSScriptRoot '..\artifacts\layout-v2-validation'
New-Item -ItemType Directory -Force -Path $folder | Out-Null
$results | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $folder 'preflight.json') -Encoding UTF8
$results | ForEach-Object { [pscustomobject]@{Slide=$_.slideId;Shapes=$_.shapes.Count;Kind=$_.constraints[0].kind;Passed=$true} } | ConvertTo-Json
