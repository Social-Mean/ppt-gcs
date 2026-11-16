$ErrorActionPreference='Stop'
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$window=$ppt.ActiveWindow
if ($window.Presentation.Slides.Count -ne 6 -or -not $window.View.Slide.Tags.Item('PPTGCS_CASE')) { throw 'Expected layout test presentation' }
$slidePane=@($window.Panes | Where-Object { $_.ViewType -eq 1 })[0]
$slidePane.Activate()
$shape=@($window.View.Slide.Shapes | Where-Object { $_.Tags.Item('PPTGCS_ID') })[0]
$probeX=$window.PointsToScreenPixelsX($shape.Left+$shape.Width/2)
$probeY=$window.PointsToScreenPixelsY($shape.Top+$shape.Height/2)
$results=@()
try {
    for($i=1; $i -le $window.Panes.Count; $i++) {
        $pane=$window.Panes.Item($i)
        $pane.Activate()
        $hitId=$null; $hitError=$null; $mapError=$null; $x=$null; $y=$null
        try { $hit=$window.RangeFromPoint($probeX,$probeY); $hitId=$hit.Id } catch { $hitError=$_.Exception.Message }
        try { $x=$window.PointsToScreenPixelsX(0); $y=$window.PointsToScreenPixelsY(0) } catch { $mapError=$_.Exception.Message }
        $results += [pscustomobject]@{Pane=$i;ViewType=$pane.ViewType;Hit=$hitId;HitError=$hitError;X=$x;Y=$y;MapError=$mapError}
    }
} finally { $slidePane.Activate() }
$results | ConvertTo-Json
