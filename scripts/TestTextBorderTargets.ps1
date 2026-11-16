[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$fixture=[IO.Path]::GetFullPath($Path)
$directory=Join-Path $root 'artifacts\text-border-fix'
New-Item -ItemType Directory -Force -Path $directory | Out-Null
Add-Type -AssemblyName System.Drawing
[void][Reflection.Assembly]::LoadFrom((Join-Path $root 'bin\PptGcs.Core.dll'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $root 'bin\PptGcs.AddIn.dll'))
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class TextBorderNativeTarget {
    [StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
    [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(Point point);
    [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr window);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr window,StringBuilder text,int size);
    public static string Class(IntPtr window) { var text=new StringBuilder(128);GetClassName(window,text,text.Capacity);return text.ToString(); }
}
"@
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p=$ppt.ActivePresentation
$window=$ppt.ActiveWindow
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
$results=[Collections.Generic.List[object]]::new()
$passed=$false;$failure=$null;$restoreFailure=$null;$mutated=$false
$before=$null;$after=$null;$originalZoom=$null;$originalPane=$null;$originalSelection=$null
function Assert-Owned {
    if($p.FullName -ne $fixture -or $ppt.ActivePresentation.FullName -ne $fixture -or $p.Tags.Item('PPTGCS_TEST') -ne 'V3'){
        throw 'The active document is not the requested owned V3 border fixture.'
    }
}
function Read-Tags($object) {
    @(for($i=1;$i -le $object.Tags.Count;$i++){[pscustomobject]@{Name=$object.Tags.Name($i);Value=$object.Tags.Value($i)}}) | Sort-Object Name
}
function Read-State {
    Assert-Owned
    $xml=$addin.GetConstraintXml()
    Assert-Owned
    $state=[pscustomobject][ordered]@{
        ConstraintXml=$xml;Tags=@(Read-Tags $p)
        Slides=@(foreach($slide in $p.Slides){[pscustomobject][ordered]@{Id=$slide.SlideID;Name=$slide.Name;Tags=@(Read-Tags $slide);Shapes=@(foreach($shape in $slide.Shapes){
            $text=$null
            if([int]$shape.HasTextFrame -ne 0){$text=[pscustomobject]@{Value=[string]$shape.TextFrame.TextRange.Text;AutoSize=$shape.TextFrame.AutoSize;AutoSize2=$shape.TextFrame2.AutoSize}}
            [pscustomobject][ordered]@{Id=$shape.Id;Name=$shape.Name;Type=$shape.Type;Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height;Rotation=$shape.Rotation;LockAspectRatio=$shape.LockAspectRatio;Tags=@(Read-Tags $shape);Text=$text}
        })}})
    }
    Assert-Owned
    return $state
}
function Json($value){ConvertTo-Json -InputObject $value -Depth 35 -Compress}
function Read-Selection {
    $selection=$window.Selection
    if($selection.HasChildShapeRange){throw 'Exit group-child editing before this read-only target check.'}
    $type=[int]$selection.Type
    if($type -notin @(0,2,3)){throw 'Use a normal canvas object/text selection before this read-only target check.'}
    $names=@()
    if($type -in @(2,3)){$names=@(for($i=1;$i -le $selection.ShapeRange.Count;$i++){[string]$selection.ShapeRange.Item($i).Name})}
    [pscustomobject]@{Type=$type;Names=$names;Start=$(if($type -eq 3){$selection.TextRange.Start}else{$null});Length=$(if($type -eq 3){$selection.TextRange.Length}else{$null})}
}
function Restore-Selection($selection) {
    $window.Selection.Unselect()
    if($selection.Type -eq 2){$window.View.Slide.Shapes.Range([object[]]$selection.Names).Select()}
    elseif($selection.Type -eq 3){$window.View.Slide.Shapes.Item($selection.Names[0]).TextFrame.TextRange.Characters([int]$selection.Start,[int]$selection.Length).Select()}
}
function Screen-Bounds($shape) {
    $left=[single]$window.PointsToScreenPixelsX([single]$shape.Left)
    $top=[single]$window.PointsToScreenPixelsY([single]$shape.Top)
    $right=[single]$window.PointsToScreenPixelsX([single]($shape.Left+$shape.Width))
    $bottom=[single]$window.PointsToScreenPixelsY([single]($shape.Top+$shape.Height))
    return [Drawing.RectangleF]::new($left,$top,$right-$left,$bottom-$top)
}
function Stable-ScreenBounds($shape) {
    # Office's Zoom/Select RPC may return before its view has repainted. Wait
    # separately from assertions, then require two equal coordinate readings.
    $previous=Screen-Bounds $shape
    for($attempt=0;$attempt -lt 5;$attempt++){
        Start-Sleep -Milliseconds 50
        $current=Screen-Bounds $shape
        if($current.Equals($previous)){return $current}
        $previous=$current
    }
    throw 'PowerPoint screen-coordinate projection did not become stable.'
}
function Read-Dpi($bounds) {
    $point=[TextBorderNativeTarget+Point]::new()
    $point.X=[int][math]::Round($bounds.Left+$bounds.Width/2)
    $point.Y=[int][math]::Round($bounds.Top+$bounds.Height/2)
    $native=[TextBorderNativeTarget]::WindowFromPoint($point)
    $class=[TextBorderNativeTarget]::Class($native)
    if($class -ne 'mdiClass'){throw ('The target center is not on the PowerPoint canvas (native class '+$class+'); make the owned fixture visible and rerun.')}
    $dpi=[int][TextBorderNativeTarget]::GetDpiForWindow($native)
    if($dpi -le 0){throw 'PowerPoint canvas DPI is unavailable.'}
    return $dpi
}
function Sample($shape,[bool]$selected,[int]$zoom,$bounds,[single]$dpi,[string]$edge,[string]$zone,[single]$x,[single]$y) {
    Assert-Owned
    $point=[Drawing.PointF]::new([single][math]::Round($x),[single][math]::Round($y))
    $nativePoint=[TextBorderNativeTarget+Point]::new();$nativePoint.X=[int]$point.X;$nativePoint.Y=[int]$point.Y
    $surface=[TextBorderNativeTarget]::WindowFromPoint($nativePoint)
    $surfaceClass=[TextBorderNativeTarget]::Class($surface)
    $hitId=$null;$hitError=$null
    try{
        $hit=$window.RangeFromPoint([int]$point.X,[int]$point.Y)
        if($null -ne $hit){$hit=[PptGcs.AddIn.ShapeCapabilities]::TopLevel($hit);$hitId=[int]$hit.Id}
    }catch{$hitError=$_.Exception.Message}
    [Nullable[bool]]$nativeMatches=$null
    if($null -ne $hitId){$nativeMatches=($hitId -eq [int]$shape.Id)}
    $border=[PptGcs.AddIn.TextPointerPolicy]::HitMoveBorder([int]$shape.Type,$false,$bounds,$point,[single](6*$dpi/96))
    $allowed=[PptGcs.AddIn.TextPointerPolicy]::CanUseBorderHit($selected,$nativeMatches)
    $text=[PptGcs.AddIn.TextPointerPolicy]::UseNativeTextInput([int]$shape.Type,$false,$bounds,$point,[single](6*$dpi/96))
    $classification=if($text){'native-text'}elseif($border -and $allowed){'gcs-move'}else{'native-other'}
    $check=$null -eq $hitError -and $surfaceClass -eq 'mdiClass'
    if($zone -eq 'center'){$check=$check -and $classification -eq 'native-text'}
    elseif($selected -or $zone -eq 'inner'){$check=$check -and $classification -eq 'gcs-move';if($zone -eq 'inner'){$check=$check -and $nativeMatches -eq $true}}
    else{$check=$check -and ($classification -ne 'gcs-move' -or $nativeMatches -eq $true)}
    $handleDistance=$null
    if($zone -ne 'center'){
        $cx=$bounds.Left+$bounds.Width/2;$cy=$bounds.Top+$bounds.Height/2
        $handles=@(@($bounds.Left,$bounds.Top),@($cx,$bounds.Top),@($bounds.Right,$bounds.Top),@($bounds.Right,$cy),@($bounds.Right,$bounds.Bottom),@($cx,$bounds.Bottom),@($bounds.Left,$bounds.Bottom),@($bounds.Left,$cy))
        $handleDistance=($handles | ForEach-Object{[math]::Sqrt([math]::Pow($point.X-$_[0],2)+[math]::Pow($point.Y-$_[1],2))}|Measure-Object -Minimum).Minimum
        $check=$check -and $handleDistance -gt (8*$dpi/96)
    }
    $results.Add([pscustomobject]@{Zoom=$zoom;Shape=[string]$shape.Name;NativeId=[int]$shape.Id;Selected=$selected;SelectionType=[int]$window.Selection.Type;Edge=$edge;Zone=$zone;Dpi=$dpi;Bounds=@{Left=$bounds.Left;Top=$bounds.Top;Right=$bounds.Right;Bottom=$bounds.Bottom};Point=@{X=$point.X;Y=$point.Y};SurfaceClass=$surfaceClass;RangeId=$hitId;RangeError=$hitError;NativeHitMatches=$nativeMatches;Border=$border;BorderAllowed=$allowed;NativeText=$text;Classification=$classification;NearestHandleDistance=$handleDistance;Passed=$check})
}
try {
    Assert-Owned
    $slide=$window.View.Slide
    $shapes=@($slide.Shapes.Item('Text_A'),$slide.Shapes.Item('Text_B'))
    if(@($shapes|Where-Object{[int]$_.Type -ne 17}).Count){throw 'Text_A and Text_B must be top-level text boxes in the current slide.'}
    $before=Read-State
    $originalZoom=[int]$window.View.Zoom;$originalPane=$window.ActivePane;$originalSelection=Read-Selection
    $mutated=$true
    $window.Activate()
    foreach($pane in $window.Panes){if([int]$pane.ViewType -eq 1){$pane.Activate();break}}
    if([int]$window.ActivePane.ViewType -ne 1){throw 'Could not activate the owned slide editing pane.'}
    Start-Sleep -Milliseconds 80
    foreach($zoom in @(75,100,125)){
        $window.View.Zoom=$zoom
        Start-Sleep -Milliseconds 500
        foreach($shape in $shapes){
            $shape.Select()
            Start-Sleep -Milliseconds 80
            $bounds=Stable-ScreenBounds $shape;$dpi=Read-Dpi $bounds;$offset=[single](3*$dpi/96)
            foreach($selected in @($true,$false)){
                if($selected){$shape.Select()}else{$window.Selection.Unselect()}
                Start-Sleep -Milliseconds 80
                $bounds=Stable-ScreenBounds $shape
                foreach($zone in @('inner','outer')){
                    $offsetDirection=if($zone -eq 'inner'){1}else{-1}
                    Sample $shape $selected $zoom $bounds $dpi 'top' $zone ($bounds.Left+$bounds.Width/4) ($bounds.Top+$offsetDirection*$offset)
                    Sample $shape $selected $zoom $bounds $dpi 'bottom' $zone ($bounds.Left+$bounds.Width/4) ($bounds.Bottom-$offsetDirection*$offset)
                    Sample $shape $selected $zoom $bounds $dpi 'left' $zone ($bounds.Left+$offsetDirection*$offset) ($bounds.Top+$bounds.Height/4)
                    Sample $shape $selected $zoom $bounds $dpi 'right' $zone ($bounds.Right-$offsetDirection*$offset) ($bounds.Top+$bounds.Height/4)
                }
                Sample $shape $selected $zoom $bounds $dpi 'center' 'center' ($bounds.Left+$bounds.Width/2) ($bounds.Top+$bounds.Height/2)
            }
        }
    }
    $bad=@($results|Where-Object{-not $_.Passed})
    if($bad.Count){throw ($bad.Count.ToString()+' native border targets failed; inspect the report samples.')}
    $passed=$true
}catch{$failure=$_.Exception.ToString()}
finally{
    if($mutated){try{Assert-Owned;$window.View.Zoom=$originalZoom;Restore-Selection $originalSelection;$originalPane.Activate()}catch{$restoreFailure=$_.Exception.ToString();$passed=$false}}
    if($null -ne $before){try{$after=Read-State;if((Json $before) -cne (Json $after)){$passed=$false;throw 'Geometry, text, tags or XML changed during the target check.'}}catch{$passed=$false;if(-not $failure){$failure=$_.Exception.ToString()}}}
    $report=[pscustomobject]@{Passed=$passed;PhysicalMouseInput=$false;ProgrammaticSelection=$true;ZoomValues=@(75,100,125);ZoomWaitMilliseconds=500;SelectionWaitMilliseconds=80;CoordinateQuiescenceMilliseconds=50;Samples=$results;SampleCount=$results.Count;DocumentStateUnchanged=$(if($null -ne $before -and $null -ne $after){(Json $before) -ceq (Json $after)}else{$false});RestoredZoom=$originalZoom;Failure=$failure;RestorationFailure=$restoreFailure;Utc=[DateTime]::UtcNow.ToString('O')}
    $report|ConvertTo-Json -Depth 15|Set-Content -LiteralPath (Join-Path $directory 'native-hit-targets.json') -Encoding UTF8
    $report|Select-Object Passed,PhysicalMouseInput,SampleCount,DocumentStateUnchanged,Failure,RestorationFailure|ConvertTo-Json -Depth 5
}
if(-not $passed){throw 'Native text border target verification failed; see native-hit-targets.json.'}
