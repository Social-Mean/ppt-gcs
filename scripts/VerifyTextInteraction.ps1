[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root 'artifacts\ui-refinement'
$fixture=Join-Path $directory 'interaction.pptx'
$baseline=Get-Content -LiteralPath (Join-Path $directory 'manual-before.json') -Raw | ConvertFrom-Json
$start=[DateTimeOffset]::Parse($baseline.Utc)
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
$p=$ppt.ActivePresentation
$undo=0;$redo=0;$passed=$false;$failure=$null;$restoreFailure=$null
$gestures=@();$heldFrames=0;$nativeTextRoutes=0;$after=$null
$checks=[Collections.Generic.List[string]]::new()

function Assert-Owned {
    if($null -eq $p -or $p.FullName -ne $fixture -or $ppt.ActivePresentation.FullName -ne $fixture -or $p.Tags.Item('PPTGCS_TEST') -ne 'V3'){
        throw 'The active document is not the owned text-interaction fixture.'
    }
}
function Read-Tags($target) {
    @(for($i=1;$i -le $target.Tags.Count;$i++){
        [pscustomobject]@{Name=$target.Tags.Name($i);Value=$target.Tags.Value($i)}
    }) | Sort-Object Name
}
function Read-State {
    Assert-Owned
    $xml=$addin.GetConstraintXml()
    Assert-Owned
    $snapshot=$addin.GetSnapshot() | ConvertFrom-Json
    Assert-Owned
    if($snapshot.presentation -ne $fixture -or $snapshot.constraintXml -cne $xml){throw 'A snapshot mixed different active documents or XML states.'}
    $state=[pscustomobject][ordered]@{
        Utc=[DateTime]::UtcNow.ToString('O');Path=$p.FullName
        PresentationTags=@(Read-Tags $p);ConstraintXml=$xml;Snapshot=$snapshot
        TextA=[string]$p.Slides.Item(1).Shapes.Item('Text_A').TextFrame.TextRange.Text
        TextB=[string]$p.Slides.Item(1).Shapes.Item('Text_B').TextFrame.TextRange.Text
        Slides=@(foreach($slide in $p.Slides){
            [pscustomobject][ordered]@{Id=$slide.SlideID;Name=$slide.Name;Tags=@(Read-Tags $slide);Shapes=@(foreach($shape in $slide.Shapes){
                $text=$null
                if([int]$shape.HasTextFrame -ne 0){$text=[pscustomobject]@{Text=[string]$shape.TextFrame.TextRange.Text;AutoSize=$shape.TextFrame.AutoSize;AutoSize2=$shape.TextFrame2.AutoSize}}
                [pscustomobject][ordered]@{NativeId=$shape.Id;Name=$shape.Name;Type=$shape.Type;Guid=[string]$shape.Tags.Item('PPTGCS_ID');Tags=@(Read-Tags $shape)
                    Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height;Rotation=$shape.Rotation;LockAspectRatio=$shape.LockAspectRatio;Text=$text}
            })}
        })
    }
    Assert-Owned
    return $state
}
function Json($value){ConvertTo-Json -InputObject $value -Depth 35 -Compress}
function Write-State([string]$name,$state){$state | ConvertTo-Json -Depth 35 | Set-Content -LiteralPath (Join-Path $directory ($name+'.json')) -Encoding UTF8}
function Geometry-Map($values) {
    $map=@{}
    foreach($shape in $values){if($map.ContainsKey([string]$shape.Id)){throw 'Duplicate geometry identity in the interaction evidence.'};$map[[string]$shape.Id]=$shape}
    return $map
}
function Snapshot-Geometry($state){@($state.Snapshot.slides | ForEach-Object {$_.geometry})}
function Assert-Geometry($actual,$expected,[string]$label) {
    $map=Geometry-Map $actual
    if($map.Count -ne @($expected).Count){throw ($label+': shape set differs')}
    foreach($shape in $expected){
        if(-not $map.ContainsKey([string]$shape.Id)){throw ($label+': missing shape')}
        foreach($key in @('Left','Top','Width','Height')){if([math]::Abs([double]$map[[string]$shape.Id].$key-[double]$shape.$key) -gt .01){throw ($label+': '+$key+' differs')}}
    }
}
function Assert-State($actual,$expected,[string]$label) {
    if($actual.Path -ne $expected.Path -or $actual.ConstraintXml -cne $expected.ConstraintXml -or (Json $actual.PresentationTags) -cne (Json $expected.PresentationTags) -or $actual.TextA -cne $expected.TextA -or $actual.TextB -cne $expected.TextB -or $actual.Slides.Count -ne $expected.Slides.Count){throw ($label+': document metadata or edited text changed')}
    Assert-Geometry (Snapshot-Geometry $actual) (Snapshot-Geometry $expected) $label
    for($i=0;$i -lt $actual.Slides.Count;$i++){
        $x=$actual.Slides[$i];$y=$expected.Slides[$i]
        if($x.Id -ne $y.Id -or $x.Name -cne $y.Name -or (Json $x.Tags) -cne (Json $y.Tags) -or $x.Shapes.Count -ne $y.Shapes.Count){throw ($label+': slide or shape set changed')}
        for($j=0;$j -lt $x.Shapes.Count;$j++){
            $a=$x.Shapes[$j];$b=$y.Shapes[$j]
            foreach($key in @('NativeId','Name','Type','Guid','Rotation','LockAspectRatio')){if($a.$key -cne $b.$key){throw ($label+': shape identity/type/format changed')}}
            if((Json $a.Tags) -cne (Json $b.Tags) -or (Json $a.Text) -cne (Json $b.Text)){throw ($label+': shape tags or text settings changed')}
            foreach($key in @('Left','Top','Width','Height')){if([math]::Abs([double]$a.$key-[double]$b.$key) -gt .01){throw ($label+': native geometry differs')}}
        }
    }
}
function With-Geometry($state,$geometry) {
    $copy=Json $state | ConvertFrom-Json
    $map=Geometry-Map $geometry
    foreach($shape in (Snapshot-Geometry $copy)){
        if(-not $map.ContainsKey([string]$shape.Id)){throw 'Recorded gesture did not contain the full constrained geometry.'}
        foreach($key in @('Left','Top','Width','Height')){$shape.$key=$map[[string]$shape.Id].$key}
    }
    foreach($slide in $copy.Slides){foreach($shape in $slide.Shapes){
        if($map.ContainsKey([string]$shape.Guid)){foreach($key in @('Left','Top','Width','Height')){$shape.$key=$map[[string]$shape.Guid].$key}}
    }}
    return $copy
}
function Assert-Relations($geometry,$relations,[string]$label) {
    $map=Geometry-Map $geometry
    foreach($relation in $relations){
        if($relation.PSObject.Properties.Name -contains 'Enabled' -and -not $relation.Enabled){continue}
        $members=@(foreach($id in $relation.ShapeIds){if(-not $map.ContainsKey([string]$id)){throw ($label+': missing relationship member')};$map[[string]$id]})
        $key=switch([string]$relation.Kind){'SameY'{'Top'} 'SameWidth'{'Width'} default{throw ($label+': unexpected relationship kind')}}
        foreach($shape in $members){if([math]::Abs([double]$shape.$key-[double]$members[0].$key) -gt .01){throw ($label+': '+$relation.Kind+' independent residual failed')}}
    }
}
function Execute-NativeHistory([string]$command) {
    Assert-Owned
    if(-not $ppt.CommandBars.GetEnabledMso($command)){throw ($command+' is not available for the next verified gesture.')}
    $ppt.CommandBars.ExecuteMso($command)
    if($command -eq 'Undo'){$script:undo++}elseif($command -eq 'Redo'){$script:redo++}
    Start-Sleep -Milliseconds 150
    Assert-Owned
}

try {
    Assert-Owned
    $after=Read-State
    Write-State 'manual-after' $after
    $stop=[DateTimeOffset]::Parse($after.Utc)
    if($baseline.Snapshot.presentation -ne $fixture -or $baseline.Snapshot.constraintXml -cne $after.ConstraintXml){throw 'The fixture relationship data changed during the manual interaction.'}
    if($after.TextA -ceq $baseline.TextA -or -not $after.TextA.Contains('测试')){throw 'Text A does not contain the requested new text.'}
    if($after.TextB -cne $baseline.TextB){throw 'Text B changed during the manual text edit.'}
    $relations=@($after.Snapshot.slides | ForEach-Object {$_.constraints})
    if($relations.Count -ne 2 -or @($relations | Where-Object Kind -eq 'SameY').Count -ne 1 -or @($relations | Where-Object Kind -eq 'SameWidth').Count -ne 1){throw 'Expected exactly the SameY and SameWidth fixture relationships.'}
    $allowed=Geometry-Map (Snapshot-Geometry $baseline)
    if($allowed.Count -ne 2){throw 'Expected the two identified text boxes only.'}
    Assert-Relations (Snapshot-Geometry $after) $relations 'final manual state'

    # Retain only this manual interval. Reject foreign-document activity before
    # exporting any log, so unrelated session/user content cannot enter evidence.
    $lines=[Collections.Generic.List[string]]::new();$events=[Collections.Generic.List[object]]::new()
    foreach($line in [IO.File]::ReadLines((Join-Path $env:LOCALAPPDATA 'PptGcs\layout-v2.log'))){
        $parts=$line -split ' ',3
        if($parts.Count -ne 3){continue}
        $time=[DateTimeOffset]::Parse($parts[0])
        if($time -lt $start -or $time -gt $stop){continue}
        $data=$parts[2] | ConvertFrom-Json
        if($data.PSObject.Properties.Name -contains 'presentation' -and $data.presentation -and $data.presentation -ne $fixture){throw 'Another document was active during the manual evidence interval; no history was changed.'}
        if($data.PSObject.Properties.Name -contains 'driverId' -and $data.driverId -and -not $allowed.ContainsKey([string]$data.driverId)){throw 'The interval includes a gesture from another document; no history was changed.'}
        if($parts[1] -like '*error'){throw ('The interaction interval contains '+$parts[1]+'; no history was changed.')}
        if($parts[1] -eq 'constraint-command'){throw 'An unexplained relationship command occurred during the text-only fixture interaction.'}
        $lines.Add($line);$events.Add([pscustomobject]@{Time=$time;Event=$parts[1];Data=$data})
    }
    $nativeTextRoutes=@($events | Where-Object {$_.Event -eq 'pointer-route' -and $_.Data.route -eq 'native-text-interior'}).Count
    if($nativeTextRoutes -lt 1){throw 'No native text-interior pointer route was recorded; no history was changed.'}
    $active=@{}
    foreach($event in $events){
        $data=$event.Data
        if($event.Event -eq 'gesture-start'){
            if($active.ContainsKey([string]$data.gestureId)){throw 'Duplicate physical gesture start.'}
            $active[[string]$data.gestureId]=[pscustomobject]@{Start=$data;Frames=[Collections.Generic.List[object]]::new();Ended=$false}
        }elseif($event.Event -eq 'gesture-frame'){
            if(-not $active.ContainsKey([string]$data.gestureId)){throw 'A frame has no gesture start inside the manual interval.'}
            $active[[string]$data.gestureId].Frames.Add($data)
        }elseif($event.Event -eq 'gesture-end'){
            if(-not $active.ContainsKey([string]$data.gestureId)){throw 'A gesture has no start inside the manual interval.'}
            $item=$active[[string]$data.gestureId]
            if($item.Ended){throw 'Duplicate physical gesture end.'};$item.Ended=$true
            if($data.frames -le 0){continue}
            if($data.test -ne $false -or $data.cancelled -ne $false -or $data.mode -ne 'Move' -or $item.Start.test -ne $false -or $item.Start.mode -ne 'Move' -or $item.Start.driverId -ne $data.driverId -or $data.reason -notin @('mouse-up','button-released')){throw 'The interval contains an unsupported, replayed or cancelled geometry edit; no history was changed.'}
            Assert-Geometry $data.before $item.Start.before 'gesture start/end baseline'
            if($item.Frames.Count -ne $data.frames){throw 'Gesture frame count does not match its full interval.'}
            foreach($frame in $item.Frames){
                if($frame.input -ne 'owned-gesture-loop' -or $frame.mode -ne 'Move' -or $frame.driverId -ne $data.driverId){throw 'Unexpected gesture frame writer.'}
                Assert-Relations $frame.geometry $relations 'held-pointer frame'
                if($frame.mouseDown -eq $true){$heldFrames++}
            }
            Assert-Relations $data.before $relations 'gesture before';Assert-Relations $data.after $relations 'gesture after'
            $gestures += [pscustomobject]@{Start=$item.Start;End=$data;Frames=$item.Frames;Time=$event.Time}
        }
    }
    if(@($active.Values | Where-Object {-not $_.Ended}).Count){throw 'A physical gesture is incomplete; no history was changed.'}
    if($gestures.Count -lt 2 -or @($gestures.End.driverId | Select-Object -Unique).Count -ne 2 -or $heldFrames -lt 2){throw 'Both text boxes need completed physical moves with held-pointer updates; no history was changed.'}
    for($i=1;$i -lt $gestures.Count;$i++){Assert-Geometry $gestures[$i].End.before $gestures[$i-1].End.after 'continuous physical geometry history'}
    Assert-Geometry (Snapshot-Geometry $after) $gestures[$gestures.Count-1].End.after 'recorded final physical geometry'
    [IO.File]::WriteAllLines((Join-Path $directory 'events.log'),$lines,[Text.UTF8Encoding]::new($false))
    $checks.Add('native text-interior routing and new Text A content; Text B unchanged')
    $checks.Add('both drivers perform complete physical Move gestures with valid held-pointer geometry')
    $checks.Add('continuous gesture geometry chain, unchanged XML and final state checked before native history')

    # Each next Undo is predicted from the current complete gesture. Never seek
    # blindly through text/IME history: stop immediately on a divergent step.
    $states=[Collections.Generic.List[object]]::new();$states.Add($after)
    for($i=$gestures.Count-1;$i -ge 0;$i--){
        Assert-State (Read-State) (With-Geometry $after $gestures[$i].End.after) 'before verified native Undo'
        Execute-NativeHistory 'Undo'
        $state=Read-State;Write-State ('manual-undo-{0:00}' -f $undo) $state
        Assert-State $state (With-Geometry $after $gestures[$i].End.before) ('native Undo '+$undo)
        $states.Add($state)
    }
    Assert-Geometry (Snapshot-Geometry $states[$states.Count-1]) $gestures[0].End.before 'first gesture baseline after all move Undos'
    $checks.Add('one native Undo per complete Move, preserving the edited text and all relationship/identity tags')
    for($i=1;$i -le $undo;$i++){
        Execute-NativeHistory 'Redo'
        $state=Read-State;Write-State ('manual-redo-{0:00}' -f $redo) $state
        Assert-State $state $states[$undo-$i] ('native Redo '+$redo)
    }
    Assert-State (Read-State) $after 'complete manual state restored'
    $p.Save();Assert-State (Read-State) $after 'saved manual state'
    $checks.Add('all native Redo steps restore exact text/tags/XML and geometry within 0.01 pt; fixture saved')
    $passed=$true
}catch{$failure=$_.Exception.ToString();throw}
finally{
    if(-not $passed -and $undo -gt $redo){
        try{
            Assert-Owned
            while($redo -lt $undo){Execute-NativeHistory 'Redo'}
            if($null -ne $after){Assert-State (Read-State) $after 'failure restoration'}
        }catch{$restoreFailure=$_.Exception.ToString()}
    }
    $result=[pscustomobject]@{Passed=$passed;PhysicalMouseInput=$true;NativeTextInteriorRoutes=$nativeTextRoutes
        PhysicalGestures=$gestures.Count;DriverIds=@($gestures.End.driverId | Select-Object -Unique);HeldFrames=$heldFrames
        NativeUndoSteps=$undo;NativeRedoSteps=$redo;Checks=@($checks);Failure=$failure;RestorationFailure=$restoreFailure
        StartUtc=$baseline.Utc;StopUtc=$(if($null -ne $after){$after.Utc}else{$null})}
    $result | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $directory 'acceptance.json') -Encoding UTF8
    $result | ConvertTo-Json -Depth 15
}
