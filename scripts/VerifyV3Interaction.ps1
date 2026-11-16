$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root 'artifacts\v3-interaction'
$capture=Join-Path $PSScriptRoot 'CaptureV3Interaction.ps1'
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
if($ppt.ActivePresentation.FullName -ne (Join-Path $directory 'interaction.pptx')){throw 'Unexpected interactive document'}
& $capture -Name 'after'
function Load-State([string]$name){Get-Content -LiteralPath (Join-Path $directory ($name+'.json')) -Raw | ConvertFrom-Json}
function Json($value){ConvertTo-Json -InputObject $value -Depth 25 -Compress}
function Same-State($a,$b) {
    if($a.ConstraintXml -cne $b.ConstraintXml -or (Json $a.PresentationTags) -cne (Json $b.PresentationTags) -or $a.Slides.Count -ne $b.Slides.Count){return $false}
    for($i=0;$i -lt $a.Slides.Count;$i++) {
        $s=$a.Slides[$i];$t=$b.Slides[$i]
        if($s.Id -ne $t.Id -or $s.Name -cne $t.Name -or $s.Shapes.Count -ne $t.Shapes.Count){return $false}
        for($j=0;$j -lt $s.Shapes.Count;$j++) {
            $x=$s.Shapes[$j];$y=$t.Shapes[$j]
            if($x.Id -ne $y.Id -or $x.Name -cne $y.Name -or (Json $x.Tags) -cne (Json $y.Tags)){return $false}
            foreach($field in @('Left','Top','Width','Height')){if([Math]::Abs($x.$field-$y.$field) -gt 0.01){return $false}}
        }
    }
    return $true
}
function Assert-Relations($state) {
    foreach($slide in $state.Snapshot.slides) {
        $map=@{};foreach($shape in $slide.geometry){$map[$shape.Id]=$shape}
        foreach($relation in $slide.constraints) {
            if(-not $relation.Enabled){continue}
            $members=@(foreach($id in $relation.ShapeIds){if(-not $map.ContainsKey($id)){throw 'Missing relationship member'};$map[$id]})
            switch($relation.Kind) {
                'SameWidth' {foreach($s in $members){if([Math]::Abs($s.Width-$members[0].Width) -gt 0.01){throw 'SameWidth residual'}}}
                'AspectRatio' {if([Math]::Abs($members[0].Width-$relation.Value*$members[0].Height) -gt 0.01){throw 'AspectRatio residual'}}
                'EqualHorizontalGap' {$gap=$members[1].Left-$members[0].Left-$members[0].Width;for($j=1;$j -lt $members.Count-1;$j++){if([Math]::Abs($members[$j+1].Left-$members[$j].Left-$members[$j].Width-$gap) -gt 0.01){throw 'EqualHorizontalGap residual'}}}
                default {throw ('Unexpected manual relation '+$relation.Kind)}
            }
        }
    }
}
$before=Load-State 'before';$after=Load-State 'after'
$start=[DateTime]::Parse($before.Utc).ToUniversalTime();$stop=[DateTime]::Parse($after.Utc).ToUniversalTime()
$events=@(foreach($line in Get-Content -LiteralPath (Join-Path $directory 'events.log')) {
    $parts=$line -split ' ',3
    if($parts.Count -ne 3){continue}
    $time=[DateTime]::Parse($parts[0]).ToUniversalTime()
    if($time -ge $start -and $time -le $stop){[pscustomobject]@{Time=$time;Event=$parts[1];Data=($parts[2] | ConvertFrom-Json)}}
})
$errors=@($events | Where-Object {$_.Event -like '*error'})
if($errors.Count){throw ('Runtime errors occurred: '+(Json $errors))}
$radial=@($events | Where-Object {$_.Event -eq 'radial-command' -and $_.Data.outcome -eq 'applied'})
foreach($kind in @('SameWidth','EqualHorizontalGap')) {
    $applied=@($radial | Where-Object {$_.Data.choice -eq $kind})
    if($applied.Count -lt 1){throw ('Missing physical radial command '+$kind)}
    foreach($command in $applied) {
        if(@($events | Where-Object {$_.Event -eq 'radial-open' -and $_.Data.gestureId -eq $command.Data.gestureId -and $_.Data.middleDown -eq $true}).Count -ne 1){throw 'Radial command has no held-middle origin'}
    }
}
$gestures=@();$active=@{}
foreach($event in $events) {
    $data=$event.Data
    if($event.Event -eq 'gesture-start'){$active[$data.gestureId]=[pscustomobject]@{Start=$data;Frames=[Collections.Generic.List[object]]::new()}}
    elseif($event.Event -eq 'gesture-frame'){$active[$data.gestureId].Frames.Add($data)}
    elseif($event.Event -eq 'gesture-end' -and $data.frames -gt 0 -and -not $data.cancelled){$gestures += [pscustomobject]@{Start=$active[$data.gestureId].Start;Frames=$active[$data.gestureId].Frames;End=$data}}
}
$coverage=@{SameWidth=@();AspectRatio=@();EqualHorizontalGap=@()}
foreach($gesture in $gestures) {
    if(@($gesture.Frames | Where-Object {$_.mouseDown -eq $true}).Count -lt 2){continue}
    foreach($relation in $gesture.Start.constraints) {
        if($coverage.ContainsKey([string]$relation.kind)){$coverage[[string]$relation.kind]+=$gesture.Start.driverId}
    }
}
if(@($coverage.SameWidth | Select-Object -Unique).Count -lt 2 -or $coverage.AspectRatio.Count -lt 1 -or $coverage.EqualHorizontalGap.Count -lt 1){throw 'Physical drag coverage is incomplete'}
Assert-Relations $after
$states=[Collections.Generic.List[object]]::new();$states.Add($after)
$undo=0;$redo=0;$passed=$false
try {
    while(-not (Same-State $states[$states.Count-1] $before)) {
        if($undo -ge 20 -or -not $ppt.CommandBars.GetEnabledMso('Undo')){throw 'Native history did not reach the initial state within retained history'}
        if($ppt.ActivePresentation.FullName -ne $after.Path){throw 'Active document changed during history verification'}
        $ppt.CommandBars.ExecuteMso('Undo');$undo++
        $name='undo-{0:00}' -f $undo;& $capture -Name $name
        $state=Load-State $name;Assert-Relations $state;$states.Add($state)
    }
    $ratios=@($states | ForEach-Object {$_.Snapshot.slides} | ForEach-Object {$_.constraints} | Where-Object Kind -eq 'AspectRatio')
    if(@($ratios | Where-Object {$_.Value -eq 2 -and $_.Enabled}).Count -eq 0 -or @($ratios | Where-Object {$_.Value -eq 1.5 -and $_.Enabled}).Count -eq 0 -or @($ratios | Where-Object {-not $_.Enabled}).Count -eq 0){throw 'Create, numeric edit and disable/enable were not all recorded in native history'}
    if(@($after.Snapshot.slides | ForEach-Object {$_.constraints} | Where-Object Kind -eq 'AspectRatio').Count -ne 0){throw 'The parameter relation was not individually deleted'}
    foreach($gesture in $gestures) {
        foreach($frame in $gesture.Frames) {
            if($frame.input -ne 'owned-gesture-loop'){throw 'Unexpected frame writer'}
            $byId=@{};foreach($shape in $frame.geometry){$byId[$shape.Id]=$shape}
            foreach($relation in $gesture.Start.constraints) {
                $m=@(foreach($id in $relation.ShapeIds){$byId[$id]})
                if($relation.kind -eq 'SameWidth'){foreach($s in $m){if([Math]::Abs($s.Width-$m[0].Width) -gt 0.01){throw 'Held width frame residual'}}}
                elseif($relation.kind -eq 'AspectRatio'){$values=@($ratios | Where-Object Id -eq $relation.Id | ForEach-Object Value | Select-Object -Unique);if(@($values | Where-Object {[Math]::Abs($m[0].Width-$_*$m[0].Height) -le 0.01}).Count -eq 0){throw 'Held ratio frame does not match any recorded parameter value'}}
                elseif($relation.kind -eq 'EqualHorizontalGap'){$gap=$m[1].Left-$m[0].Left-$m[0].Width;for($j=1;$j -lt $m.Count-1;$j++){if([Math]::Abs($m[$j+1].Left-$m[$j].Left-$m[$j].Width-$gap) -gt 0.01){throw 'Held gap frame residual'}}}
            }
        }
    }
    for($step=1;$step -le $undo;$step++) {
        if(-not $ppt.CommandBars.GetEnabledMso('Redo')){throw 'Redo unavailable'}
        $ppt.CommandBars.ExecuteMso('Redo');$redo++
        $name='redo-{0:00}' -f $step;& $capture -Name $name
        $state=Load-State $name
        if(-not (Same-State $state $states[$undo-$step])){throw ('Native redo diverged at '+$step)}
        Assert-Relations $state
    }
    $ppt.ActivePresentation.Save();& $capture -Name 'saved'
    $passed=$true
}
finally {
    # Restore the recorded final user test state even if an assertion fails.
    if(-not $passed -and $ppt.ActivePresentation.FullName -eq $after.Path){for($i=$redo;$i -lt $undo;$i++){if(-not $ppt.CommandBars.GetEnabledMso('Redo')){break};$ppt.CommandBars.ExecuteMso('Redo')}}
    $result=[pscustomobject]@{Passed=$passed;PhysicalMouseInput=$true;RadialCommands=@($radial | ForEach-Object {$_.Data.choice});PhysicalGestures=$gestures.Count;HeldFrames=@($gestures | ForEach-Object {$_.Frames} | Where-Object {$_.mouseDown -eq $true}).Count;NativeUndoSteps=$undo;NativeRedoSteps=$redo;Coverage=$coverage;UserConfirmedPaneVisible=$true;Checks=@('real middle-button radial creation','two-way width and ratio/gap held-pointer geometry','pane create/update/disable/enable/delete parameters','full native metadata and geometry history','final state restored')}
    $result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'acceptance.json') -Encoding UTF8
    $result | ConvertTo-Json -Depth 10
}
