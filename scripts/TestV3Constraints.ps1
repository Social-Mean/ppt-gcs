[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$evidenceRoot=Join-Path $root 'artifacts\v3-constraint-validation'
$directory=Join-Path $evidenceRoot ('run-'+[DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfff'))
New-Item -ItemType Directory -Force -Path $directory | Out-Null
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$returnWindow=$ppt.ActiveWindow
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
$fixture=Join-Path $directory 'twenty-constraints.pptx'
$p=$null;$passed=$false;$failure=$null;$results=@()
$cases=@(
    @{Kind='SameX';Count=2;Parameters=@{};Mode='Move';Dx=24;Dy=0},
    @{Kind='SameY';Count=2;Parameters=@{};Mode='Move';Dx=0;Dy=24},
    @{Kind='SameWidth';Count=2;Parameters=@{};Mode='ResizeE';Dx=24;Dy=0},
    @{Kind='SameHeight';Count=2;Parameters=@{};Mode='ResizeS';Dx=0;Dy=24},
    @{Kind='EqualHorizontalGap';Count=3;Parameters=@{};Mode='Move';Dx=24;Dy=0},
    @{Kind='EqualVerticalGap';Count=3;Parameters=@{};Mode='Move';Dx=0;Dy=24},
    @{Kind='RightAligned';Count=2;Parameters=@{};Mode='Move';Dx=24;Dy=0},
    @{Kind='BottomAligned';Count=2;Parameters=@{};Mode='Move';Dx=0;Dy=24},
    @{Kind='CenterXAligned';Count=2;Parameters=@{};Mode='Move';Dx=24;Dy=0},
    @{Kind='CenterYAligned';Count=2;Parameters=@{};Mode='Move';Dx=0;Dy=24},
    @{Kind='CentersCoincident';Count=2;Parameters=@{};Mode='Move';Dx=24;Dy=18},
    @{Kind='FixedWidth';Count=1;Parameters=@{value=180};Mode='ResizeS';Dx=0;Dy=24},
    @{Kind='FixedHeight';Count=1;Parameters=@{value=130};Mode='ResizeE';Dx=24;Dy=0},
    @{Kind='FixedPosition';Count=1;Parameters=@{value=170;second=190};Mode='ResizeE';Dx=24;Dy=0},
    @{Kind='AspectRatio';Count=1;Parameters=@{value=1.6};Mode='ResizeE';Dx=24;Dy=0},
    @{Kind='FixedHorizontalGap';Count=3;Parameters=@{value=35};Mode='Move';Dx=24;Dy=0},
    @{Kind='FixedVerticalGap';Count=3;Parameters=@{value=30};Mode='Move';Dx=0;Dy=24},
    @{Kind='Offset';Count=2;Parameters=@{value=260;second=70};Mode='Move';Dx=24;Dy=18},
    @{Kind='SymmetricVertical';Count=2;Parameters=@{value=480};Mode='Move';Dx=24;Dy=18},
    @{Kind='SymmetricHorizontal';Count=2;Parameters=@{value=270};Mode='Move';Dx=24;Dy=18}
)
function Assert-TestDocument {
    if($null -eq $p -or $ppt.ActivePresentation.FullName -ne $fixture -or $p.Tags.Item('PPTGCS_TEST') -ne 'V3'){throw 'Active document changed; V3 constraint test stopped.'}
}
function Tags($object) {
    @(for($i=1;$i -le $object.Tags.Count;$i++){[pscustomobject]@{Name=$object.Tags.Name($i);Value=$object.Tags.Value($i)}}) | Sort-Object Name
}
function Read-State {
    Assert-TestDocument
    [pscustomobject][ordered]@{
        PresentationTags=@(Tags $p)
        ConstraintXml=$addin.GetConstraintXml()
        Slides=@(foreach($slide in $p.Slides){[pscustomobject][ordered]@{Id=$slide.SlideID;Shapes=@(foreach($shape in $slide.Shapes){
            [pscustomobject][ordered]@{NativeId=$shape.Id;Name=$shape.Name;Tags=@(Tags $shape);Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height;Rotation=$shape.Rotation}
        })}})
        Snapshot=($addin.GetSnapshot() | ConvertFrom-Json)
    }
}
function Snapshot([string]$name) {
    $state=Read-State
    [pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('O');ProcessId=@(Get-Process POWERPNT | ForEach-Object Id);State=$state} | ConvertTo-Json -Depth 35 | Set-Content -LiteralPath (Join-Path $directory ($name+'.json')) -Encoding UTF8
    return $state
}
function Same-State($actual,$expected){($actual | ConvertTo-Json -Depth 35 -Compress) -ceq ($expected | ConvertTo-Json -Depth 35 -Compress)}
function Native-History([string]$command,[string]$name,$expected) {
    Assert-TestDocument
    if(-not $ppt.CommandBars.GetEnabledMso($command)){throw "$command is unavailable before $name"}
    $ppt.CommandBars.ExecuteMso($command)
    Start-Sleep -Milliseconds 100
    $actual=Snapshot $name
    if(-not (Same-State $actual $expected)){throw "$name did not restore the complete native shape/tag/XML state"}
}
function Residual($relation,$byId) {
    $members=@(foreach($id in $relation.ShapeIds){if(-not $byId.ContainsKey($id)){throw 'Relation references a missing shape'};$byId[$id]})
    $a=$members[0]
    $r=[Collections.Generic.List[double]]::new()
    switch($relation.Kind) {
        'FixedWidth' {$r.Add($a.Width-$relation.Value)}
        'FixedHeight' {$r.Add($a.Height-$relation.Value)}
        'FixedPosition' {$r.Add($a.Left-$relation.Value);$r.Add($a.Top-$relation.SecondaryValue)}
        'AspectRatio' {$r.Add($a.Width-$relation.Value*$a.Height)}
        'Offset' {$r.Add($members[1].Left-$a.Left-$relation.Value);$r.Add($members[1].Top-$a.Top-$relation.SecondaryValue)}
        {$_ -in @('EqualHorizontalGap','EqualVerticalGap','FixedHorizontalGap','FixedVerticalGap')} {
            $horizontal=$relation.Kind -in @('EqualHorizontalGap','FixedHorizontalGap')
            $gaps=@(for($i=1;$i -lt $members.Count;$i++){
                if($horizontal){$members[$i].Left-$members[$i-1].Left-$members[$i-1].Width}else{$members[$i].Top-$members[$i-1].Top-$members[$i-1].Height}
            })
            $target=if($null -ne $relation.Value){$relation.Value}else{$gaps[0]}
            foreach($gap in $gaps){$r.Add($gap-$target)}
        }
        {$_ -in @('SymmetricVertical','SymmetricHorizontal')} {
            $b=$members[1];$r.Add($a.Width-$b.Width);$r.Add($a.Height-$b.Height)
            if($relation.Kind -eq 'SymmetricVertical'){$r.Add($a.Top-$b.Top);$r.Add($a.Left+$a.Width/2+$b.Left+$b.Width/2-2*$relation.Value)}
            else{$r.Add($a.Left-$b.Left);$r.Add($a.Top+$a.Height/2+$b.Top+$b.Height/2-2*$relation.Value)}
        }
        default {
            foreach($b in @($members | Select-Object -Skip 1)){
                switch($relation.Kind){
                    'SameX' {$r.Add($b.Left-$a.Left)}
                    'SameY' {$r.Add($b.Top-$a.Top)}
                    'SameWidth' {$r.Add($b.Width-$a.Width)}
                    'SameHeight' {$r.Add($b.Height-$a.Height)}
                    'RightAligned' {$r.Add($b.Left+$b.Width-$a.Left-$a.Width)}
                    'BottomAligned' {$r.Add($b.Top+$b.Height-$a.Top-$a.Height)}
                    'CenterXAligned' {$r.Add($b.Left+$b.Width/2-$a.Left-$a.Width/2)}
                    'CenterYAligned' {$r.Add($b.Top+$b.Height/2-$a.Top-$a.Height/2)}
                    'CentersCoincident' {$r.Add($b.Left+$b.Width/2-$a.Left-$a.Width/2);$r.Add($b.Top+$b.Height/2-$a.Top-$a.Height/2)}
                    default {throw ('Unknown constraint formula: '+$relation.Kind)}
                }
            }
        }
    }
    if($r.Count -eq 0){throw ('No formula evaluated for '+$relation.Kind)}
    return ($r | ForEach-Object {[Math]::Abs($_)} | Measure-Object -Maximum).Maximum
}
function Assert-Constraints($state) {
    $maximum=0.0
    foreach($slide in $state.Snapshot.slides){
        $byId=@{};foreach($shape in $slide.geometry){$byId[$shape.Id]=$shape;if($shape.Width -lt 0.99 -or $shape.Height -lt 0.99){throw 'Invalid minimum shape size'}}
        foreach($relation in $slide.constraints){
            if(-not $relation.Enabled){continue}
            $errorValue=Residual $relation $byId
            if($errorValue -gt 0.01){throw ($relation.Kind+' residual is '+$errorValue+' pt')}
            $maximum=[Math]::Max($maximum,$errorValue)
        }
    }
    return $maximum
}
try {
    $p=$ppt.Presentations.Add();$p.Tags.Add('PPTGCS_TEST','V3')
    $p.PageSetup.SlideWidth=960;$p.PageSetup.SlideHeight=540
    for($i=0;$i -lt $cases.Count;$i++) {
        $case=$cases[$i];$slide=$p.Slides.Add($i+1,12);$slide.Name=$case.Kind;$slide.Tags.Add('PPTGCS_CASE',$case.Kind)
        $positions=@(@(100,130,110,70),@(420,280,150,100),@(700,200,95,80))
        for($j=0;$j -lt $case.Count;$j++) {
            $g=$positions[$j];$shape=$slide.Shapes.AddShape(1,$g[0],$g[1],$g[2],$g[3]);$shape.Name='Case_'+$case.Kind+'_'+([char](65+$j));$shape.LockAspectRatio=0;$shape.TextFrame.AutoSize=0
        }
    }
    $p.SaveAs([string]$fixture,24);Start-Sleep -Milliseconds 250;$p.Close()
    $p=$ppt.Presentations.Open([string]$fixture,0,0,-1)
    $null=Snapshot 'before'
    for($i=0;$i -lt $cases.Count;$i++) {
        $case=$cases[$i];$prefix='{0:00}-{1}' -f ($i+1),$case.Kind
        $ppt.ActiveWindow.View.GotoSlide($i+1);$slide=$p.Slides.Item($i+1)
        $indices=@(for($j=1;$j -le $case.Count;$j++){$j});$slide.Shapes.Range($indices).Select()
        $id=$addin.AddConstraintCommand([string]$case.Kind,($case.Parameters | ConvertTo-Json -Compress))
        if(-not $id){throw ('Constraint was not created: '+$case.Kind)}
        $state=Snapshot ($prefix+'-added')
        $residual=Assert-Constraints $state
        $relation=@($state.Snapshot.slides[$i].constraints | Where-Object Id -eq $id)[0]
        if($null -eq $relation -or $relation.Kind -ne $case.Kind){throw 'Added constraint kind does not match requested case'}
        $drivers=@(if($case.Count -eq 1){$relation.ShapeIds[0]}else{$relation.ShapeIds[0];$relation.ShapeIds[-1]})
        $driverResults=@()
        for($driverIndex=0;$driverIndex -lt $drivers.Count;$driverIndex++) {
            $driver=[string]$drivers[$driverIndex];$sign=if($driverIndex -eq 0){1}else{-1}
            $dx=[int]($case.Dx*$sign);$dy=[int]($case.Dy*$sign)
            $offsets='[[{0},{1}],[{2},{3}],[{4},{5}]]' -f ([int]($dx/3)),([int]($dy/3)),([int]($dx*2/3)),([int]($dy*2/3)),$dx,$dy
            $before=$state
            $null=$addin.ReplayGestureForTest($driver,[string]$case.Mode,$offsets)
            $name=$prefix+'-driver-'+($driverIndex+1)
            $state=Snapshot ($name+'-after')
            if(Same-State $state $before){throw ($name+' produced no actual geometry change')}
            if($state.ConstraintXml -cne $before.ConstraintXml){throw 'Geometry edit changed constraint data'}
            $residual=[Math]::Max($residual,(Assert-Constraints $state))
            Native-History 'Undo' ($name+'-undo') $before
            Native-History 'Redo' ($name+'-redo') $state
            $driverResults += [pscustomobject]@{DriverId=$driver;Mode=$case.Mode;NativeUndo=$true;NativeRedo=$true}
        }
        $pane=$addin.GetPaneState() | ConvertFrom-Json
        if(@($pane.Relations | Where-Object {$_.Id -eq $id -and $_.Status -eq '已满足'}).Count -ne 1){throw ('Pane does not report satisfied '+$case.Kind)}
        $results += [pscustomobject]@{Kind=$case.Kind;ConstraintId=$id;Members=$relation.ShapeIds;Parameters=$case.Parameters;Drivers=$driverResults;MaximumResidualPt=$residual;Passed=$true}
        Write-Output ('PASS '+$case.Kind+': '+$drivers.Count+' driven member(s), independent geometry formula and native history')
    }
    $p.Save();$null=Snapshot 'saved';$passed=$true
} catch {$failure=$_.Exception.ToString();$failure | Set-Content -LiteralPath (Join-Path $directory 'failure.txt') -Encoding UTF8;throw}
finally {
    if($null -ne $p){try{if($p.FullName -eq $fixture){$p.Save();$p.Close()}}catch{}}
    if($null -ne $returnWindow){try{$returnWindow.Activate()}catch{}}
    $report=[pscustomobject]@{Passed=$passed;PhysicalMouseInput=$false;Cases=$results;ExpectedCases=$cases.Count;CompletedCases=$results.Count;Failure=$failure;Directory=$directory;Utc=[DateTime]::UtcNow.ToString('O')}
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $directory 'acceptance.json') -Encoding UTF8
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $evidenceRoot 'latest.json') -Encoding UTF8
    $report | ConvertTo-Json -Depth 12
}
