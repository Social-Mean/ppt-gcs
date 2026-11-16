[CmdletBinding()]
param([string]$VerifyReopenedDirectory)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$evidenceRoot=Join-Path $root 'artifacts\shape-support'
if($VerifyReopenedDirectory){
    $directory=[IO.Path]::GetFullPath($VerifyReopenedDirectory)
    if(-not (Test-Path -LiteralPath (Join-Path $directory 'saved.json')) -or -not (Test-Path -LiteralPath (Join-Path $directory 'shapes.pptx'))){throw 'The requested run must contain saved.json and shapes.pptx.'}
}else{
    $directory=Join-Path $evidenceRoot ('run-'+[DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfff'))
    New-Item -ItemType Directory -Force -Path $directory | Out-Null
}
$fixture=Join-Path $directory 'shapes.pptx'
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$returnWindow=$null
if($ppt.Windows.Count -gt 0){$returnWindow=$ppt.ActiveWindow}
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
$p=$null;$passed=$false;$failure=$null;$results=@();$checks=@()
$cases=@(
    @{Kind='Ellipse';AutoShape=9;Type=1},
    @{Kind='RoundedRectangle';AutoShape=5;Type=1},
    @{Kind='Triangle';AutoShape=7;Type=1},
    @{Kind='Diamond';AutoShape=4;Type=1},
    @{Kind='RightArrow';AutoShape=33;Type=1},
    @{Kind='TextBox';Type=17},
    @{Kind='Picture';Type=13},
    @{Kind='LinkedPicture';Type=11},
    @{Kind='Graphic';Type=28},
    @{Kind='Freeform';Type=5},
    @{Kind='Group';Type=6}
)
function Assert-Owned {
    if($null -eq $p -or $ppt.ActivePresentation.FullName -ne $fixture -or $p.Tags.Item('PPTGCS_TEST') -ne 'V3'){throw 'Shape test lost its owned fixture; no command was sent.'}
}
function Read-Tags($object) {
    @(for($i=1;$i -le $object.Tags.Count;$i++){[pscustomobject]@{Name=$object.Tags.Name($i);Value=$object.Tags.Value($i)}}) | Sort-Object Name
}
function Read-Shape($shape) {
    $text=$null
    if([int]$shape.HasTextFrame -ne 0){
        $text=[pscustomobject]@{Text=$shape.TextFrame.TextRange.Text;AutoSize=$shape.TextFrame.AutoSize;AutoSize2=$shape.TextFrame2.AutoSize}
    }
    [pscustomobject][ordered]@{
        NativeId=$shape.Id;Name=$shape.Name;Type=$shape.Type
        AutoShapeType=$(if([int]$shape.Type -eq 1){$shape.AutoShapeType}else{$null})
        Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height;Rotation=$shape.Rotation
        LockAspectRatio=$shape.LockAspectRatio;Tags=@(Read-Tags $shape);Text=$text
        Children=@(if([int]$shape.Type -eq 6){for($j=1;$j -le $shape.GroupItems.Count;$j++){Read-Shape $shape.GroupItems.Item($j)}})
    }
}
function Read-State {
    Assert-Owned
    $xml=$addin.GetConstraintXml()
    Assert-Owned
    $tags=@(Read-Tags $p)
    $slides=@(foreach($slide in $p.Slides){[pscustomobject][ordered]@{Id=$slide.SlideID;Name=$slide.Name;Tags=@(Read-Tags $slide);Shapes=@(foreach($shape in $slide.Shapes){Read-Shape $shape})}})
    Assert-Owned
    $snapshot=$addin.GetSnapshot() | ConvertFrom-Json
    if($snapshot.presentation -cne $fixture){throw 'The snapshot came from a different presentation; nothing was captured.'}
    Assert-Owned
    if($snapshot.constraintXml -cne $xml){throw 'Authoritative XML changed during capture; nothing was captured.'}
    [pscustomobject][ordered]@{
        PresentationTags=$tags;ConstraintXml=$xml;Slides=$slides;Snapshot=$snapshot
    }
}
function Save-State([string]$name) {
    $state=Read-State
    [pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('O');ProcessId=@(Get-Process POWERPNT | ForEach-Object Id);State=$state} | ConvertTo-Json -Depth 45 | Set-Content -LiteralPath (Join-Path $directory ($name+'.json')) -Encoding UTF8
    return $state
}
function Same-State($a,$b){($a | ConvertTo-Json -Depth 45 -Compress) -ceq ($b | ConvertTo-Json -Depth 45 -Compress)}
function Native-History([string]$command,[string]$name,$expected) {
    Assert-Owned
    if(-not $ppt.CommandBars.GetEnabledMso($command)){throw "$command unavailable before $name"}
    $ppt.CommandBars.ExecuteMso($command);Start-Sleep -Milliseconds 120
    if(-not (Same-State (Save-State $name) $expected)){throw "$name did not restore all shape properties, child membership, tags and XML"}
}
function Assert-Relations($state) {
    foreach($slide in $state.Snapshot.slides){
        $map=@{};foreach($s in $slide.geometry){$map[$s.Id]=$s}
        foreach($r in $slide.constraints){
            $members=@(foreach($id in $r.ShapeIds){if(-not $map.ContainsKey($id)){throw 'Missing relation member'};$map[$id]})
            if($r.Kind -notin @('SameY','SameWidth','SameHeight')){throw ('Unexpected test relation '+$r.Kind)}
            $key=if($r.Kind -eq 'SameY'){'Top'}elseif($r.Kind -eq 'SameWidth'){'Width'}else{'Height'}
            foreach($s in $members){if([math]::Abs($s.$key-$members[0].$key) -gt .01){throw ($r.Kind+' independent formula failed')}}
        }
    }
}
function Shape-Properties($shapes) {
    @(foreach($s in $shapes){[pscustomobject]@{Id=$s.NativeId;Name=$s.Name;Type=$s.Type;AutoShapeType=$s.AutoShapeType;Rotation=$s.Rotation;Aspect=$s.LockAspectRatio;Tags=$s.Tags;Text=$s.Text;Children=@(Shape-Properties $s.Children)}})
}
function Assert-NonGeometry($a,$b) {
    if($a.ConstraintXml -cne $b.ConstraintXml -or ($a.PresentationTags|ConvertTo-Json -Depth 10 -Compress) -cne ($b.PresentationTags|ConvertTo-Json -Depth 10 -Compress)){throw 'Gesture changed relation data or presentation tags'}
    for($i=0;$i -lt $a.Slides.Count;$i++){
        $x=Shape-Properties $a.Slides[$i].Shapes;$y=Shape-Properties $b.Slides[$i].Shapes
        if(($x|ConvertTo-Json -Depth 35 -Compress) -cne ($y|ConvertTo-Json -Depth 35 -Compress)){throw 'Geometry edit changed type, text, AutoSize, aspect lock, child structure or identity'}
    }
}
function Assert-Reopened($a,$b) {
    Assert-NonGeometry $a $b
    for($i=0;$i -lt $a.Slides.Count;$i++){
        if($a.Slides[$i].Id -ne $b.Slides[$i].Id -or $a.Slides[$i].Name -cne $b.Slides[$i].Name){throw 'Persisted slide changed'}
        Compare-Geometry $a.Slides[$i].Shapes $b.Slides[$i].Shapes
    }
}
function Compare-Geometry($a,$b) {
    if($a.Count -ne $b.Count){throw 'Persisted shape count changed'}
    for($j=0;$j -lt $a.Count;$j++){
        foreach($key in @('Left','Top','Width','Height')){if([math]::Abs($a[$j].$key-$b[$j].$key) -gt .01){throw 'Persisted shape geometry changed'}}
        Compare-Geometry $a[$j].Children $b[$j].Children
    }
}
function Add-TestShape($slide,$case,[string]$name) {
    if($case.ContainsKey('AutoShape')){$shape=$slide.Shapes.AddShape([int]$case.AutoShape,100,120,150,100)}
    else {switch($case.Kind){
        'TextBox' {$shape=$slide.Shapes.AddTextbox(1,100,120,150,100);$shape.TextFrame.TextRange.Text='Text box sizing test'}
        'Picture' {$shape=$slide.Shapes.AddPicture([string]$png,0,-1,100,120,150,100)}
        'LinkedPicture' {$shape=$slide.Shapes.AddPicture([string]$png,-1,0,100,120,150,100)}
        'Graphic' {$shape=$slide.Shapes.AddPicture([string]$svg,0,-1,100,120,150,100)}
        'Freeform' {
            $builder=$slide.Shapes.BuildFreeform(1,100,120)
            $builder.AddNodes(0,1,250,135);$builder.AddNodes(0,1,230,220);$builder.AddNodes(0,1,120,200);$builder.AddNodes(0,1,100,120)
            $shape=$builder.ConvertToShape()
        }
        'Group' {
            $first=$slide.Shapes.AddShape(9,100,120,80,80);$first.Name=$name+'_ellipse'
            $second=$slide.Shapes.AddTextbox(1,190,145,90,60);$second.Name=$name+'_text';$second.TextFrame.TextRange.Text='Grouped text'
            $shape=$slide.Shapes.Range(@($first.Name,$second.Name)).Group()
        }
        default {throw ('Unknown fixture shape '+$case.Kind)}
    }}
    $shape.Name=$name
    if([int]$shape.Type -ne [int]$case.Type){throw ($case.Kind+' produced unexpected Office type '+$shape.Type)}
    return $shape
}
if($VerifyReopenedDirectory){
    $verificationPassed=$false;$verificationFailure=$null;$priorProcessIds=@();$currentProcessIds=@()
    try{
        $savedCapture=Get-Content -LiteralPath (Join-Path $directory 'saved.json') -Raw | ConvertFrom-Json
        $acceptance=Get-Content -LiteralPath (Join-Path $directory 'acceptance.json') -Raw | ConvertFrom-Json
        if(-not $acceptance.Passed){throw 'New-process verification requires a previously passed shape run.'}
        $p=$ppt.ActivePresentation
        Assert-Owned
        $priorProcessIds=@($savedCapture.ProcessId)
        $currentProcessIds=@(Get-Process POWERPNT | ForEach-Object Id)
        if($priorProcessIds.Count -ne 1 -or $currentProcessIds.Count -ne 1 -or $priorProcessIds[0] -eq $currentProcessIds[0]){throw 'Expected exactly one PowerPoint process with a different PID from saved.json.'}
        $reopened=Read-State
        Assert-Reopened $reopened $savedCapture.State
        Assert-Relations $reopened
        $verificationPassed=$true
    }catch{
        $verificationFailure=$_.Exception.ToString()
        throw
    }finally{
        $verification=[pscustomobject]@{
            Passed=$verificationPassed;ReadOnly=$true;PhysicalMouseInput=$false;PersistenceNewProcess=$true
            SavedProcessId=$priorProcessIds;ReopenedProcessId=$currentProcessIds
            Checks=@('different native process from saved.json','all native types, text, AutoSize, aspect locks and child membership preserved','geometry within 0.01 pt, exact identity tags and XML','independent relation formulas')
            Failure=$verificationFailure;Directory=$directory;Utc=[DateTime]::UtcNow.ToString('O')
        }
        $verification | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'new-process-persistence.json') -Encoding UTF8
        $verification | ConvertTo-Json -Depth 10
    }
    return
}
try {
    Add-Type -AssemblyName System.Drawing
    $png=Join-Path $directory 'fixture.png';$svg=Join-Path $directory 'fixture.svg'
    $bitmap=New-Object Drawing.Bitmap 160,100;$graphics=[Drawing.Graphics]::FromImage($bitmap)
    try{$graphics.Clear([Drawing.Color]::SteelBlue);$graphics.FillEllipse([Drawing.Brushes]::Gold,30,15,100,70);$bitmap.Save($png,[Drawing.Imaging.ImageFormat]::Png)}finally{$graphics.Dispose();$bitmap.Dispose()}
    [IO.File]::WriteAllText($svg,'<svg xmlns="http://www.w3.org/2000/svg" width="160" height="100" viewBox="0 0 160 100"><path d="M10 80L60 10L100 70L150 20V90H10Z" fill="#1987b5"/></svg>')
    $p=$ppt.Presentations.Add();$p.Tags.Add('PPTGCS_TEST','V3');$p.PageSetup.SlideWidth=960;$p.PageSetup.SlideHeight=540
    for($i=0;$i -lt $cases.Count;$i++){
        $case=$cases[$i];$slide=$p.Slides.Add($i+1,12);$slide.Name=$case.Kind
        $null=Add-TestShape $slide $case ('Case_'+$case.Kind+'_A')
        $b=$slide.Shapes.AddShape(9,500,290,210,100);$b.Name='Case_'+$case.Kind+'_B'
    }
    $slide=$p.Slides.Add($cases.Count+1,12);$slide.Name='Unsupported'
    $s=$slide.Shapes.AddShape(1,100,120,120,80);$s.Name='Unsupported_Peer'
    $s=$slide.Shapes.AddShape(9,400,120,130,90);$s.Name='Unsupported_Rotated';$s.Rotation=20
    $s=$slide.Shapes.AddLine(100,350,500,350);$s.Name='Unsupported_Line'
    $s=$slide.Shapes.AddTable(2,2,600,300,180,120);$s.Name='Unsupported_Table'
    $p.SaveAs([string]$fixture,24);$p.Close();$p=$ppt.Presentations.Open([string]$fixture,0,0,-1)
    for($i=0;$i -lt $cases.Count;$i++){
        $case=$cases[$i];$prefix='{0:00}-{1}' -f ($i+1),$case.Kind
        $ppt.ActiveWindow.View.GotoSlide($i+1);$slide=$p.Slides.Item($i+1)
        $a=$slide.Shapes.Item('Case_'+$case.Kind+'_A');$b=$slide.Shapes.Item('Case_'+$case.Kind+'_B')
        $slide.Shapes.Range(@($a.Name,$b.Name)).Select()
        $yId=$addin.AddConstraintCommand('SameY','{}');$wId=$addin.AddConstraintCommand('SameWidth','{}')
        if(-not $yId -or -not $wId){throw 'Mixed shape relation was not created'}
        $state=Save-State ($prefix+'-linked');Assert-Relations $state
        $drivers=@($a.Tags.Item('PPTGCS_ID'),$b.Tags.Item('PPTGCS_ID'));$edits=@()
        for($driver=0;$driver -lt 2;$driver++){
            foreach($mode in @('Move','ResizeE')){
                $before=$state;$sign=if($driver -eq 0){1}else{-1};$dx=if($mode -eq 'Move'){0}else{24*$sign};$dy=if($mode -eq 'Move'){24*$sign}else{0}
                $offsets='[[{0},{1}],[{2},{3}],[{4},{5}]]' -f ([int]($dx/3)),([int]($dy/3)),([int]($dx*2/3)),([int]($dy*2/3)),$dx,$dy
                $null=$addin.ReplayGestureForTest([string]$drivers[$driver],$mode,$offsets)
                $name=$prefix+'-driver-'+($driver+1)+'-'+$mode;$state=Save-State ($name+'-after')
                if(Same-State $state $before){throw ($name+' did not change geometry')}
                Assert-Relations $state;Assert-NonGeometry $state $before
                # Independent checks also catch native AutoSize/group side effects
                # that do not violate the two relation residuals themselves.
                for($page=0;$page -lt $state.Snapshot.slides.Count;$page++){
                    $oldGeometry=$before.Snapshot.slides[$page].geometry
                    foreach($shape in $state.Snapshot.slides[$page].geometry){
                        $prior=@($oldGeometry | Where-Object Id -eq $shape.Id)[0]
                        $unchanged=if($page -ne $i){@('Left','Top','Width','Height')}elseif($mode -eq 'Move'){@('Width','Height')}else{@('Left','Top','Height')}
                        foreach($field in $unchanged){if([math]::Abs($shape.$field-$prior.$field) -gt .01){throw ($name+' unexpectedly changed '+$field)}}
                    }
                }
                Native-History 'Undo' ($name+'-undo') $before;Native-History 'Redo' ($name+'-redo') $state
                $edits += [pscustomobject]@{Driver=$driver+1;Mode=$mode;NativeUndo=$true;NativeRedo=$true}
            }
        }
        $results += [pscustomobject]@{Kind=$case.Kind;OfficeType=$case.Type;Passed=$true;Edits=$edits}
        Write-Output ('PASS '+$case.Kind+': both drivers move/resize, properties preserved and native Undo/Redo')
    }
    $ppt.ActiveWindow.View.GotoSlide($cases.Count+1);$slide=$p.Slides.Item($cases.Count+1)
    foreach($bad in @('Unsupported_Rotated','Unsupported_Line','Unsupported_Table')){
        $slide.Shapes.Range(@('Unsupported_Peer',$bad)).Select();$before=Save-State ('reject-'+$bad+'-before');$rejected=$false
        try{$null=$addin.AddConstraintCommand('SameY','{}')}catch{$rejected=$true}
        if(-not $rejected -or -not (Same-State (Save-State ('reject-'+$bad+'-after')) $before)){throw ($bad+' was not rejected atomically')}
        $checks += $bad+' rejected without geometry, tag or XML changes'
    }
    $ppt.ActiveWindow.View.GotoSlide(1);$slide=$p.Slides.Item(1);$beforeCopy=Save-State 'copy-before'
    $copies=$slide.Shapes.Range(@('Case_Ellipse_A','Case_Ellipse_B')).Duplicate();$copies.Select()
    $deadline=[DateTime]::UtcNow.AddSeconds(3)
    do{$copyIds=@($copies.Item(1).Tags.Item('PPTGCS_ID'),$copies.Item(2).Tags.Item('PPTGCS_ID'));$originalIds=@($slide.Shapes.Item('Case_Ellipse_A').Tags.Item('PPTGCS_ID'),$slide.Shapes.Item('Case_Ellipse_B').Tags.Item('PPTGCS_ID'));$ready=($copyIds[0] -and $copyIds[1] -and $copyIds[0] -ne $copyIds[1] -and @($copyIds | Where-Object {$_ -in $originalIds}).Count -eq 0);if(-not $ready){Start-Sleep -Milliseconds 50}}while(-not $ready -and [DateTime]::UtcNow -lt $deadline)
    if(-not $ready){throw 'Copied ellipses did not receive independent IDs'}
    $copied=Save-State 'copy-after'
    if($copied.ConstraintXml -cne $beforeCopy.ConstraintXml){throw 'Copy altered original relation data'}
    foreach($old in $beforeCopy.Slides[0].Shapes){$actual=@($copied.Slides[0].Shapes | Where-Object NativeId -eq $old.NativeId)[0];if(($old|ConvertTo-Json -Depth 30 -Compress) -cne ($actual|ConvertTo-Json -Depth 30 -Compress)){throw 'Copy changed an original shape'}}
    $copies.Select();$newId=$addin.AddConstraintCommand('SameHeight','{}');$linked=Save-State 'copy-linked'
    $relation=@($linked.Snapshot.slides[0].constraints | Where-Object Id -eq $newId)[0]
    if($null -eq $relation -or @($relation.ShapeIds | Where-Object {$_ -notin $copyIds}).Count -ne 0){throw 'Copied shapes relation references originals'}
    Native-History 'Undo' 'copy-link-undo' $copied;Native-History 'Redo' 'copy-link-redo' $linked
    $checks += 'Copied nonrectangles have distinct IDs; originals unchanged; new relation Undo/Redo passes'
    Assert-Relations $linked;$p.Save();$saved=Save-State 'saved';$p.Close();$p=$ppt.Presentations.Open([string]$fixture,0,0,-1)
    $reopened=Save-State 'reopened';Assert-Reopened $reopened $saved;Assert-Relations $reopened
    $checks += 'Same-process save/reopen preserves all types, native properties, geometry, group membership and relation data'
    $passed=$true
}catch{$failure=$_.Exception.ToString();$failure | Set-Content -LiteralPath (Join-Path $directory 'failure.txt') -Encoding UTF8;throw}
finally{
    if($null -ne $p){try{if($p.FullName -eq $fixture){$p.Save();$p.Close()}}catch{}}
    if($null -ne $returnWindow){try{$returnWindow.Activate()}catch{}}
    $report=[pscustomobject]@{Passed=$passed;PhysicalMouseInput=$false;PersistenceNewProcess=$false;Cases=$results;ExpectedCases=$cases.Count;CompletedCases=$results.Count;Checks=$checks;Failure=$failure;Directory=$directory;Utc=[DateTime]::UtcNow.ToString('O')}
    $report | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $directory 'acceptance.json') -Encoding UTF8
    $report | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $evidenceRoot 'latest.json') -Encoding UTF8
    $report | ConvertTo-Json -Depth 15
}
