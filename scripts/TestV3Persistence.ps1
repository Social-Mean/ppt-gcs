[CmdletBinding()]
param(
    [ValidateSet('Capture','Verify')][string]$Action='Capture',
    [Parameter(Mandatory=$true)][string]$Path,
    [string]$Baseline,
    [string]$OutputDirectory
)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if(-not $OutputDirectory){$OutputDirectory=Join-Path $root 'artifacts\v3-persistence'}
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(-not $Baseline){$Baseline=Join-Path $OutputDirectory 'before.json'}
$Baseline=[IO.Path]::GetFullPath($Baseline)
$expectedPath=[IO.Path]::GetFullPath($Path)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p=$ppt.ActivePresentation
if($p.FullName -ne $expectedPath -or $p.Tags.Item('PPTGCS_TEST') -ne 'V3'){throw 'Open the requested owned V3 test document before this read-only check.'}
$processIds=@(Get-Process POWERPNT | ForEach-Object Id)
if($processIds.Count -ne 1){throw 'Persistence evidence requires exactly one PowerPoint process.'}
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
$expectedKinds=@('SameX','SameY','SameWidth','SameHeight','EqualHorizontalGap','EqualVerticalGap','RightAligned','BottomAligned','CenterXAligned','CenterYAligned','CentersCoincident','FixedWidth','FixedHeight','FixedPosition','AspectRatio','FixedHorizontalGap','FixedVerticalGap','Offset','SymmetricVertical','SymmetricHorizontal')
function Gcs-Tags($object) {
    @(for($i=1;$i -le $object.Tags.Count;$i++){
        $name=$object.Tags.Name($i)
        if($name -like 'PPTGCS_*'){[pscustomobject]@{Name=$name;Value=$object.Tags.Value($i)}}
    }) | Sort-Object Name
}
function Read-State {
    $snapshot=$addin.GetSnapshot() | ConvertFrom-Json
    $xml=$addin.GetConstraintXml()
    if($snapshot.constraintXml -cne $xml){throw 'Snapshot and authoritative XML disagree.'}
    [pscustomobject][ordered]@{
        Utc=[DateTime]::UtcNow.ToString('O')
        ProcessId=$processIds
        Path=$p.FullName
        SlideWidth=$p.PageSetup.SlideWidth
        SlideHeight=$p.PageSetup.SlideHeight
        PresentationTags=@(Gcs-Tags $p)
        ConstraintXml=$xml
        Slides=@(foreach($slide in $p.Slides){
            [pscustomobject][ordered]@{Id=$slide.SlideID;Name=$slide.Name;Tags=@(Gcs-Tags $slide);Shapes=@(foreach($shape in $slide.Shapes){
                [pscustomobject][ordered]@{NativeId=$shape.Id;Name=$shape.Name;Type=$shape.Type;Tags=@(Gcs-Tags $shape);Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height;Rotation=$shape.Rotation}
            })}
        })
        Snapshot=$snapshot
    }
}
function Relation-Residual($relation,$byId) {
    $members=@(foreach($id in $relation.ShapeIds){if(-not $byId.ContainsKey($id)){throw 'Persisted relation references a missing shape'};$byId[$id]})
    $a=$members[0];$r=[Collections.Generic.List[double]]::new()
    switch($relation.Kind){
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
                    default {throw ('Unknown persisted constraint: '+$relation.Kind)}
                }
            }
        }
    }
    if($r.Count -eq 0){throw ('No persistence formula for '+$relation.Kind)}
    return ($r | ForEach-Object {[Math]::Abs($_)} | Measure-Object -Maximum).Maximum
}
function Check-Constraints($state) {
    $kinds=@();$maximum=0.0;$count=0;$enabled=0;$ids=@()
    foreach($slide in $state.Snapshot.slides){
        $byId=@{}
        foreach($shape in $slide.geometry){
            if($byId.ContainsKey($shape.Id)){throw 'Duplicate geometry GUID on a slide'}
            $byId[$shape.Id]=$shape;$ids+=$shape.Id
            foreach($field in @('Left','Top','Width','Height')){if([double]::IsNaN($shape.$field) -or [double]::IsInfinity($shape.$field)){throw 'Nonfinite geometry'}}
            if($shape.Width -lt 0.99 -or $shape.Height -lt 0.99){throw 'Invalid persisted shape size'}
        }
        foreach($relation in $slide.constraints){
            $count++;$kinds+=$relation.Kind
            if(-not $relation.Enabled){continue}
            $enabled++;$residual=Relation-Residual $relation $byId
            if($residual -gt 0.01){throw ($relation.Kind+' does not hold after persistence: '+$residual+' pt')}
            $maximum=[Math]::Max($maximum,$residual)
        }
    }
    if(@($ids | Select-Object -Unique).Count -ne $ids.Count){throw 'A GUID is reused by persisted shapes'}
    $uniqueKinds=@($kinds | Select-Object -Unique)
    if($uniqueKinds.Count -ne 20 -or @($expectedKinds | Where-Object {$_ -notin $uniqueKinds}).Count -ne 0){throw 'Persistence fixture does not cover all twenty constraint kinds'}
    if($enabled -ne $count){throw 'Persistence fixture contains a disabled relation; all twenty kinds must be effective'}
    [pscustomobject]@{Kinds=$uniqueKinds;Constraints=$count;Enabled=$enabled;MaximumResidualPt=$maximum;TaggedShapes=$ids.Count}
}
function Assert-Number([double]$actual,[double]$expected,[string]$label) {
    if([double]::IsNaN($actual) -or [double]::IsInfinity($actual) -or [Math]::Abs($actual-$expected) -gt 0.01){throw ('Persistence geometry mismatch: '+$label)}
}
function Assert-Metadata($actual,$expected,[string]$label) {
    if(($actual | ConvertTo-Json -Depth 25 -Compress) -cne ($expected | ConvertTo-Json -Depth 25 -Compress)){throw ('Persistence metadata mismatch: '+$label)}
}
function Compare-Documents($actual,$expected) {
    if($actual.Path -ne $expected.Path){throw 'Persistence document path changed'}
    if($actual.ConstraintXml -cne $expected.ConstraintXml){throw 'Authoritative constraint XML changed'}
    Assert-Metadata $actual.PresentationTags $expected.PresentationTags 'presentation tags'
    Assert-Number $actual.SlideWidth $expected.SlideWidth 'slide width'
    Assert-Number $actual.SlideHeight $expected.SlideHeight 'slide height'
    if($actual.Slides.Count -ne $expected.Slides.Count){throw 'Slide count changed'}
    for($i=0;$i -lt $expected.Slides.Count;$i++){
        $a=$actual.Slides[$i];$b=$expected.Slides[$i]
        if($a.Id -ne $b.Id -or $a.Name -cne $b.Name){throw 'Slide identity or name changed'}
        Assert-Metadata $a.Tags $b.Tags ('slide '+$b.Id+' tags')
        if($a.Shapes.Count -ne $b.Shapes.Count){throw ('Shape count changed on slide '+$b.Id)}
        for($j=0;$j -lt $b.Shapes.Count;$j++){
            $x=$a.Shapes[$j];$y=$b.Shapes[$j]
            if($x.NativeId -ne $y.NativeId -or $x.Name -cne $y.Name -or $x.Type -ne $y.Type){throw 'Native shape identity, name, order or type changed'}
            Assert-Metadata $x.Tags $y.Tags ('shape '+$b.Id+'/'+$y.NativeId+' tags')
            foreach($field in @('Left','Top','Width','Height','Rotation')){Assert-Number $x.$field $y.$field ('shape '+$b.Id+'/'+$y.NativeId+' '+$field)}
        }
    }
}
$state=Read-State
$constraintCheck=Check-Constraints $state
if($Action -eq 'Capture') {
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Baseline) | Out-Null
    $state | ConvertTo-Json -Depth 35 | Set-Content -LiteralPath $Baseline -Encoding UTF8
    [pscustomobject]@{Passed=$true;Action='Capture';Baseline=$Baseline;Path=$state.Path;ProcessId=$processIds;ConstraintCheck=$constraintCheck;ReadOnly=$true} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'capture-check.json') -Encoding UTF8
    Write-Output ('PASS read-only persistence baseline: '+$constraintCheck.Constraints+' constraints / '+$constraintCheck.TaggedShapes+' shapes')
} else {
    $state | ConvertTo-Json -Depth 35 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'after.json') -Encoding UTF8
    $before=Get-Content -LiteralPath $Baseline -Raw | ConvertFrom-Json
    if(@($before.ProcessId).Count -ne 1 -or $before.ProcessId[0] -eq $processIds[0]){throw 'PowerPoint was not restarted into a different process'}
    Compare-Documents $state $before
    $report=[pscustomobject]@{Passed=$true;ReadOnly=$true;PhysicalMouseInput=$false;Path=$state.Path;BeforeProcessId=$before.ProcessId;AfterProcessId=$processIds;GeometryTolerancePt=0.01;MetadataExact=$true;ConstraintCheck=$constraintCheck;Checks=@('different PowerPoint process','all twenty constraint kinds effective','independent relation formulas','all native IDs and shape names','GUIDs and ownership tags','all PPTGCS tags','authoritative XML exact','geometry within 0.01 pt')}
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'acceptance.json') -Encoding UTF8
    Write-Output 'PASS new-process persistence: metadata exact, geometry preserved, all twenty kinds effective'
    $report | ConvertTo-Json -Depth 12
}
