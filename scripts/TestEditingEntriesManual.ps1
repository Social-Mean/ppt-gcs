#requires -Version 5.1
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$FixtureDirectory, [switch]$KeyboardAndPosition)

$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'OfficeComSession.ps1')
$repository=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=[IO.Path]::GetFullPath($FixtureDirectory).TrimEnd('\','/')
$artifacts=[IO.Path]::GetFullPath((Join-Path $repository 'artifacts')).TrimEnd('\','/')
$leaf=[IO.Path]::GetFileName($directory);$guid=[guid]::Empty
if([IO.Path]::GetDirectoryName($directory) -ine $artifacts -or $leaf -notmatch '^editing-entries-manual-([0-9a-fA-F-]{36})$' -or
   -not [guid]::TryParseExact($Matches[1],'D',[ref]$guid) -or $guid -eq [guid]::Empty -or -not (Test-Path -LiteralPath $directory -PathType Container)) {
    throw 'FixtureDirectory must be an existing direct artifacts/editing-entries-manual-<GUID> directory in this repository.'
}
$manifest=Get-Content -LiteralPath (Join-Path $directory 'fixture.json') -Raw|ConvertFrom-Json
$baseline=Get-Content -LiteralPath (Join-Path $directory 'beforeFrames.json') -Raw|ConvertFrom-Json
$targetPages=if($KeyboardAndPosition){@(1,3,5,8,9)}else{@(1..12)}
$fixture=[IO.Path]::GetFullPath((Join-Path $directory 'EditingEntriesManual.pptx'))
$token=$guid.ToString('D').ToUpperInvariant();$tag='PPTGCS_EDITING_ENTRIES_MANUAL'
if([IO.Path]::GetFullPath([string]$manifest.Path) -ine $fixture -or $manifest.Token -cne $token -or $manifest.TestTag -cne $tag -or
   $manifest.Baseline -cne 'beforeFrames.json' -or $manifest.ExpectedDataEdits -ne 12 -or $manifest.ExpectedNativeUndoCount -ne 12 -or
   $manifest.ExpectedNativeRedoCount -ne 12 -or @($manifest.Cases).Count -ne 12 -or @($baseline.Rows).Count -ne 45 -or @($baseline.Slides).Count -ne 12) {
    throw 'The frozen 12-page/45-row manual fixture manifest and baseline are required.'
}
if(Test-Path -LiteralPath (Join-Path $directory 'manual-result.json')) { throw 'This fixture already has a manual result. Do not repeat history against an unknown current Undo position.' }
$ppt=$null;$document=$null;$owned=$null;$window=$null;$addin=$null;$filter=$null;$failure=$null
$historyStarted=$false;$preflightPassed=$false;$persisted=$false;$filterRestored=$false;$traceEndByte=$null
$otherDocuments=@();$operations=@();$results=New-Object 'Collections.Generic.List[object]';$cleanupErrors=New-Object 'Collections.Generic.List[string]'
$byBefore=@{};$byFinal=@{};$caseByPage=@{};$caseByGuid=@{};$caseBySlide=@{}
$tolerance=.01

function Write-Evidence([string]$Name,$Value) {
    $Value|ConvertTo-Json -Depth 35|Set-Content -LiteralPath (Join-Path $directory ($Name+'.json')) -Encoding UTF8
}
function Get-TagValue($Tags,[string]$Name) {
    if($Tags -is [Collections.IDictionary]) { return [string]$Tags[$Name] }
    $property=$Tags.PSObject.Properties[$Name];if($null -eq $property){return $null};return [string]$property.Value
}
function Get-Tags($Owner) {
    $values=@{};for($index=1;$index -le $Owner.Tags.Count;$index++){$values[[string]$Owner.Tags.Name($index)]=[string]$Owner.Tags.Value($index)}
    $ordered=[ordered]@{};foreach($name in @($values.Keys|Sort-Object)){$ordered[$name]=$values[$name]};return $ordered
}
function Tags-Json($Tags) {
    $ordered=[ordered]@{}
    $names=if($Tags -is [Collections.IDictionary]){@($Tags.Keys)}else{@($Tags.PSObject.Properties.Name)}
    foreach($name in @($names|Sort-Object)){$ordered[$name]=Get-TagValue $Tags $name};return ($ordered|ConvertTo-Json -Compress)
}
function Assert-Owned([switch]$Active) {
    if($null -eq $document -or -not (Test-OfficeComIdentity $document $owned) -or
       [IO.Path]::GetFullPath([string]$document.FullName) -ine $fixture -or [string]$document.Tags.Item('PPTGCS_TEST') -cne 'V3' -or
       [string]$document.Tags.Item($tag) -cne $token) { throw 'The exact owned V3 document/path/nonce must remain valid.' }
    foreach($other in $otherDocuments){if(Test-OfficeComIdentity $document $other.Document){throw 'Another open document cannot be mutated.'}}
    if($Active -and (-not (Test-OfficeComIdentity $ppt.ActivePresentation $document) -or
       -not (Test-OfficeComIdentity $ppt.ActiveWindow $window) -or [int]$window.ViewType -ne 9)) { throw 'Keep the captured owned document/window active in normal view.' }
}
function Get-Snapshot([string]$Name) {
    Assert-Owned -Active
    $rows=New-Object 'Collections.Generic.List[object]';$slides=New-Object 'Collections.Generic.List[object]'
    function Read-Shape($Shape,[string]$Key,[int]$ParentId) {
        $text=$null;if([int]$Shape.HasTextFrame -ne 0){$text=[string]$Shape.TextFrame.TextRange.Text}
        $rows.Add([pscustomobject]@{Key=$Key;NativeId=[int]$Shape.Id;ParentNativeId=$ParentId;Name=[string]$Shape.Name;Type=[int]$Shape.Type;
            Child=[int]$Shape.Child;Tags=Get-Tags $Shape;Text=$text;Left=[double]$Shape.Left;Top=[double]$Shape.Top;
            Width=[double]$Shape.Width;Height=[double]$Shape.Height;Rotation=[double]$Shape.Rotation;LockAspectRatio=[int]$Shape.LockAspectRatio})
        if([int]$Shape.Type -eq 6){for($i=1;$i -le $Shape.GroupItems.Count;$i++){$child=$Shape.GroupItems.Item($i);Read-Shape $child ($Key+'/'+[string]$child.Name) ([int]$Shape.Id)}}
    }
    for($page=1;$page -le $document.Slides.Count;$page++){
        $slide=$document.Slides.Item($page);$slides.Add([pscustomobject]@{Index=$page;Id=[int]$slide.SlideID;Name=[string]$slide.Name;Tags=Get-Tags $slide})
        for($index=1;$index -le $slide.Shapes.Count;$index++){$shape=$slide.Shapes.Item($index);Read-Shape $shape ($page.ToString()+'/'+[string]$shape.Name) 0}
    }
    Assert-Owned -Active
    $selection=$window.Selection;$ids=@();$children=@();$hasChild=[bool]$selection.HasChildShapeRange
    if([int]$selection.Type -in @(2,3)){$ids=@(for($i=1;$i -le $selection.ShapeRange.Count;$i++){[int]$selection.ShapeRange.Item($i).Id})}
    if($hasChild){$children=@(for($i=1;$i -le $selection.ChildShapeRange.Count;$i++){[int]$selection.ChildShapeRange.Item($i).Id})}
    $value=[pscustomobject]@{Utc=[datetime]::UtcNow.ToString('O');ConstraintXml=[string]$addin.GetConstraintXml();DocumentTags=Get-Tags $document;
        SlideWidth=[double]$document.PageSetup.SlideWidth;SlideHeight=[double]$document.PageSetup.SlideHeight;SnapToGrid=[int]$document.SnapToGrid;
        GridDistance=[double]$document.GridDistance;Slides=@($slides.ToArray());Rows=@($rows.ToArray());ActiveSlideId=[int]$window.View.Slide.SlideID;
        ActiveSlideIndex=[int]$window.View.Slide.SlideIndex;Selection=[pscustomobject]@{Type=[int]$selection.Type;HasChild=$hasChild;NativeIds=$ids;ChildNativeIds=$children}}
    Write-Evidence $Name $value;return $value
}
function Assert-Near($Actual,$Expected,[string]$Name) {
    if($null -eq $Actual -or $null -eq $Expected -or [double]::IsNaN([double]$Actual) -or [double]::IsInfinity([double]$Actual) -or
       [math]::Abs([double]$Actual-[double]$Expected) -gt $tolerance){throw "$Name expected $Expected, got $Actual"}
}
function Assert-SameSnapshot($Actual,$Expected,[string]$Name) {
    foreach($property in @('ConstraintXml','SlideWidth','SlideHeight','GridDistance')){if($Actual.$property -cne $Expected.$property){throw "$Name changed $property"}}
    if((Tags-Json $Actual.DocumentTags) -cne (Tags-Json $Expected.DocumentTags)){throw "$Name changed document tags"}
    if(@($Actual.Slides).Count -ne @($Expected.Slides).Count){throw "$Name changed slide count"}
    for($index=0;$index -lt @($Expected.Slides).Count;$index++){
        $a=$Actual.Slides[$index];$b=$Expected.Slides[$index]
        foreach($property in @('Index','Id','Name')){if($a.$property -cne $b.$property){throw "$Name changed slide $index $property"}}
        if((Tags-Json $a.Tags) -cne (Tags-Json $b.Tags)){throw "$Name changed slide tags"}
    }
    if(@($Actual.Rows).Count -ne @($Expected.Rows).Count){throw "$Name changed recursive shape count"}
    $expectedRows=@{};foreach($row in $Expected.Rows){if($expectedRows.ContainsKey($row.Key)){throw 'Duplicate expected shape key'};$expectedRows[$row.Key]=$row}
    $seen=@{};foreach($row in $Actual.Rows){
        if($seen.ContainsKey($row.Key) -or -not $expectedRows.ContainsKey($row.Key)){throw "$Name changed shape membership: $($row.Key)"};$seen[$row.Key]=$true;$want=$expectedRows[$row.Key]
        foreach($property in @('NativeId','ParentNativeId','Name','Type','Child','Text','LockAspectRatio')){if($row.$property -cne $want.$property){throw "$Name changed $($row.Key).$property"}}
        if((Tags-Json $row.Tags) -cne (Tags-Json $want.Tags)){throw "$Name changed shape tags: $($row.Key)"}
        foreach($axis in @('Left','Top','Width','Height','Rotation')){Assert-Near $row.$axis $want.$axis "$Name $($row.Key).$axis"}
    }
    # UTC, active view/selection and SnapToGrid are observed but excluded from history/persistence equality.
}
function Copy-Snapshot($Value){return ($Value|ConvertTo-Json -Depth 35|ConvertFrom-Json)}
function Assert-Relations($Snapshot,[string]$Name) {
    [xml]$xml=$Snapshot.ConstraintXml;$byId=@{}
    foreach($row in $Snapshot.Rows){$id=Get-TagValue $row.Tags 'PPTGCS_ID';if($id){if($byId.ContainsKey($id)){throw 'Duplicate tagged shape ID'};$byId[$id]=$row}}
    $relations=@($xml.DocumentElement.ChildNodes|Where-Object LocalName -eq 'constraint')
    if($relations.Count -ne 16){throw "$Name must contain exactly 16 fixture relations"}
    foreach($relation in $relations){
        if($relation.GetAttribute('enabled') -eq 'false'){throw "$Name disabled a fixture relation"}
        $axis=switch($relation.GetAttribute('kind')){'SameX'{'Left'}'SameY'{'Top'}'SameWidth'{'Width'}'SameHeight'{'Height'}default{throw 'Unexpected fixture relation kind'}}
        $members=@(foreach($shape in $relation.ChildNodes){if($shape.LocalName -eq 'shape'){$id=$shape.GetAttribute('id');if(-not $byId.ContainsKey($id)){throw "$Name missing relation member $id"};$byId[$id]}})
        if($members.Count -ne 2){throw 'Fixture relations must have two members'}
        foreach($member in $members){Assert-Near $member.$axis $members[0].$axis "$Name relation $($relation.GetAttribute('kind'))"}
    }
}
function Allow-Axes($Map,[string]$Key,[string[]]$Axes){$Map[$Key]=$Axes}
function Assert-Final($Final) {
    foreach($row in $Final.Rows){$byFinal[$row.Key]=$row}
    $allowed=@{}
    foreach($case in $manifest.Cases){
        $p=[int]$case.Page;if($p -notin $targetPages){continue};$a=$byBefore["$p/A"];$now=$byFinal["$p/A"]
        switch($p){
            1 {foreach($name in @('A','B')){Allow-Axes $allowed "$p/$name" @('Top')};if($now.Top -ge $a.Top-.01){throw 'Page1 did not move A upward'}}
            {$_ -in @(2,3)} {
                foreach($name in @('A','D','ZeroWidthLine')){Allow-Axes $allowed "$p/$name" @('Left','Top')}
                Allow-Axes $allowed "$p/B" $(if($p -eq 2){@('Top')}else{@('Left')})
                $dx=$now.Left-$a.Left;$dy=$now.Top-$a.Top
                if(($p -eq 2 -and $dy -ge -.01) -or ($p -eq 3 -and $dx -le .01)){throw "Page$p driver did not move in the intended direction"}
                if($p -eq 3){Assert-Near $dy 0 'Page3 keyboard vertical delta'}
                foreach($name in @('D','ZeroWidthLine')){$old=$byBefore["$p/$name"];$n=$byFinal["$p/$name"];Assert-Near ($n.Left-$old.Left) $dx "Page$p $name dx";Assert-Near ($n.Top-$old.Top) $dy "Page$p $name dy"}
            }
            {$_ -in @(4,5)} {
                $old=$byBefore["$p/AB/A"];$n=$byFinal["$p/AB/A"]
                Allow-Axes $allowed "$p/AB/A" $(if($p -eq 4){@('Left','Top')}else{@('Top')})
                Allow-Axes $allowed "$p/AB" $(if($p -eq 4){@('Left','Top','Width','Height')}else{@('Top','Height')})
                Allow-Axes $allowed "$p/C" $(if($p -eq 4){@('Top','Width','Height')}else{@('Top','Height')})
                if($n.Top -ge $old.Top-.01){throw "Page$p child A did not move upward"}
            }
            default {
                $axis=switch($p){6{'Width'}7{'Height'}8{'Left'}9{'Top'}10{'Width'}11{'Height'}12{'Width'}}
                $cm=switch($p){6{6}7{4}8{6}9{4}10{7}11{4.5}12{6}};$target=[double]$cm*72/2.54
                if($p -in @(8,9)){
                    $half=if($p -eq 8){$baseline.SlideWidth/2}else{$baseline.SlideHeight/2}
                    if([math]::Abs($now.$axis-$target) -gt .01 -and [math]::Abs($now.$axis-($target+$half)) -gt .01){throw "Page$p scalar target does not match TopLeft or current Center origin"}
                    $target=$now.$axis
                }
                foreach($name in $(if($p -eq 12){@('A','B','C')}else{@('A','B')})){Allow-Axes $allowed "$p/$name" @($axis);Assert-Near $byFinal["$p/$name"].$axis $target "Page$p $name scalar target"}
                if([math]::Abs($now.$axis-$a.$axis) -le .01){throw "Page$p driver scalar axis did not change"}
            }
        }
    }
    $expected=Copy-Snapshot $baseline
    foreach($row in $expected.Rows){if($allowed.ContainsKey($row.Key)){foreach($axis in $allowed[$row.Key]){$row.$axis=$byFinal[$row.Key].$axis}}}
    Assert-SameSnapshot $Final $expected 'Final allowed edits (all other geometry/text/tags/IDs unchanged)'
    Assert-Relations $Final 'Final'
}
function Assert-IdSet($Actual,$Expected,[string]$Name) {
    if(((@($Actual)|ForEach-Object{[string]$_}|Sort-Object)-join ',') -cne ((@($Expected)|ForEach-Object{[string]$_}|Sort-Object)-join ',')){throw "$Name has unexpected selected/driver IDs"}
}
function Assert-TraceGeometry($Geometry,$Snapshot,[int]$Page,[string]$Name,[switch]$NativeIds) {
    $rows=@($Snapshot.Rows|Where-Object {$_.Key.StartsWith($Page.ToString()+'/',[StringComparison]::Ordinal)})
    foreach($value in @($Geometry|Where-Object {$null -ne $_})){
        $matches=@($rows|Where-Object {if($NativeIds){$_.NativeId -eq $value.NativeId}else{(Get-TagValue $_.Tags 'PPTGCS_ID') -ieq $value.Id}})
        if($matches.Count -ne 1){throw "$Name cannot map trace geometry to one native fixture shape"}
        foreach($axis in @('Left','Top','Width','Height')){Assert-Near $value.$axis $matches[0].$axis "$Name $($matches[0].Key).$axis"}
    }
}
function Assert-SameTraceGeometry($Before,$After,[string]$Name,[switch]$NativeIds) {
    $left=@($Before|Where-Object {$null -ne $_});$right=@($After|Where-Object {$null -ne $_})
    if($left.Count -ne $right.Count){throw "$Name changed trace membership"}
    foreach($value in $left){
        $matches=@($right|Where-Object {if($NativeIds){$_.NativeId -eq $value.NativeId}else{$_.Id -ceq $value.Id}})
        if($matches.Count -ne 1){throw "$Name has missing/duplicate trace members"}
        foreach($axis in @('Left','Top','Width','Height')){Assert-Near $matches[0].$axis $value.$axis "$Name $axis"}
    }
}
function Get-EventCase($Data) {
    $pages=@{}
    foreach($id in @($Data.driverId)+@($Data.driverIds)+@($Data.selectedDriverIds)+@($Data.RelatedIds)+@($Data.geometry|ForEach-Object{$_.Id})+@($Data.before|ForEach-Object{$_.Id})){
        if($id -and $caseByGuid.ContainsKey([string]$id)){$pages[[int]$caseByGuid[[string]$id].Page]=$true}
    }
    if($null -ne $Data.slideId -and $caseBySlide.ContainsKey([int]$Data.slideId)){$pages[[int]$caseBySlide[[int]$Data.slideId].Page]=$true}
    if($pages.Count -gt 1){throw 'One native trace event unexpectedly refers to multiple manual pages'}
    if($pages.Count -eq 1){return $caseByPage[@($pages.Keys)[0]]};return $null
}
function Read-ActualTrace($Final) {
    $expectedLog=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PptGcs/layout-v2.log'
    if([IO.Path]::GetFullPath([string]$manifest.Trace.LogPath) -ine [IO.Path]::GetFullPath($expectedLog)){throw 'The manifest must bind the production native log'}
    $stream=[IO.File]::Open($expectedLog,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    try{
        $start=[long]$manifest.Trace.StartByte;$script:traceEndByte=$stream.Length
        if($start -lt 0 -or $start -gt $traceEndByte){throw 'The native log was truncated or the manifest byte offset is invalid'}
        [void]$stream.Seek($start,[IO.SeekOrigin]::Begin);$bytes=New-Object byte[] ([int]($traceEndByte-$start));$read=0
        while($read -lt $bytes.Length){$count=$stream.Read($bytes,$read,$bytes.Length-$read);if($count -eq 0){throw 'Incomplete native trace read'};$read+=$count}
        $text=([Text.UTF8Encoding]::new($false,$true)).GetString($bytes)
    }finally{$stream.Dispose()}
    $events=New-Object 'Collections.Generic.List[object]';$index=0;$operationCases=@{}
    foreach($line in ($text -split '\r?\n')){
        if(-not $line){continue};$index++
        if($line -notmatch '^(\S+) ([a-zA-Z0-9-]+) (\{.*\})$'){throw "Malformed native trace at line $index"}
        $utc=[datetime]::Parse($Matches[1],[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::RoundtripKind)
        $kind=$Matches[2];$data=$Matches[3]|ConvertFrom-Json;$case=Get-EventCase $data
        $operationId=if($data.gestureId){[string]$data.gestureId}elseif($kind -match 'keyboard-nudge' -and $data.id){[string]$data.id}else{$null}
        if($null -eq $case -and $operationId -and $operationCases.ContainsKey($operationId)){$case=$operationCases[$operationId]}
        if($null -ne $case){
            if($utc -lt [datetime]::Parse($manifest.ReadyUtc,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::RoundtripKind)){throw 'An owned trace event predates the fixture ready marker'}
            if($operationId){$operationCases[$operationId]=$case}
            $events.Add([pscustomobject]@{Index=$index;Utc=$utc.ToString('O');Kind=$kind;Page=[int]$case.Page;Data=$data})
        }
    }
    Write-Evidence 'manual-actual-trace' @($events.ToArray())
    $ops=New-Object 'Collections.Generic.List[object]';$starts=@{};$frames=@{};$beforeFields=@{};$usedEnds=@{};$focusFields=@{}
    foreach($event in $events){
        $data=$event.Data;$case=$caseByPage[$event.Page];$kind=$event.Kind
        if($kind -like 'test-*'){throw 'Programmatic test/replay events cannot count as physical manual acceptance'}
        if($kind -match '(error|rejected|mismatch|drift)$'){throw "Native trace contains $kind on page $($case.Page)"}
        if($kind -eq 'native-field-focus'){$focusFields[($case.Page.ToString()+'/'+$data.field)]=$event;continue}
        if($kind -in @('gesture-start','keyboard-nudge-start')){
            if($null -eq $data.PSObject.Properties['test'] -or $data.test -ne $false){throw 'A native start must explicitly have test=false'}
            $id=if($kind -eq 'gesture-start'){[string]$data.gestureId}else{[string]$data.id}
            if(-not $id -or $starts.ContainsKey($id)){throw 'Duplicate/missing native operation ID'};$starts[$id]=$event;$frames[$id]=New-Object 'Collections.Generic.List[object]'
        }
        elseif($kind -in @('gesture-frame','keyboard-nudge-frame')){
            $id=if($kind -eq 'gesture-frame'){[string]$data.gestureId}else{[string]$data.id}
            if(-not $starts.ContainsKey($id)){throw 'Native frame has no bound actual start'};$frames[$id].Add($event)
        }
        elseif($kind -in @('gesture-end','keyboard-nudge-end')){
            $id=if($kind -eq 'gesture-end'){[string]$data.gestureId}else{[string]$data.id}
            if(-not $starts.ContainsKey($id) -or $usedEnds.ContainsKey($id)){throw 'Native end has missing/duplicate start'};$usedEnds[$id]=$true
            $begin=$starts[$id];$list=@($frames[$id].ToArray())
            if($null -eq $data.PSObject.Properties['test'] -or $data.test -ne $false){throw 'A native end must explicitly have test=false'}
            if([int]$data.frames -eq 0){
                if($list.Count -ne 0){throw 'Zero-frame native end has frame records'}
                Assert-SameTraceGeometry $data.before $data.after 'Selection-only geometry'
                Assert-SameTraceGeometry $data.beforeExtras $data.afterExtras 'Selection-only free geometry' -NativeIds
                Assert-SameTraceGeometry $data.childBefore $data.childAfter 'Selection-only child geometry' -NativeIds
                continue
            }
            $minimumFrames=if($KeyboardAndPosition -and $kind -eq 'keyboard-nudge-end'){1}else{2}
            if($data.cancelled -ne $false -or $list.Count -lt $minimumFrames -or [int]$data.frames -ne $list.Count){throw "Native operation needs at least $minimumFrames complete noncancelled actual frames"}
            for($i=0;$i -lt $list.Count;$i++){if([int]$list[$i].Data.frame -ne $i+1){throw 'Native frame sequence is incomplete'};Assert-Near $list[$i].Data.residualPt 0 'Native frame residual'}
            if($kind -eq 'keyboard-nudge-end'){Assert-IdSet $begin.Data.driverIds $case.DriverIds 'Keyboard driver scope'}
            Assert-TraceGeometry $begin.Data.before $baseline $case.Page 'Operation baseline'
            Assert-TraceGeometry $data.before $baseline $case.Page 'Operation end baseline'
            Assert-TraceGeometry $data.after $Final $case.Page 'Operation final'
            if($kind -eq 'gesture-end'){
                if($case.Entry -notin @('pointer-selection','pointer-child') -or $begin.Data.mode -cne 'Move' -or $data.mode -cne 'Move'){throw 'Unexpected pointer entry/mode'}
                Assert-IdSet @($begin.Data.driverId) $case.DriverIds 'Pointer driver'
                $down=@($list|Where-Object {$_.Data.mouseDown -eq $true -and $_.Data.input -ceq 'owned-gesture-loop'})
                if($down.Count -lt 2){throw 'Slow drag did not record two actual button-held owned-loop frames'}
                if($case.Entry -eq 'pointer-selection'){
                    Assert-IdSet $begin.Data.selectedNativeIds $case.ExpectedSelectedNativeIds 'Pointer selected members';Assert-IdSet $data.selectedNativeIds $case.ExpectedSelectedNativeIds 'Pointer final selected members'
                    Assert-TraceGeometry $begin.Data.extras $baseline $case.Page 'Pointer free baseline' -NativeIds
                    Assert-TraceGeometry $data.afterExtras $Final $case.Page 'Pointer free final' -NativeIds
                }else{
                    Assert-IdSet $begin.Data.childNativeIds $case.ExpectedSelectedNativeIds 'Pointer child members';Assert-IdSet $data.childNativeIds $case.ExpectedSelectedNativeIds 'Pointer child final members'
                    Assert-TraceGeometry $begin.Data.childBefore $baseline $case.Page 'Pointer child baseline' -NativeIds
                    Assert-TraceGeometry $data.childAfter $Final $case.Page 'Pointer child final' -NativeIds
                }
            }else{
                if($case.Entry -notin @('keyboard','keyboard-selection','keyboard-child') -or $data.reason -cne 'key-up'){
                    throw 'Expected a keyboard edit closed by the queued key-up'
                }
                $key=if($case.Page -eq 3){39}else{38};if($begin.Data.key -ne $key -or $data.key -ne $key -or $data.undoStarted -ne $true){throw 'Wrong keyboard direction or missing native history boundary'}
                Assert-IdSet $begin.Data.selectedNativeIds $case.ExpectedSelectedNativeIds 'Keyboard selected members';Assert-IdSet $data.selectedNativeIds $case.ExpectedSelectedNativeIds 'Keyboard final selected members'
                if(($case.Entry -eq 'keyboard-child') -ne ($begin.Data.groupChild -eq $true)){throw 'Wrong keyboard child scope'}
                foreach($frame in $list){if($frame.Data.control -ne ($case.Page -eq 3)){throw 'Keyboard Ctrl state does not match the case'}}
                Assert-TraceGeometry $begin.Data.extras $baseline $case.Page 'Keyboard free/child baseline' -NativeIds
                Assert-TraceGeometry $data.afterExtras $Final $case.Page 'Keyboard free/child final' -NativeIds
            }
            $ops.Add([pscustomobject]@{Page=$case.Page;Entry=$case.Entry;StartIndex=$begin.Index;EndIndex=$event.Index;Id=$id;Frames=$list.Count;Kind='ActualNativeGesture';Passed=$true})
        }
        elseif($kind -eq 'native-field-before'){
            if($beforeFields.ContainsKey([int]$case.Page)){throw 'More than one numeric write was attempted on one page'};$beforeFields[[int]$case.Page]=$event
        }
        elseif($kind -eq 'native-field-committed'){
            if(-not $beforeFields.ContainsKey([int]$case.Page)){throw 'Numeric commit has no matching before record'};$begin=$beforeFields[[int]$case.Page]
            $field=switch([int]$case.Page){6{'width'}7{'height'}8{'left'}9{'top'}10{'width'}11{'height'}12{'width'}default{throw 'Unexpected numeric case'}}
            $source=switch([int]$case.Page){7{'tab'}9{'tab'}10{'pointer-before-focus-loss'}11{'pointer-before-focus-loss'}default{'enter'}}
            if($data.field -cne $field -or $begin.Data.field -cne $field -or $data.source -cne $source -or $begin.Data.source -cne $source){throw 'Numeric field/commit source does not match the case'}
            Assert-IdSet $begin.Data.selectedNativeIds $case.ExpectedSelectedNativeIds 'Numeric selected members'
            Assert-TraceGeometry $begin.Data.before $baseline $case.Page 'Numeric baseline';Assert-TraceGeometry $begin.Data.beforeExtras $baseline $case.Page 'Numeric free baseline' -NativeIds
            Assert-TraceGeometry $data.geometry $Final $case.Page 'Numeric final';Assert-TraceGeometry $data.extras $Final $case.Page 'Numeric free final' -NativeIds
            $axis=switch($field){'width'{'Width'}'height'{'Height'}'left'{'Left'}'top'{'Top'}}
            if($case.Page -in @(8,9)){
                $focusKey=$case.Page.ToString()+'/'+$field
                if(-not $focusFields.ContainsKey($focusKey)){throw 'Position field lacks an independently recorded current origin'}
                $origin=$focusFields[$focusKey].Data.origin
                if($origin -cne '左上角' -and $origin -cne '居中'){throw 'Position field current origin is unknown'}
                $expectedPoints=$(if($case.Page -eq 8){6.0}else{4.0})*72/2.54
                if($origin -ceq '居中'){$expectedPoints+=$(if($case.Page -eq 8){$baseline.SlideWidth/2}else{$baseline.SlideHeight/2})}
                Assert-Near $data.points $expectedPoints 'Position field target from independent focus origin'
            }
            Assert-Near $data.points $byFinal[($case.Page.ToString()+'/A')].$axis 'Numeric actual requested target'
            $ops.Add([pscustomobject]@{Page=$case.Page;Entry=$case.Entry;StartIndex=$begin.Index;EndIndex=$event.Index;Id=$null;Frames=$null;Kind='ActualNativeFieldCommit';Source=$source;Passed=$true})
        }
        elseif($kind -eq 'native-field-child-committed'){throw 'This manual matrix has no numeric child-field operation'}
    }
    foreach($id in $starts.Keys){if(-not $usedEnds.ContainsKey($id)){throw 'A native operation has no end record'}}
    foreach($case in $manifest.Cases){if(@($ops|Where-Object Page -eq $case.Page).Count -ne [int]($case.Page -in $targetPages)){throw "Unexpected actual data operation count on page $($case.Page)"}}
    if($ops.Count -ne $targetPages.Count){throw "Expected exactly $($targetPages.Count) actual native data edits"}
    $ordered=@($ops|Sort-Object StartIndex)
    for($i=1;$i -lt $ordered.Count;$i++){if($ordered[$i].StartIndex -le $ordered[$i-1].EndIndex){throw 'Native data operations overlap or interleave'}}
    return $ordered
}
function Expected-Mix($Final,[Collections.Generic.HashSet[int]]$FinalPages) {
    $value=Copy-Snapshot $baseline
    foreach($row in $value.Rows){$page=[int]($row.Key.Split('/')[0]);if($FinalPages.Contains($page)){$source=$byFinal[$row.Key];foreach($axis in @('Left','Top','Width','Height','Rotation')){$row.$axis=$source.$axis}}}
    return $value
}
function Prepare-History([int]$Page) {
    Assert-Owned -Active;$window.View.GotoSlide($Page)
    Assert-Owned -Active
    for($i=1;$i -le $window.Panes.Count;$i++){if([int]$window.Panes.Item($i).ViewType -eq 1){Assert-Owned -Active;$window.Panes.Item($i).Activate();break}}
    Assert-Owned -Active;$window.Selection.Unselect();Assert-Owned -Active
    $name=if($Page -in @(4,5)){'AB'}else{'A'};$document.Slides.Item($Page).Shapes.Item($name).Select()
    Assert-Owned -Active
    if([int]$window.ActivePane.ViewType -ne 1 -or [int]$window.Selection.Type -ne 2 -or [bool]$window.Selection.HasChildShapeRange){throw 'Native document history must target a top-level shape in the slide pane'}
}
function Native-History([string]$Command,[string]$Name) {
    Assert-Owned -Active;if(-not $ppt.CommandBars.GetEnabledMso($Command)){throw "$Command is disabled before $Name"}
    Assert-Owned -Active;$ppt.CommandBars.ExecuteMso($Command);Start-Sleep -Milliseconds 250
    return Get-Snapshot $Name
}

try{
    if($PSVersionTable.PSEdition -eq 'Core' -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA'){throw 'Use powershell.exe -STA -File scripts/TestEditingEntriesManual.ps1 -FixtureDirectory <exact directory>.'}
    foreach($row in $baseline.Rows){if($byBefore.ContainsKey($row.Key)){throw 'Duplicate baseline shape key'};$byBefore[$row.Key]=$row}
    foreach($case in $manifest.Cases){
        $p=[int]$case.Page;if($p -lt 1 -or $p -gt 12 -or $caseByPage.ContainsKey($p) -or $case.ExpectedDataEdits -ne 1){throw 'Malformed manual case page/count'}
        if(@($baseline.Slides|Where-Object {$_.Index -eq $p -and $_.Id -eq $case.SlideId}).Count -ne 1){throw 'Case SlideID does not match baseline'}
        $caseByPage[$p]=$case;$caseBySlide[[int]$case.SlideId]=$case
        foreach($row in $baseline.Rows|Where-Object {$_.Key.StartsWith($p.ToString()+'/',[StringComparison]::Ordinal)}){$id=Get-TagValue $row.Tags 'PPTGCS_ID';if($id){$caseByGuid[$id]=$case}}
        foreach($path in $case.SelectionPaths){if(-not $byBefore.ContainsKey($p.ToString()+'/'+$path)){throw 'Missing manifest selection path'}}
    }
    Assert-Relations $baseline 'Saved baseline'
    $filter=New-Object PptGcs.Testing.OfficeMessageFilterScope
    # Do not create PowerPoint or a presentation when the required fixture is absent.
    $ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    $document=$ppt.ActivePresentation;$owned=$document;$window=$ppt.ActiveWindow
    Assert-Owned -Active
    $otherDocuments=@(for($i=1;$i -le $ppt.Presentations.Count;$i++){$item=$ppt.Presentations.Item($i);if(-not (Test-OfficeComIdentity $item $owned)){[pscustomobject]@{Document=$item;Name=[string]$item.Name;Path=[string]$item.FullName;Saved=[int]$item.Saved}}})
    $addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object;if($null -eq $addin){throw 'The production add-in must be loaded'}
    Start-Sleep -Milliseconds 250;$final=Get-Snapshot 'manual-final'
    Assert-Final $final;$operations=@(Read-ActualTrace $final);Write-Evidence 'manual-operation-order' $operations
    $preflightPassed=$true;$results.Add([pscustomobject]@{Name='FinalAndActualTrace';Passed=$true;Operations=$operations.Count})
    $remaining=New-Object 'Collections.Generic.HashSet[int]';foreach($page in $targetPages){[void]$remaining.Add([int]$page)}
    $expected=$final
    for($i=$operations.Count-1;$i -ge 0;$i--){
        $op=$operations[$i];Prepare-History $op.Page
        $prepared=Get-Snapshot ('manual-before-undo-'+('{0:D2}' -f ($operations.Count-$i)))
        Assert-SameSnapshot $prepared $expected 'Canvas/history preparation must not commit additional edits'
        $historyStarted=$true;$step=$operations.Count-$i;$actual=Native-History 'Undo' ('manual-undo-'+('{0:D2}' -f $step))
        [void]$remaining.Remove([int]$op.Page);$expected=Expected-Mix $final $remaining
        Assert-SameSnapshot $actual $expected "One native Undo $step page $($op.Page)";Assert-Relations $actual "Undo $step"
        $results.Add([pscustomobject]@{Name='Undo';Step=$step;Page=$op.Page;Passed=$true})
    }
    foreach($op in $operations){
        Prepare-History $op.Page;$step=$remaining.Count+1
        $prepared=Get-Snapshot ('manual-before-redo-'+('{0:D2}' -f $step));Assert-SameSnapshot $prepared $expected 'Redo preparation'
        $actual=Native-History 'Redo' ('manual-redo-'+('{0:D2}' -f $step));[void]$remaining.Add([int]$op.Page);$expected=Expected-Mix $final $remaining
        Assert-SameSnapshot $actual $expected "One native Redo $step page $($op.Page)";Assert-Relations $actual "Redo $step"
        $results.Add([pscustomobject]@{Name='Redo';Step=$step;Page=$op.Page;Passed=$true})
    }
    Assert-SameSnapshot $actual $final 'All Redo returns the exact accepted final deck'
    Assert-Owned -Active;$document.Save();$saved=Get-Snapshot 'manual-saved';Assert-SameSnapshot $saved $final 'Save'
    Assert-Owned -Active;$document.Close();$document=$null;$window=$null
    $document=$ppt.Presentations.Open($fixture,0,0,-1);$owned=$document;$window=$ppt.ActiveWindow
    Assert-Owned -Active;Start-Sleep -Milliseconds 250;$reopened=Get-Snapshot 'manual-reopened'
    Assert-SameSnapshot $reopened $final 'Save/close/reopen';Assert-Relations $reopened 'Reopened';$persisted=$true
    $results.Add([pscustomobject]@{Name='SaveCloseReopen';Passed=$true})
}catch{$failure=$_}
finally{
    foreach($other in $otherDocuments){try{if([string]$other.Document.FullName -cne $other.Path -or [int]$other.Document.Saved -ne $other.Saved){$cleanupErrors.Add('Another open document changed state: '+$other.Name)}}catch{$cleanupErrors.Add('Cannot verify another open document: '+$other.Name)}}
    if($null -ne $filter){try{$filter.Dispose();$filterRestored=$true}catch{$cleanupErrors.Add('OLE filter restoration: '+$_.Exception.Message)}}
    Write-Evidence 'manual-result' ([pscustomobject]@{Passed=($null -eq $failure -and $cleanupErrors.Count -eq 0);Fixture=$fixture;Token=$token;
        PreflightPassed=$preflightPassed;HistoryStarted=$historyStarted;Persisted=$persisted;Operations=$operations;Checks=@($results.ToArray());
        Error=$(if($null -ne $failure){$failure.Exception.Message}else{$null});ErrorDetails=$(if($null -ne $failure){$failure.ToString()}else{$null});
        CleanupErrors=@($cleanupErrors.ToArray());OleMessageFilterRestored=$filterRestored;TraceStartByte=$manifest.Trace.StartByte;TraceEndByte=$traceEndByte;
        EvidenceLimits=@('The KeyboardAndPosition run permits single-frame arrow presses on pages 1, 3 and 5.',
            'Numeric logs prove the field and commit source; Ribbon/pane placement for width/height follows the page instructions.',
            'Pages10/11 share pointer-before-focus-loss; blank-canvas versus nonclient-window-border targeting relies on the user following those page instructions, not separate hit-test metadata.');
        Note='Actual PowerPoint UI traces from keyboard or pointer input. Undo/Redo through PowerPoint commands. No Replay APIs or geometry fixes. Failure preserves the current owned document for inspection.'})
}
if($null -ne $failure){throw $failure};if($cleanupErrors.Count -gt 0){throw ($cleanupErrors -join '; ')}
Write-Output ("PASS $($operations.Count) actual native operations, $($operations.Count) single Undo/Redo steps and save/reopen. Evidence: $directory")
