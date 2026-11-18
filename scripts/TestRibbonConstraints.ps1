[CmdletBinding()]
param([string]$OutputDirectory = (Join-Path $PSScriptRoot '..\artifacts\ribbon-ui'))
$ErrorActionPreference='Stop'
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
$directory=Join-Path $OutputDirectory ('run-'+[DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfff'))
New-Item -ItemType Directory -Force -Path $directory|Out-Null
$fixture=Join-Path $directory 'ribbon.pptx'
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$returnWindow=if($ppt.Presentations.Count -gt 0){$ppt.ActiveWindow}else{$null}
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
$p=$null;$passed=$false;$failure=$null;$cases=@()
function Assert-Owned {
    if($null -eq $p -or $p.FullName -ne $fixture -or $ppt.ActivePresentation.FullName -ne $fixture -or $p.Tags.Item('PPTGCS_TEST') -ne 'V3'){throw 'The active document is not the owned Ribbon fixture.'}
}
function Read-Tags($item) {
    @(for($i=1;$i -le $item.Tags.Count;$i++){[pscustomobject]@{Name=$item.Tags.Name($i);Value=$item.Tags.Value($i)}})|Sort-Object Name
}
function Read-State {
    Assert-Owned
    $xml=$addin.GetConstraintXml()
    Assert-Owned
    [pscustomobject]@{Xml=$xml;Tags=@(Read-Tags $p);Slides=@(foreach($s in $p.Slides){
        [pscustomobject]@{Id=$s.SlideID;Shapes=@(foreach($shape in $s.Shapes){[pscustomobject]@{
            Id=$shape.Id;Name=$shape.Name;Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height;Tags=@(Read-Tags $shape)
        }})}
    })}|ConvertTo-Json -Depth 25 -Compress
}
function Invoke-VerifiedHistory([string]$command,[string]$expected) {
    Assert-Owned
    if(-not $ppt.CommandBars.GetEnabledMso($command)){throw ($command+' is unavailable')}
    $ppt.CommandBars.ExecuteMso($command);Start-Sleep -Milliseconds 120
    if((Read-State) -cne $expected){throw ($command+' did not restore full geometry, identities and XML')}
}
try {
    $p=$ppt.Presentations.Add();$p.Tags.Add('PPTGCS_TEST','V3')
    $p.PageSetup.SlideWidth=960;$p.PageSetup.SlideHeight=540
    [xml]$ribbon=$addin.GetCustomUI('Microsoft.PowerPoint.Presentation')
    $buttons=@($ribbon.SelectNodes('//*[local-name()="button" and @tag]'))
    if($buttons.Count -ne 32){throw 'Expected 32 actual Ribbon constraint buttons'}
    $ribbon.Save((Join-Path $directory 'ribbon.xml'))
    for($i=0;$i -lt $buttons.Count;$i++){
        $slide=$p.Slides.Add($i+1,12);$slide.Name=$buttons[$i].GetAttribute('tag')
        for($j=0;$j -lt 3;$j++){$s=$slide.Shapes.AddShape(1,100+250*$j,150+80*$j,100+20*$j,70+10*$j);$s.Name='Member_'+$j;$s.TextFrame.AutoSize=0}
    }
    $p.SaveAs([string]$fixture,24);$p.Close();$p=$ppt.Presentations.Open([string]$fixture,0,0,-1)
    for($i=0;$i -lt $buttons.Count;$i++){
        $kind=$buttons[$i].GetAttribute('tag');$ppt.ActiveWindow.View.GotoSlide($i+1)
        $slide=$p.Slides.Item($i+1)
        $option=@(($addin.GetPaneState()|ConvertFrom-Json).Options|Where-Object Kind -eq $kind)[0]
        $names=@(for($j=0;$j -lt $option.MinimumMembers;$j++){'Member_'+$j})
        $container=$kind.StartsWith('Container',[StringComparison]::Ordinal)
        $slide.Shapes.Range([object[]]$names).Select()
        $option=@(($addin.GetPaneState()|ConvertFrom-Json).Options|Where-Object Kind -eq $kind)[0]
        $before=Read-State
        # Same dispatch routine as CreateConstraint's native IRibbonControl callback.
        $addin.BeginConstraint([string]$kind)
        $parameterized=-not [string]::IsNullOrEmpty($option.ValueLabel)
        if($parameterized){
            if((Read-State) -cne $before){throw ($kind+' wrote a relationship before parameter confirmation')}
            $diag=$addin.GetPaneDiagnostics()|ConvertFrom-Json
            if($diag.title -ne '约束' -or -not $diag.visible -or $diag.editorKind -ne $kind){throw ($kind+' did not open the matching parameter editor')}
            $values=@{value=[double]$option.DefaultValue}
            if(-not [string]::IsNullOrEmpty($option.SecondLabel)){$values.second=[double]$option.DefaultSecond}
            if($container){
                $reference='shape:{0}/{1}' -f $slide.SlideID,$slide.Shapes.Item($names[-1]).Id
                $null=$addin.AddContainerConstraintCommand([string]$kind,$values.value,$reference)
            }else{$null=$addin.AddConstraintCommand([string]$kind,($values|ConvertTo-Json -Compress))}
        }
        $snapshot=$addin.GetSnapshot()|ConvertFrom-Json;Assert-Owned
        $relations=@($snapshot.slides[$i].constraints)
        if($snapshot.presentation -ne $fixture -or $relations.Count -ne 1 -or $relations[0].Kind -ne $kind -or $relations[0].ShapeIds.Count -ne $names.Count){throw ($kind+' created an unexpected relationship')}
        $actualIds=@(foreach($name in $names){$slide.Shapes.Item($name).Tags.Item('PPTGCS_ID')})
        if(@($relations[0].ShapeIds|Where-Object {$_ -notin $actualIds}).Count){throw ($kind+' references objects outside the selection')}
        if($container -and $relations[0].ShapeIds[0] -cne $actualIds[-1]){throw ($kind+' did not preserve the designated container')}
        if($parameterized -and ([math]::Abs($relations[0].Value-$option.DefaultValue) -gt .00001)){throw ($kind+' did not retain its entered parameter')}
        $after=Read-State
        Invoke-VerifiedHistory 'Undo' $before
        Invoke-VerifiedHistory 'Redo' $after
        $cases += [pscustomobject]@{Kind=$kind;Label=$buttons[$i].GetAttribute('label');ParameterEditor=$parameterized;NoWriteBeforeConfirmation=$parameterized;CorrectMembers=$true;NativeUndo=$true;NativeRedo=$true;Passed=$true}
        Write-Output ('PASS '+$kind+' Ribbon dispatch and native Undo/Redo')
    }
    $p.Save();$passed=$true
}catch{$failure=$_.Exception.ToString();throw}
finally {
    if($null -ne $p){try{if($p.FullName -eq $fixture){$p.Save();$p.Close()}}catch{}}
    if($null -ne $returnWindow){try{$returnWindow.Activate()}catch{}}
    $report=[pscustomobject]@{Passed=$passed;PhysicalMouseInput=$false;ActualRibbonClicks=$false;RibbonControlTagsAndSharedDispatch=$true;CompletedCases=$cases.Count;Cases=$cases;Failure=$failure;Directory=$directory;Utc=[DateTime]::UtcNow.ToString('O')}
    $report|ConvertTo-Json -Depth 12|Set-Content (Join-Path $directory 'acceptance.json') -Encoding UTF8
    $report|ConvertTo-Json -Depth 12|Set-Content (Join-Path $OutputDirectory 'latest.json') -Encoding UTF8
}
