$ErrorActionPreference='Stop'
$directory=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\artifacts\native-history-probe'))
New-Item -ItemType Directory -Force -Path $directory | Out-Null
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$userWindow=$ppt.ActiveWindow
$probe=$null
$cases=@('Name','AlternativeText','Title','Text','ShapeTag','SlideTag','PresentationTag','CustomXml')
$results=@()
try {
 foreach($case in $cases) {
  $probe=$ppt.Presentations.Add()
  $slide=$probe.Slides.Add(1,12)
  $shape=$slide.Shapes.AddShape(1,30,30,100,60)
  $shape.Name='MetadataProbe'
  $shape.AlternativeText='baseline-alt'
  $shape.Title='baseline-title'
  $shape.TextFrame.TextRange.Text='baseline-text'
  $shape.Tags.Add('HISTORY_PROBE','baseline-tag')
  $slide.Tags.Add('HISTORY_PROBE','baseline-slide')
  $probe.Tags.Add('HISTORY_PROBE','baseline-presentation')
  $path=Join-Path $directory ($case+'.pptx')
  $probe.SaveAs([string]$path,24)
  Start-Sleep -Milliseconds 250
  $probe.Close()
  $probe=$ppt.Presentations.Open([string]$path,0,0,-1)
  $slide=$probe.Slides.Item(1)
  $shape=$slide.Shapes.Item(1)
  function Read-Value {
   switch($case) {
    'Name' {$shape.Name}
    'AlternativeText' {$shape.AlternativeText}
    'Title' {$shape.Title}
    'Text' {$shape.TextFrame.TextRange.Text}
    'ShapeTag' {$shape.Tags.Item('HISTORY_PROBE')}
    'SlideTag' {$slide.Tags.Item('HISTORY_PROBE')}
    'PresentationTag' {$probe.Tags.Item('HISTORY_PROBE')}
    'CustomXml' {$probe.CustomXMLParts.SelectByNamespace('urn:history-probe').Count}
   }
  }
  $before=Read-Value
  $ppt.StartNewUndoEntry()
  switch($case) {
   'Name' {$shape.Name='ChangedProbe'}
   'AlternativeText' {$shape.AlternativeText='changed-alt'}
   'Title' {$shape.Title='changed-title'}
   'Text' {$shape.TextFrame.TextRange.Text='changed-text'}
   'ShapeTag' {$shape.Tags.Add('HISTORY_PROBE','changed-tag')}
   'SlideTag' {$slide.Tags.Add('HISTORY_PROBE','changed-slide')}
   'PresentationTag' {$probe.Tags.Add('HISTORY_PROBE','changed-presentation')}
   'CustomXml' {$null=$probe.CustomXMLParts.Add('<metadata xmlns="urn:history-probe">changed</metadata>')}
  }
  $ppt.StartNewUndoEntry()
  Start-Sleep -Milliseconds 100
  $after=Read-Value
  $undoEnabled=$ppt.CommandBars.GetEnabledMso('Undo')
  if($undoEnabled) {$ppt.CommandBars.ExecuteMso('Undo');Start-Sleep -Milliseconds 100}
  $undone=Read-Value
  $redoEnabled=$ppt.CommandBars.GetEnabledMso('Redo')
  if($redoEnabled) {$ppt.CommandBars.ExecuteMso('Redo');Start-Sleep -Milliseconds 100}
  $redone=Read-Value
  $result=[pscustomobject]@{Property=$case;Before=$before;After=$after;UndoEnabled=$undoEnabled;Undone=$undone;RedoEnabled=$redoEnabled;Redone=$redone;Pass=($before -eq $undone -and $after -eq $redone -and $before -ne $after)}
  $results+=$result
  $result | ConvertTo-Json -Compress
  $probe.Save()
  $probe.Close(); $probe=$null
 }
} finally {
 if($null -ne $probe){$probe.Save();$probe.Close()}
 $userWindow.Activate()
 $results | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'property-results.json') -Encoding UTF8
}
