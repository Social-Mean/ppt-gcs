$ErrorActionPreference='Stop'
$directory=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\artifacts\native-history-probe'))
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$userWindow=$ppt.ActiveWindow
$probe=$null
$checks=@()
function To-Hex([string]$value){[BitConverter]::ToString([Text.Encoding]::UTF8.GetBytes($value)).Replace('-','')}
try {
 $probe=$ppt.Presentations.Add()
 $slide=$probe.Slides.Add(1,12)
 $shape=$slide.Shapes.AddShape(1,30,30,100,60)
 $shape.Name='ConstraintGeometry'
 $path=Join-Path $directory 'PresentationTagTransactions.pptx'
 $probe.SaveAs([string]$path,24)
 Start-Sleep -Milliseconds 250
 $probe.Close()
 $probe=$ppt.Presentations.Open([string]$path,0,0,-1)
 $shape=$probe.Slides.Item(1).Shapes.Item(1)
 function Read-State {
  [pscustomobject]@{Data=$probe.Tags.Item('PPTGCS_PROBE_DATA');Token=$probe.Tags.Item('PPTGCS_PROBE_TOKEN');ShapeTag=$shape.Tags.Item('PROBE_GUID');Width=$shape.Width;Height=$shape.Height;Left=$shape.Left;Top=$shape.Top}
 }
 function Check-State($expected,[string]$label) {
  $actual=Read-State
  if(($actual | ConvertTo-Json -Compress) -ne ($expected | ConvertTo-Json -Compress)){throw "History mismatch: $label"}
  $script:checks += $label
  Write-Output "PASS $label"
 }
 $states=@((Read-State))
 $data=@('<constraints><constraint Kind="SameWidth" Label="中文混合Case" /></constraints>', '<constraints><constraint Kind="FixedWidth" Value="180" /></constraints>', '<constraints><constraint Kind="FixedWidth" Value="180" Enabled="false" /></constraints>')
 $ppt.StartNewUndoEntry()
 $probe.Tags.Add('PPTGCS_PROBE_DATA',(To-Hex $data[0]))
 $probe.Tags.Add('PPTGCS_PROBE_TOKEN','TEST-TOKEN')
 $shape.Tags.Add('PROBE_GUID','TEST-GUID')
 $shape.Width=140; $shape.Height=70
 $ppt.StartNewUndoEntry()
 $states+=Read-State
 $ppt.StartNewUndoEntry()
 $probe.Tags.Add('PPTGCS_PROBE_DATA',(To-Hex $data[1]))
 $shape.Width=180
 $ppt.StartNewUndoEntry()
 $states+=Read-State
 $ppt.StartNewUndoEntry()
 $probe.Tags.Add('PPTGCS_PROBE_DATA',(To-Hex $data[2]))
 $ppt.StartNewUndoEntry()
 $states+=Read-State
 $ppt.StartNewUndoEntry()
 $probe.Tags.Delete('PPTGCS_PROBE_DATA')
 $ppt.StartNewUndoEntry()
 $states+=Read-State
 for($step=1;$step -le 4;$step++) {
  if(-not $ppt.CommandBars.GetEnabledMso('Undo')){throw "Undo disabled at step $step"}
  $ppt.CommandBars.ExecuteMso('Undo');Start-Sleep -Milliseconds 100
  Check-State $states[4-$step] "native Undo $step/4"
 }
 for($step=1;$step -le 4;$step++) {
  if(-not $ppt.CommandBars.GetEnabledMso('Redo')){throw "Redo disabled at step $step"}
  $ppt.CommandBars.ExecuteMso('Redo');Start-Sleep -Milliseconds 100
  Check-State $states[$step] "native Redo $step/4"
 }
 $sizeChecks=@()
 foreach($size in @(1024,16384,65536,1048576)) {
  $text=('中文Case-'+('abcdefXYZ123'*[Math]::Ceiling($size/12.0))).Substring(0,$size)
  $hex=To-Hex $text
  $probe.Tags.Add('PPTGCS_PROBE_LARGE',$hex)
  $actual=$probe.Tags.Item('PPTGCS_PROBE_LARGE')
  if($actual -cne $hex){throw "Tag truncation/case mutation at $size chars"}
  $sizeChecks += [pscustomobject]@{TextChars=$size;HexLength=$hex.Length;Exact=$true}
  Write-Output "PASS exact large tag: $size chars / $($hex.Length) hex chars"
 }
 $beforeSave=Read-State
 $probe.Save()
 Start-Sleep -Milliseconds 250
 $probe.Close()
 $probe=$ppt.Presentations.Open([string]$path,0,0,-1)
 $shape=$probe.Slides.Item(1).Shapes.Item(1)
 Check-State $beforeSave 'save/reopen transaction data and geometry'
 if($probe.Tags.Item('PPTGCS_PROBE_LARGE') -cne $hex){throw 'Large tag changed after save/reopen'}
 $checks += 'largest tag save/reopen'
 [pscustomobject]@{Passed=$true;Checks=$checks;Sizes=$sizeChecks;States=$states;ProcessId=(Get-Process POWERPNT).Id} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'presentation-tag-transactions.json') -Encoding UTF8
 $probe.Close();$probe=$null
} finally {
 if($null -ne $probe){$probe.Save();$probe.Close()}
 $userWindow.Activate()
}
