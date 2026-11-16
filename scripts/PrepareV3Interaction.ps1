$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root 'artifacts\v3-interaction'
New-Item -ItemType Directory -Force -Path $directory | Out-Null
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
if($ppt.Presentations.Count -ne 1 -or $ppt.ActivePresentation.Tags.Item('PPTGCS_TEST') -ne 'V3'){throw 'Expected one owned V3 test presentation'}
$old=$ppt.ActivePresentation
if($old.Saved -eq 0){$old.Save()}
$old.Close()
$p=$ppt.Presentations.Add()
$p.Tags.Add('PPTGCS_TEST','V3')
$p.PageSetup.SlideWidth=960
$p.PageSetup.SlideHeight=540
function Text($slide,$name,$text,$y,$size) {
    $shape=$slide.Shapes.AddTextbox(1,30,$y,900,52)
    $shape.Name=$name;$shape.TextFrame.TextRange.Text=$text;$shape.TextFrame.TextRange.Font.Size=$size
}
function Rectangle($slide,$name,$letter,$x,$y,$w,$h,$color) {
    $shape=$slide.Shapes.AddShape(1,$x,$y,$w,$h)
    $shape.Name=$name;$shape.LockAspectRatio=0;$shape.TextFrame.AutoSize=0
    $shape.Fill.ForeColor.RGB=$color
    $shape.TextFrame.TextRange.Text=$letter;$shape.TextFrame.TextRange.Font.Size=30
    $shape.TextFrame.TextRange.Font.Color.RGB=16777215
}
$slide=$p.Slides.Add(1,12);$slide.Name='Wheel'
Text $slide '说明标题' '1  中键轮盘：建立同宽' 20 26
Text $slide '操作说明' 'A、B 已选中：在画布空白处按住中键，移向右侧“同宽”，再松开。' 65 18
Rectangle $slide 'A' 'A' 160 220 160 100 14125352
Rectangle $slide 'B' 'B' 540 260 220 140 9135426
$slide=$p.Slides.Add(2,12);$slide.Name='Pane'
Text $slide '说明标题' '2  约束窗格：参数与宽高比' 20 26
Text $slide '操作说明' '选中 A，在右侧新增“固定宽高比”，输入 2；创建后拉动右边中点，再修改比例。' 65 18
Rectangle $slide 'A' 'A' 290 220 200 120 14125352
$slide=$p.Slides.Add(3,12);$slide.Name='Gap'
Text $slide '说明标题' '3  中键轮盘：水平等间距' 20 26
Text $slide '操作说明' '选中 A、B、C，用轮盘“水平等距”建立关系；左右拖动 B，检查间距联动。' 65 18
Rectangle $slide 'A' 'A' 110 250 120 90 14125352
Rectangle $slide 'B' 'B' 380 230 160 110 9135426
Rectangle $slide 'C' 'C' 720 260 100 80 6536550
$path=[string](Join-Path $directory 'interaction.pptx')
$p.SaveAs($path,24)
$p.Close()
$p=$ppt.Presentations.Open($path,0,0,-1)
$ppt.ActiveWindow.View.GotoSlide(1)
$ppt.ActiveWindow.View.Zoom=65
$p.Slides.Item(1).Shapes.Range(@('A','B')).Select()
$null=$ppt.COMAddIns.Item('PptGcs.Connect').Object.OpenConstraintPane()
& (Join-Path $PSScriptRoot 'CaptureV3Interaction.ps1') -Name 'before'
$ppt.COMAddIns.Item('PptGcs.Connect').Object.GetPaneDiagnostics()
