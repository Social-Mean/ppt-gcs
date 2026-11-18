#requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OfficeComSession.ps1')
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$nonce = [guid]::NewGuid().ToString('D').ToUpperInvariant()
$directory = Join-Path $repository ('artifacts/editing-entries-manual-' + $nonce.ToLowerInvariant())
$fixture = [IO.Path]::GetFullPath((Join-Path $directory 'EditingEntriesManual.pptx'))
$tag = 'PPTGCS_EDITING_ENTRIES_MANUAL'
$session = $null; $ppt = $null; $document = $null; $owned = $null; $window = $null; $addin = $null
$saved = $false; $ready = $false; $failure = $null
$cases = New-Object 'Collections.Generic.List[object]'

function Write-Json([string]$Name, $Value) {
    $Value | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath (Join-Path $directory ($Name + '.json')) -Encoding UTF8
}
function Assert-Owned([switch]$Active) {
    if ($null -eq $document -or -not (Test-OfficeComIdentity $document $owned)) { throw 'The newly created fixture reference is required.' }
    foreach ($original in $session.OriginalDocuments) {
        if (Test-OfficeComIdentity $document $original.Document) { throw 'Original documents cannot be modified by this fixture script.' }
    }
    if ([string]$document.Tags.Item('PPTGCS_TEST') -cne 'V3' -or [string]$document.Tags.Item($tag) -cne $nonce -or
        ($saved -and [IO.Path]::GetFullPath([string]$document.FullName) -ine $fixture)) { throw 'The exact owned fixture path and tags are required.' }
    if ($Active -and (-not (Test-OfficeComIdentity $ppt.ActivePresentation $document) -or
        -not (Test-OfficeComIdentity $ppt.ActiveWindow $window) -or [int]$window.ViewType -ne 9)) { throw 'Keep the owned fixture window active in normal view.' }
}
function New-Case([string]$Name, [string[]]$Steps, [string[]]$Selection, [string]$Entry, [string]$RequestedInput, [string[]]$Relations) {
    Assert-Owned -Active
    $page = $cases.Count + 1
    $slide = $document.Slides.Add($page, 12); $slide.Name = $Name
    $instructions = ('第 {0} 页｜{1}' -f $page, $Name.Substring(3)) + "`r`n" + ($Steps -join "`r`n")
    $hint = $slide.Shapes.AddTextbox(1, 25, 10, 910, 100); $hint.Name = 'Instructions'
    $hint.TextFrame.AutoSize = 0; $hint.TextFrame2.AutoSize = 0; $hint.TextFrame.WordWrap = -1
    $hint.TextFrame.MarginTop = 1; $hint.TextFrame.MarginBottom = 1; $hint.TextFrame.MarginLeft = 3; $hint.TextFrame.MarginRight = 3
    $hint.TextFrame.TextRange.Text = $instructions; $hint.TextFrame.TextRange.Font.Size = 14
    $hint.TextFrame.TextRange.Paragraphs(1).Font.Size = 18; $hint.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
    $cases.Add([pscustomobject]@{
        Page = $page; SlideId = [int]$slide.SlideID; Name = $Name; SelectionPaths = $Selection
        Entry = $Entry; RequestedInput = $RequestedInput; Relations = $Relations; ExpectedDataEdits = 1; Instructions = $instructions
    })
    return $slide
}
function Rect($Slide, [string]$Name, [double]$Left, [double]$Top, [double]$Width, [double]$Height) {
    Assert-Owned -Active
    $shape = $Slide.Shapes.AddShape(1, $Left, $Top, $Width, $Height); $shape.Name = $Name
    $shape.LockAspectRatio = 0; $shape.TextFrame.AutoSize = 0; $shape.TextFrame2.AutoSize = 0
    $shape.TextFrame.TextRange.Text = $Name; $shape.TextFrame.TextRange.Font.Size = 24
    $shape.Fill.ForeColor.RGB = $(if ($Name -in @('B', 'C')) { 0x8B6542 } elseif ($Name -eq 'D') { 0xBEBEBE } else { 0xD88928 })
    return $shape
}
function Zero-Line($Slide, [double]$Left, [double]$Top, [double]$Bottom) {
    Assert-Owned -Active
    $line = $Slide.Shapes.AddLine($Left, $Top, $Left, $Bottom); $line.Name = 'ZeroWidthLine'; $line.Line.Weight = 3
    if ([math]::Abs([double]$line.Width) -gt .0001) { throw 'The native free line must have zero width.' }
}
function Relate($Slide, [string[]]$Names, [string]$Kind) {
    Assert-Owned -Active
    $window.View.GotoSlide([int]$Slide.SlideIndex)
    for ($paneIndex = 1; $paneIndex -le $window.Panes.Count; $paneIndex++) {
        if ([int]$window.Panes.Item($paneIndex).ViewType -eq 1) { $window.Panes.Item($paneIndex).Activate(); break }
    }
    $window.Selection.Unselect(); $Slide.Shapes.Range([object[]]$Names).Select()
    Assert-Owned -Active
    $id = [string]$addin.AddConstraintCommand($Kind, '{}')
    if ([string]::IsNullOrEmpty($id)) { throw ('Cannot create fixture constraint ' + $Kind) }
}
function Group-Page($Slide) {
    $null = Rect $Slide 'A' 140 180 130 70; $null = Rect $Slide 'B' 320 350 110 60
    $window.View.GotoSlide([int]$Slide.SlideIndex)
    $group = $Slide.Shapes.Range([object[]]@('A', 'B')).Group(); $group.Name = 'AB'
    $null = Rect $Slide 'C' 600 $group.Top $group.Width $group.Height
    foreach ($kind in @('SameY', 'SameWidth', 'SameHeight')) { Relate $Slide @('AB', 'C') $kind }
}
function Tags($Owner) {
    $values = @{}
    for ($index = 1; $index -le $Owner.Tags.Count; $index++) { $values[[string]$Owner.Tags.Name($index)] = [string]$Owner.Tags.Value($index) }
    $ordered = [ordered]@{}
    foreach ($name in @($values.Keys | Sort-Object)) { $ordered[$name] = $values[$name] }
    return $ordered
}
function Baseline {
    Assert-Owned -Active
    $rows = New-Object 'Collections.Generic.List[object]'; $slides = New-Object 'Collections.Generic.List[object]'
    function Read-Shape($Shape, [string]$Key, [int]$ParentId) {
        $text = $null; if ([int]$Shape.HasTextFrame -ne 0) { $text = [string]$Shape.TextFrame.TextRange.Text }
        $rows.Add([pscustomobject]@{
            Key = $Key; NativeId = [int]$Shape.Id; ParentNativeId = $ParentId; Name = [string]$Shape.Name
            Type = [int]$Shape.Type; Child = [int]$Shape.Child; Tags = Tags $Shape; Text = $text
            Left = [double]$Shape.Left; Top = [double]$Shape.Top; Width = [double]$Shape.Width; Height = [double]$Shape.Height
            Rotation = [double]$Shape.Rotation; LockAspectRatio = [int]$Shape.LockAspectRatio
        })
        if ([int]$Shape.Type -eq 6) {
            for ($index = 1; $index -le $Shape.GroupItems.Count; $index++) {
                $child = $Shape.GroupItems.Item($index); Read-Shape $child ($Key + '/' + [string]$child.Name) ([int]$Shape.Id)
            }
        }
    }
    for ($page = 1; $page -le $document.Slides.Count; $page++) {
        $slide = $document.Slides.Item($page)
        $slides.Add([pscustomobject]@{ Index = $page; Id = [int]$slide.SlideID; Name = [string]$slide.Name; Tags = Tags $slide })
        for ($index = 1; $index -le $slide.Shapes.Count; $index++) {
            $shape = $slide.Shapes.Item($index); Read-Shape $shape ($page.ToString() + '/' + [string]$shape.Name) 0
        }
    }
    Assert-Owned -Active
    $selection = $window.Selection; $ids = @(); $childIds = @()
    if ([int]$selection.Type -in @(2, 3)) { $ids = @(for ($index = 1; $index -le $selection.ShapeRange.Count; $index++) { [int]$selection.ShapeRange.Item($index).Id }) }
    if ([bool]$selection.HasChildShapeRange) { $childIds = @(for ($index = 1; $index -le $selection.ChildShapeRange.Count; $index++) { [int]$selection.ChildShapeRange.Item($index).Id }) }
    return [pscustomobject]@{
        Utc = [datetime]::UtcNow.ToString('O'); ConstraintXml = [string]$addin.GetConstraintXml(); DocumentTags = Tags $document
        SlideWidth = [double]$document.PageSetup.SlideWidth; SlideHeight = [double]$document.PageSetup.SlideHeight
        SnapToGrid = [int]$document.SnapToGrid; GridDistance = [double]$document.GridDistance
        Slides = @($slides.ToArray()); Rows = @($rows.ToArray()); ActiveSlideId = [int]$window.View.Slide.SlideID; ActiveSlideIndex = [int]$window.View.Slide.SlideIndex
        Selection = [pscustomobject]@{ Type = [int]$selection.Type; HasChild = [bool]$selection.HasChildShapeRange; NativeIds = $ids; ChildNativeIds = $childIds }
    }
}

try {
    New-Item -ItemType Directory -Path $directory | Out-Null
    $session = Start-OfficeComSession; $ppt = $session.Application
    $originals = @($session.OriginalDocuments | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path; Saved = $_.Saved } })
    Write-Json 'original-documents' $originals
    $addin = $ppt.COMAddIns.Item('PptGcs.Connect').Object
    if ($null -eq $addin) { throw 'The installed production PptGcs.Connect add-in must be loaded.' }
    $document = $ppt.Presentations.Add(-1); $owned = $document; $window = $ppt.ActiveWindow
    $document.Tags.Add('PPTGCS_TEST', 'V3'); $document.Tags.Add($tag, $nonce)
    Assert-Owned -Active
    $document.PageSetup.SlideWidth = 960; $document.PageSetup.SlideHeight = 540; $document.SnapToGrid = 0
    $window.WindowState = 1 # Only this new document window: make its border resizable before any typed field edit.

    $slide = New-Case '01_单对象键盘移动' @('单击 A 的边框，确认没有文字光标。按住 ↑ 约 1 秒，然后松开。', 'B 应与 A 保持顶端对齐，并在按住期间实时跟随。', '本页只做这一次；切页后按下一页说明重新选择，先不要撤销、重做或重复操作。') @('A') 'keyboard' 'Up hold ~1s' @('SameY')
    $null = Rect $slide 'A' 160 220 150 100; $null = Rect $slide 'B' 550 220 150 100; Relate $slide @('A', 'B') 'SameY'
    Assert-Owned -Active
    if (Test-Path -LiteralPath $fixture) { throw 'The unique fixture path must not already exist.' }
    $document.SaveAs($fixture, 24); $saved = $true

    $slide = New-Case '02_多选鼠标整体移动' @('Ctrl 单击 A、D 和竖线，完成三选；然后松开 Ctrl。', '从 A 内部空白处向上慢拖约 1 秒后松手，避开文字和缩放点。', 'A、D、竖线应整体平移，B 实时跟随 A 保持顶端对齐；本页只拖一次，先不要撤销。') @('A', 'D', 'ZeroWidthLine') 'pointer-selection' 'Up drag ~1s' @('SameY')
    $null = Rect $slide 'A' 140 210 140 90; $null = Rect $slide 'B' 540 210 140 90; $null = Rect $slide 'D' 320 370 100 70
    Zero-Line $slide 465 355 465; Relate $slide @('A', 'B') 'SameY'

    $slide = New-Case '03_多选键盘整体移动' @('Ctrl 单击 A、D 和竖线，完成三选，确认没有文字光标。', '保持 Ctrl，按住 → 约 1 秒；先松开 →，再松开 Ctrl。', '三个选中对象应整体右移，B 实时跟随 A 保持左对齐；只做一次，先不要撤销。') @('A', 'D', 'ZeroWidthLine') 'keyboard-selection' 'Ctrl+Right hold ~1s' @('SameX')
    $null = Rect $slide 'A' 160 210 140 90; $null = Rect $slide 'B' 160 390 140 80; $null = Rect $slide 'D' 415 220 95 70
    Zero-Line $slide 570 195 305; Relate $slide @('A', 'B') 'SameX'

    $slide = New-Case '04_组内成员鼠标移动' @('先单击 AB 组合，再单击组内 A 的边框；确认只选中 A，且没有文字光标。', '从 A 内部空白处向上慢拖约 1 秒后松手，避开文字和缩放点。', 'B 应留在原位，外部 C 跟随 AB 的顶边、宽度、高度；只拖一次，先不要撤销。') @('AB/A') 'pointer-child' 'Child A up drag ~1s' @('SameY', 'SameWidth', 'SameHeight')
    Group-Page $slide
    $slide = New-Case '05_组内成员键盘移动' @('先单击 AB 组合，再单击组内 A 的边框；确认只选中 A，且没有文字光标。', '按住 ↑ 约 1 秒，然后松开；本页只做这一次。', 'B 应留在原位，外部 C 跟随 AB 的顶边、宽度、高度；先不要撤销。') @('AB/A') 'keyboard-child' 'Child A Up hold ~1s' @('SameY', 'SameWidth', 'SameHeight')
    Group-Page $slide

    $slide = New-Case '06_功能区宽度回车提交' @('单击 A 的边框，打开“形状格式”功能区，点击“宽度”数值框。', 'Ctrl+A 后直接输入 6 cm，按一次 Enter；不要使用旁边的微调箭头。', 'A、B 应同宽，其他尺寸不变；本页只提交一次，先不要撤销。') @('A') 'ribbon-width-enter' '6 cm; Enter' @('SameWidth')
    $null = Rect $slide 'A' 140 220 130 85; $null = Rect $slide 'B' 540 220 130 115; Relate $slide @('A', 'B') 'SameWidth'
    $slide = New-Case '07_功能区高度制表提交' @('单击 A 的边框，打开“形状格式”功能区，点击“高度”数值框。', 'Ctrl+A 后直接输入 4 cm，按一次 Tab；不要使用旁边的微调箭头。', 'A、B 应同高，宽度各自保持不变；本页只提交一次，先不要撤销。') @('A') 'ribbon-height-tab' '4 cm; Tab' @('SameHeight')
    $null = Rect $slide 'A' 140 220 150 70; $null = Rect $slide 'B' 540 220 190 70; Relate $slide @('A', 'B') 'SameHeight'
    $slide = New-Case '08_原生窗格水平位置回车' @('单击 A 的边框；右键打开“设置形状格式”，展开“大小与属性 → 位置”。', '保持“从”的当前选项，不切换原点；在“水平位置”框 Ctrl+A 输入 6 cm，按一次 Enter。', 'A、B 应保持左对齐；本页只提交这一次，先不要撤销。') @('A') 'pane-left-enter' '6 cm; Enter; origin unchanged' @('SameX')
    $null = Rect $slide 'A' 110 200 160 80; $null = Rect $slide 'B' 110 375 120 70; Relate $slide @('A', 'B') 'SameX'
    $slide = New-Case '09_原生窗格垂直位置制表' @('单击 A 的边框；打开“设置形状格式”，展开“大小与属性 → 位置”。', '保持“从”的当前选项，不切换原点；在“垂直位置”框 Ctrl+A 输入 4 cm，按一次 Tab。', 'A、B 应保持顶端对齐；本页只提交这一次，先不要撤销。') @('A') 'pane-top-tab' '4 cm; Tab; origin unchanged' @('SameY')
    $null = Rect $slide 'A' 150 155 160 80; $null = Rect $slide 'B' 550 155 120 80; Relate $slide @('A', 'B') 'SameY'
    $slide = New-Case '10_宽度点击空白提交' @('单击 A 的边框，在“形状格式”功能区点击“宽度”数值框。', 'Ctrl+A 输入 7 cm，不按 Enter/Tab；直接单击幻灯片空白处一次以提交。', 'A、B 应同宽，其他尺寸不变；本页只提交一次，先不要撤销。') @('A') 'ribbon-width-click-away' '7 cm; blank-canvas click' @('SameWidth')
    $null = Rect $slide 'A' 140 220 145 90; $null = Rect $slide 'B' 540 220 145 110; Relate $slide @('A', 'B') 'SameWidth'
    $slide = New-Case '11_高度窗口边框失焦提交' @('输入前先确认 PPT 窗口不是最大化，边框可拖动；必要时先还原窗口。', '重新单击 A 边框；在“形状格式”功能区“高度”框 Ctrl+A 输入 4.5 cm。', '不按 Enter/Tab，不点其他区域；直接拖动当前 PPT 窗口右边或下边框，小幅缩放一次后松开。', 'A、B 应同高；只做这一次，再切到第 12 页，先不要撤销。') @('A') 'ribbon-height-window-border' '4.5 cm; resize own window border' @('SameHeight')
    $null = Rect $slide 'A' 140 220 160 82; $null = Rect $slide 'B' 540 220 130 82; Relate $slide @('A', 'B') 'SameHeight'
    $slide = New-Case '12_混合多选宽度裸数值提交' @('Ctrl 单击 A 和 B，完成两选后松开 Ctrl；C 不要选。B 是自由对象，A 与 C 约束同宽。', '在“形状格式”功能区点击“宽度”框；因 A/B 宽度不同，框内应为空。', 'Ctrl+A 后直接输入裸数值 6（不带单位），按一次 Enter；A、B、C 应一起变为 6 厘米宽。', '本页只提交一次；完成后停止操作，等待统一读取，再统一撤销 12 次、重做 12 次。') @('A', 'B') 'ribbon-width-mixed-enter' '6 (bare, current cm UI); Enter' @('SameWidth')
    $null = Rect $slide 'A' 120 220 120 80; $null = Rect $slide 'B' 400 220 160 110; $null = Rect $slide 'C' 700 220 120 100
    Relate $slide @('A', 'C') 'SameWidth'

    Assert-Owned -Active
    $window.View.GotoSlide(1)
    for ($index = 1; $index -le $window.Panes.Count; $index++) { if ([int]$window.Panes.Item($index).ViewType -eq 1) { $window.Panes.Item($index).Activate(); break } }
    $window.Selection.Unselect(); $document.Slides.Item(1).Shapes.Item('A').Select()
    $document.Save()
    $before = Baseline; Write-Json 'beforeFrames' $before
    foreach ($case in $cases) {
        $selected = @(); $driverIds = @()
        foreach ($path in $case.SelectionPaths) {
            $row = $before.Rows | Where-Object Key -ceq ($case.Page.ToString() + '/' + $path)
            $rootPath = $case.Page.ToString() + '/' + ($path -split '/')[0]
            $rootRow = $before.Rows | Where-Object Key -ceq $rootPath
            if ($null -eq $row -or $null -eq $rootRow) { throw 'The manual selection path is missing from the saved baseline.' }
            $selected += [int]$row.NativeId
            if ($rootRow.Tags['PPTGCS_ID']) { $driverIds += [string]$rootRow.Tags['PPTGCS_ID'] }
        }
        $case | Add-Member NoteProperty ExpectedSelectedNativeIds @($selected)
        $case | Add-Member NoteProperty DriverIds @($driverIds | Select-Object -Unique)
    }
    $logPath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PptGcs/layout-v2.log'
    $logStart = if (Test-Path -LiteralPath $logPath) { (Get-Item -LiteralPath $logPath).Length } else { 0 }
    $originalAfter = @(foreach ($item in $session.OriginalDocuments) { [pscustomobject]@{ Name = [string]$item.Document.Name; Path = [string]$item.Document.FullName; Saved = [int]$item.Document.Saved } })
    Write-Json 'fixture' ([pscustomobject]@{
        Path = $fixture; Token = $nonce; TestTag = $tag; ReadyUtc = [datetime]::UtcNow.ToString('O'); Baseline = 'beforeFrames.json'
        ExpectedDataEdits = 12; ExpectedNativeUndoCount = 12; ExpectedNativeRedoCount = 12; AcceptanceStatus = 'NotPerformed'
        InitialWindowResizable = ([int]$window.WindowState -eq 1); PerPageAutoSelection = $false; Cases = @($cases.ToArray())
        OriginalDocuments = $originals; OriginalDocumentsAfterCreation = $originalAfter
        Unsupported = @('native spinner clicks', 'coordinate origin combo changes', 'AutoSize text-driven geometry')
        Trace = [pscustomobject]@{
            LogPath = $logPath; StartByte = $logStart
            PointerFrames = 'gesture-frame'; PointerEnd = 'gesture-end'; PointerRequirement = 'gesture-end.test=false and frames>=2; join by gestureId; require a move frame with mouseDown=true and input=owned-gesture-loop; match driverId/childNativeIds/selectedNativeIds'
            KeyboardFrames = 'keyboard-nudge-frame'; KeyboardEnd = 'keyboard-nudge-end'; KeyboardRequirement = 'keyboard-nudge-end.test=false and frames>=2 for the held key; join by id; driverIds are in keyboard-nudge-start, selectedNativeIds identify the selected members'
            NumericCommit = 'native-field-committed'; NumericSources = @('enter', 'tab', 'pointer-before-focus-loss')
            Note = 'Fixture creation only. No Replay API, synthetic input, event observer or completed manual acceptance is claimed.'
        }
        HistoryPlan = 'Record the final whole document after one data edit per page; compare one native Undo per page in reverse order, then one Redo per page in forward order. Selection is recorded, not treated as persistent history.'
    })
    # Separate setup from the first manual edit. No passive page-change handler
    # or automatic per-page mutation is installed by this script.
    Assert-Owned -Active; $ppt.StartNewUndoEntry(); $ready = $true
    Write-Output ('已打开 12 页手动验收文稿：' + $fixture)
    Write-Output ('逐页只做一次说明中的数据编辑；完成前先不要撤销或重做。基线与清单：' + $directory)
}
catch { $failure = $_ }
finally {
    try {
        if (-not $ready -and $null -ne $document) {
            if (-not (Test-OfficeComIdentity $document $owned)) { throw 'Refusing to clean up a non-owned document.' }
            foreach ($original in $session.OriginalDocuments) { if (Test-OfficeComIdentity $document $original.Document) { throw 'Refusing to close an original document.' } }
            if ($saved) { Assert-Owned }
            $document.Saved = -1; $document.Close()
            if ($null -ne $session.OriginalWindow) { $session.OriginalWindow.Activate() }
        }
    }
    finally { if ($null -ne $session) { Stop-OfficeComSession $session } }
}
if ($null -ne $failure) { throw $failure }
