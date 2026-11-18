# 组件复制和命名参数验证

2026-11-18 在 Windows 桌面 PowerPoint 上验收。同页复制使用 PowerPoint 的形状复制方法，跨页使用形状复制和目标页粘贴，沿用 Office 对实际对象格式和剪贴板的处理。文稿中的命名参数使用磅（pt），关系只保存参数引用编号。副本的参数可独立，也可由用户显式选择共享。[ShapeRange.Copy](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.copy) · [Shapes.Paste](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.paste)

## 自动检查与 PowerPoint 命令

`Build.ps1` 通过 562 项自动检查。新增核心检查覆盖命名参数解析和引用完整性、更新后的求解值、复制时顺序及停用状态、参数去重复制和共享。窗格检查覆盖默认独立选择、目标页、明确共享、参数选择、刷新保留编辑及禁止删除仍被引用的参数。

安装构建产物并打开 PowerPoint 后，可运行真实宿主回归。脚本创建并识别自己的文稿，默认结果写入 `artifacts/component-validation/`。测试使用 PowerPoint 剪贴板进行跨页复制。

```powershell
powershell.exe -NoProfile -STA -File .\scripts\TestComponentCommands.ps1
```

- 复制矩形、椭圆和未绑定形状到另一页，检查可编辑文本、全新形状编号和关系编号。已停用内部关系仍停用，指向未复制对象的对齐关系留在原页。独立副本的宽度和间距引用同一个新参数。同页复制也通过单步撤销及重做。
- 显式共享的副本继续绑定原参数。原参数从 84 改到 110 磅时，源页和共享页的宽度、间距同时更新，独立页保持 84 磅。复制和更新各自通过一步原生 Undo、Redo，且核对了原件和副本的几何与身份。
- 在共享页增加固定宽度及同宽关系，使参数变更发生冲突。命令拒绝提交，所有页面、参数库和历史的当前位置保持原状态。仍被引用的参数也无法删除。保存后重开，关系 XML、身份与几何一致。
- 使用同一构建重跑幻灯片与容器验收，12 类关系、跨页重排、原生历史和保存重开继续通过。[参考框约束验证](validation-reference-constraints.md)

## 实际界面操作

利用 Windows 原生窗口截图定位控件，向实际 PowerPoint 窗口发送鼠标和键盘输入。先从第 1 页用功能区和窗格复制到第 4 页，保留默认独立设置。再通过命名参数编辑器把原参数从 110 改为 126 磅，四页的宽度和间距分别读回 126、84、126、110 磅。在第 4 页使用 **参数来源**，把椭圆的固定高度绑定到该页的独立参数，读回 110 磅。三次操作都实际点击快速访问栏撤销和重做，并比较全篇快照。

保存后退出 PowerPoint，在新进程中打开同一个文件。核对 XML、编号、文本和几何。通过加载项命令把第 4 页的独立参数改为 118 磅，只有该页的宽、间距及新绑定的高一起变为 118 磅，一次原生撤销恢复初始状态。

本机实际状态和截图位于 `%LOCALAPPDATA%/PptGcs/Sessions/components-20261118/30F62D61-67F9-44F3-AA67-4E4463D07B89/`。`acceptance.json` 汇总命令验收，`ui-result.json` 汇总窗口验收，`copy-target.png` 和 `binding-choice.png` 记录目标页与参数选择。
