# v0.4 更多形状与 Office 内置图标验证

日期：2026-11-16。环境为 Windows、64 位桌面 PowerPoint 16.0、.NET Framework 4.8，以及当前用户注册的 PPT GCS COM 加载项。

## 结论与证据边界

本轮程序化检查通过：129 项离线检查、25 个命令的内置图标加载、11 类代表对象的 44 次双向轨迹及逐次原生 Undo/Redo、非矩形复制后新建关系，以及同进程和新进程保存重开。

**新增异形对象的真实鼠标操作和界面视觉效果尚未通过人工验收。** 本轮形状报告保持 `PhysicalMouseInput=false`，不将程序化轨迹算作物理拖动。桌面截图通道连续返回 `FrameArrived` / `window capture timed out`，因此不能用本轮记录声称已目视确认每个内置图标、异形拖动或组合命中。历史 [v0.3 真实鼠标验收](validation-v3.md) 覆盖当时的矩形、轮盘和窗格路径，不替代本轮新增范围。

## 形状能力与布局含义

统一的 `ShapeCapabilities` 用于命令校验、窗格模型及鼠标接管，支持同页未旋转的顶层对象：

- 自选图形，包括矩形、椭圆、圆角矩形、三角形、菱形和箭头等。
- 文本框、图片、链接图片、SVG 图形和链接图形。
- 自由曲线，以及作为一个对象操作的组合整体。

这些对象共用现有 20 类布局约束，求解依据是 Office 的 Left、Top、Width、Height 外接框。不同类型可互相建立关系；同宽统一外接框宽度，间距为相邻外接框边缘的有符号距离。轴对称只使两个等尺寸外接框关于固定轴对称，不翻转图片、箭头或曲线内容。

对象宽高至少为 1 pt。旋转、线条、连接符、表格、图表、SmartArt 和占位符不在本轮支持范围。组合内部编辑时保留原生操作，不能给子成员直接建立外部布局约束；组内命中按父级归属提升到顶层组合。黄色调节点、路径和顶点编辑不作为布局缩放。

文本框保留原生文本编辑入口，文本光标状态和双击不被布局手势接管。写回不临时切换并恢复文本自动适应，因为重新启用自动适应会触发再次排版；宽高和坐标按 Office 即时读回补写，再检查完整目标几何。图片等对象原有的锁定宽高比状态保留。

## Office 图标

[图标诊断](../artifacts/office-icons/diagnostics.json) 记录 25 个命令映射均 `loaded=true`，本机获取的图像为 24 × 24。这里的 25 指命令映射数量，一些命令共用同一内置图标。

功能区使用 `imageMso`；窗格与轮盘在 PowerPoint 进程内调用 `CommandBars.GetImageMso`，将图像复制为托管资源并按尺寸缓存。绘制期间不发起 COM 取图调用，关闭时释放缓存。对齐和等间距使用 PowerPoint 相应命令的图标；同宽、同高使用 Office 平均分布列宽/行高的图标，文字仍明确显示实际约束。具体样式由当前安装的 Office 提供。

微软公开接口见 [CommandBars.GetImageMso](https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.getimagemso)，命令标识来源为 [OfficeDev 官方命令列表](https://github.com/OfficeDev/office-fluent-ui-command-identifiers)。加载诊断证明资源可取得，不等于完整视觉验收。

## 真实 PowerPoint 中的程序化检查

[形状报告](../artifacts/shape-support/run-20261116T083737058/acceptance.json) 为 `Passed=true`，11 个代表对象全部完成：Ellipse、RoundedRectangle、Triangle、Diamond、RightArrow、TextBox、Picture、LinkedPicture、Graphic、Freeform、Group。链接 SVG 类型由能力判定支持，但不在这 11 类实际 fixture 中。

每页以目标类型和另一个对象建立同 Y、同宽关系，分别由两个成员执行 Move 和 ResizeE，共 44 条程序化轨迹。每次轨迹后调用真实 PowerPoint 原生 Undo、Redo，比较几何、编号、标签、权威 XML 和原生属性，独立核对关系公式。检查保留文本、AutoSize、锁定宽高比和组合子成员数据，不用清空这些属性来绕过差异。

额外检查包括：

- 旋转对象、线条和表格的新增命令被拒绝，几何、标签及 XML 不变。
- 非矩形副本获得独立编号，原件保持；副本之间可以建立关系，并逐步原生撤销和重做。
- 保存并在同一进程重新打开后，形状类型、原生属性、几何、组合成员及关系数据一致。

脚本为 [TestShapeSupport.ps1](../scripts/TestShapeSupport.ps1)，只创建和关闭自己的测试文稿，并恢复之前的活动窗口。以上输入是加载项测试入口驱动，不是模拟鼠标消息，也不证明所有原生命中路径已被用户遍历。

[新进程持久化报告](../artifacts/shape-support/run-20261116T083737058/new-process-persistence.json) 为 `Passed=true`、`PersistenceNewProcess=true`。测试文稿保存后完全退出 PowerPoint，在不同原生进程重新打开，按保存快照只读核对：类型、文本、自动适应、宽高比锁和组合成员保持；编号标签与 XML 精确一致；几何在 0.01 pt 内；独立关系公式有效。实际进程号保留在报告中。

[命令回归报告](../artifacts/v3-command-validation/run-20261116T084038776/acceptance.json) 检查现有共享命令及窗格事务，通过原生新增、修改、停用、启用和删除的历史检查。20 类求解器语义沿用已有实现，本轮 11 类形状试验不表示对所有形状与所有约束的笛卡尔组合逐项穷举。

## 构建检查

`Build.ps1` 的 129 项离线检查为：42 项 CoreTests、11 项 IdentityTests、11 项 NativeConstraintStoreTests、16 项 PaneTests、18 项 RadialTests、23 项 ShapeSupportTests、8 项 OfficeIconTests。随后构建原生 x64 ATL 窗格宿主。

离线检查覆盖能力边界、嵌套组命中归属及图标映射等行为；与真实 PowerPoint 报告、物理鼠标操作和视觉验收分别计算。
