# v0.3 约束、轮盘与窗格验证

日期：2026-11-16。环境为 Windows、64 位桌面 PowerPoint 16.0、.NET Framework 4.8、MathNet.Numerics 5.0.0，以及 Visual Studio C++/ATL 原生任务窗格宿主。

## 最终结论

**v0.3 本机验收通过：20 类约束的程序化驱动、代表性的真实中键轮盘/窗格操作、完整原生历史、复制后建立关系的历史链，以及新进程保存重开均取得通过证据。** 全部 20 类关系由程序化宿主检查覆盖；最终物理鼠标试验覆盖同宽、宽高比和水平等间距等代表路径，不表示逐个人工遍历了所有类型和每个轮盘扇区。

最终真实交互记录了两次轮盘命令、6 次受约束手势、349 帧鼠标按住期间的几何更新，以及 14 次原生 Undo 和 14 次 Redo。新进程持久化重新检查了 20 条启用关系和 40 个已标记矩形，权威数据保持一致。

本记录区分纯核心/存储测试、真实 PowerPoint 中的程序化命令和轨迹回放、物理鼠标操作。脚本报告中的 `PhysicalMouseInput=false` 保持原义，不把程序化轨迹计作用户拖动。

仓库中的交互日志只保留 `before.json` 至 `saved.json` 的本次验收时段，逐行内容未改写；之前的会话记录不随仓库发布。裁剪后重新执行历史审计，结果仍通过。

## 约束范围与程序化驱动

[20 类约束报告](../artifacts/v3-constraint-validation/latest.json) 为 `Passed=true`，最终成功运行位于 `artifacts/v3-constraint-validation/run-20261116T081512592/`。20 个独立页面覆盖全部枚举：

- 原有六种：SameX、SameY、SameWidth、SameHeight、EqualHorizontalGap、EqualVerticalGap。
- 边缘和中心：RightAligned、BottomAligned、CenterXAligned、CenterYAligned、CentersCoincident。
- 单对象参数：FixedWidth、FixedHeight、FixedPosition、AspectRatio。
- 多对象参数：FixedHorizontalGap、FixedVerticalGap、Offset、SymmetricVertical、SymmetricHorizontal。

共有 36 次程序化驱动。多对象关系由两个不同成员驱动，单对象约束使用允许的尺寸编辑方向。每次都要求实际几何变化，逐步调用 PowerPoint 原生 Undo、Redo，并比较所有页面的原生几何、编号、归属、标签和权威 XML。

验收器使用独立的坐标、尺寸、边缘间距、比例和镜像公式，不以窗格的“已满足”文字代替数值检查。全部有效关系残差不超过 0.01 pt；本轮最大残差约为 0.0000916 pt。窗格模型中相应关系也显示已满足。

复核脚本为 [TestV3Constraints.ps1](../scripts/TestV3Constraints.ps1)。它创建自己的 20 页 `PPTGCS_TEST=V3` 文稿，不修改当前用户文稿，并明确标记程序化输入。

## 原生约束管理历史

[命令与窗格报告](../artifacts/v3-command-validation/latest.json) 为 `Passed=true`，最终成功运行位于 `artifacts/v3-command-validation/run-20261116T081507443/`。

从没有任何约束编号的原生矩形开始，依次执行固定宽度新增、参数修改、停用、启用、删除，形成五个事务。连续五次 Undo 和五次 Redo 均恢复对应完整状态，包含文稿标签、原生形状几何、GUID、owner 和 XML。停用、启用、删除这些仅修改关系数据的操作也各有独立原生历史项。

额外检查通过：

- 同宽关系连接的两个对象不能同时满足不同固定宽度，冲突命令被拒绝，完整状态保持。
- 按关系成员选择对象后，PowerPoint 原生选择和窗格模型一致。
- 删除一个被引用对象后，窗格仍显示失效关系；可以只删除该关系，保留其他关系。
- 对失效关系清理执行原生 Undo/Redo，也能完整恢复数据状态。

这些结果来自 [TestV3Commands.ps1](../scripts/TestV3Commands.ps1) 的实际 COM 加载项入口和 PowerPoint 原生命令，不是直接改写测试期望值。

## 复制后建立关系的历史链

[复制历史报告](../artifacts/v3-copy-validation/latest.json) 为 `Passed=true`，运行位于 `artifacts/v3-copy-validation/run-20261116T075410883/`。测试在独立六页文稿中原生 Duplicate 两个已有约束的矩形，选中副本后自动分配新的 GUID；所有原件的几何、标签及已有关系不变。

随后只给两个副本添加同高关系。副本高度原本就相同，所以这次添加没有任何形状几何或标签变化，仅修改权威关系数据。依次执行 Undo 新关系、Undo 复制、Redo 复制、Redo 新关系，四步均精确恢复所有原生几何、全部标签和 XML。

这直接复测了 v0.2 曾失败的“复制后仅新增元数据关系”场景，证明缺口已补齐，没有通过额外几何变动制造历史项。脚本 [TestV3CopyHistory.ps1](../scripts/TestV3CopyHistory.ps1) 仅操作自建测试文稿，保存关闭后恢复原活动窗口；该项为真实 PowerPoint 程序化测试，`PhysicalMouseInput=false`。

## 存储机制与回滚

探针确认 `Presentation.Tags` 可进入 PowerPoint 原生 Undo/Redo，而旧 `CustomXMLParts` 数据更新不能。详见 [属性探针](../artifacts/native-history-probe/property-results.json) 和 [标签事务探针](../artifacts/native-history-probe/presentation-tag-transactions.json)。后者覆盖四个连续事务、4 次 Undo/Redo，以及大标签保存重开；探针文稿均与用户文稿隔离。

`NativeConstraintStore` 将权威 XML 编成 UTF-8 大写 HEX，写入 `PPTGCS_CONSTRAINTS`，避免标签大小写处理改变 XML。读写有 1 MiB UTF-8 边界、严格编码和 XML 校验、写后精确读回以及失败恢复。空约束文档仍是明确的编码标签值，可以作为正常元数据事务撤销。

共享命令先求解再提交，将数据标签、对象身份和几何放在同一原生边界内。异常路径保存原始标签值，分别尝试恢复各项数据，汇总恢复错误，不以重新序列化空文档替代原始无标签状态。没有隐藏形状、无意义几何修改、盲目回退原生历史或替换 Ctrl+Z 的私有历史栈。

旧 CustomXML/v0.1 运行时兼容路径不保留。当前会话的用户工作文稿已在备份后一次性保留全部 14 条原有关系并另存为新存储；这是本次明确的文件操作，不是产品中的自动迁移或 fallback。最终重新打开七页工作文稿，并与更新前的备份核对：已存储的形状名称、几何、原生 ID、GUID 和关系数据保持不变，新的约束窗格可用。用户文件与备份不在仓库附件中。

## 原生窗格与轮盘

原生窗格通过 `ICustomTaskPaneConsumer`、`ICTPFactory.CreateCTP` 激活原生 ATL ActiveX 宿主，再承载托管 WinForms 视图。实际诊断为 `factory=true`、`created=true`、`visible=true`、右侧停靠，并有非零宿主和内容窗口句柄，见 [窗格诊断](../artifacts/v3-command-validation/run-20261116T081507443/pane-diagnostics.json)。微软提供的 [CreateCTP 接口](https://learn.microsoft.com/en-us/office/vba/api/office.ictpfactory.createctp) 是此方案的公开扩展点；未修改内置选择窗格。

窗格实现对象列表、关系列表、当前页/仅所选对象过滤、双向选择、参数创建与修改、启停、单条删除，以及缺失对象状态。等间距保留一条多成员关系及其顺序。Ribbon、轮盘、窗格共用 `ConstraintCommands`。

窗格绑定所属的原生文稿窗口，延迟回调执行前核对活动窗口身份，避免切换文稿后误操作另一窗口中的对象。最终命令和 20 类约束回归均在用户工作文稿保持打开时运行，并在结束后恢复原活动窗口。

中键轮盘有八个固定扇区：左边对齐、顶部对齐、同宽、同高、水平等距、垂直等距、更多约束、查看约束。先选对象，再在画布按住中键并划向目标，松开确认；中心区域释放或 Esc 取消。不满足数量要求的扇区置灰，预览不写入文稿，更多约束进入参数编辑器。滚轮事件仍交给 PowerPoint。

最终 [真实交互报告](../artifacts/v3-interaction/acceptance.json) 为 `Passed=true`、`PhysicalMouseInput=true`，用户确认窗格可见，并完成：

- 中键轮盘创建 SameWidth 与 EqualHorizontalGap 两种关系。
- 对同宽的两个不同对象、宽高比对象和等间距对象进行真实拖动/缩放，共 6 个受约束手势、349 帧鼠标按住期间更新。
- 从窗格创建宽高比 2，修改为 1.5，再停用、启用及删除关系。
- 完整执行 14 次原生 Undo 和 14 次 Redo，恢复全部元数据与几何。

[历史审计](../artifacts/v3-interaction/history-audit.json) 将 14 个历史步骤对应为 6 个完整插件手势、7 个关系元数据事务，以及 1 个对象无约束时的原生编辑。每个手势和元数据命令都对应完整的一步，没有拆成多帧历史；最终状态已恢复。人工最终区间覆盖 Move、ResizeE 和 ResizeSE，未逐项遍历其余缩放方向、所有扇区与取消分支；其他约束类型仍以独立程序化检查为证。

## 新进程保存重开

[TestV3Persistence.ps1](../scripts/TestV3Persistence.ps1) 的只读 Capture/Verify 两阶段检查已完成，[持久化报告](../artifacts/v3-persistence/acceptance.json) 为 `Passed=true`。保存 20 页测试文稿后完全退出 PowerPoint，进程由 35736 切换到 57128，再打开同一路径核对。

权威 XML 和全部 PPTGCS 标签精确一致，40 个已标记矩形的原生 ID、名称、次序和归属保持，几何误差在 0.01 pt 内。独立公式重新核对 20 条启用关系全部有效，最大残差约为 0.0000610 pt。此项是完整产品文稿的新进程检查，不依赖早期标签探针推断。

## 构建与适用边界

`scripts/Build.ps1` 构建 net48 加载项、核心库并运行 98 项检查：42 项 CoreTests、11 项 IdentityTests、11 项 NativeConstraintStoreTests、16 项 PaneTests 和 18 项 RadialTests；随后通过 `BuildPaneHost.ps1` 构建原生 x64 ATL 宿主。窗格与轮盘的离线行为测试不模拟物理输入。需要 Visual Studio MSBuild/Roslyn、C++ x64 工具、Windows SDK、ATL，以及已安装 Office/PIA。`Install.ps1` 注册当前用户的加载项与原生宿主，构建/安装前需关闭 PowerPoint。

当前支持同页未旋转、未分组的普通矩形，尺寸至少 1 pt。固定间距使用有符号边缘距离，允许负值和重叠；轴对称是关于固定轴的完整矩形镜像。普通拖动和八方向缩放可触发求解，宽高比编辑会遵守相应锚点。

任意旋转、曲线相切、三维装配、不重叠不等式、完整自由度诊断，以及键盘微移/原生数值面板/多选拖动联动不属于当前实现。关系满足与“完全约束”不同，保留整体移动自由度属于正常布局行为。

历史 [v0.2 真实鼠标记录](validation-v2.md) 和 [复制编号记录](copy-identity-validation.md) 保留，当前版本结果以本页链接的 v0.3 报告为准。日志记录的是几何写回采样，不是渲染 FPS 或端到端延迟测试。
