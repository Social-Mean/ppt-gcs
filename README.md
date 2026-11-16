# PPT GCS

Windows 桌面 PowerPoint 的几何约束插件原型。v0.2 支持同 X、同 Y、同宽、同高，以及水平、垂直等间距；关系可组合，编辑任意成员都会求解它所在的约束连通分量。

**v0.2 已在本机通过真实鼠标端到端验收。** 六种约束分别由两个不同成员驱动，共 13 次移动/缩放、567 次鼠标按住期间的几何更新；每帧约束和独立端点检查通过。13 次原生 Undo、13 次 Redo 逐步恢复完整布局，保存并重启 PowerPoint 后几何、GUID 和约束 XML 一致。结果见 [验收报告](artifacts/layout-v2-validation/acceptance.json) 和 [验证记录](docs/validation-v2.md)。这些日志记录实际几何写回，不是渲染 FPS。

编译、28 项核心测试、交互几何和消息队列检查也已通过。额外的 [程序化回放](artifacts/layout-v2-replay/report.json) 覆盖 12 个主用例（36 帧）、每例额外 15 px 释放端点变化，以及 14 次原生 Undo 和 14 次 Redo（含两步历史链）；它是补充证据，不计入上面的真实鼠标帧数。v0.1 顶部联动结果保留在 [历史记录](docs/validation.md)。

## 构建与安装

当前仅面向 **Windows + 64 位桌面 PowerPoint**。开发机需要 Visual Studio 的 MSBuild、Roslyn 编译器，以及已安装的 Microsoft Office / Office PIA。项目使用 .NET Framework 4.8 原生 COM 加载项，不依赖 VSTO 项目模板或 .NET SDK。

关闭 PowerPoint 后，在项目根目录的 64 位 PowerShell 中运行：

```powershell
.\scripts\Build.ps1
.\scripts\Install.ps1
```

`Build.ps1` 编译两个项目并运行核心测试，产物位于 `bin/`。首次构建从 NuGet 恢复固定版本 `Microsoft.NETFramework.ReferenceAssemblies 1.0.3` 和 `MathNet.Numerics 5.0.0`，缓存于 `.packages/`；无需在系统中安装 net48 targeting pack。MathNet 使用 MIT 许可证，项目引用其 net48 程序集。官方 Extensibility 和 Office PIA 的 COM 接口元数据由编译器嵌入插件。

`Install.ps1` 只写入当前用户的 COM 类与 PowerPoint 加载项注册表项，不需要管理员权限。重新打开 PowerPoint 后应出现 **PPT GCS** 功能区。注册指向本项目的 `bin/PptGcs.AddIn.dll`；移动项目目录后需重新运行安装脚本。重新构建前请关闭 PowerPoint，避免加载中的 DLL 被占用。

也可以用 Visual Studio 打开 `PptGcs.sln` 编辑源码。安装脚本使用的统一输出目录由 `Build.ps1` 生成。

## 建立与编辑约束

1. 在同一页选中至少两个未旋转、未分组的普通矩形，点击 **同 X（左边对齐）**、**同 Y（顶部对齐）**、**同宽** 或 **同高**。
2. 建立等间距时至少选中三个矩形，点击 **水平等间距** 或 **垂直等间距**。建立时按对应坐标排序，之后始终按这个已记录的成员顺序计算。
3. 从矩形内部拖动可以移动它；先选中一个矩形，再拖动原生选择框的八个缩放点，可以调整大小。拖动西侧、北侧或角点时，以相对一侧为锚点计算新几何。
4. 同宽或同高关系会传播对应尺寸。位置和间距约束根据更新后的尺寸重新求解，其余位置尽量少移动。
5. 选中关系中的对象，点击 **移除选中对象的约束** 可移除相关关系；**状态** 显示当前载入情况和日志路径。
6. **新建六组演示** 会生成六页测试文稿，每页写有操作说明，也可打开 [layout-v2.pptx](artifacts/layout-v2.pptx)。

X、Y 分别指矩形的左侧、顶部坐标。间距按相邻矩形的边缘计算，不是中心间距。等间距的数值没有固定，也没有固定首尾：拖动目标按编辑位置确定，其他位置采用最小平方位移解。

间距是**有符号的边缘间距**，允许变为负数而出现重叠；这项约束不等于“不重叠”或“固定间距”。拖动过程中不重新排序成员。

普通鼠标拖动和缩放由插件接管。按住 Ctrl、Alt 或 Shift 时保留 PowerPoint 原生手势，该手势不会由插件维持约束；键盘微移、位置/尺寸面板编辑、多选拖动和复制粘贴的联动也尚未纳入当前实现。

## 数据与验证

- `PptGcs.Core`：不依赖 Office 的约束模型、XML 校验、同尺寸分组传播，以及使用 MathNet SVD 的线性位置约束求解。每次只修改被编辑对象所在的约束连通分量；新增约束会先试算，冲突时不写入几何。
- `PptGcs.AddIn`：PowerPoint COM 接入、功能区、鼠标移动与八方向缩放、原生 Undo 边界及文档存储。
- 约束保存在演示文稿的 `CustomXMLParts`，命名空间为 `urn:ppt-gcs:v2`；矩形通过 `PPTGCS_ID` GUID 标签关联。文件仍是普通 `.pptx`。
- v0.2 不读取或迁移 v0.1 约束格式。旧 [two-rectangles.pptx](artifacts/two-rectangles.pptx) 和对应验收证据是历史归档，当前演示请使用六页 v0.2 文稿。
- 日志写入 `%LOCALAPPDATA%/PptGcs/layout-v2.log`。实际验收快照保存在 `artifacts/layout-v2-validation/`。

`scripts/TestLayoutDocument.ps1` 需通过 Windows PowerShell 5.1 运行，用于六页测试文稿的创建、选页、保存、重开和快照。写入和关闭前会核对六页的测试标记。

`python scripts/VerifyLayoutEvidence.py` 核对已有的真实操作日志、每一步撤销/重做快照，以及重启前后的文档状态。缺少任何必需证据都会失败；它不会自动控制 PowerPoint 或生成鼠标拖动数据。

## 卸载

关闭 PowerPoint 后运行：

```powershell
.\scripts\Uninstall.ps1
```

脚本只删除该插件的当前用户注册表项。PPT 文件与其中的约束数据保留；未加载插件时，矩形仍按普通 PowerPoint 形状显示。
