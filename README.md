# PPT GCS

Windows 桌面 PowerPoint 的几何约束实验插件。当前目标是让同一页上的两个矩形保持顶部坐标相同：拖动任意一个，另一个随之上下移动，并验证撤销和保存重开。

首个端到端版本已在本机通过验收：从保存文件重新打开后，A/B 双向拖动共记录 48 次鼠标按住时的联动更新，原生撤销和重做各两步正确恢复完整几何，再次保存并重启 PowerPoint 后几何与约束数据保持一致。实际鼠标拖动由用户操作，插件日志与 COM 快照交叉核对。详见 [验证记录](docs/validation.md) 和 [机器验收结果](artifacts/acceptance.json)。

## 构建与安装

当前仅面向 **Windows + 64 位桌面 PowerPoint**。开发机需要 Visual Studio 的 MSBuild、Roslyn 编译器，以及已安装的 Microsoft Office / Office PIA。项目使用 .NET Framework 4.8 原生 COM 加载项，**不是 VSTO 项目**，不依赖 VSTO 项目模板或 .NET SDK。

关闭 PowerPoint 后，在项目根目录的 64 位 PowerShell 中运行：

```powershell
.\scripts\Build.ps1
.\scripts\Install.ps1
```

`Build.ps1` 编译两个项目并运行核心测试，产物位于 `bin/`。首次构建会从 NuGet 恢复固定版本 `Microsoft.NETFramework.ReferenceAssemblies 1.0.3` 及其 net48 参考程序集，缓存于 `.packages/`；无需在系统中安装 net48 targeting pack。官方 `Extensibility` 和 Office PIA 的 COM 接口元数据由编译器嵌入插件，不使用手写的接口 ABI。

`Install.ps1` 只写入当前用户的 COM 类与 PowerPoint 加载项注册表项，不需要管理员权限。重新打开 PowerPoint 后，应出现 **PPT GCS** 功能区。注册指向本项目的 `bin/PptGcs.AddIn.dll`；移动项目目录后需要重新运行安装脚本。重新构建前请关闭 PowerPoint，避免正在加载的 DLL 被占用。

也可以用 Visual Studio 打开 `PptGcs.sln` 查看和编辑源码。安装脚本使用的统一输出目录由 `Build.ps1` 生成。

## 试用流程

1. 点击 **PPT GCS → 新建演示**，生成已配对的两个矩形；或在同一页选择两个普通矩形，点击 **顶部联动**。
2. 从矩形内部拖动 A，再拖动 B，观察另一矩形的顶部是否持续对齐。不要从边框或缩放控制点开始拖动。
3. 使用 PowerPoint 的撤销，检查两个矩形能否一起恢复到操作前的位置。
4. 保存为 `.pptx`，关闭并重新打开，再拖动两个矩形，检查约束是否仍然有效。
5. 选择配对中的矩形，点击 **移除联动** 可删除关系；**状态** 显示当前载入情况和日志路径。

已验证的示例为 [two-rectangles.pptx](artifacts/two-rectangles.pptx)。运行日志默认写入 `%LOCALAPPDATA%/PptGcs/addin.log`。`python scripts/VerifyEvidence.py` 可重新核对本次保存的验收证据；它不会自动控制 PowerPoint 或重新执行拖动。

## 实现与范围

- `PptGcs.Core`：不依赖 Office 的数据模型、XML 校验，以及单个 `A.Top = B.Top` 等式的直接精确求解。尚未引入完整 Cassowary 求解器，也不支持通用约束网络。
- `PptGcs.AddIn`：PowerPoint COM 接入、功能区、受约束拖动交互与文档存储。使用 .NET Framework 4.8。
- 约束关系保存在演示文稿的 `CustomXMLParts` 中，命名空间为 `urn:ppt-gcs:v1`；矩形通过 `PPTGCS_ID` GUID 标签关联。文件仍是普通 `.pptx`。
- 当前仅支持同一页、未旋转、未分组的普通矩形；每个矩形只能加入一组顶部配对。仅覆盖矩形内部拖动，不承诺缩放、键盘微移、位置面板编辑、复制粘贴或多选拖动的联动行为。
- 同宽、同高、X 坐标、间距等约束不在当前版本范围内。

## 卸载

关闭 PowerPoint 后运行：

```powershell
.\scripts\Uninstall.ps1
```

脚本只删除该插件的当前用户注册表项。PPT 文件与其中保存的约束数据会保留；未加载插件时，矩形仍按普通 PowerPoint 形状显示。
