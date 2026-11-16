# 复制矩形编号修复

日期：2026-11-16。Windows 桌面 PowerPoint 16.0；约束存储仍为 `urn:ppt-gcs:v2`。

## 原因与行为

PowerPoint 复制形状时连同插件的 `PPTGCS_ID` 标签一起复制，原实现只检测重复并拒绝建立关系。现在由 `ShapeIdentityManager` 保留原件编号，为副本生成新的 GUID。副本可以建立独立约束，不会自动加入原件的关系。

文稿通过 `PPTGCS_DOCUMENT` 保存令牌，形状通过 `PPTGCS_OWNER` 保存“文稿令牌 / Slide ID / 原生 Shape ID”。原件与记录相符时保留编号；副本的实际归属不同，会换号。处理先计划再写标签，失败时逆序恢复；不改写约束 XML、位置、尺寸或形状名称。

处理连接到 PowerPoint 的 `WindowSelectionChange`、`AfterPresentationOpen` 事件，以及建立/移除约束和手势开始的编辑入口。空闲计时器只读文档状态，不轮询写入标签。事件定义取自本机 Office `MSPPT.OLB`，通过 .NET Framework 的 `ComEventsHelper` 订阅。微软的 [WindowSelectionChange 文档](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.windowselectionchange) 说明此事件覆盖界面和代码导致的选择变化；外部 COM 测试使用有界等待，避免在异步事件送达前读取副本编号。

对缺少归属标签的既有矩形，唯一编号可以直接记录归属；已有重复编号但没有约束引用时，可以安全分配独立编号。若某个已引用编号在同页对应多个对象，且没有归属证据，则拒绝猜测原件，不会静默转移原有约束。

## 验证

构建和 28 项核心测试、11 项编号测试通过。编号测试覆盖原件与 XML 保持、多代复制、重排、跨页/跨文稿、歧义拒绝、部分写入失败回滚和重复调用不写入。

实际 PowerPoint 对象模型测试通过，使用原生 Duplicate、Copy/Paste、选择事件、插件约束入口和原生 Undo/Redo 命令；这是程序化宿主测试，不是新增的人工鼠标测试：

- 批量复制两个受约束矩形，副本自动得到不同 GUID；全部原件编号、几何、XML 保持。
- 原生撤销复制，再重做复制，副本 GUID 和 owner 完整恢复。
- 两个副本能建立同宽关系，分别从两个副本驱动尺寸；缩放的原生撤销/重做通过，原矩形不变化。
- 原件同 Y 关系继续联动，副本不跟随；撤销恢复原状态。
- 再次复制副本会产生新的 GUID；对这个未建立关系的副本执行移除约束，不会删除已有关系。
- 跨页 Copy/Paste 后编号独立，能够建立新的同高约束，原有关系全部保留。
- 跨文稿粘贴生成独立文稿令牌与形状编号，不导入源文稿关系。
- 保存后完全退出 PowerPoint，再从进程 10544 切换到新进程 55248 重开，所有形状、原生 ID、GUID、归属标签、几何、约束 XML 和文稿令牌均完全一致。

主证据：[基本验收](../artifacts/copy-identity-validation/acceptance.json)、[扩展验收](../artifacts/copy-identity-validation/variants/acceptance.json)、[新进程重开验收](../artifacts/copy-identity-validation/variants/reopen-check.json)。测试脚本为 `scripts/TestCopyIdentity.ps1` 和 `scripts/TestCopyIdentityVariants.ps1`，需使用 Windows PowerShell 5.1，并只操作带有路径检查的测试文稿。重开后使用后一个脚本的 `-VerifyReopen` 参数进行只读核对。

用户当前文稿另外保存在本机工作副本中：7 页、17 个已标记矩形，两个旧副本重新编号，原矩形编号、全部几何及原有 6 条约束保持。该工作副本也已保存并在新进程中重新打开，快照完全一致。用户文稿及备份不纳入仓库验收附件。

## 独立发现的原生历史限制

新增或删除 `CustomXMLParts` 约束数据本身不会产生可用的 PowerPoint 原生撤销项。这是原有约束创建/移除历史的缺口，不是复制编号处理失败。测试中对尺寸本已相同的副本添加同宽关系后，一次 Undo 直接撤销了复制，XML 中的新关系仍在；Redo 恢复副本及相同的新编号后，关系再次有效。

相关失败证据保留在 `artifacts/copy-identity-validation/constraint-creation-undo-failure/`；报告明确标记 `ConstraintCreationUndoSupported=false`。本次通过的是复制历史，以及已建立关系后的几何修改历史，不宣称新增/移除关系的原生历史已通过，也没有用无意义的几何写入伪造撤销项。
