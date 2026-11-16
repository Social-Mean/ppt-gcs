# v0.2 多约束验证记录

日期：2026-11-16。环境：Windows，64 位桌面 PowerPoint 16.0.17932.21056，.NET Framework 4.8，MathNet.Numerics 5.0.0。

## 最终结论

**v0.2 本轮真实鼠标端到端验收通过。** [acceptance.json](../artifacts/layout-v2-validation/acceptance.json) 为 `passed=true`，检查容差为 0.01 pt。成功区间记录 13 次实际移动/缩放、567 次几何更新，全部 567 次都发生在鼠标按住期间。每帧约束、从鼠标位置和缩放比例独立计算的驱动目标、最终释放端点及无关几何保持均通过检查。

六类约束均由两个不同矩形驱动，覆盖如下：

- 同 Y：2 次 Move，97 帧。
- 同 X：2 次 Move，50 帧。
- 同宽：3 次 ResizeSE，146 帧；A 操作两次，B 操作一次。
- 同高：2 次 ResizeS，139 帧。
- 水平等间距：2 次 Move，62 帧。
- 垂直等间距：2 次 Move，73 帧。

13 次 PowerPoint 原生 Undo 和 13 次 Redo 均逐步恢复对应完整手势的六页布局，形状标识和 XML 未变。北京时间 13:33:50 保存后完全退出进程 52948，13:34:04 在新进程 53568 中重新打开同一文件；几何、GUID、Slide ID、约束 XML 均一致。新进程的插件状态确认已载入 6 条约束并正常运行，见 [保存前快照](../artifacts/layout-v2-validation/before-restart.json) 和 [重开后快照](../artifacts/layout-v2-validation/after-restart.json)。

其他已完成检查：

- 最新控制器正式构建成功，28 项核心测试通过。
- 10 项交互几何检查通过，覆盖移动及八方向缩放的目标几何计算；另有 10 项消息队列检查通过。
- [程序化回放报告](../artifacts/layout-v2-replay/report.json) 为 `passed=true`，覆盖六种约束、两个驱动成员、额外释放端点变化及真实 PowerPoint 原生历史，详情见下节。
- 六页演示的 14 个矩形及 112 个缩放点预检通过，相关输出保存在 [preflight.json](../artifacts/layout-v2-validation/preflight.json)。
- [before.json](../artifacts/layout-v2-validation/before.json) 记录北京时间 13:28:11 的初始几何与 XML；[after.json](../artifacts/layout-v2-validation/after.json) 记录全部手势结束后的布局；[gestures.log](../artifacts/layout-v2-validation/gestures.log) 保留真实操作轨迹。

最终成功的鼠标测试覆盖 Move、右下角缩放 ResizeSE 和下边缩放 ResizeS；八个缩放方向的目标计算均做了独立检查，但本轮最终成功区间没有逐一人工重测其余六个方向。组合约束的求解已由核心测试覆盖，本轮六页人工验收使用单类关系。

## 最新控制器的程序化回放

[report.json](../artifacts/layout-v2-replay/report.json) 明确标注 `evidenceType="programmatic-replay-native-undo"`、`provesPhysicalMouseInput=false`。它使用同一控制器和真实 PowerPoint 对象模型，覆盖六种约束的两个驱动成员，共 12 个主用例、36 次几何写回。

每个用例的最终释放点都比上一移动点额外变化 15 px，独立核对实际几何是否到达精确目标，并检查每个程序化帧的约束。12 个主用例加上额外的两步历史链，共核对 14 次 PowerPoint 原生 Undo 和 14 次 Redo，全部通过；六页几何和 XML 的保留也通过检查。

这证明同一控制器在程序化调用中的求解、释放端点写回及原生历史行为。程序化调用不计入真实鼠标证明；后续实际角点缩放及六类关系的完整操作已由上面的最终真实验收单独核对。

## 先前真实交互测试的问题与修正

下列输入接管、释放端点及撤销分组缺陷已在最终实现中修正，并通过上述成功区间复测。失败记录继续保留；PowerPoint 的原生历史容量属于宿主限制，本轮 13 次修改在容量内完整验证。

第一轮人工测试记录了 23 次非取消修改，六种关系都有双向驱动，所有已记录帧都满足约束。但还发现 8 次活动窗格导致的坐标转换错误，以及同宽页 3 次未被接管的原生尺寸变化，因此该轮完整验收失败。原始快照、日志和文稿保存在 `artifacts/layout-v2-failed-run-20261116-125200/`，不能按 23 个日志手势直接推算原生撤销次数。

`scripts/TestPaneProjection.ps1` 在实际 PowerPoint 中复现：缩略图或备注窗格处于活动状态时，RangeFromPoint/PointsToScreenPixels 接口不可用；激活 ViewType=ppViewSlide 的窗格后恢复正常。

后续修正版已编译并在新进程 50184 载入：只接管实际 mdiClass 画布，在该画布输入开始时激活幻灯片窗格；操作结束释放捕获并 Unselect/Select 刷新原生选区；命中半径按 DPI 缩放；处理双击和文本容器选择。新增指针路由及原生选择范围日志。六页所有形状及 112 个缩放点所属窗口的预检通过。

## 第二轮发现的撤销分组问题

第二轮及补充的 B 同高操作共记录 23 次非取消修改，六种约束的双向覆盖齐全。该轮主操作区间没有运行错误或未记录的几何变化，各帧约束、手势间基线连续性和最终快照均通过检查。

原生撤销第 1–11 步正确恢复对应完整手势。第 12 步处理同宽页 A 的右下角缩放（`ResizeSE`，手势 `ec7a2f3d-ef39-447d-80d8-0069d6aef54f`）时失败：它只撤销了鼠标释放回调中的第 29 帧，停在第 28 帧的宽度约 211.162750 pt、高度约 118.604645 pt，而不是操作前的宽度约 209.612366 pt、高度 100 pt。日志表明第 28 帧 `mouseDown=true`，第 29 帧 `mouseDown=false`；释放回调中的几何写回形成了额外撤销项。

初次通过桌面控制工具发送 Ctrl+Z 时，焦点仍在浏览器，因此该次按键未计入 PowerPoint 验证。后续实际步骤通过已验证身份的六页测试文稿执行 PowerPoint 官方 `CommandBars.ExecuteMso("Undo")` / `ExecuteMso("Redo")`，操作的是 PowerPoint 原生历史，每一步再用只读 COM 快照核对几何。

发现第 12 步失败后，连续执行 12 次原生 Redo，恢复后的完整状态与本轮 `after.json` 一致。原始日志、失败快照、恢复快照、二进制哈希及文稿保留于本地归档 `artifacts/layout-v2-undo-failure-20261116-131207/`，该轮不能标记为通过。

随后曾尝试仅在 `WM_MOUSEMOVE` 中写回几何，鼠标释放和定时器结束路径不再补写，并核对最后一次移动坐标和释放坐标。该方案在第三轮出现下面的释放端点问题，因此最终改为专属手势消息循环。

## 第三轮释放端点与历史容量检查

第三轮共记录 27 次有几何修改的手势，其中出现 3 次 `gesture-release-mismatch-error`：最终释放位置对应的几何与最后一次鼠标移动写回不一致。错误由明确的释放位置检查暴露，不能通过忽略末尾误差来判定通过。

原生 Undo 第 1–20 步的完整几何均与预期一致；第 21 步命令不可用，触及本机 PowerPoint 的 20 项原生历史容量，因而无法完成该轮全部 27 次操作的历史验证。之后执行 20 次原生 Redo，恢复状态与该轮最终快照精确一致。

该轮日志和快照保留于本地归档 `artifacts/layout-v2-release-failure-20261116-132316/validation/`。这轮同时存在释放端点不一致和撤销覆盖不足，不能标记为完整通过。

最新实现改为在最初鼠标按下回调内运行专属手势消息循环。移动和最终释放位置的几何写回都在同一次回调内完成，并继续用实际 COM 读回验证释放端点。日志帧的来源字段固定为 `input="owned-gesture-loop"`；验收器不接受旧的 `mouse-move` 来源作为当前版本证据。

普通内部点击增加拖动阈值：`frames=0` 且 `dragThresholdPassed=false` 时必须保持全部几何不变，不计入修改或撤销次数。有写回的手势仍全部进入历史链；双向覆盖仍要求正确操作模式、目标属性实际改变以及至少两次鼠标按住期间的更新。释放误差、错误事件、缺帧、断裂基线和取消后的修改均继续使验收失败。最终实现以这套严格检查通过了 13 次真实手势及完整原生历史复测。

测试文稿：[layout-v2.pptx](../artifacts/layout-v2.pptx)。

## 真实操作验收流程

演示共六页，顺序为 SameY、SameX、SameWidth、SameHeight、EqualHorizontalGap、EqualVerticalGap。用户按各页说明对两个不同成员分别操作一次，形成至少 12 次有效手势：

1. 同 Y：分别上下移动两个矩形。
2. 同 X：分别左右移动两个矩形。
3. 同宽：先选中矩形，再拖动其左右缩放点改变宽度；对另一成员重复。
4. 同高：先选中矩形，再拖动其上下缩放点改变高度；对另一成员重复。
5. 水平等间距：分别左右移动不同成员，检查相邻边缘间距持续相等。
6. 垂直等间距：分别上下移动不同成员，检查相邻边缘间距持续相等。

本轮普通手势不按 Ctrl、Alt 或 Shift。每次应有多次鼠标按住期间的实际写回，不能只检查松手后的终点。

完成后，逐步执行 PowerPoint 原生 Undo，再逐步 Redo，每步读取六页完整几何。当前自动验收通过 `CommandBars.ExecuteMso` 调用这两个原生命令，并在执行前核对测试文稿身份。之后保存、完全退出测试 PowerPoint，启动新进程并打开同一文件，比较几何、形状 GUID、Slide ID 和约束 XML。

## 证据文件约定

所有证据位于 `artifacts/layout-v2-validation/`：

- `before.json`、`after.json`：实际手势序列前后快照。
- `gestures.log`：日志事件包括 gesture-start、gesture-frame、gesture-release 和 gesture-end；每帧包含鼠标按下状态及真实 COM 几何读回。
- `undo-01.json` 至 `undo-NN.json`：每一步原生撤销后的完整快照。
- `redo-01.json` 至 `redo-NN.json`：每一步原生重做后的完整快照。
- `before-restart.json`、`after-restart.json`：保存并重启前后快照，进程 ID 应不同。
- `acceptance.json`：检查脚本生成的最终机器结论；本轮为 `passed=true`。

复核命令：

```powershell
python scripts/VerifyLayoutEvidence.py
```

检查脚本要求六种约束分别由至少两个不同形状驱动，每次有效手势至少有两帧 `mouseDown=true`。它检查每帧约束、无关尺寸或坐标保持、完整撤销和重做顺序，以及重启后的几何和元数据；缺失证据或陈旧快照会失败。

## 语义和范围

同 X、同 Y 指左侧和顶部坐标相同。同宽、同高分别传播对应尺寸。约束可以组合，并通过连通分量传播。

等间距建立时记录按轴坐标排序的成员顺序，之后按固定成员顺序计算有符号边缘间距。其数值可以变化，也可以为负而产生重叠。没有固定首尾或固定间距值；求解在满足驱动目标和等式的条件下，使其余位置的平方位移总和最小。

当前支持同页未旋转、未分组普通矩形的内部拖动及八个缩放点。Ctrl/Alt/Shift 手势由 PowerPoint 原生处理，不执行插件约束；键盘和位置/尺寸面板编辑的联动尚未实现。

v0.2 使用 `urn:ppt-gcs:v2` XML 格式，不读取或迁移旧 v0.1 约束。v0.1 已完成的双矩形验收保留在 [历史验证记录](validation.md)，不能直接当作 v0.2 的通过证据。

日志中的实时帧是实际几何写回采样，不是渲染 FPS 或端到端鼠标延迟测试。
