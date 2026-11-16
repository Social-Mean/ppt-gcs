# 双矩形联动验证记录

日期：2026-11-16。环境：Windows，64 位桌面 PowerPoint 16.0.17932.21056，.NET Framework 4.8。

## 结论

本轮两矩形顶部约束的最小端到端验收通过：双向实时拖动、跟随对象 X 不变、一次原生撤销/重做恢复一次完整拖动、约束随 PPT 保存并在新进程重开后恢复。

- 编译成功，9 项核心测试通过。
- A、B 两次拖动分别记录 23、25 次实际写回，48 帧均为 `mouseDown=True`。每帧两个 Top 在 0.001 pt 的日志精度内相等。
- A 拖动时 B.Left 保持 580 pt；B 拖动时 A.Left 保持 178.4496 pt。两形状尺寸一直为 200 × 110 pt。
- Ctrl+Z 两次、Ctrl+Y 两次，每步都通过实际 COM 几何快照核对。每次撤销恢复完整一次拖动，不是只退回一帧或仅恢复跟随矩形。
- 保存后完全退出 PowerPoint，进程从 47500 变为 12572。重开同一文件，全部位置、尺寸、形状 GUID、Slide ID、约束 XML 一致，插件恢复 1 条约束，两个矩形的命中预检都通过。

机器检查结果：[acceptance.json](../artifacts/acceptance.json)。复核命令：`python scripts/VerifyEvidence.py`。

## 操作与证据

桌面控制工具的截图连续出现 `window capture timed out`，坐标输入出现 `coordinate input geometry is unavailable`。实际鼠标拖动由用户操作；插件记录鼠标状态和每次写回的真实坐标，COM 快照用于交叉核对。Ctrl+Z 和 Ctrl+Y 通过桌面控制工具发送给 PowerPoint，未通过脚本重设坐标来伪造撤销。

所有日志时间为 UTC，下列说明使用北京时间。

1. 测试文件先保存为 [two-rectangles.pptx](../artifacts/two-rectangles.pptx)，关闭此前进程；10:31:31 在进程 47500 中重新加载插件，10:31:37 从该已保存文件恢复约束。该进程内没有重新创建演示或重新添加关系。
2. 10:40:53–10:40:55 拖动 A：两个 Top 从 180 pt 变为约 114.108505 pt，A.Left 从 180 pt 变为约 178.4496 pt。
3. 10:40:56–10:40:58 拖动 B：两个 Top 变为约 203.255829 pt，B.Left 从 580 pt 变为约 567.5969 pt。
4. 撤销 B 后，两个 Top 一起恢复为 114.108505 pt，B.Left 恢复 580 pt；撤销 A 后完整恢复初始几何。随后重做 A、B，分别恢复对应终点。
5. 10:42:27 保存最终状态并退出；10:43:01 在新进程 12572 中重新打开，保存前后快照一致。

成功拖动本身是在第一次“保存 → 退出 → 新进程打开”之后执行，因而验证了持久化约束重新载入后确实能驱动拖动。第二次重启核对了更新后的几何与元数据；第二次重启后未额外重复鼠标拖动。

证据文件：

- [逐帧日志](../artifacts/owned-drag.log)
- [拖动前](../artifacts/owned-drag-before.json)、[拖动后](../artifacts/owned-drag-after.json)
- [撤销 B](../artifacts/undo-b.json)、[撤销 A](../artifacts/undo-a.json)
- [重做 A](../artifacts/redo-a.json)、[重做 B](../artifacts/redo-b.json)
- [最终保存前](../artifacts/before-restart.json)、[新进程重开后](../artifacts/after-restart.json)

PPTX 包内直接检查到了 `customXml/item1.xml` 约束及两个 `ppt/tags/` 形状 GUID 标签；不是外部 JSON 文件代替文档存储。JSON 只用于测试快照。

## 实现与验收边界

最终实现采用线程级 WH_MOUSE 接管受约束矩形内部拖动，在 PowerPoint 主线程更新两个真实形状，并在手势开始/结束建立 Undo 边界。空闲时只读取变化，不重新施加几何，避免破坏原生撤销。

仅验收同一幻灯片、未旋转、未分组的两个普通矩形的顶部约束。缩放、键盘微移、位置面板修改、复制粘贴、多选移动、复杂约束网络和跨平台不在本轮范围。Escape、失焦及捕获丢失的取消路径已实现，但尚未做完整人工验收。

日志帧表示应用实际写回采样，不等于渲染 FPS，也不是鼠标输入延迟基准。此次结果针对上述 PowerPoint 版本和测试场景，不外推到尚未测试的 Office 版本。

## 开发过程中修正的问题

原生拖动只在松手后提交坐标，直接轮询会产生延后联动，并把跟随写回变成单独撤销项，见 `native-polling-failure.*`；该逻辑已移除。

手写 COM 接口 ABI 曾造成加载失败，已删除并改为嵌入官方 PIA。初版误用了不存在的 DocumentWindow.HWND，后改为使用实际鼠标窗口和 RangeFromPoint；`hook-preflight-failure.*` 保留失败证据。最终验收区间没有记录到运行错误。
