# v0.4.2 文本框边框命中修复

## 已观察到的失败

用户在 v0.4.1 实际点击输入并拖动两个文本框，反馈“边框拖动有时未联动”。原始时段日志为 [failed-interaction.log](../artifacts/text-border-fix/failed-interaction.log)，摘要为 [failure-observed.json](../artifacts/text-border-fix/failure-observed.json)，[最终状态](../artifacts/text-border-fix/failed-state.json) 保留该测试文稿的文字与几何。

六次漏接的起点在外接框外约 0.42–3.46 设备像素，原生已命中边框，但 `HitBody` 只接受外接框内部，所以整段移动落到原生，另一对象没有跟随。例如光标 `(1006,736)` 距 A 下边框外沿 3.416 px；A.Top 从 190 变为 159.707，B.Top 仍为 190。

另一次未选中边框命中位于内侧约 3.55 px，旧策略把所有未选中文本框点击交给原生，也漏掉了首次边框拖动。文字输入本身已进入原生编辑，这次失败没有按成功验收处理，也没有盲目撤销混杂的原生历史。

## 修复规则

缩放点仍优先。文本框边框按外接框内外各 6 DIP 命中，内部文字输入区域保持原生完整点击序列。

- 已选对象：边框命中时，允许原生命中本对象或原生命中为空；后者覆盖外侧选框。
- 未选对象：必须由 PowerPoint 原生命中本对象，才允许直接从边框起拖。
- 原生命中另一个顶层对象：不接管选中对象的边框，避免抢占覆盖对象。
- 普通形状不扩大命中范围；正在编辑形状文字时采用同一边框规则。

每次只查询一次原生命中，不合成或重放鼠标消息。

## 检查范围

最终构建的 192 项离线检查通过：42 Core、11 Identity、11 NativeConstraintStore、29 Pane、18 Radial、23 ShapeSupport、11 OfficeIcon、47 TextPointer。

47 项文本策略检查包含真实失败坐标、四边内外带、6 px 精确边界、边界外空白、已选/未选与本对象/空白/其他对象六种组合、125%/150% DPI 以及小文本框的文字内区。真实日志的屏幕比例为 1.85，测试直接使用对应设备坐标，不把缩放百分比与像素比例混同。

[原生命中报告](../artifacts/text-border-fix/native-hit-targets.json) 通过 75%/100%/125% 三档缩放、A/B 选中及未选中的 108 个边框和中心样本；原生文字、几何、标签和 XML 在检查前后完全一致。脚本 [TestTextBorderTargets.ps1](../scripts/TestTextBorderTargets.ps1) 不生成鼠标输入。初跑未等待缩放绘制完成，原生 hit 缓存尚未同步，[初跑报告](../artifacts/text-border-fix/native-hit-targets-initial.json) 保留；最终在投影和绘制稳定后检查，全部通过，没有放宽命中判定。

## 复测

独立 [interaction.pptx](../artifacts/text-border-fix/interaction.pptx) 保留用户已输入的文字和最后位置，建立新的人工基线，不把 v0.4.1 的失败轨迹混入修复后的连续历史。

复测要求在 A 追加“复测”，再分别从 A 上边框、B 下边框拖动一次，B 不必预先选中。`VerifyTextInteraction.ps1 -EvidenceDirectory artifacts/text-border-fix -RequiredText 复测` 会验证文字、两个驱动、完整真实手势、逐帧关系、原生 Undo/Redo，并保存该时段的独立报告。

用户完成复测并反馈“两次均实时联动，文字可编辑”。[真实交互验收](../artifacts/text-border-fix/acceptance.json) 为 `Passed=true`、`PhysicalMouseInput=true`：3 次原生文字内部路由，两个不同对象各一次完整移动，562 帧按住鼠标期间更新，逐帧约束均满足。两次原生 Undo 和两次 Redo 精确恢复文字、标签、XML，几何误差在 0.01 pt 内，最终状态已恢复并保存。

[路由复核](../artifacts/text-border-fix/manual-route-audit.json) 进一步确认 A 起拖时处于文字编辑状态；B 起拖前选中的是 A，且光标在 B 下边框外约 2.675 设备像素，两次均为 `owned-Move`。这直接覆盖了已观察到的两个漏接原因。实际输入和逐帧记录为证，不将离线策略或程序化命中样本当作物理拖动。

桌面捕获重试两次均返回 `window capture timed out`，因此本项使用人工输入和原生逐帧日志，不声称完成了桌面截图或所有边框方向的逐项物理遍历。
