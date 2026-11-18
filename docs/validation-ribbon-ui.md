# v0.5.0 四列按钮、完整 Ribbon 与轮盘透明度验证

日期：2026-11-16；Windows、64 位桌面 PowerPoint、.NET Framework 4.8。

## 实现

新建约束类型区改为四列五行，按 Catalog 顺序从左到右逐行填充。保留类型资格置灰、修改关系锁定类型、参数输入保持、键盘顺序及窄窗格横向滚动。Ribbon 入口名为“约束窗格”，原生侧栏标题为“约束”。

对齐标签统一为“左对齐、水平居中、右对齐、顶端对齐、垂直居中、底端对齐”。窗格、轮盘与 Ribbon 使用同一份 Catalog 标签；XML 中仍保留原有英文约束 kind，不改变存储格式或求解含义。

Ribbon 包含全部 20 类按钮，按位置、尺寸、间距、偏移与对称分组。`IRibbonControl.Tag` 将 kind 交给同一入口；11 类无参数约束直接创建，9 类参数约束打开并定位到窗格编辑器，确认前不写关系。标签、图标映射及参数分类共用既有定义。接口依据为微软 [IRibbonControl.Tag](https://learn.microsoft.com/en-us/office/vba/api/office.iribboncontrol.tag)。

## 紫色边框根因与修复

旧轮盘使用洋红 `TransparencyKey`，抗锯齿边缘先与洋红底色混合；混色像素不再等于透明键，所以留下紫色锯齿。旧透明键路径已删除，改用 [UpdateLayeredWindow](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-updatelayeredwindow) 与 `ULW_ALPHA`。

完整透明帧用 `Format32bppPArgb` 绘制并复制到缓存的顶向下 32 位 DIB，通过原生逐像素透明度合成。文字改为 GDI+ 抗锯齿绘制，避免 GDI 文字破坏 alpha；Office straight-alpha 图标由 GDI+ 正确合成到最终预乘帧。不通过换透明键颜色、关闭抗锯齿或二值裁剪掩盖问题。

窗口继续不激活、鼠标穿透，捕获仍归原画布；在扇区变化时重画，帧、DIB 和 DC 复用，关闭释放资源。

## 自动与原生呈现检查

构建通过 232 项检查：42 Core、11 Identity、11 NativeConstraintStore、29 Pane、18 Radial、23 ShapeSupport、11 OfficeIcon、47 TextPointer、15 Ribbon、25 RadialRendering。

Ribbon 检查确认所有枚举各有一个按钮，无重复 ID，标签与图标映射一致，9/11 参数分流正确，管理名称及回调正确。窗格检查确认四列五行和逐行顺序，保留窄窗格及参数编辑断言。

透明渲染在 100%、150%、200% 三档测试：

- 轮盘外 alpha 为零，曲边存在半透明像素，预乘 BGRA 合法，无洋红污染。
- 同一帧可在浅色和深色背景正确合成；完整 Catalog 标签不越出外圆。
- 实际隐藏 HWND 的 `UpdateLayeredWindow` 呈现成功，窗口无透明键，保持不激活与鼠标穿透。
- 12 次重复刷新复用缓存，无 GDI 句柄增长，未改变前台窗口或捕获。

另外用最终 DLL 在所有物理屏幕之外执行真实 `Show → OnShown → WM_PAINT → Present → Hide → Dispose`，[生命周期报告](../artifacts/radial-alpha/native-show-check.json) 确认框架未覆盖逐像素透明度，窗口样式、前台与捕获保持正确，原生 DIB/DC 释放；没有显示屏内窗口或操作 PowerPoint。

[轮盘浅色图](../artifacts/radial-alpha/radial-100-light.png)、[深色图](../artifacts/radial-alpha/radial-150-dark.png) 及透明图来自最终 DLL 的独立渲染组件。使用测试图案验证图像合成，不是 Office 图标全集或 PowerPoint 屏幕截图。

最终 DLL 的 [380 宽四列图](../artifacts/ribbon-ui/grid-final-380.png)、[250 宽](../artifacts/ribbon-ui/grid-final-250.png)、[窄窗格右滚](../artifacts/ribbon-ui/grid-final-250-right.png) 已逐张复核。正常宽度完整显示 20 项、无横条；窄宽度参数和创建/取消不裁切。布局图只映射了实际导出的四种 Office 图标，其余留空；[25 个图标映射诊断](../artifacts/ribbon-ui/office-icons.json) 单独确认实际 PowerPoint 中全部加载成功。

## PowerPoint 中的 20 类入口

[TestRibbonConstraints.ps1](../scripts/TestRibbonConstraints.ps1) 自建 20 页文稿，读取加载项实际提供的 Ribbon XML，再调用与原生 Ribbon 回调相同的入口分流，逐类验证。

[v0.5 报告](../artifacts/ribbon-ui/run-20261116T105550933/acceptance.json) 的 20 类全部通过。9 类参数入口在打开编辑器后，全部几何、标签和 XML 都保持不变，原生窗格可见、标题为“约束”，且类型定位正确；输入参数后才创建。11 类无参数入口直接创建正确类型和成员关系。每类都执行一次原生 Undo 和 Redo，精确比较所有页面的原生几何、身份标签及 XML，共 20 次 Undo、20 次 Redo。

该检查为真实 PowerPoint 中的程序化入口检查，`PhysicalMouseInput=false`、`ActualRibbonClicks=false`，不声称逐个人工点击全部 Ribbon 按钮。此前几何联动、文本边框实际鼠标历史验收见 [v0.4.2 记录](validation-text-border.md)；当前轮盘的像素及 Win32 呈现检查不作为所有中键物理交互都已重新人工遍历的证明。

原工作文稿在备份后重新打开并恢复到原页面，原有几何、标识和约束 XML 已与开始前快照核对一致。私人文稿与备份不进入仓库。
