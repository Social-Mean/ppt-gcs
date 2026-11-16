# v0.4.1 文本交互、约束按钮与图标验证

日期：2026-11-16；Windows、64 位桌面 PowerPoint、.NET Framework 4.8。

## 修复行为

文本框内部的完整按下/松开消息交给 PowerPoint，因此首次点击可以进入原生文字输入流程，空白区可定位光标、可双击和拖动选字。选中的文本框只在边框带接管移动，缩放点优先；已有文字编辑状态下也区分内部与边框。未选中的文本框边缘先交给原生选择。没有合成点击或重放鼠标事件。

根因是旧流程即使未越过移动阈值也吞掉整个首次点击，单独放过双击消息不足以恢复原生点击序列。交互遵循 PowerPoint 的 [点击内部输入文字](https://support.microsoft.com/en-gb/powerpoint/add-text-to-a-slide) 与 [拖动边框移动文本框](https://support.microsoft.com/en-au/office/graphics-visuals/move-a-text-box-wordart-or-shape) 模式。

新建约束移除类型下拉框，使用每列四个、共四行五列的原生按钮，图标和名称同时显示；支持键盘、选中高亮、人数/对象资格置灰、修改关系锁定类型，以及参数值保持。窄窗格只横向滚动类型区，编辑器高度随换行内容调整，创建/取消按钮不会被裁切。

## 图标像素证据

取自本机 PowerPoint 进程的原始样本和统计位于 [icons-before](../artifacts/ui-refinement/icons-before/samples.json)。四个命令各取得 16、24、32 像素样本；导出方法仅用于本次诊断，未保留在正式 API 中。

Office 的图像实际为 straight BGRA。宽高比图标 16/24/32 档的半透明像素分别有 44/76/100 个，全部存在 RGB 分量大于 alpha。将它误标为预乘 alpha 会错误合成边缘；现在用 `Format32bppArgb` 保留原始 RGB/alpha。32 位全透明图像也保持透明，不把零 alpha 强制改为不透明。

水平等距图标的 16 和 32 档没有半透明插值像素，24 档有 170 个。轮盘基础尺寸改为 32 逻辑像素；窗格采用 16。两者按设备 DPI 向 Office 请求尺寸并原样绘制，窗格在原生宿主附着后读取实际 DPI，DPI 变化时重新取图。没有增加锐化或其他图像滤镜。

真实接口的 [25 个命令图标诊断](../artifacts/ui-refinement/office-icons.json) 全部加载成功。11 项像素测试使用实际 Office 边缘样本，验证浅色/深色背景合成、透明背景、缓存和释放。

## 自动检查与布局复核

最终构建通过 159 项检查：42 Core、11 Identity、11 NativeConstraintStore、29 Pane、18 Radial、23 ShapeSupport、11 OfficeIcon、14 TextPointer。

窗格检查覆盖按钮纵向顺序、每列四个、选中切换、资格、锁定类型、刷新保持参数、窄窗格内容高度。文本策略检查覆盖首次/再次点击、已有文字选择、四边拖动、普通形状/图片拖动和 DPI。它们是离线组件与指针策略检查，不代表真实鼠标手势已完成。

布局复核使用最终加载项 DLL 创建独立、不可见的 WinForms 组件并离屏渲染，不截取或操作 PowerPoint 屏幕。380 像素和 250 像素窄窗格分别检查按钮区和完整参数区；最终图片为 [380 宽](../artifacts/ui-refinement/grid-final-380.png)、[250 宽](../artifacts/ui-refinement/grid-final-250.png) 和 [窄窗格右滚](../artifacts/ui-refinement/grid-final-250-right.png)。[RenderConstraintGrid.ps1](../scripts/RenderConstraintGrid.ps1) 可复现检查；只对本次实际导出的四种图标按相同 ImageMso 映射，其余图标留空，实际 25 个映射的完整加载由 Office 进程内诊断检查。

## 真实文本交互状态

实际 PowerPoint 已加载 v0.4.1 并打开独立的 [interaction.pptx](../artifacts/ui-refinement/interaction.pptx)。A、B 两个文本框已有同 Y、同宽关系；已请求人工在 A 输入“测试”，再分别拖 A、B 边框移动。

**截至本次提交，人工操作尚待反馈，不能宣称这项真实鼠标验收已通过。** [最后只读状态](../artifacts/ui-refinement/manual-status.json) 记录尚无文字输入或位置变化。[VerifyTextInteraction.ps1](../scripts/VerifyTextInteraction.ps1) 会先核对输入文字、原生文字路由、两个驱动的完整真实手势、逐帧残差和连续几何，再检查逐次原生 Undo/Redo；预检不通过时不撤销。记录只截取该测试时段，不发布完整用户会话日志。

用户原工作文稿已先保存并在本机另存备份，私人文稿与备份不进入仓库。测试文稿保持打开以便继续人工检查。
