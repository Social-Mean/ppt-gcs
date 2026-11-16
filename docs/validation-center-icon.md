# 中心重合图标修正

日期：2026-11-16；版本：v0.5.2。

用户要求“一个矩形中间一个圆形”。此前仅凭“效果”控件的名称选择 `ThemeEffectsGallery` 不足以验证外观：本机 PowerPoint 实际返回的是只有圆形的资源。

现在保留 Office 原生圆形，按尺寸添加灰色矩形外框。窗格直接使用合成图，Ribbon 的 `getImage` 回调使用同一图像，通过拥有 DIB 句柄的 `IPictureDisp` 返回；半透明边缘不先与背景混色。图片与 COM 句柄由应用级缓存统一释放。

最终样本：[16 像素](../artifacts/center-icon-correction/chosen-center-16.png)、[32 像素](../artifacts/center-icon-correction/chosen-center-32.png)、[64 像素](../artifacts/center-icon-correction/chosen-center-64.png)。[Ribbon 图片回传样本](../artifacts/center-icon-correction/ribbon-center-32.png)与 32 像素样本的 SHA256 相同。

验证结果：

- 完整构建与 236 项检查通过，包括外框、透明间隙、RGBA 原样回传和图片缓存复用。构建过程中，窗格测试曾在全部断言通过后，于进程退出阶段触发 `UIAutomationCore.dll` 异常；最终完整重跑正常退出，未增加绕过或重试逻辑。
- 最终 DLL 已在真实 PowerPoint 进程中载入；25 个图标均加载，中心重合使用带外框资源，Ribbon 定义使用对应图片回调，见 [加载诊断](../artifacts/center-icon-correction/final-diagnostics.json)。
- 重开原文稿后，对象 GUID、名称、几何和约束 XML 与备份一致，保留 16 条关系。备份和文稿快照仅保存在用户本地会话目录。
- [380 像素窗格组件渲染](../artifacts/center-icon-correction/grid-final-380.png)、[250 像素窄窗格](../artifacts/center-icon-correction/grid-final-250.png)及[横向滚动](../artifacts/center-icon-correction/grid-final-250-right.png)通过；这些是组件离屏渲染，不是 PowerPoint 桌面截图。

本次未取得可用的桌面截图，不将资源导出、组件渲染或回调定义检查表述为原生 Ribbon 屏幕显示的人工验收。
