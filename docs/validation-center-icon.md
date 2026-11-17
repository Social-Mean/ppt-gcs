# 中心重合内置图标

日期：2026-11-17；版本：v0.5.3。

按用户最终指定，中心重合使用 `DiagramTargetInsertClassic`。功能区直接设置 `imageMso`，窗格和轮盘通过 `GetImageMso` 使用同一个标识，按尺寸缓存 Office 原始图像。旧的效果图标补框、专用 `getImage` 回调及图片传输代码已移除。

真实 PowerPoint 进程内已成功导出 16、32、64 像素图像，仓库保留[16 像素](../artifacts/center-icon-correction/diagram-target-16.png)和[32 像素](../artifacts/center-icon-correction/diagram-target-32.png)样本。候选图标探测、截图裁切及资源调查均留在本机私有会话，未作为插件资源或实验代码提交。

验证结果：

- 完整构建与 232 项检查通过；旧补框和自定义图片传输的四项检查随其代码移除，原生像素透明度、边缘合成和缓存检查继续通过。
- 用户保存并关闭 PowerPoint 后已安装新版；新 PowerPoint 进程载入 v0.5.3，25 个图标全部加载，中心映射及 Ribbon 的 `imageMso` 均为指定标识，见[加载诊断](../artifacts/center-icon-correction/final-diagnostics.json)。
- 重开用户保存后的原工作文稿，载入 7 页、26 个带编号对象和 16 条约束；检查期间文件 SHA256 保持不变，Saved=-1，独立窗格右侧停靠属性正常。
- [380 像素窗格组件渲染](../artifacts/center-icon-correction/grid-final-380.png)、[250 像素窄窗格](../artifacts/center-icon-correction/grid-final-250.png)及[横向滚动](../artifacts/center-icon-correction/grid-final-250-right.png)通过；这些是组件离屏渲染。

本次未进行原生 Ribbon 屏幕显示的人工验收，不将资源导出和组件渲染表述为桌面截图。
