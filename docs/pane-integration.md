# 图标选择与原生窗格切换区

日期：2026-11-16；当前版本 v0.5.2。

## 指定的 Office 图标

“固定宽高比”使用 `PowerPointPageSetup`。“中心重合”以 `ThemeEffectsGallery` 为基础添加矩形外框。标识取自 [微软当前 PowerPoint 控件清单](https://github.com/OfficeDev/office-fluent-ui-command-identifiers/blob/main/Microsoft%20365/Current%20Channel/powerpointcontrols.xlsx)，分别对应设计选项卡中的“幻灯片大小”和“变体 → 效果”。

v0.5.1 只核对了 `GetLabelMso` 名称和图标加载，未验证图形匹配。本机导出的 `ThemeEffectsGallery` 只有圆形，缺少用户要求的矩形外框。v0.5.2 补齐外框，并让功能区和窗格共用同一图像；236 项检查通过，实际资源、组件渲染和加载诊断见 [修正验证](validation-center-icon.md)。

只替换图标资源，约束名称、参数和执行逻辑保持原意，不执行被借用图标的 PowerPoint 命令。布局渲染脚本将历史图标样本绑定到采集时的真实 ImageMso 标识，避免把旧宽高比样本误当作新图标。

## 图示右侧垂直切换区的能力边界

微软将这个宿主界面称为 [Tabbed Panes](https://learn.microsoft.com/en-us/officeupdates/semi-annual-enterprise-channel-archived)。当前插件使用 COM `CustomTaskPane`；公开接口没有提供将自定义窗格注册进原生切换条、设置该条图标或指定与某个原生窗格共用堆叠分组的扩展点。

- [`CustomTaskPane` 完整成员](https://learn.microsoft.com/en-us/office/vba/api/overview/library-reference/customtaskpane-members-office) 提供 DockPosition、DockPositionRestrict、Visible、Width、Height、Window 等属性和相关事件。
- [`ICTPFactory.CreateCTP`](https://learn.microsoft.com/en-us/office/vba/api/office.ictpfactory.createctp) 的参数只有 ActiveX 控件标识、标题和父窗口。
- Office.js 的 [`TaskpaneId`](https://learn.microsoft.com/en-us/javascript/api/manifest/action?view=common-js-preview) 用于复用插件自身的容器；`SupportsPinning` 属于 Outlook，也不能作为加入 PowerPoint 原生切换条的接口。

因此，当前实现保留公开支持的独立停靠/浮动窗格，由 Ribbon 和轮盘打开。本次没有通过私有窗口重挂接模拟原生堆叠。

这个结论限定为“没有可指定该行为的公开 API”。文档没有承诺自定义 COM 窗格会自动进入原生切换条；本次也未对所有 Office 版本的宿主自动归组行为做实机验证，不据此断言任何版本都绝不可能自动收纳。
