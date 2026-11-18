# 幻灯片与容器约束验证

2026-11-18 在 Windows 桌面 PowerPoint 中验证。相对幻灯片和相对容器各有六类关系，包括四边边距和水平、垂直中心偏移。全部数值使用磅（pt）。容器关系将显式选择的容器编号保存在参与者第一位，一条关系可覆盖多个成员。页面宽高从打开的演示文稿读取，XML 仅保存关系参数与成员。

## 构建和宿主检查

`Build.ps1` 通过 551 项自动检查。新增核心测试核对两侧边距推导尺寸、用新页面宽高求解同一 XML、根容器移动和缩放、两个成员保持内边距、重复成员集合和循环引用拒绝。窗格测试核对显式容器选择、刷新保持草稿、带单位的页面尺寸输入和冲突反馈。

安装构建产物并打开 PowerPoint 后，在项目根目录运行以下检查。测试文稿、原生几何快照及结果分别写入各自的运行目录。

```powershell
powershell.exe -NoProfile -STA -File .\scripts\TestReferenceConstraints.ps1
powershell.exe -NoProfile -STA -File .\scripts\TestRibbonConstraints.ps1
```

- 参考约束脚本校验 12 类关系的数值，覆盖三页文稿。页面尺寸从 960 × 540 改为 1040 × 610 磅时，两侧边距仍各为 45 磅，多成员保持边距，第三页的绝对固定位置也保持。自由对象按 PowerPoint 的页面调整行为随画布平移。一次原生 Undo 恢复全部页面的尺寸和位置，一次 Redo 恢复变更，保存重开后全部几何、编号和 XML 匹配。
- 将测试页请求宽度改为 72 磅时，两侧各 45 磅的边距无法同时满足。命令在写入前拒绝并报告冲突，页面尺寸、几何、XML 与撤销历史均保持原状态。
- 功能区脚本检查实际加载项的 XML 和共享分派入口，32 类关系均通过创建、成员编号以及逐次原生 Undo、Redo。21 类数值关系还核对确认参数前无写入。六类容器关系指定选择中的最后一个对象作为容器，验证顺序写入正确。本机运行用 `-OutputDirectory` 将报告保存在 `%LOCALAPPDATA%/PptGcs/Sessions/frame-constraints-20261118/ribbon-ui/`。

## 实际界面操作

使用 Windows 原生窗口截图读取画面，在 PowerPoint 窗口实际发送鼠标与键盘输入。以下几何来自宿主读回，历史恢复误差在 0.005 磅以内。

1. 从透明容器边框内侧拖动，Frame、A、B、C 各沿两轴移动约 29.915 磅。拉右侧控制点使 Frame 加宽约 69.231 磅，A、B 保持 22/34 磅左右内边距并等量变宽，C 保持中心偏移。移动和缩放各自只占一步撤销和重做。
2. 点击功能区的 **幻灯片尺寸…**，在窗格输入 1100 磅并应用。由第二页触发修改，第一、三页也按自己的约束重新求解。通过快速访问栏分别撤销、重做，核对全份文稿的尺寸和对象状态。
3. 从功能区的 **相对容器** 菜单打开左内边距，显式选择 Frame，为 C 输入 206 磅并提交。读取到 Frame 为参照，C 为成员。一次撤销恢复几何并移除关系，一次重做恢复两者。
4. 保存后退出 PowerPoint，启动新进程并打开同一文件。检查页面尺寸、对象身份、关系 XML 和几何。再次拉宽 Frame 约 35.898 磅，A、B、C 都按保存的关系变宽。随后拉高约 29.915 磅，A 保持上边距，B 跟随底部，C 随垂直中心移动约 14.957 磅。两次缩放分别单步撤销，恢复重开时的布局。

原生截图和状态位于 `%LOCALAPPDATA%/PptGcs/Sessions/frame-constraints-20261118/8F714DD8-8662-4FF4-88E9-55598AAA8D78/`，其中 `ui-result.json` 汇总实际界面检查，`container-rule-created.png` 记录窗格创建结果，`reopened-resized.png` 记录新进程中再次拉宽的画面。

保持约束的全局页面调整通过 **PPT GCS → 幻灯片尺寸…** 完成。它使用可读写的 PowerPoint 原生页面宽高属性。PowerPoint 自带的 **设计 → 幻灯片大小** 使用 Office 自己的尺寸与内容缩放规则。[SlideWidth](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pagesetup.slidewidth) · [SlideHeight](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pagesetup.slideheight) · [更改幻灯片大小](https://support.microsoft.com/en-us/powerpoint/change-the-size-of-your-powerpoint-slides)
