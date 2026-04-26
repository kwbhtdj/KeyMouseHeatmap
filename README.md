# KeyMouse Heatmap

KeyMouse Heatmap 是一个 Windows 键盘和鼠标使用热力图工具，用来统计按键、鼠标点击、滚轮、鼠标移动距离、输入速度、峰值时段、应用按键使用情况，并通过主窗口和悬浮窗实时显示。

当前版本：v1.6.1

## 主要功能

- 键盘热力图：真实 104 键布局，支持左右 Shift / Ctrl / Alt / Win 分开统计。
- 鼠标热力图：支持左键、右键、中键、侧键、滚轮上滚/下滚。
- 鼠标移动距离：按屏幕像素估算并显示为 m / km。
- 总览页：同时展示键盘和鼠标统计。
- 排行页：按键次数排行，并显示占总量的统计条。
- 分组页：字母区、数字区、功能键区、方向键区、小数字键盘、修饰键、鼠标等分组统计，并显示统计条。
- 应用详情：按应用汇总按键使用，一应用一行，显示应用图标、总次数、按键种类、占比统计条、常用按键摘要。
- 峰值页：按小时统计键盘和鼠标峰值。
- 速度页：显示当前分钟、当前小时和平均输入速度。
- 悬浮窗：可显示最近按键、按住状态和输入速度。
- 导出：支持 CSV、PNG、HTML 报告。
- 主题：支持浅色、深色、赛博蓝、极简。

## 数据位置

从 v1.5.0 起，数据固定保存到：

```text
%LocalAppData%\KeyMouseHeatmap\data
```

这样不同版本、不同发布包会共用同一份数据，更新程序不会清空原来的统计。

首次启动新版本时，程序会自动尝试把旧版本 exe 同目录下的 `data/*.json` 迁移到上面的固定目录。迁移是复制，不会删除旧数据。

## 打包和运行方式

项目提供多种发布方式，运行：

```bat
build_publish.cmd
```

会生成：

```text
publish\framework-dependent-win-x64
publish\self-contained-win-x64
publish\singlefile-win-x64
publish\self-contained-win-x86
```

推荐给普通用户使用：

```text
publish\singlefile-win-x64\KeyMouseHeatmap.exe
```

这个版本是 self-contained，不需要用户安装 .NET 9。

如果用户是 32 位 Windows，使用：

```text
publish\self-contained-win-x86
```

如果用户已经安装 .NET 9 Desktop Runtime，可以使用更小的：

```text
publish\framework-dependent-win-x64
```

## 编译

需要 .NET 9 SDK：

```powershell
dotnet restore
dotnet build
```

发布：

```powershell
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

## 更新记录

### v1.6.1

- 在 hook 层过滤长按产生的重复 KeyDown，重复按键不再进入 WinForms UI 线程，降低长按按键时鼠标移动卡顿。
- 后台、最小化到托盘时不再执行主窗口 UI 重绘，只保留统计和必要保存。
- 鼠标移动距离刷新降频到最多 250ms 更新一次界面，减少小窗口热力图频繁重绘。
- 分组页只在界面可见时即时刷新，后台统计不再触发列表重绘。
- 版本号更新为 v1.6.1。

### v1.6.0

- 应用详情第一列改为手绘应用图标和应用名，修复 OwnerDraw 下图标不显示的问题。
- 新增程序版本号，状态栏左侧显示当前版本。
- 项目文件写入 `Version`、`AssemblyVersion`、`FileVersion`。
- README 重写为正常中文，并补充数据目录、打包方式和更新记录。

### v1.5.0

- 数据目录迁移到 `%LocalAppData%\KeyMouseHeatmap\data`，多个版本共用同一份统计数据。
- 启动时自动迁移旧版本 exe 同目录下的 `data/*.json`。
- 排行、应用详情、分组统计条改为占总量百分比，而不是相对最大项百分比。
- 分组页新增统计条。
- 前台应用识别增加 250ms 缓存，减少按键高频查询进程信息。

### v1.4.0

- 应用详情改为一应用一行，显示总次数、按键种类、常用按键摘要。
- 应用详情常用按键列加宽，摘要扩展到 Top 50。
- 排行页和应用详情新增统计条。
- 列表页刷新节流，减少按键时 UI 重绘。

### v1.3.0

- 鼠标移动距离显示为 m / km。
- 鼠标移动距离统计改为定时读取 `Cursor.Position`，不再用全局 mouse move hook，改善鼠标移动丝滑度。
- 鼠标热力图底部说明改为两行并预留空间，避免文字被遮挡。

### v1.2.0

- 新增应用按键统计。
- CSV 和 HTML 报告补充应用统计与鼠标移动距离。
- 应用详情支持按任意日期和范围汇总。

### v1.1.0

- 调整总览布局，减少键盘热力图、鼠标热力图标题遮挡。
- 优化鼠标热力图区尺寸。

## 隐私说明

本工具只统计：

- 按键名称
- 鼠标按键名称
- 次数
- 时间分布
- 应用进程名和可执行文件路径（用于应用详情图标和应用统计）

本工具不记录：

- 具体输入文本内容
- 密码
- 聊天内容
- 网页内容
- 剪贴板内容

所有数据默认保存在本机 `%LocalAppData%\KeyMouseHeatmap\data`。

## 注意

本工具使用 Windows 全局键盘 / 鼠标 Hook 统计输入事件。部分安全软件可能会提示风险，这是输入统计工具常见情况。

请不要将本工具用于未经允许的监控、记录或侵犯他人隐私的场景。
