# KeyMouse Heatmap

KeyMouse Heatmap 是一个 Windows 键盘 / 鼠标使用热力图工具，用来统计键盘按键、鼠标按键、滚轮、打字速度、峰值数据，并通过主窗口和悬浮窗实时显示输入反馈。

## 功能特点

### 键盘热力图

- 支持真实 104 键键盘布局
- 支持主键区、功能键区、编辑区、方向键、小数字键盘分区显示
- 支持左右 Shift / Ctrl / Alt / Win 分开统计
- 支持按键按下时实时高亮
- 长按按键只统计一次

### 鼠标热力图

- 支持左键、右键、中键
- 支持侧键 Forward / Back
- 支持滚轮上滚 Wheel Up / 下滚 Wheel Down 分开统计
- 支持鼠标按键实时高亮
- 鼠标按键次数直接显示在对应按键上

### 总览页面

- 键盘和鼠标可以在同一个页面同时查看
- 支持窗口大小自适应
- 小窗口下内容会自动缩放

### 统计功能

- 今日统计
- 最近 7 天统计
- 最近 30 天统计
- 全部统计
- 键盘 / 鼠标 / 全部统计模式切换
- 排行页查看按键次数排名
- 分组统计：字母区、数字区、功能键区、方向键、小数字键盘、修饰键、鼠标等

### 速度统计

- 本分钟输入次数
- 当前小时输入次数
- 平均打字速度
- 长按不会重复计入速度

### 峰值图

- 按小时统计键盘和鼠标峰值
- 横轴支持半小时刻度线
- 刻度文字会根据窗口宽度自动精简

### 悬浮窗

- 支持显示最近按键
- 支持显示正在按住的按键
- 支持显示打字速度
- 支持调节悬浮窗大小
- 支持调节背景透明度，文字保持清晰不透明
- 支持置顶显示

### 其他功能

- 最小化到后台
- 托盘运行
- 开机启动
- 自动开始记录
- 导出 CSV
- 导出图片
- 导出 HTML 报告
- 主题切换

## 运行环境

- Windows 10 / Windows 11
- .NET 9 Runtime 或 .NET 9 SDK

如果只运行发布好的 exe，通常只需要 .NET Runtime。  
如果需要自己编译，需要安装 .NET 9 SDK。

查看本机 SDK：

```powershell
dotnet --list-sdks
```

如果能看到类似：

```text
9.0.xxx
```

说明已经安装 .NET 9 SDK。

## 编译方法

进入项目目录后运行：

```powershell
dotnet restore
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=false
```

也可以直接双击项目里的：

```text
build_publish.cmd
```

编译成功后，程序会生成在：

```text
bin\Release\net9.0-windows\win-x64\publish\KeyMouseHeatmap.exe
```

## 打包发布

推荐只打包 `publish` 文件夹里的内容：

```text
bin\Release\net9.0-windows\win-x64\publish
```

不要把下面这些目录一起打包：

```text
bin
obj
data
.vs
.git
```

特别注意：`data` 文件夹里是统计数据。  
如果你把 `data` 一起发出去，别人打开软件时会带着你的旧统计次数。

## 数据存储位置

程序会在 exe 同目录下创建：

```text
data
```

里面保存每日统计数据，例如：

```text
usage-2026-04-26.json
```

如果想清空统计，可以在软件里点击“清空”，或者关闭软件后删除 `data` 文件夹。

## GitHub 上传注意事项

### 1. 不要上传编译产物

不要上传：

```text
bin/
obj/
```

这两个文件夹是编译生成文件，不属于源码。

### 2. 不要上传个人统计数据

不要上传：

```text
data/
```

里面可能包含你的键盘、鼠标使用统计。

### 3. 推荐使用 `.gitignore`

本项目已经附带 `.gitignore`，会自动忽略常见编译产物和统计数据。

### 4. 推荐仓库结构

```text
KeyMouseHeatmap/
├─ Program.cs
├─ KeyMouseHeatmap.csproj
├─ build_publish.cmd
├─ README.md
├─ LICENSE
└─ .gitignore
```

### 5. 发布 exe 的推荐方式

如果想让别人直接下载 exe，推荐使用 GitHub Releases：

1. 打开 GitHub 仓库
2. 进入 Releases
3. 点击 Create a new release
4. 填写版本号，例如 `v1.0.0`
5. 把 `publish` 文件夹压缩成 zip 后上传
6. 发布 Release

这样源码和可执行文件可以分开管理。

## 隐私说明

本工具只统计：

- 按键名称
- 鼠标按键名称
- 次数
- 时间分布
- 打字速度统计

本工具不记录：

- 具体输入文字内容
- 密码
- 聊天内容
- 网页内容
- 剪贴板内容

所有数据默认保存在本地 `data` 文件夹中。

## 杀毒软件误报说明

本工具使用 Windows 全局键盘 / 鼠标 Hook 统计按键次数。  
部分杀毒软件可能会对这类程序产生误报，这是输入统计工具常见情况。

本工具不记录具体输入文本，不联网，不上传数据。

## 常见问题

### 为什么长按只算一次？

为了避免长按时 Windows 自动重复发送按键事件导致统计失真，程序会在按键松开前只统计一次。

### 为什么有些特殊键显示为 VK？

某些键盘厂商的自定义按键没有标准名称，Windows 只能提供虚拟键码，因此可能显示为 VK 开头的名称。

### 为什么打包给别人后统计不是 0？

大概率是把 `data` 文件夹一起打包了。  
重新打包时不要包含 `data` 文件夹即可。

### 为什么编译失败？

先确认安装了 .NET 9 SDK：

```powershell
dotnet --list-sdks
```

然后在项目目录运行：

```powershell
dotnet restore
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=false
```

如果仍然失败，把红色报错截图发给维护者。

## 免责声明

本工具仅用于个人输入统计和可视化分析。  
请勿将其用于未经允许的监控、记录或侵犯他人隐私的场景。
