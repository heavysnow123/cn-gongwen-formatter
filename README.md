# 中文文档智能排版工具

一个开源的中文文档（Word / TXT / Markdown）智能排版工具，适用于公文、报告、论文、
材料、总结等各类中文文档。它能自动识别标题层级、统一字体版式、标准化中英文标点，
并生成可导航的大纲，把格式混乱的文档一键整理成规范排版。公文只是其中一类典型场景，
软件并无公文专属限制——任何需要规范中文排版的文档都能用。

## 功能

- **主/副标题识别**：文档开头连续居中、字体字号相同的段落自动识别为主标题；字号不同者判为副标题。
- **标题层级识别**：
  - 一级 `一、`（中文数字 + 顿号）
  - 二级 `（一）`
  - 三级 `1.` / `1、`
  - 四级 `(1)`
- **图/表标题**：居中且以「图/表」开头的段落自动识别。
- **附件处理**：识别「附件1 / 附件：」并可选段前分页。
- **表格格式化**（可选）：表头加粗、统一边框、自动列宽、智能对齐，行高/字号可调。
- **符号标准化 / 标点全半角**（可选）：中英文标点统一为全角（逗号、句号、冒号、分号、问号、叹号、括号、引号等），
  保留数字内的半角点（如 `1.5`）、千分位逗号（如 `1,000`），英文语境下的全角标点自动转回半角，避免误伤。
- **序号风格统一**（可选）：将 `1.` / `1、` / `（一）` / `(1)` 等行首序号归一为目标风格
  （`一、` / `1.` / `1、` / `（一）` / `（1）`），仅改序号标记、保留正文。
- **中文换行禁则**（可选）：去除禁则字符邻接处的多余空格（降低破版），并在 `settings.xml` 启用 Word 的
  `kinsoku`（避头/避尾）排版规则。
- **处理模式**：
  - **全量排版**（默认）：字体、页边距、页码、表格等全部处理。
  - **仅修标点**：保留原文档的字体与段落格式，只做标点 / 序号 / 禁则的文本修正（对标“只修标点混用”场景）。
- **标题中英字体细分**（可选）：一级至四级标题、主/副标题可分别指定西文字体（数字/字母），留空则沿用全局西文字体。
- **页面背景色**（可选）：通过 `w:background` 设置页面底色（填 hex，如 `F2F2F2`）。
- **TXT / Markdown 导入**：清理 Markdown 标记，支持空行保留/删单等模式，直接生成排版后 docx（强制 A4）。
- **旧格式转换**：`.doc` / `.wps` 经 Word/WPS 或 LibreOffice 兜底转换为 `.docx`。
- **页面与页码**：页边距、A4、居中的页码（PAGE 域）。
- **大纲级别**：自动设置，生成可在左侧导航栏浏览的目录。
- **中文字体检查**：排版前检测系统是否安装常用中文字体（仿宋_GB2312 / 楷体_GB2312 等），
  缺失时醒目提示。**本工具只检查、不安装字体**，由用户自行准备并安装到系统（详见 `fonts/README.txt`）。
- **中文字体下拉选择**：标题 / 正文 / 表格 / 页眉页脚等字段均为下拉框，内置公文常用字体，也可直接输入自定义字体名。

## 文档模板

左侧「快捷工具 → 模板」打开模板面板，一键套用预设版式并生成标准骨架文档：

- **公文**：套用公文排版预设（版心页边距、仿宋正文、黑体层级标题、居中页码），生成红头 + 文号 + 主送机关 + 正文 + 落款的标准公文骨架。
- **报告**：套用报告排版预设（黑体标题层级、宋体正文、页码 + 页眉），生成大标题 + 副标题 + 章节占位骨架。
- **红头**：仅生成红色发文机关标志 + 红线 + 文号占位。
- **信函**：套用信函预设（黑体标题、宋体正文、首行缩进），生成标题 + 称呼 + 正文 + 右对齐落款的信函骨架。
- **通知**：套用通知预设（黑体层级、宋体正文、页眉），生成标题 + 发文单位 + 正文 + 落款的骨架。
- **会议纪要**：套用纪要预设（黑体层级、宋体正文、表格），生成标题 + 会议要素（时间/地点/参会人）+ 正文骨架。

> 模板生成的均为占位文本，请替换为真实内容；红头仅作版式占位，单位印章需自行加盖。
> 模板种类与生成函数集中注册在 `templates.py` 的 `TEMPLATES` / `GENERATORS` 中，新增模板只需补两处映射。

## PDF 导出

左侧「快捷工具 → 导出 PDF」将排版后的 `.docx` 一键导出为 PDF。导出后端自动探测：

1. Microsoft Word（版式与 Word 完全一致）
2. WPS Office
3. LibreOffice（headless）
4. **内置引擎（reportlab，纯 Python）** —— 本工具自带，无需安装任何办公软件即可导出 PDF。

勾选「优先用内置引擎（不依赖 Office）」可跳过外部程序、直接用内置引擎导出；未勾选时
若本机没有 Word/WPS/LibreOffice，也会自动回退到内置引擎，保证一定能导出。
内置引擎以通用排版规则还原字体、字号、对齐、行距、页边距、表格与红头颜色，
并做了分页优化——文档开头的红头 / 标题整体保持同页（不被割裂到两页）、每个表格
整体保持（行不被拆散）。与 Word 的精确分页仍可能有细微差异，但内容、字体与版式要素完整可打印。

## 排版质检

左侧「快捷工具 → 排版质检」对文档做合规检查并给出综合评分与改进建议，涵盖：

- **字体完整性**：扫描文档所用字体，比对系统已安装字体，列出缺失项（常用中文字体缺失会导致版式不达标）。
- **页边距（版心）**：判断是否符合常规 A4 或公文版心（上 37 / 下 35 / 左 28 / 右 26 mm）。
- **最小字号 / 行距 / 页码**：提示过小字号、未设行距、缺页码等隐患。

## 大文件处理

采用 lxml 流式读写引擎处理超大文档（数百页标书、数十 MB 以上）：

- **原理**：用 `etree.iterparse` 边读 `word/document.xml` 边改边写，每处理完一个顶层块
  （段落 / 表格）即序列化输出并释放，**内存恒定，与文件总大小无关**。
- **实测**：6000+ 段落、含表格与分节符的文档，进程峰值内存仅约 1.8 MB。
- **格式完整**：分节符完整保留（横向附录页、页眉/页脚分节不丢）；标题 / 正文 / 表格 / 附件分页
  统一修正照常生效；页眉 / 页脚 / 页码在流式模式下于 zip 层注入，只新增小部件、不重排原包。
- **自动降级**：单文件超过 `large_file_threshold_mb`（默认 50MB）自动走流式模式；
  也可在 GUI「其他」页签手动强制开启。

## 界面与批量处理

- **亮 / 暗主题切换**：GUI「其他」页签的「界面主题」可在「浅色 / 深色」间切换，切换即时生效并自动持久化（存入 `~/.word_formatter_default.json`）。Windows 与国产系统共用同一套设置。
- **高分屏（HiDPI）自适应**：Windows 下启动时启用 `customtkinter.WindowsDPIAware()`，高分屏文字与控件更清晰；国产系统由桌面环境自行处理缩放。
- **目录树批量处理**：顶部「目录树」按钮可递归收纳所选目录下所有子目录的中文文档，输出时**保持原始相对目录结构**（即 `子目录/文件_formatted.docx` 落在输出目录的同名子目录下），适合批量排版一整个资料树。普通「+ 文件夹」仍为平铺（仅当前目录）。

## 中文字体检查说明

中文规范排版常用仿宋_GB2312 作正文、楷体_GB2312 作部分标题；系统缺失时排版软件会回退替代字体，
导致版式不达标。本工具**只检查、不安装**——字体安装属于系统级修改且涉及授权分发，由用户自行完成。

- **检测**：跨平台识别系统字体——Windows 读注册表，Linux / macOS 用 `fc-list`（无则扫描字体目录）；按别名匹配，并兼容国产系统开源替代（FandolFang / AR PL UKai / Noto Serif CJK 等）。
- **触发**：「字体」页点「检查字体状态」查看本机情况；开始排版时也会预检并提示缺失。
- **授权说明**：仿宋_GB2312 / 楷体_GB2312 为微软专有字体，本工具不内置、不安装；
  请使用您已合法获得的字体并自行安装到系统。大标题用的思源宋体（Source Han Serif）为
  SIL OFL 开源字体，可免费获取安装。

## 国产系统 / Linux 支持

本工具核心逻辑（python-docx 排版、reportlab 内置 PDF 引擎）均为纯 Python，**跨平台可用**，
已在以下场景验证思路：统信 UOS、银河麒麟等国产 Linux 发行版（x86_64 / ARM64 / LoongArch）。

- **字体检测跨平台**：自动识别系统字体——Windows 读注册表；Linux / macOS 用 `fc-list`
  （无则扫描 `/usr/share/fonts` 等目录）。除 Windows 专有「仿宋_GB2312 / 楷体_GB2312」外，
  还会匹配国产系统常见开源替代（**FandolFang 方政仿宋、文鼎 AR PL UKai 楷体、
  Noto Serif CJK / 思源宋体**），因此无需强装 GB2312 专有字体也能正确排版与导出。
- **PDF 导出在 Linux 上**：Word / WPS 的 COM 后端在 Linux 不可用（会自动跳过），
  依次回退到 **LibreOffice headless**（国产系统多预装）或**内置 reportlab 引擎**（纯 Python，
  自带 CID 中文字体，零依赖即可出中文 PDF）。导出时可勾选「优先用内置引擎」跳过外部程序。
- **一键运行（无需打包）**：项目根目录 `run_linux.sh` 会自动建虚拟环境、安装运行依赖并启动 GUI：
  ```bash
  chmod +x run_linux.sh
  ./run_linux.sh
  ```
  前置：系统需有 `python3`、`python3-venv`、`python3-tk`（GUI 依赖 tkinter）。
- **旧格式（.doc / .wps）转换**：Windows 走 Word / WPS；Linux 上由 LibreOffice 兜底转换，
  请确保已安装 LibreOffice。
- **打包交付**：Windows 提供单文件 EXE；国产系统提供两种免安装分发：
  - `build_linux.sh` 产出 Linux 单文件 `dist/WordFormatterPro`（无需安装 Python）；
  - `package_appimage.sh` 进一步封装为 **AppImage**（`WordFormatterPro-<架构>.AppImage`，双击即跑，
    已含 Python 与 tkinter，不再依赖系统 Python）。

## 下载

- **Windows 单文件 EXE（无需安装 Python）**：[WordFormatterPro.exe v1.1.0](https://github.com/heavysnow123/cn-gongwen-formatter/releases/download/v1.1.0/WordFormatterPro.exe)
- 所有版本与更新说明：[Releases 页面](https://github.com/heavysnow123/cn-gongwen-formatter/releases)

> 运行前请确认系统已安装「仿宋_GB2312」与「楷体_GB2312」（公文 / 规范中文排版常用字体）；
> 大标题用的思源宋体（Source Han Serif）为 SIL OFL 开源字体，可免费获取安装。

## 安装与运行

```bash
# 创建虚拟环境（需含 tkinter 的 Python 3.11+）
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt

# 启动图形界面
.venv\Scripts\python.exe -m word_formatter.gui

# 命令行
.venv\Scripts\python.exe -m word_formatter.cli --help
.venv\Scripts\python.exe -m word_formatter.cli -o out *.docx
.venv\Scripts\python.exe -m word_formatter.cli --text "一、标题
正文..." -o result.docx
```

## 构建（Windows 版 / 国产系统版）

两个版本**共用同一套源码**（`src/`），功能与界面完全一致，仅在窗口标题 / 关于中
标注「Windows 版」或「国产系统版」，以及各自的打包产物不同。

- **Windows 版**（单文件 EXE，双击即用）：
  ```bash
  .venv\Scripts\python.exe build.py
  # 产物：dist/WordFormatterPro.exe
  ```
- **国产系统版**（统信 UOS / 银河麒麟等 Linux，单文件可执行）：
  ```bash
  chmod +x build_linux.sh
  ./build_linux.sh
  # 产物：dist/WordFormatterPro（Linux 单文件，无需安装 Python）
  ```
  > PyInstaller 不支持交叉编译：x86_64 / ARM64 / LoongArch 需在该架构的本机
  > （或对应架构的容器）上构建。
- **国产系统版 AppImage（免安装单文件）**：
  ```bash
  chmod +x package_appimage.sh
  ./package_appimage.sh
  # 产物：WordFormatterPro-<架构>.AppImage（内含 Python 与 tkinter，双击即运行）
  ```
  > 运行 AppImage 需要 FUSE；若目标机无 FUSE，可执行
  > `./WordFormatterPro-<架构>.AppImage --appimage-extract` 解包后运行 `squashfs-root/AppRun`。

## 目录结构

```
word-formatter-pro/
├── src/word_formatter/
│   ├── config.py     # 全部可配置参数（默认值面向中文公文）
│   ├── core.py       # 排版引擎（与界面解耦，可测试 / 可无头运行）
│   ├── stream.py     # 大文件 lxml 流式引擎
│   ├── legacy.py     # .doc / .wps 转换 + COM 预处理
│   ├── gui.py        # 图形界面
│   ├── cli.py        # 命令行入口
│   ├── templates.py  # 文档模板：排版预设 + 骨架 / 红头生成
│   ├── checker.py    # 排版质检：字体/页边距/字号/页码评分
│   ├── export_pdf.py # PDF 导出：Word/WPS/LibreOffice 后端探测
│   └── __init__.py
├── tests/            # 无头单元测试
├── build.py          # Windows 版构建（产出 WordFormatterPro.exe）
├── build_linux.sh    # 国产系统版构建（产出 Linux 单文件 WordFormatterPro）
├── package_appimage.sh # 国产系统版 AppImage 封装（产出免安装单文件 AppImage）
├── run_linux.sh      # 国产系统一键启动（建 venv + 装依赖 + 启动 GUI）
├── launcher.py       # EXE 启动入口
└── pyproject.toml
```

## 许可证

本项目以 MIT 许可证开源。字体文件各自遵循其原授权（详见 `fonts/README.txt`）。
