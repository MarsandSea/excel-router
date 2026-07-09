<div align="center">

# ExcelRouter · Excel 智能拆分工具

**整个文件夹一键拆完：按部门、区域、工号等字段自动拆分，打包分发**

保留原格式 · 跨文件自动合并 · 单个文件也能拆 · 可同时拆到每个人

<br>

[![立即下载](https://img.shields.io/badge/%E2%AC%87%20%E7%AB%8B%E5%8D%B3%E4%B8%8B%E8%BD%BD-%E5%85%8D%E8%B4%B9%20%C2%B7%20%E5%85%8D%E5%AE%89%E8%A3%85%20%C2%B7%20Windows-1E7F4B?style=for-the-badge)](../../releases)

**[📖 3 分钟上手指引](docs/使用指引.md)** · **[❓ 常见问题](docs/FAQ.md)** · [下载哪个文件？](#-直接下载使用无需安装-python--download)

🔒 免费开源（MIT）· 数据仅在本机处理，不上传任何服务器 · 零编程，面向普通办公人员

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

*Batch-split a whole folder of Excel files by any field (department / region / ID), keep original
formatting, merge across files, optional per-person output — free, open-source, no coding required.*

作者 / Author：**AbeLin** · 觉得好用请点 ⭐ Star！

</div>

---

## ✨ 功能特点 / Features

- **选一个字段就能拆** —— 自动识别表头，下拉选「拆分字段」，每个取值拆成一个文件，无需预先列出。
  *Pick one column and split — auto-detect header, choose a column, one file per value.*
- **三步卡片式界面** —— ①选表格 ②选字段 ③开始拆分，主按钮固定在底部；不常用的设置收进
  「▸ 高级设置」默认折叠，界面单屏不用切标签。
  *Three-step card UI — pick table → pick field → split; advanced options collapse by default.*
- **跨文件合并** —— 同一个取值出现在多个源文件里，自动**合并到同一个文件**（修掉了旧版互相覆盖的问题）。
  *Cross-file merge — same value across multiple files is merged into one.*
- **二级拆分（按人分发）** —— 可选再按第二列细分（如 部门 → 姓名），一次产出汇总 + 到人双份结果。
  *Optional secondary split (e.g. Department → Person), producing both summary and per-person outputs.*
- **保留格式** —— 表头与数据行的字体、颜色、边框、数字格式、合并表头完整保留。
- **智能识别表头** —— 表头不在第一行也能自动找到（前 15 行启发式扫描）。
- **公式显示真实值** —— 读公式缓存值，并对「未计算的公式」提前预警。
- **多 Sheet / 兼容 .xls / 批量递归** —— 一次处理整个文件夹。
- **大文件不假死** —— 几万行的单文件也有实时进度与心跳日志，界面全程响应（见 [FAQ](docs/FAQ.md#处理大文件几万行时界面卡住像死机了一样是不是崩溃了)）。

> **关于格式保留的两点限制 / Limitations：**
> ① `.xls` 转换后无法保留原格式（仅保留数据）；
> ② 数据区的合并单元格暂不保留（表头的合并单元格正常保留）；
> ③ 跨文件合并按**列位置**追加，最适合「同一套模板的多个表」。
>
> 更多问题见 **[FAQ](docs/FAQ.md)**（杀毒软件误报怎么办、公式列为什么是空的、.xls 支持范围等）。

---

## 🖼️ 演示截图 / Screenshots

**数据源（3 行复杂表头——大标题 / 分组标签 / 列名，工具自动识别表头行）**

![数据源示例](docs/screenshot_source.jpg)

**拆分结果（按区域自动分组，每组含汇总 + 到人，整包 ZIP 可直接发负责人）**

![输出结果示例](docs/screenshot_output.jpg)

---

## 🚀 直接下载使用（无需安装 Python）/ Download

普通用户请直接下载打包好的程序：👉 **[前往 Releases 下载](../../releases)**

每个版本提供两种产物，**优先选 ZIP**：

| 产物 | 说明 | 适用场景 |
|---|---|---|
| `ExcelRouter-vX.X.X-win64.zip` | 文件夹形式，解压后双击里面的 exe | **推荐**，启动更快，极少触发杀毒软件误报 |
| `ExcelRouter-vX.X.X.exe` | 单文件版，下载即用无需解压 | 图方便，但个别杀毒软件可能误报（[why?](docs/FAQ.md#杀毒软件误报)） |

---

## 🧭 三步上手 / Quick Start

界面就是三张卡片，从上到下做完即可：

1. **①选择要拆的表格** —— 点「📄 选一个 Excel 文件」或「📁 选整个文件夹（批量拆）」。
2. **②按哪个字段拆分** —— 点「🔄 扫描字段」，在「拆分字段」下拉里选（如「部门」）；
   选文件夹时还能勾「同时拆到人」，再单独按第二列（如姓名）产出个人文件。
3. **③开始拆分** —— 输出位置不用改（自动放进「拆分结果」文件夹），点「▶ 开始拆分」，
   完成后每个取值一个文件，自动打开输出目录。

不常用的设置（表头识别策略、只拆部分取值、取值归并、跳过值等）收在
**「▸ 高级设置（一般用不到）」**里，默认折叠，一般流程用不到点开它。

想试一下？仓库自带样本：

```bash
python examples/make_sample.py   # 生成 5 个月份的虚拟员工明细（1月A分公司明细.xlsx … 5月A分公司明细.xlsx）
```

用「📁 选整个文件夹」，输入选 `examples/`，拆分字段选「所属部门」，勾选「同时拆到人」，
点「▶ 开始拆分」——每个部门的 5 个月数据自动**跨文件合并**，同时产出按员工姓名细分的个人文件，
整体打包成 ZIP 可直接发给对应负责人。

---

## 🛠️ 开发者运行 / Run from source

```bash
git clone https://github.com/MarsandSea/excel-router.git
cd excel-router
pip install -r requirements.txt
python main.py
```

运行测试：

```bash
pip install pytest
pytest -q
```

---

## 📦 自行打包 / Build

直接运行 `build.bat`（onedir + onefile 双产物，已含必需参数），或参考
**[发版手册](docs/RELEASING.md)** 了解 CI 自动发版流程与防误报细节。

> ⚠️ `--collect-all customtkinter` 与 `--add-data "config;config"` 两个参数缺一不可，
> 否则打包后的 exe 会启动崩溃或找不到默认配置。

---

## 📂 项目结构 / Structure

```
excel-router/
├── main.py                  # 入口
├── config/default_config.json
├── core/
│   ├── splitter.py          # 核心拆分逻辑（表头识别 / 列枚举 / 跨文件合并）
│   └── utils.py             # 文本清理 / 取值归并 / 文件名净化
├── gui/app.py               # 三步卡片式图形界面
├── examples/make_sample.py  # 样本生成器
├── docs/                    # FAQ / 发版手册 / 截图
└── tests/                   # pytest 测试
```

---

## ❓ 遇到问题 / Support

先看 **[FAQ](docs/FAQ.md)**，多数问题（杀毒误报、公式列空白、.xls 限制等）都有解释。
没解决再提 [Issue](../../issues)，模板会引导你附上必要信息，处理更快。

**想提建议？** 程序内点「💬 反馈建议」可匿名反馈（1 分钟，不需要注册任何账号），
你的真实使用场景是这个工具迭代的主要依据。

---

## 📄 开源协议 / License

[MIT License](LICENSE)。可自由使用、修改、分发，请保留原作者署名。

## 👤 作者 / Author

**AbeLin** · 有问题欢迎提 [Issue](../../issues)
