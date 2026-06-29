# CLAUDE.md

本文件是给 Claude Code 的项目说明与工作契约。请在每次开始工作前阅读。

---

## 项目是什么

一个 Windows 桌面工具（v2.0「通用拆分」）：把一批 Excel 表格，**按用户指定的任意一列的取值**
拆分成多个文件（每个取值一个文件），可选再按第二列做二级拆分，并保留原始表头格式。
面向不会编程的普通办公人员，界面分「简单 / 进阶 / 专家」三层递进。

> v1.0 原是「网格化管理」专用（按网格 + 工号识别）。v2.0 已**通用化**：表头自动识别、
> 按列名选列、自动枚举取值、跨文件合并。作者原来的网格 + 岗位工作流用「专家模式」
> （header_mode=keyword + 主列=网格 + 二级列=岗位 + value_alias_map 复现岗位归并）即可复现。

**作者：Leo。开源协议 MIT。目标是开源传播 + 建立个人声誉，不收费。**
所有源码文件顶部都有 `Copyright (c) 2025 Leo · MIT License` 署名，**请勿删除任何文件的版权头**。

---

## 运行环境

- **目标平台：Windows**（打包出的 exe 给 Windows 用户）
- Python 3.9+
- 依赖见 `requirements.txt`：customtkinter、openpyxl、pandas、xlrd

---

## 目录结构与各文件职责

```
excel_splitter/
├── main.py                    # 入口，只有几行，调用 gui.app.run()
├── CLAUDE.md                  # 本文件
├── README.md                  # 中英双语说明（含截图占位 docs/screenshot.png）
├── LICENSE                    # MIT 许可证
├── requirements.txt           # 运行依赖
├── requirements-dev.txt       # 开发/测试依赖（pytest、pyinstaller）
├── conftest.py                # 让 pytest 能 import core 包
├── .gitignore                 # 忽略 dist/build/用户配置等
├── build.bat                  # 一键打包脚本（含必需的打包参数，推荐直接用）
├── app.ico                    # 程序图标（占位，用户可替换）
├── version.txt                # Windows 版本信息（PyInstaller --version-file 用，v2.0）
├── config/
│   └── default_config.json    # 默认配置（通用空配置，新字段见下「数据模型」）
├── user_config.json           # 用户保存的配置（运行时生成在程序目录，不入库）
├── core/
│   ├── __init__.py
│   ├── splitter.py            # 核心拆分逻辑（最重要的文件）
│   └── utils.py               # 文本清理 + 取值归并 + 文件名净化
├── gui/
│   ├── __init__.py
│   └── app.py                 # customtkinter 三层递进界面（含打包路径适配）
├── examples/
│   ├── make_sample.py         # 可复现样本生成器（演示跨文件合并）
│   └── sample_团队A/B.xlsx    # 生成的演示样本
├── docs/                      # README 截图 / GIF 资源（待补 screenshot.png）
└── tests/
    ├── test_utils.py          # utils 单元测试
    └── test_splitter.py       # splitter 集成测试（单文件/合并、到人双产出、格式保留）
```

### 数据模型（config 关键字段）

| 字段 | 含义 | 默认 |
|---|---|---|
| `header_mode` | 表头识别：`auto` / `row`(指定行号) / `keyword`(旧关键词法) | `auto` |
| `header_row` | header_mode=row 时的 1 基行号 | `1` |
| `grid_keys` / `id_keys` | header_mode=keyword 时表头需同时含的两类关键词（任一空则不要求） | `[]` |
| `split_column` | 主拆分列（按列名） | `""` |
| `selected_values` | 只拆这些主取值；空 = 自动枚举该列所有取值 | `[]` |
| `to_person` | 是否在汇总之外**附加产出到人**（仅文件夹批量有效） | `false` |
| `person_column` | 到人按哪列拆 | `""` |
| `person_file_filter` | 只对文件名命中这些关键词的表做到人；空 = 全部 | `[]` |
| `make_zip` | 批量时按主取值打包 ZIP | `true` |
| `value_alias_map` | 取值归并 `{规范值:[别名...]}`（旧 position_map 的通用化身） | `{}` |
| `skip_values` | 拆分列中要忽略的取值（合计/小计/空等） | `["合计","小计","总计","平均",""]` |
| `merge_across_files` | 汇总：同取值跨源文件是否合并到一个文件（到人始终按人合并） | `true` |
| `exact_match` `preserve_format` `auto_open_output` | 同义保留 | |

### core/splitter.py 的关键函数

- `detect_header_row_auto(rows, max_scan=15)` —— 通用启发式：非空多、文字多、下一行像数据 → 表头行
- `find_header_row(rows, grid_keys, id_keys)` —— 关键词法（专家兜底，接收**值矩阵**而非 ws）
- `resolve_header_row(rows, config)` —— 按 header_mode 分发 auto / row / keyword
- `list_columns(file_path, config)` / `list_values(file_path, config, column)` —— 供 GUI 下拉与多选
- `detect_uncalculated_formulas(work_path)` —— 检测「有公式但无缓存值」，采样前 200 行预警
- `_cached_copy(style_obj, cache)` —— 按源样式 id 缓存复制，保留格式提速的关键（**缓存按单个源文件作用域**）
- `_OutputBook` —— 一个输出文件的内存工作簿；`get_or_create_sheet` + `_append_rows` 支持**跨源文件追加**
- `_summary_key_path` / `_person_key_path` —— 计算「汇总 / 到人」两套输出树的 key 与路径
- `_matches_filter(name, keywords)` —— 文件名是否命中到人范围关键词
- `normalize_to_xlsx(file_path, log_fn)` —— .xls 转临时 .xlsx
- `process_file(..., single_file, ...)` —— **始终产出汇总、可选附加到人**，追加进共享 outputs 注册表
- `run_split(config, ...)` —— 主流程：**输入可为单文件或目录** → 累积 → 统一保存 → 按主取值打 ZIP

**架构要点：pandas 负责过滤数据（向量化 mask），openpyxl 负责复制格式。**

**一次运行两套产出（v2.1）：**
- **汇总**（始终）：按 `split_column` 拆。单文件 / `merge=True` → 扁平 `{主取值}.xlsx`（跨文件合并）；
  目录 + `merge=False` → `{主取值}/汇总/{原文件名}.xlsx`（原文件拆分）。
- **到人**（可选）：`to_person` 开 + 有 `person_column` + 文件名命中 `person_file_filter` →
  `{主取值}/到人/{姓名}.xlsx`（同一人跨文件按 sheet 名合并）。
- 批量时对每个生成了文件夹的主取值打 `{主取值}.zip`。

**跨文件合并（v2.0 核心，仍在）：** 输出按 key 在内存累积、统一保存；第一个产生某 key 的源文件
定义表头/格式，后续文件**按列位置追加**（到人则按 sheet 名分别累积）。从根本上修掉 v1.0「同名输出互相覆盖」bug。

**数据行两种模式，由 `config['preserve_format']` 控制（默认 True）：**
- **保留格式**：用 `wb_src` 逐行复制「值 + 完整格式」。按【原始行号】读取、加法式写入
  （`rows_df.index` 即原始 0 基行号，`excel_row = src_idx + 1`），**绝不删行**。
- **快速模式**：数据行只写值（来自 pandas），最快。

表头格式（含合并单元格、列宽）两种模式都保留。`wb_src` 在 `process_file` 末尾统一 `close()`。

**限制（已在 README/界面注明）：** ① `.xls` 转换后无法保留原格式；
② 数据区合并单元格暂不保留；③ 跨文件合并按**列位置**追加，最适合「同一套模板的多个表」。

### gui/app.py（v2.1：单屏自适应，已去三层分档）

- 输入用 `CTkSegmentedButton`「单个文件 / 文件夹」单选；`_on_input_type` 控制**批量区** `_batch_frame` 显隐
  （选文件夹才显示「每组 ZIP」「同时到人 + 列 + 关键词」）。**注意**：`_on_input_type` 只在路径与新类型
  不符时才清空，避免清掉启动时已保存的有效路径。
- 「▸ 高级」`_toggle_adv` 折叠 `_adv_frame`：表头识别 / 跳过值 / 取值归并 / 精确匹配 / 跨文件合并 / 保留格式。
- `_detect_columns` 在**子线程**读模板文件（单文件=它本身、文件夹=首个 Excel）、回主线程填主列/到人列下拉。
- `_collect_config()` 产出 v2.1 字段；拆分在子线程跑，`self.after(0, ...)` 回主线程更新 UI。
- 配置存到 `user_config.json`；`load_config` 用 FALLBACK 补齐缺键，旧版配置不会报错。

---

## ⚠️ 已知历史 Bug（重要，勿重蹈覆辙）

早期版本（V33）有过一个**行号错位 Bug**：在删除行之后，用删除后的行号去索引 pandas
镜像 DataFrame，导致公式列（如「到账金额」）取到错误的值。

**当前架构已规避**：本版本不做「先删行再回填」，而是 pandas 直接 `mask` 过滤出要保留的行，
表头格式单独从缓存写入。如果未来要改回「删行」逻辑，**务必先按原始行号缓存镜像数据，
删除后再按缓存顺序回填**，不要用删除后的行号索引。

### ⚠️ v1.0 的「多表覆盖」Bug（v2.0 已修，勿回退）

v1.0 逐文件独立处理、每个输出文件都新建 Workbook 后 `save` 覆盖，输出文件名又不含源表名——
同一个人/取值出现在多个源表时，后一个表直接覆盖前一个，导致数据丢失。
**v2.0 改为按 key 在内存累积、统一保存的 `_OutputBook` 机制**（见上）。
未来若改输出写法，**务必保持「同 key 跨文件追加」语义**，不要回到「逐文件新建并覆盖同名文件」。

## ⚠️ 公式列依赖缓存值（实测确认）

本工具用 pandas 读公式列的**缓存计算结果**（不是重新计算）。实测结论：

- 真实 Excel/WPS 另存的文件 → 带缓存值 → 公式列**正常**显示真实数字 ✅
- 程序(openpyxl)生成、或公式从未被 Excel 计算过的文件 → 无缓存值 → 公式列读成空白 ❌

`core/splitter.py` 里的 `detect_uncalculated_formulas()` 会提前检测这种情况并在日志里预警，
提示用户「用 Excel 打开另存后再处理」。**不要删掉这个检测**，它直接防止公式列静默丢数据。
.xls 路径不受影响（xlrd 读 .xls 时本来就只读值，转换后已是纯值）。

---

## 当前阶段任务（打包与部署）

设计/调试阶段与 v2.0 通用化重构（三层界面、跨文件合并、测试/示例/README）已完成。
剩下交给 Claude Code 的是**打包和部署**。按以下顺序执行，每步做完告诉用户结果：

### 1. 本地验证

```bash
pip install -r requirements.txt
python main.py            # 解释器在本机是 py 启动器：py main.py
pytest -q                 # 跑测试（需先 pip install -r requirements-dev.txt）
```

确认 GUI 能正常打开、能识别列、能跑通拆分；测试全绿。
（注意：GUI 需要图形环境，若在无显示器的环境，跳过实际运行，只做语法检查
`python -m py_compile main.py core/*.py gui/*.py` + `pytest -q`）

### 2. 用 PyInstaller 打包

**强烈建议直接运行项目里的 `build.bat`**，它已经包含所有必需参数。
如果手动执行，命令如下（⚠️ 两个参数缺一不可，否则 exe 会崩溃）：

```bash
pyinstaller --onefile --noconsole ^
  --name "Excel通用拆分工具v2.0" ^
  --icon app.ico ^
  --version-file version.txt ^
  --collect-all customtkinter ^
  --add-data "config;config" ^
  main.py
```

> ⚠️ **两个致命坑（实测确认，务必保留）：**
>
> 1. **`--collect-all customtkinter` 必须有。** customtkinter 依赖 `assets/themes/*.json`
>    和字体文件，PyInstaller 默认不收集，缺了会导致 exe **一启动就崩溃**
>    （报错找不到 blue.json 之类）。
>
> 2. **`--add-data "config;config"` 必须有。** 否则打包后找不到默认配置文件。
>    Windows 用分号 `;` 分隔（Linux/Mac 是冒号 `:`）。
>
> 代码已做路径适配：`gui/app.py` 里用 `sys._MEIPASS` 读打包进去的默认配置，
> 用户配置 `user_config.json` 存到 exe 所在目录（持久、用户可见可编辑）。
> **不要把配置读取改回纯相对路径**，那样打包后会失效。
>
> `app.ico` 和 `version.txt` 已生成（app.ico 是占位图标，用户可替换）。

### 3. （可选）代码混淆

本项目目标是**开源**，通常**不需要**混淆。除非用户明确要求，否则跳过这步。
若要做，注意 PyArmor 和 PyInstaller 的 `--collect-all customtkinter` 需要兼容处理。

### 4. Git 与 GitHub

```bash
git init
git add .
git commit -m "feat: v2.0 通用拆分（按任意列拆分、跨文件合并、三层界面）"
git branch -M main
git remote add origin https://github.com/<用户ID>/excel-splitter.git
git push -u origin main
```

- **源码推到仓库，打包好的 exe 上传到 GitHub Releases**（不要把 exe 提交进 git 仓库）
- 提交信息用语义化前缀：`feat:` `fix:` `docs:` 等
- `.gitignore` 已配置好，会自动忽略 `dist/` `build/` `*.spec` 和用户配置

### 5. 完善 README

- README 里有截图占位，提醒用户**补一张运行截图**（有截图的项目 Star 数明显更高）
- 把 README 和 Releases 里的下载链接替换成真实地址

---

## 代码风格约定

- 中文注释，函数加简短 docstring
- 所有源码文件保留版权头，**不要删**
- 保持模块化：核心逻辑在 `core/`，界面在 `gui/`，不要把业务逻辑写进 `main.py`
- 错误处理：拆分过程中单个文件/网格出错不要中断整体，记录到日志继续跑

---

## 给 Claude Code 的协作提示

- 改动前先 `git status` / `git log` 看清当前状态
- 大改动前先 commit，方便回滚（`git checkout` / `git revert`）
- 修改核心逻辑后，至少做 `python -m py_compile` 语法检查
- 不确定的破坏性操作（删文件、改架构），先问用户
