# CLAUDE.md

本文件是给 Claude Code 的项目说明与工作契约。请在每次开始工作前阅读。

---

## 项目是什么

**ExcelRouter**：一个 Windows 桌面工具，把一批 Excel
表格，**按用户指定的任意一个字段的取值**拆分成多个文件（每个取值一个文件），可选再按第二个
字段做二级拆分（「到人」），并保留原始表头格式。面向不会编程的普通办公人员，v2.1 起界面已改为
**单屏自适应**（不再是三层分档 Tab），v2.4 后期进一步重构为**三步卡片式**：
①选表格 → ②选字段 → ③开始拆分，主按钮固定底部，高级设置与日志默认折叠。
当前版本 **v2.4**，GitHub 仓库：`MarsandSea/excel-router`。

> v1.0 原是「网格化管理」专用（按网格 + 工号识别）。v2.0 已**通用化**：表头自动识别、
> 按列名选列、自动枚举取值、跨文件合并。作者原来的网格 + 岗位工作流用「专家模式」
> （header_mode=keyword + 主列=网格 + 二级列=岗位 + value_alias_map 复现岗位归并）即可复现。

**作者：Abelin。开源协议 MIT。目标是开源传播 + 建立个人声誉，不收费。**
所有源码文件顶部都有 `Copyright (c) 2026 Abelin · MIT License` 署名，**请勿删除任何文件的版权头**。

> **品牌名变更（2026-07，v2.2）**：项目对外品牌名从「Excel 通用拆分工具 / Excel Splitter」
> 改为 **ExcelRouter**。GitHub 仓库已从 `excel-splitter` rename 为 `excel-router`
> （GitHub 自动保留旧链接 301 跳转）。当时只改了**用户可见的品牌文案**：窗口标题、
> `version.txt` 元数据、`build.bat` 产物名、README。旧版本号 v2.1 → 新版本号 v2.2。
> **2026-07-03 补充**：内部残留的旧代号也已清理——本地项目目录从 `Excel_splitter` 改名为
> `excel-router`（与 GitHub 仓库名一致），文档目录树同步更新。`core/splitter.py`、`gui/` 等
> 文件/包名描述的是**功能**（拆分、界面）而非旧品牌，**保持不变**，不要为了品牌一致去改它们。
>
> **品牌文案定稿（2026-07-03，v2.4 后）**：采用「品牌名 + 直白副标 + 差异化标语」三层结构，
> 中文副标从「Excel 业务数据自动分发工具」改为「Excel 批量拆分工具」（新人 3 秒看懂 +
> 保留"Excel/拆分"搜索词；Router 的"分发"含义降级到标语里承载）。标语「整个文件夹一键拆完：
> 按部门、区域、工号等字段自动拆分，打包分发」；特性行「保留原格式 · 跨文件自动合并 · 单个
> 文件也能拆」；界面术语「拆分字段」「🔄 扫描字段」（不再用「按这列拆 / 识别列」）。
>
> **品牌文案二次修订（2026-07-03）**：中文副标再改为 **「Excel 智能拆分工具」**——作者要求
> 突出智能拆分能力（自动识别表头、自动枚举取值、跨文件智能合并等），「批量」只体现数量维度，
> 不能传达工具的智能识别能力。统一文案：窗口标题 `ExcelRouter · Excel 智能拆分工具`；
> README H1、`version.txt` FileDescription 同步。标语/特性行/界面术语不变。
> **此为最新定稿**，「Excel 批量拆分工具」已废弃，勿再混用。

---

## 运行环境

- **目标平台：Windows**（打包出的 exe 给 Windows 用户）
- Python 3.9+
- 依赖见 `requirements.txt`：customtkinter、openpyxl、pandas、xlrd

---

## 目录结构与各文件职责

```
excel-router/
├── main.py                    # 入口，只有几行，调用 gui.app.run()
├── CLAUDE.md                  # 本文件
├── README.md                  # 中英双语说明（含实拍截图 docs/screenshot_*.jpg）
├── LICENSE                    # MIT 许可证
├── requirements.txt           # 运行依赖
├── requirements-dev.txt       # 开发/测试依赖（pytest、pyinstaller、ruff）
├── conftest.py                # 让 pytest 能 import core 包
├── ruff.toml                  # Ruff 静态检查配置（全仓库 lint；core/tests 存量风格已按文件豁免）
├── pyrightconfig.json         # Pyright 类型检查配置（检查范围：gui/ + main.py，core 存量未纳入）
├── .gitignore                 # 忽略 dist/build/用户配置等
├── build.bat                  # 一键打包脚本：本地同时产出 onedir（推荐）+ onefile 两个产物
├── app.ico                    # 程序图标（占位，用户可替换）
├── version.txt                # Windows 版本信息（PyInstaller --version-file 用，随版本号同步更新）
├── config/
│   └── default_config.json    # 默认配置（通用空配置，新字段见下「数据模型」）
├── user_config.json           # 用户保存的配置（运行时生成在程序目录，不入库）
├── core/
│   ├── __init__.py
│   ├── splitter.py            # 核心拆分逻辑（最重要的文件）
│   └── utils.py               # 文本清理 + 取值归并 + 文件名净化
├── gui/
│   ├── __init__.py
│   └── app.py                 # customtkinter 三步卡片式界面（含打包路径适配、队列泵）
├── examples/
│   ├── make_sample.py         # 可复现样本生成器：5 个月份 × 55 名虚拟员工，3 行合并表头
│   └── {1-5}月A分公司明细.xlsx # 生成的演示样本（跨文件合并 + 到人演示用）
├── docs/
│   ├── FAQ.md                 # 常见问题（杀毒误报、公式空白、.xls限制、大文件进度等）
│   ├── RELEASING.md           # 维护者发版手册（CI 流程、误报处理、双产物说明）
│   └── screenshot_*.jpg       # README 用截图
├── .github/
│   ├── workflows/release.yml  # tag push v* 触发：测试→双 PyInstaller 构建→打包→发 Release
│   └── ISSUE_TEMPLATE/        # Bug/Question 结构化表单，config.yml 禁用空白 issue
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
| `split_column` | 拆分字段（按字段名，GUI 里叫「拆分字段」） | `""` |
| `selected_values` | 只拆这些主取值；空 = 自动枚举该字段所有取值 | `[]` |
| `to_person` | 是否在汇总之外**附加产出到人**（仅文件夹批量有效） | `false` |
| `person_column` | 到人按哪个字段拆 | `""` |
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

### gui/app.py（v2.4 后期：三步卡片式重构；队列泵机制不变）

**布局**（self 的 grid 行）：row0 品牌区（品牌名+副标同行、标语、特性行，作者信息在页脚）→
row1 `CTkScrollableFrame` 步骤区（weight=1）：①②卡片 + 「▸ 高级设置」折叠 →
row2 **固定操作区**③卡片（输出目录 + 主按钮 + 进度 + 状态行，永远可达）→
row3 工具条（「▸ 处理详情」日志折叠钮 | 「💬 反馈建议」「保存配置」）→ row4 日志框（默认收起，
**失败时自动展开**）→ row5 页脚（🔒 数据仅在本机处理 · 作者 · MIT）。
`_step_card()` 生成带编号圆徽章的卡片；`_ghost_button`/`_flat_button` 是次要/纯文字按钮工厂
（勿改回 `**dict` 解包样式——pyright 无法对异构 dict 解包做类型匹配）。

- **① 选表格**：「📄 选一个 Excel 文件 / 📁 选整个文件夹」两个直选按钮（替代旧 SegmentedButton+浏览）；
  输入类型由 `os.path.isdir` 自动推断，`_update_input_ui` 控制批量区 `_batch_frame` 显隐；
  路径框可手动粘贴（Return/FocusOut → `_on_path_edited`，`_scanned_path` 防重复扫描）。
- **选完输入自动三连**（`_after_pick`）：显隐批量区 + `_suggest_output` 自动推荐输出目录
  + `_scan_input` 后台扫描。**保存位置是可选项**：不填也能开始（`_start` 里会自动补默认值）。
- **输出默认值规则**（`_auto_output_for`）：文件→同目录`拆分结果`；文件夹→**文件夹里面的**
  `拆分结果`（结果永远出现在数据旁边；放输入里是安全的，run_split 会整体跳过 output_root
  子树）。`_out_auto` 记录推荐值，用户手改过就不再覆盖；启动时若上次保存的输出恰是自动值，
  也识别为自动值（换输入时跟着更新，不会钉在旧位置）。
- `_scan_input`（替代旧 `_detect_columns`）在**子线程**统计 Excel 数、找模板、`list_columns`，
  经 UI 泵消息 `("scan", (cols, n, tpl_name, is_dir, p))` 回主线程；`_on_scan` 先校验 p 未过期。
  **计数镜像 run_split 的跳过规则**（当前输出目录子树不算），结果放输入里时数字才不虚高。
  `_on_columns` 填充下拉，保存值失效时用 `_recommend_split`/`_recommend_person` 按关键词智能预选。
- **② 选字段**：拆分字段下拉 + 🔄 扫描字段；批量选项子卡片在②内（仅文件夹显示）：每组 ZIP、
  同时拆到人（未勾选时子控件置灰，`_update_person_state`）。
- **③ 开始拆分**：`⏹ 停止` 只在运行时出现、`📂 打开输出文件夹` 成功后出现（`_last_output`）；
  `_run_status` 状态行运行时**镜像最新一条日志**（截 70 字符），完成 ✅ 绿 / 停止 ⏹ 橙 / 失败 ❌ 红。
- `_start` 校验链含**三道拦截**：① 目录模式下「输出 == 输入」或「输出是输入的上层」直接阻止
  （`run_split` 会把 output_root 前缀整体跳过，这两种选法会导致“找不到任何文件”）；
  ② `_find_stale_results`：输入里有历史拆分结果目录（`拆分结果`/`*_拆分结果`/`\d{8}结果`）
  且不被本次输出覆盖时弹确认——否则旧结果会被当数据重复拆；③ 开跑前 `os.makedirs` 预检
  保存位置可写，只读/无权限当场友好报错而不是跑一半失败。
- `_on_done` 成功时**静默 `save_config(cfg)`**（失败运行不存，避免记坏参数）；启动时若上次输入
  路径仍存在则自动扫描，实现「打开即用」。
- `_collect_config()` 字段与 v2.1 完全一致（config schema 未变）；`load_config` 用 FALLBACK
  补齐缺键，旧版配置不会报错。
- **匿名反馈入口（v2.3 起）**：模块级常量 `APP_VERSION`、`FEEDBACK_URL`（当前指向 WPS 匿名问卷）。
  「💬 反馈建议」按钮 `_open_feedback` 打开问卷 + 把版本号复制进剪贴板；`_on_done` 仅在**成功完成**
  时在日志追加一行反馈引导（失败路径不加，避免像推卸责任）。**不做任何遥测/自动上报**——
  「数据不出本地」是产品卖点，反馈只能是用户主动点开的外部链接，不要改成程序内嵌表单或自动上传。

**拆分线程 → 主线程通信（v2.4，重要，勿回退）：** 子线程把日志/进度/完成信号 `put` 进
`self._ui_q`（`queue.Queue`），主线程 `_pump_ui` 每 100ms 用 `after(100, self._pump_ui)` 自我调度、
一次性 `get_nowait()` 排空队列后批量刷新 UI。**不要改回 `self.after(0, self._log, ...)` 这种子线程直接
排程主线程回调的写法**——大文件（几万行）时 `core/splitter.py` 会高频报告进度/心跳日志，逐条
`after(0, ...)` 会把 Tk 事件队列打满，表现为界面卡死/黑条纹无响应（v2.3 用户实测反馈的 bug）。
点击「开始拆分」的瞬间要给出即时反馈：进度条先切 `mode="indeterminate"` 播放滚动动画 + 立刻打一条
日志，第一条真实进度值到达后 `_pump_ui` 自动切回 `determinate`（见 `_stop_indeterminate`）。
`core/splitter.py` 侧配合：`process_file` 的 `tick_fn` 报告文件内部阶段进度（读取→读格式→拆分→保存），
`_append_rows` 每 200 行调一次 `heartbeat()`（内含 `time.sleep(0.001)` 让出 GIL + 达到 2000 行报一次
日志心跳）——**这两个回调不要删**，否则单文件几万行场景又会退回“进度条一开始就 100%、中间像假死”。

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

## 项目当前状态与发版流程

初始打包/部署阶段（三层界面时代的 v2.0）早已完成，GitHub 仓库、CI 自动发版、FAQ、Issue 模板、
匿名反馈入口均已上线。**日常开发不需要手动打包/手动建 Release**——正确流程是：

### 1. 本地验证（改动核心逻辑后必做）

```bash
pip install -r requirements.txt
pip install -r requirements-dev.txt
py main.py                 # 本机解释器是 py 启动器，见 [[python-launcher-gotcha]]
ruff check .               # 静态检查，必须全绿（配置见 ruff.toml，存量豁免勿扩大）
npx pyright --pythonpath "$(py -c 'import sys; print(sys.executable)')"
                           # 类型检查（范围 gui/+main.py；必须传 --pythonpath，
                           # 因为 PATH 里的 python 是商店 stub，pyright 自己找不到依赖）
py -m py_compile main.py core/*.py gui/*.py
pytest -q
```

GUI 需要图形环境；无显示器时只做静态检查 + `pytest -q`。

### 2. 发版（真正会打包发布时）

```bash
git add -A && git commit -m "feat: ..."
git push origin main
git tag vX.Y.Z && git push origin vX.Y.Z    # push tag 才会触发 CI
```

推送 `v*` 格式的 tag 会自动触发 `.github/workflows/release.yml`：跑测试 → 语法检查 →
分别用 `--onedir`（推荐，防杀毒误报）和 `--onefile` 两种模式跑 PyInstaller → 打包 →
`softprops/action-gh-release` 自动创建 GitHub Release 并上传双产物。**改版本号时
`version.txt` 里的 `filevers`/`prodvers`/`FileVersion`/`ProductVersion` 四处要同步改**，
否则 exe 属性里的版本号和 tag 对不上。详细的发版规范、误报处理、CI 排障历史见
`docs/RELEASING.md`（例如 Windows runner 默认 `pwsh` 不展开 glob、需要给用到 `*.py` 的
step 加 `shell: bash`；`Rename-Item` 是原地重命名不是移动，后面别再接多余的 `Move-Item`）。

本地手动打包（不走 CI）用 `build.bat`，同样会产出 onedir + onefile 两个产物，两个
PyInstaller 致命坑仍然成立、缺一不可：

> 1. **`--collect-all customtkinter` 必须有**：customtkinter 依赖 `assets/themes/*.json`
>    和字体文件，缺了会导致 exe 一启动就崩溃（报错找不到 blue.json 之类）。
> 2. **`--add-data "config;config"` 必须有**：否则打包后找不到默认配置文件
>    （Windows 分号 `;` 分隔，Linux/Mac 是冒号 `:`）。
>
> 代码已做路径适配：`gui/app.py` 用 `sys._MEIPASS` 读打包进去的默认配置，用户配置
> `user_config.json` 存到 exe 所在目录。**不要把配置读取改回纯相对路径**，打包后会失效。

### 3. 代码混淆

项目目标是开源，**不需要**混淆，除非用户明确要求。

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
