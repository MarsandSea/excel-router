# 发版手册（维护者用）

给项目维护者（当前是 Abelin）看的操作手册：怎么发新版本、怎么防杀毒软件误报、
误报真的发生了该怎么处理。目标是**发版全程自动化，人只需要打一个 tag**。

---

## 1. 日常发版流程

1. 改动代码，本地跑一遍验证：
   ```bash
   pytest -q
   python -m py_compile main.py core/*.py gui/*.py
   ```
2. 更新版本号（三处要一致，缺一会导致 exe 属性 / 标题栏 / 文档互相对不上）：
   - `gui/app.py` 的 `APP_VERSION` 常量（窗口标题、反馈按钮里都引用它）
   - `version.txt` 的 `filevers` / `prodvers` / `FileVersion` / `ProductVersion` 四个字段
   - `CLAUDE.md` 开头「当前版本」那一行
3. Commit、push 到 `main`。
4. 打 tag 并推送，**这一步会自动触发 CI 发版**（`vX.X.X` 换成实际版本号，
   比如 `v2.5.0`；千万别抄本手册里的旧示例，早已被占用会直接 push 失败）：
   ```bash
   git tag vX.X.X
   git push origin vX.X.X
   ```
5. 去仓库 **Actions** 页看 `Release` 工作流跑完（约 3-5 分钟），跑绿后
   **Releases** 页会自动出现两个产物：
   - `ExcelRouter-vX.X.X-win64.zip`（onedir，推荐分发）
   - `ExcelRouter-vX.X.X.exe`（onefile，备选）
6. 下载其中一个到本机，脱离开发环境（换个目录）双击冒烟测试：能正常打开、
   识别列、跑通一次拆分。

CI 配置见 `.github/workflows/release.yml`。它会先跑 `pytest -q` 拦住测试不过的版本，
测试失败则整个发版流程停止，不会出坏的 Release。

---

## 2. 为什么要两种打包形态

`--onefile` 单文件版每次启动都会把 Python 解释器和依赖**解压到系统临时目录再执行**，
这个「自解压后执行」的行为模式与蠕虫/木马的启动方式高度相似，是杀毒软件（尤其
Windows Defender）**启发式误报**的主要触发点之一；另一个触发点是 PyInstaller
官方预编译 Bootloader 的哈希值被广泛用于打包各类程序（含恶意程序），本身已经
被部分厂商标记为可疑。

`--onedir` 版把依赖以普通文件夹形式摊开，运行时不需要自解压，能大幅降低这类
启发式误报概率，因此**默认推荐用户下载 onedir 版**（见 README 的下载区）。
onefile 版仍然提供，图个方便，但风险自担并在文档里提前说明。

---

## 3. 每次发版后：VirusTotal 例行检查

发布新版本后，顺手查一次两个产物在 VirusTotal 上的检出情况：

1. 打开 https://www.virustotal.com/
2. 分别上传 `ExcelRouter-vX.X.X-win64.zip`（或解压后的 exe）和
   `ExcelRouter-vX.X.X.exe`
3. 记录检出引擎数量（如 `2/70`），对比两种形态的差异
4. 如果某个杀毒引擎误报，记下厂商名字——下一步要用

这一步不阻塞发版，是**观测**，用于判断要不要走第 4 步的误报申诉。

**验收标准：国内主流杀毒软件（Windows Defender、360、QQ电脑管家、火绒等）干净即可**，
目标用户是国内普通办公人员，这几家才是真实影响他们下载/使用信任的关键。冷门/海外引擎
（如 Bkav、Gridinsoft、Yandex 等）零星误报**不需要逐个申诉**——v2.5.0 实测 3/64、
全部是这类小众引擎，国内主流全部干净，判定为可接受，未做任何申诉动作。

---

## 4. 真的遇到误报：处理流程

按性价比从高到低排序，**不要一上来就上重武器**：

### 4.1 向对应厂商提交误报申诉（首选，免费）

- **Windows Defender / Microsoft**：提交到
  [Microsoft Security Intelligence 误报反馈门户](https://www.microsoft.com/en-us/wdsi/filesubmission)，
  选择 "Software developer" → 上传文件 → 说明是开源项目 + 附 GitHub 链接。
  通常数天内会更新云端库，之后同一份文件哈希不再被拦截。
- **其他厂商**（火绒、360、QQ管家等）：各家一般都有类似的「文件误报申诉」入口，
  搜「厂商名 + 误报申诉」即可找到，操作方式类似。

### 4.2 如果同一版本反复被新误报（触发式，非默认执行）

说明公共 Bootloader 的哈希污染已经比较严重，可以考虑**自编译 PyInstaller
Bootloader**：从源码克隆 PyInstaller，用本地 C 编译器（GCC）重新编译
Bootloader，产出的二进制哈希是唯一的，不在任何公共特征库里。参考
[PyInstaller 官方文档 - Building the Bootloader](https://pyinstaller.org/en/stable/bootloader-building.html)。

这一步有实际的时间成本（需要装 C 编译环境、重新验证打包流程），**只有在
第 4.1 步不够用、且误报持续影响用户信任时才做**，不是每次发版的标配。

### 4.3 不做的事（成本不划算，明确记录以免以后纠结）

- **不用 Nuitka 重写打包链路**：真编译能进一步降低误报，但迁移成本高，
  与「零维护」的项目定位冲突。
- **不买 EV 代码签名证书**：数千元/年，本项目是免费开源工具，不构成收益模型，
  不值得为此掏钱。

---

## 5. 内部版本（如果需要，不走公共 Release）

如果需要一份带公司内部署名的定制版本（例如 `CompanyName` 改成内部邮箱），
**不要**把这类版本上传到公共 GitHub Release ——会有隐私/合规风险，也会让
公共受众困惑「这是不是带后门的企业定制版」。

正确做法：本地单独跑一次 `pyinstaller`，把 `--version-file` 换成一份内部专用
的版本信息文件，产物只发到公司内部群/网盘，不进 CI、不进公共 Release。
