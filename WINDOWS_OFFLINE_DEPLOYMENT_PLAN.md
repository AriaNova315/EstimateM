# EstimateM Windows 单机离线部署交付方案（已批准，方案阶段记录）

检查日期：2026-10-24。检查基线：`e46fb5514cf6c09c8010f4ca5b69d17d09bdee79`。

后续状态：用户已明确批准按方案实施。以下保留方案阶段的边界与推荐，实施结果见 `DEPLOYMENT_VERIFICATION.md`，正式操作见 `DEPLOY_WINDOWS_OFFLINE.md`。当前设备无法执行 Windows，依赖锁定改为按 Windows 目标交叉解析并核对目标 wheels，已明确保留实机验收门槛。

本文件是方案，不是已经实施的部署说明。本文出现的新增文件、代码修改、下载、安装和命令均为批准后执行的建议。本次只新增本文件，没有启动应用、初始化数据库、安装依赖、修改现有文件或执行 Git 写操作。

## 结论

现有 FastAPI + SQLite + openpyxl 架构适合各电脑独立运行。推荐采用“源码目录 + Python 完整安装程序 + Windows wheelhouse + 本地启动脚本”，不引入打包 EXE、Docker、Windows 服务或 CI。

交付前必须处理：Alpine.js 公网引用、监听地址、开发文档页面的公网引用、完整离线依赖锁定与 Windows 安装验证、启动及部署说明。数据库仍由每台目标机首次运行自行创建，不迁移任何现有数据。

用户已明确：不另做空白输入模板，使用人员复制现有 `项目相关信息表.xlsx` 后修改。用户计划以后将 GitHub 仓库改为公开；当前保持私有，公开操作和可能的历史清理不包含在本阶段，也不因批准部署实施而自动获得授权。

## 1. 当前状态与检查证据

### 1.1 项目边界

- 项目目录和已检查的各级父目录均未发现 `AGENTS.md`；读取了 `CLAUDE.md`、`DEV_PLAN.md`、依赖清单、入口、全部运行模块及现有 Excel 格式说明中的相关约定。
- 业务范围为项目配置项规模及工作量估算、Excel 导入、4 人估算、10 Sheet Excel 导出。开发计划记载 Phase 1—5 已通过，但该记录不代表本次 Windows 验收通过。
- 现有约定是 macOS 系统 Python、`python3 run.py`、不用 venv。建议仅对 Windows 交付改用独立 venv；批准后更新说明，避免两种启动要求冲突。
- 未发现部署说明、README、项目内测试套件或 Python 版本声明。部分设计/测试文档只引用本机 `~/.gstack/...` 路径，交付不能依赖它们。

### 1.2 已满足

| 项目 | 检查结果与依据 |
|---|---|
| 服务架构 | `run.py` 启动 `app.main:app`，单进程、无 reload，端口 8001 |
| 数据库 | `app/database.py` 将数据库定位在项目根目录，使用 aiosqlite；启动时建表，启用外键，退出时关闭连接 |
| 初次建库 | DDL 已包含 `phase_wbs`；共 10 张业务表，不需要随包携带旧数据库或独立种子 SQL |
| 导出 | `excel_builder.py` 新建 Workbook，在 BytesIO 中生成；`export.py` 返回下载响应，不依赖磁盘输出模板或 `generated/` 中的历史文件 |
| 运行模块路径 | 数据库及静态目录使用 `__file__` 定位和 `os.path.join`，运行代码未发现依赖同级 ConfigM 源码 |
| Office 集成 | 未发现 COM、win32com、AppleScript、LibreOffice 命令行或其他 Office 自动化调用 |
| 网络业务依赖 | 前端 fetch 使用同源 API；运行代码未发现外部业务 API、在线模型、自动下载、在线字体或后台更新逻辑 |
| Git | 当前目录本身为独立仓库，不是父目录仓库；已有对应私有 GitHub 仓库 |

### 1.3 已发现的问题

| 问题 | 证据与影响 |
|---|---|
| 业务页面依赖公网 JS | `app/static/index.html:7` 引用 jsDelivr Alpine.js 3.14.1；空缓存断网时交互不可用 |
| 监听局域网接口 | `run.py:4` 为 `0.0.0.0:8001`，不符合只允许本机访问的要求 |
| API 开发文档引用公网 | `FastAPI(...)` 默认开放 `/docs` 和 `/redoc`；本机安装的相同版本 FastAPI 源码确认 Swagger/ReDoc 和 favicon 的外部默认地址 |
| 离线依赖未完整固定 | `requirements.txt` 固定了 5 个直接依赖，未固定全部传递依赖；不同时间准备可能解析出不同版本 |
| 开发辅助脚本使用 macOS 绝对路径 | `analyze_template.py` 的 SAMPLE，`gen_layout.py` 的 SAMPLE/OUT 指向 `/Users/...`；不在业务启动链中，不应作为部署步骤 |
| 缺少部署规范 | 尚无 Windows 启动脚本、Python/架构说明、wheelhouse 准备与断网安装说明 |

### 1.4 尚未验证与限制

- 本机为 macOS，系统 Python 为 3.14.3；没有将它的已安装环境或 pip freeze 当作 Windows 离线交付环境。
- 已完成运行源文件按 Python 3.10 语法解析的小范围只读检查；这说明语法下限至少为 3.10，不代表完整依赖或业务已在该版本通过。
- 未在 Windows 实测：安装程序完全断网安装、cp313 win_amd64 wheel 完整性、首次建库、停止/重启、中文路径、浏览器下载、Excel/WPS 重算及打印。
- 未启动当前应用、未读取或迁移业务数据库内容、未执行全流程业务测试。本次没有将历史开发计划中的勾选重新标记为验证通过。
- Excel ZIP 结构只读检查发现：输入样例有 5 Sheet，输出参照有 10 Sheet；两者未发现 `xl/externalLinks/` 项或 VBA 项。这不是完整的公开发布内容审计。

## 2. 推荐环境与选择依据

| 项目 | 推荐 |
|---|---|
| Windows | Windows 11 x64，使用组织仍支持的版本；Windows 10、Windows Server、32 位或 ARM64 需另行确认和验收 |
| Python | 标准 CPython 3.13.17 x64（常规 GIL 版本），使用 python.org 的完整 Windows installer (64-bit) |
| 环境隔离 | 应用目录下 `.venv`，联网准备机与目标机使用相同 Python 小版本、架构，最好相同补丁版本 |
| 浏览器 | 目标机已有的受支持 Microsoft Edge，或可离线安装的 Chrome；不需要浏览器账号或联网扩展 |
| 运行目录 | 当前使用人员具有修改权限的本地目录，例如 `%LOCALAPPDATA%\EstimateM`；不放 Program Files、网络共享或只读介质 |
| 地址/端口 | `http://127.0.0.1:8001`，单进程、单实例，不开放入站防火墙规则 |
| Office | 程序运行不要求 Excel/Word。查看、重算、打印输出时推荐使用本机已安装且可离线使用的 Excel；WPS 作为需验收的替代 |

选择 3.13 系列是兼顾现有类型注解和依赖兼容范围、Windows 安装包可得性，并避免直接将最新大版本用于旧依赖。3.13.17 的官方页面列出 2026-10-21 发布及带安装程序的特别安全发布背景。这里是候选基线；在 Windows 上解析全部依赖、下载 wheels 并安装通过后才冻结交付版本，不承诺尚未验证的兼容性。[Python 官方发行说明](https://www.python.org/downloads/release/python-31317/)

不使用 Microsoft Store 在线入口、仅联网引导的安装管理器或 embeddable ZIP 代替完整安装程序。安装时包含 pip、venv 所需标准库，避免勾选需额外下载的调试资源。若基础 Windows 缺少运行库，只按完整安装程序或现场错误的实际要求准备微软官方离线运行库；目前没有证据要求额外交付编译器、Node.js、Git、Word、独立 SQLite 服务或 Java。

openpyxl 写入公式而不计算公式，因此“生成成功”不等于“所有公式已有缓存结果”。用户需要在 Excel/WPS 中打开并重算，必要时保存；输入工作簿中的公式也必须已有正确缓存，因为导入器采用 `data_only=True`。如果要求完全没有电子表格软件、所有公式仍立即可读，则属于另外的需求取舍，不在当前最小方案内。[openpyxl 公式说明](https://openpyxl.readthedocs.io/en/stable/simple_formulae.html)

## 3. 必要修改清单

### 3.1 必须完成（批准后）

| 项目 | 涉及文件 | 原因、建议与预期影响 |
|---|---|---|
| 本地化 Alpine.js | 修改 `app/static/index.html`；新增 `app/static/vendor/alpinejs-3.14.1.min.js`、许可证/来源记录 | 保持 3.14.1，改用 `/static/vendor/...`，不升级框架；记录原始来源、版本和 SHA-256。断网空缓存可用 |
| 限制监听 | 修改 `run.py` | 将 host 改为 `127.0.0.1`，保留 8001、单进程及 reload=False；不提供局域网访问 |
| 消除开发文档公网资源 | 修改 `app/main.py` | 推荐设置 `docs_url=None, redoc_url=None`，保留业务 API 与 `/health`；不额外打包 Swagger/ReDoc。OpenAPI JSON 可保留，若以后需要离线交互文档再单独本地化 |
| 固定完整依赖 | 保留现有 5 项版本，新增 `requirements-win-py313.lock.txt` | 在干净 Windows 3.13 环境解析并冻结直接及传递依赖，不从本机开发环境拷贝清单；含 Pydantic、pydantic-core、Starlette、AnyIO、Click、h11、et-xmlfile 等实际解析包；以最终锁文件为准 |
| Windows 启动 | 新增 `start_windows.cmd` | 使用 `%~dp0`、`cd /d`、带引号的 `.venv\Scripts\python.exe`，显式 UTF-8 环境；检查解释器，保留错误输出，不自动安装依赖、不终止占用端口的其他进程 |
| 部署说明和交付记录 | 新增 `DEPLOY_WINDOWS_OFFLINE.md`、发布清单/校验说明；同步 `CLAUDE.md`、`DEV_PLAN.md` 的部署约定 | 区分准备机与目标机，写明运行目录、模板使用、全新初始化、停止/重启及排障；记录经验证 Python、OS、依赖和提交版本 |
| 交付排除与复核 | 按明确允许列表制作源码包；按实际需要补充 `.gitignore` | 只交付代码/资源，不复制整个开发目录。新增规则可包含 wheelhouse、交付压缩包、验收产物、SQLite `.sqlite3-wal/.sqlite3-shm` 等；不得删除本地现有文件 |

FastAPI 官方提供离线自托管开发文档资源的方法；本方案选择关闭这两个非业务页面，以减少交付资源和维护工作。[FastAPI 官方说明](https://fastapi.tiangolo.com/how-to/custom-docs-ui-assets/)

锁定文件负责可重复解析，发布文件清单负责文件完整性；建议记录 wheel 和安装程序的 SHA-256。可以在后续交付中为锁文件增加 pip hashes，但不为此引入额外依赖管理工具。若现有版本在候选 Windows 环境安装失败，先记录具体阻碍，再提出最小依赖调整，不自动进行全面升级。

### 3.2 Windows 兼容性检查与最小处理

| 检查项 | 当前判断与处理 |
|---|---|
| 路径 | 运行链路径可跨平台；启动脚本固定工作目录。保留数据库在应用根目录，先不重构为多目录配置；必须确保整个目录可写，便于 SQLite 创建 journal 等侧文件 |
| 绝对路径/同级项目 | 两个开发分析脚本与部分文档有 macOS 路径，但业务不依赖；部署说明明确“不运行分析脚本”。可选以后支持相对路径或参数，不作为本次运行阻碍 |
| 文件句柄 | 导入通过上传字节和 BytesIO；`peek_project_code` 使用 read_only=True，却未显式关闭 workbook，主解析也未显式关闭。未发现源 Excel 路径长期占用的直接证据，建议小范围补 finally 关闭 workbook 并关闭上传文件资源，作为可选可靠性修补；重复制导入与 Windows 文件锁测试失败时再升为必须 |
| 临时文件 | 业务没有命名临时文件后重开/rename 的实现；上传由框架管理，可能落到系统 TEMP，需默认用户 TEMP 可写；无需新增自建临时目录或替换流程 |
| 文件替换/覆盖 | 导出在内存生成，由浏览器选择保存位置；未发现 Windows 专有 rename 覆盖问题，不计划引入服务器磁盘导出逻辑 |
| 中文下载名 | 已用 UTF-8 filename*；项目名只替换斜线和反斜线，Windows 其他保留字符/设备名仍需实测；推荐按验收结果做最小文件名净化，不扩大到其他业务字段 |
| 编码/换行 | HTML 有 UTF-8；Python 源码可按 UTF-8 解析。新脚本建议 ASCII 提示或明确 UTF-8、CRLF，现有代码不批量转换换行或格式化；中文路径/空格路径实测 |
| 数据库生命周期 | lifespan 打开全局 aiosqlite 连接，正常退出关闭；单机维持单 worker，不增加连接池/并发架构。建议 finally 确保异常退出关闭为可选改进 |
| 初始化迁移 | `CREATE TABLE IF NOT EXISTS` 后尝试重复加 phase_wbs，捕获了所有异常；全新库无需该 ALTER。先保留兼容逻辑，验证新建与重启，不为部署做迁移框架重构 |
| 系统命令 | 业务运行模块未发现 shell、macOS 专用命令或同级程序调用；Windows 不要求安装 Unix 工具 |

### 3.3 可选且默认不实施

- 开发分析脚本改为相对路径或 CLI 参数；运行数据搬到独立目录；一键离线安装脚本；双击自动打开浏览器；关闭 OpenAPI JSON；API 文档全部本地化。
- 以上只有确有使用需要或验收发现阻碍时再选取。默认不引入 EXE 打包、Windows 服务、注册表修改、自动开机启动、自动升级、外网入口、数据库迁移、CI/CD 或无关业务修复。

## 4. 交付文件与职责

### 4.1 项目方交付

| 内容 | 交付方式与用途 |
|---|---|
| `app/**`、`run.py` | 完整业务代码和静态资源，包含本地 Alpine.js |
| `requirements.txt`、Windows 完整锁文件 | 直接依赖说明及正式离线安装依据 |
| `start_windows.cmd`、部署说明 | 普通用户操作与部署人员排障 |
| `项目相关信息表.xlsx` | 沿用现有输入样例，使用人员先复制后修改；不新增空白输入模板。输入包含项目/人员/估算人等信息，填写说明须指出哪些列需替换、必需 Sheet、代码模块可选及导入覆盖规则 |
| `规模及工作量估计表.xlsx` | 可随参考资料交付，用于输出格式对照；不是程序运行时模板，缺少它不影响代码构建 Workbook |
| `CLAUDE.md`、`DEV_PLAN.md`、格式说明及分析脚本 | 源码资料可保留，运行说明不依赖文档内的作者电脑路径；开发提示文件不属于终端用户必读内容 |
| 第三方许可证、版本/文件校验清单 | 记录静态资源来源、依赖版本、安装程序候选及源码提交号 |

没有业务种子文件、独立迁移目录或必须从现有数据库导出的基础配置。`app/database.py` 已包含首次建库结构。

### 4.2 部署人员在联网 Windows 准备并携带

- 对应版本 Python 完整 x64 安装程序及校验记录；Windows 相同 Python/架构生成的完整 wheelhouse。
- 在断网演练中验证过的锁文件、依赖版本记录；必要时是浏览器/电子表格软件自身的离线安装介质和组织许可。
- 不将整个准备机 venv 复制到目标机，目标机创建自己的 venv 并离线安装。wheelhouse 和安装程序作为交付介质资源，默认不提交到 GitHub。

### 4.3 必须排除

`estimatem.db`、其他业务数据库及 journal/WAL/SHM；`generated/` 与历史导出；`.venv/`；缓存、日志、临时验收数据库和产物；`.env`、本机工具设置、凭据；`.git/` 和准备机独有文件。原有文件留在本机，不删除。

现有 `.gitignore` 已排除实际 `estimatem.db` 及其侧文件模式、generated、venv、缓存和本地设置；没有误排 `app/static` 或两个 xlsx。当前跟踪清单未包含数据库或历史导出。广泛的 `*.db` 等规则未来若遇到必要种子资源要逐项判断，目前没有这种资源。对 `.sqlite`/`.sqlite3` 的侧文件及任意目录中的临时导出仍需交付清单复核，不能把 `.gitignore` 当作数据审计或打包清单。

## 5. GitHub 处理方式及未来公开事项

- 本地根目录：`/Users/gaoyiman/Desktop/Code/5000B/EstimateM`，与 `git rev-parse --show-toplevel` 一致。
- 分支 `main`，跟踪 `origin/main`。检查开始时工作区干净；本次完成后预期只新增本方案文件，保持未提交。
- origin：`https://github.com/AriaNova315/EstimateM.git`。
- GitHub：`AriaNova315/EstimateM`，当前 PRIVATE，默认分支 main；通过只读 API 确认远程 main 为 `e46fb5514cf6c09c8010f4ca5b69d17d09bdee79`，与本地一致。无需初始化、创建或重新关联。
- 批准部署实施后：在此仓库完成修改、验证、审阅允许提交的文件，再提交推送，核对远程提交及私有状态。不在本阶段切分支、提交、推送或修改权限。
- 用户以后会自行或另行要求公开。公开前检查两个 xlsx 的单元格、隐藏 Sheet/行列、批注、文档属性、人员/项目标识，以及 Markdown、开发提示、Git 历史中的内容。它们目前已在首次提交中；“不另做空白模板”不等于确认已有内容均可公开。
- 若确认现有样例全部允许公开，原样沿用；若有不可公开信息，优先另行讨论脱敏并检查历史可见性。仅修改当前文件无法从历史中移除内容，历史改写需单独明确批准，当前不执行，也不默认 force push。

待确认的信息仅为目标 Windows/架构和公开前内容适用性；仓库归属与现有远程已明确。部署实施批准不等于公开批准。

## 6. 部署说明拟包含内容和关键命令

以下为拟写入正式部署说明的 PowerShell 示例，依赖正式锁文件且需实施后验证，不是本阶段执行记录。

### 6.1 联网 Windows：准备 wheelhouse

先安装选定的完整 Python 3.13 x64，确认普通 CPython，打开新终端。项目方在干净 Windows 环境生成完整锁文件，部署人员以发布锁文件下载，不重新自由解析。

```powershell
py -3.13 -c "import sys,struct; print(sys.version); print(struct.calcsize('P')*8)"
py -3.13 -m venv .prepare-venv
.\.prepare-venv\Scripts\python.exe -m pip download --only-binary=:all: --dest .\wheelhouse -r .\requirements-win-py313.lock.txt
Get-ChildItem .\wheelhouse\*.whl | Get-FileHash -Algorithm SHA256
```

锁文件的首次生成流程：在干净的 Windows 3.13 venv 安装现有 `requirements.txt`，确认实际应用兼容后 `pip freeze` 导出为锁文件，复核没有本地路径、开发包或不适用于目标机的条目；重新下载锁文件，再用另一个干净 venv 进行纯 wheel 安装演练。不得从本机 macOS 环境直接生成交付 lock。

默认同平台下载，不建议在 macOS 交叉准备。必须跨平台时才使用 pip 的 `--platform win_amd64 --python-version 3.13 --implementation cp --abi cp313 --only-binary=:all:`，仍以真实 Windows 离线安装为最终验收。`pydantic-core` 含平台二进制，不能只带 pure Python wheel 或 macOS wheel。[pip download 官方说明](https://pip.pypa.io/en/stable/cli/pip_download/)

### 6.2 目标机：完全离线安装

安装完整 Python（含 pip），把应用及 wheelhouse 解压到当前用户可写目录；不要直接在 ZIP、U 盘只读目录或网络共享运行。无需目标机访问 GitHub 或安装 Git。

```powershell
Set-Location 'C:\Users\部署用户\AppData\Local\EstimateM'
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip --isolated --disable-pip-version-check install --no-index --no-cache-dir --only-binary=:all: --find-links=.\wheelhouse -r .\requirements-win-py313.lock.txt
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe -c "import fastapi,uvicorn,aiosqlite,openpyxl,pydantic; print('imports OK')"
.\start_windows.cmd
```

最终说明用实际用户目录代替示意路径。安装命令不运行在线 pip upgrade，也不使用激活脚本，避免 PowerShell 执行策略影响。缺 wheel 时在联网准备机补齐并重新演练，不在目标机尝试源码编译或临时联网安装。[pip 离线安装说明](https://pip.pypa.io/en/latest/user_guide/#installing-from-local-packages)

### 6.3 启动、使用和排障

- 双击启动脚本，保持窗口打开；确认 `/health` 后访问 `http://127.0.0.1:8001`。不要用 file:// 打开 index.html。
- 复制 `项目相关信息表.xlsx`，修改项目、人员、配置项、代码模块和进度计划，再导入。说明同编号导入更新以及重新导入会清理该项目估算 Session，操作前先保留输入文件副本。
- 在界面生成估算、修改结果、下载 Excel，使用 Excel/WPS 打开重算；浏览器下载目录需可写，避免覆盖正在被 Office 打开的同名文件。
- Ctrl+C 正常停止，待进程退出再关闭窗口；仅运行一份应用，端口占用时先查已有实例，不盲目终止进程或开放防火墙。
- 首次启动创建空数据库，不带旧数据库或样例业务记录；验收数据只在专用测试目录初始化，不混入正式交付包。
- 排障包括 Python 位数/路径、缺 wheel、DLL 错误、端口占用、数据库只读/锁定、TEMP 权限、页面静态资源 404、中文下载名和公式未重算。日志/截图分享前排除业务内容。
- 全新部署不含旧数据迁移；上线后若需备份，停服务后复制数据库到用户指定位置，勿在服务运行时只拷贝单一 DB 文件。未规划自动备份或升级工具。

## 7. 验证方案

| 验证 | 方法 | 验收条件 |
|---|---|---|
| 清洁交付目录 | 从明确文件清单构建新目录，不拷旧 DB/历史导出/venv | 启动前无业务 DB、文件齐全、与发布清单一致 |
| Windows 离线安装 | 同版 x64 Python，在关闭外网且清空 pip 缓存的干净 venv，按 --no-index 命令安装 | 没有在线请求、源码编译或下载；pip check 与模块导入成功 |
| 首次初始化 | 在独立临时应用副本首次启动 | 自动生成 10 张业务表，含 phase_wbs；项目列表为空；真实开发库未变化 |
| 停止和重启 | Ctrl+C 停止、再次启动 | 无残留实例或数据库锁；临时测试数据正确保留，无重复初始化副作用 |
| 离线页面 | 新浏览器配置/无缓存，阻断外网但保留 loopback | 5 Tab、导入、编辑、生成、下载均正常；业务及保留页面不请求公网资源；关闭的 /docs、/redoc 不加载外部资源 |
| 本机隔离 | 检查监听地址，尝试从另一台局域网机器访问 | 仅 127.0.0.1:8001 监听，局域网地址无法访问；不新增入站规则 |
| 核心业务 | 测试目录中导入原样例副本，必要时仅在临时副本改用虚构信息 | 项目/配置项/模块/任务导入，估算 Session 生成、编辑保存、重新导入、Session 删除按现有业务完成；无新增业务范围 |
| 输出 | 下载并检查 10 Sheet、关键公式、跨 Sheet 引用；Excel 打开重算、保存后读取缓存 | 无文件修复提示、关键汇总正确、中文名正常；WPS 若作为正式选择则执行相同验收 |
| Windows 文件边界 | 应用路径含中文/空格；重复导入下载；源 xlsx 在 Office 打开时读取副本；下载文件同名/打开占用；项目名含 Windows 特殊字符 | 路径与临时文件无锁冲突；失败有明确处理方法；确有缺陷才实施最小修补 |
| 公共发布前检查 | 只在获得单独授权后检查完整公开候选内容及历史 | 确认样例、文档、属性和历史适合公开；本阶段不改变私有状态 |
| Git 一致性 | 实施后提交前复核暂存文件，推送后只读核对远程 SHA/权限 | 正确仓库、无业务 DB/历史输出/凭据、远程提交一致，仍私有 |

Windows 实测是发布门槛。本机静态检查和过去开发记录不能替代它。若暂时没有 Windows 设备，可先完成获批修改与资料，但只能标注“待 Windows 验收”，不能称已完成离线交付。

## 8. 建议执行顺序与阶段验收

1. **方案批准**：确认环境与样例交付约定。验收：明确批准实施；本阶段到此停止。
2. **最小离线修改**：本地 JS、许可证、loopback、关闭开发文档、Windows 启动脚本。验收：资源没有公网引用，未变更估算/导入/导出业务规则。
3. **Windows 依赖基线**：干净 Windows 解析现有依赖、生成 lock，下载 wheelhouse。验收：选定 Python/架构可纯 wheel 安装，全部传递依赖固定，记录校验信息。
4. **清洁目录验收**：全新数据库、完整流程、断网页面及 Excel 重算、文件边界。验收：第 7 节相关用例通过，所有失败有处理结果；测试不触碰真实数据。
5. **交付资料与打包**：正式部署说明、原输入样例使用说明、允许列表与发布清单。验收：另一位部署人员可按说明断网安装，不需要作者电脑路径或现有 DB。
6. **提交与推送**：在正确私有仓库提交获批范围代码/资源/文档，校对远程 SHA。验收：工作区状态、发布基线和远程一致，仓库仍私有。不添加/触发 GitHub Actions。
7. **后续公开**：用户另行决定时间和授权，先复核公开内容及历史。验收：用户确认哪些内容允许公开，再执行单独批准的权限/历史操作。

## 9. 需要用户决定的事项与默认建议

当前只需要批准或调整整份方案，无需逐项决定常规实现细节。

- **目标环境是否符合 Windows 11 x64？** 默认按该环境实施；若实际是 Windows 10、32 位或 ARM64，需要先提供版本和架构，调整运行基线。
- **输出查看环境**：默认以可离线使用的 Excel 为验收基准。若目标用户只用 WPS，请在 Windows 验收前说明；若目标机不安装任何电子表格软件，可完成服务和导出验收，但公式查看/重算验收需在另外的电脑进行。
- **公开仓库前的内容授权**：保留并直接使用原输入样例的决定已记录。公开前仍需确认样例人员/项目等内容可公开；现在不要求改模板，也不公开仓库。

已确定而不再询问：沿用 AriaNova315/EstimateM；当前保持私有；每台独立、仅本机访问；不迁移旧数据；部署人员准备完整 Python 安装程序和 Windows wheelhouse；不新增空白输入模板；批准后才修改、安装、初始化测试、提交或推送。
