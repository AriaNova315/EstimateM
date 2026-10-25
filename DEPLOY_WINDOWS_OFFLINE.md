# EstimateM：Windows 单机离线部署

适用：Windows 11 x64、标准 CPython 3.13.17 x64，每台电脑独立使用。服务只监听 `127.0.0.1:8001`。不需要访问互联网，不迁移旧数据库，也不需要目标机安装 Git、Node.js、Word 或数据库服务。

**验收状态：目前完成 Mac 隔离回归与 Windows wheels 交叉解析/下载核实；尚未在 Windows 实机安装运行，也未完成 Excel/WPS 重算验收。正式使用前执行本文末尾的验收清单。** 实际结果见 [验证记录](DEPLOYMENT_VERIFICATION.md)。

## 一、两方各自准备什么

项目方交付源码 ZIP（带本地 Alpine.js、许可证、启动脚本、依赖锁文件、两个现有 Excel 样例和本说明）。部署人员在联网 Windows 准备 Python **完整安装程序**及 wheelhouse，携带到目标机。正式包不包含旧数据库、业务数据、历史导出、开发机 venv 或账号凭据。

Python 完整安装程序：从 [Python 3.13.17 官方页面](https://www.python.org/downloads/release/python-31317/)选择 **Windows installer (64-bit)**。选标准 CPython，不选 free-threaded、ARM64、32 位、embeddable ZIP 或需要联网引导的安装方式。若目标不是 Windows 11 x64，先重新确认支持范围，不能复用此 wheelhouse。

安装时包含 pip 和标准库/venv，不选择在线下载调试符号等组件。推荐按当前用户安装。记录安装程序文件名、版本和 SHA-256，并在断网 Windows 上验证该完整安装程序可用。不需要在线升级 pip。

## 二、联网 Windows：准备离线依赖

在联网准备机安装与目标相同版本/架构的 Python，把源码包解压到一个目录，用 PowerShell 进入该目录。以下命令均在应用根目录运行。使用显式解释器路径，不必激活 venv，也无需调整 PowerShell 执行策略。

```powershell
py -3.13 -c "import sys,struct; print(sys.version); print(struct.calcsize('P')*8)"
py -3.13 -m venv .prepare-venv
.\.prepare-venv\Scripts\python.exe -m pip --isolated --disable-pip-version-check download --no-cache-dir --only-binary=:all: --require-hashes --dest .\wheelhouse -r .\requirements-win-py313.lock.txt
Get-ChildItem .\wheelhouse\*.whl | Get-FileHash -Algorithm SHA256
```

确认输出版本 3.13.17、64 位。wheelhouse 必须包含完整锁文件中的所有包，尤其是与 Windows x64/CPython 3.13 匹配的 `pydantic_core` wheel。`py3-none-any` 通用 wheel 可用；macOS/Linux 的二进制 wheel 不能带入。不能复制准备机的 `.prepare-venv` 作为目标机环境。

当前锁文件使用目标 Windows 环境解析并对已下载 wheels 生成 SHA-256，保留原有 5 个直接依赖版本；它尚未经 Windows 实机验证，不能称为 Windows 已验收锁文件。准备机先按下一节创建一个独立 `.verify-venv`，断网后执行相同 `--no-index` 安装与业务验收。成功后保留锁文件，不在每台目标机重新解析或运行 `pip freeze` 覆盖它。

代理或证书失败时检查组织网络配置。缺 wheel、哈希失败或 DLL 错误时停止安装，回联网准备机确认版本/架构和来源；不要删除哈希校验、允许源码编译或随意升级所有依赖。

携带的介质目录建议如下：

```text
delivery/
  python-3.13.17-amd64.exe       # 部署人员从官方获取，核对实际文件名
  Python-installer-SHA256.txt
  EstimateM/                   # 解压的源码交付包
    app/
    run.py
    start_windows.cmd
    requirements.txt
    requirements-win-py313.lock.txt
    项目相关信息表.xlsx
    规模及工作量估计表.xlsx
    DELIVERY_MANIFEST.json
    DEPLOY_WINDOWS_OFFLINE.md
    ...
    wheelhouse/                # 部署人员下载，不在源码 ZIP 或 Git 中
      *.whl
```

## 三、目标 Windows：完全离线安装

1. 安装完整 Python，含 pip；重新打开终端。确认版本和架构，不使用 Microsoft Store 下载入口。
2. 将 EstimateM 目录及 wheelhouse 复制到当前用户可写的**本地目录**，例如 `%LOCALAPPDATA%\EstimateM`。不从 ZIP、只读介质、Program Files 或网络共享直接运行。
3. 检查源码包的清单/哈希及 Python 安装程序校验记录。首次交付目录不应有 `estimatem.db`、旧导出或 venv。不要为此删除任何已有业务目录；如有同名目录，应换用新的空目录。
4. 在 PowerShell 运行：

```powershell
Set-Location "$env:LOCALAPPDATA\EstimateM"
py -3.13 -c "import sys,struct; print(sys.version); print(struct.calcsize('P')*8)"
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip --isolated --disable-pip-version-check install --no-index --no-cache-dir --only-binary=:all: --require-hashes --find-links=.\wheelhouse -r .\requirements-win-py313.lock.txt
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe -c "import fastapi,uvicorn,aiosqlite,openpyxl,pydantic; print('imports OK')"
```

这组命令不访问包索引、不检查在线 pip 更新、不安装源码包。安装失败时不要省略参数重试在线安装。也可用 Python 安装目录下的完整 `python.exe` 路径替代 `py -3.13`，所有涉及空格的路径须加引号。

若验证环境使用 `.verify-venv`，上述 `.venv` 路径相应替换；它仅用于演练，日常启动脚本固定使用 `.venv`。

## 四、启动与日常使用

双击 `start_windows.cmd`。窗口显示本机网址，保持窗口打开。在浏览器访问：

```text
http://127.0.0.1:8001
http://127.0.0.1:8001/health
```

健康检查应返回 `{"status":"ok"}`。不要用文件管理器直接打开 `app/static/index.html`。无需新增防火墙入站规则，其他局域网电脑不能访问此服务。API 开发页面 `/docs` 和 `/redoc` 已关闭，业务界面/API 不受影响。

首次启动自动在应用根目录创建 `estimatem.db` 和表结构，项目列表为空。整个应用目录须有写权限，以便数据库创建 journal 等侧文件；系统 TEMP 和浏览器下载目录也须可写。只启动一份实例，使用同一个 Windows 用户账户操作。

1. **复制**现有 `项目相关信息表.xlsx` 为自己的输入文件，再修改内容。不另提供空白模板，也不要把原样例直接当成自己项目的数据。
2. 保留“项目信息、项目进度计划、配置项清单”这些必需 Sheet 和字段表头；“代码模块”按需填写；“基线组成”是现有样例的组成部分，本工具当前不读取它。不要只改项目名而遗留示例项目编号、人员、估算人、进度和配置项内容。
3. 从界面导入，检查项目信息、配置项、代码模块及任务类型；生成估算 Session，再查看或编辑结果，下载 Excel。
4. 相同项目编号的智能导入更新原项目；**重新导入会覆盖配置项、模块和任务，并清除该项目已有估算 Session**。重新导入前保留输入文件和需要的输出副本。
5. `规模及工作量估计表.xlsx` 是输出格式参照，不是运行时读取的模板。程序在内存里生成 10 Sheet xlsx，再由浏览器下载；`generated/` 不是当前导出必需目录。

程序不调用 Excel/Word。查看、重算及打印推荐使用已安装且能够离线使用的 Excel；WPS 须在现场验证后使用。openpyxl 只写公式，不计算结果：用 Excel/WPS 打开后重算并保存，检查汇总。输入文件若有公式，也须先重算保存，以便导入能读取缓存值。若没有任何电子表格软件，可以生成并下载，但本机无法完成公式重算验收。

完成后在启动窗口按 Ctrl+C，等待进程结束再关窗口。重新启动后本机数据保留。需要备份时，先正常停止应用，再复制数据库和自行保存的输入/导出文件到用户指定位置；不要在服务运行中仅复制一个数据库文件。

## 五、常见问题

| 现象 | 处理 |
|---|---|
| 提示缺 `.venv\Scripts\python.exe` | 按第三节在应用目录安装环境；不要直接用其他项目 venv |
| No matching distribution / hash mismatch | 检查 Python 3.13 x64、完整 wheelhouse 和锁文件是否为同一发布；联网准备机重新核对来源，目标机不临时联网 |
| DLL load failed | 核实 Python/Windows 架构和 wheel 标签；按安装程序或错误提示确认是否缺微软运行库，使用官方离线介质；不安装开发编译器碰运气 |
| 8001 已占用 | 检查是否已经启动一份；可用 `Get-NetTCPConnection -LocalPort 8001` 查看。不要直接杀死未知进程；关闭已确认的旧实例或协调端口冲突 |
| 数据库只读/无法打开 | 应用必须放在当前用户有修改权限的本地目录，不能只给 DB 文件权限而禁止创建侧文件 |
| database is locked | 确保没有第二份实例、数据库编辑器或网络共享干扰；不要删除数据库或锁文件来强行恢复 |
| 页面空白/按钮失效 | 通过 HTTP 访问，检查 `/static/vendor/alpinejs-3.14.1.min.js` 是否返回 JS；从完整发布包恢复缺失资源，不能靠浏览器缓存 |
| 下载名异常或同名文件已打开 | 输入项目名避免 `< > : " / \ | ? *`、尾随空格/句点和 Windows 保留名称；使用“另存为”，关闭占用输出文件的 Excel/WPS 后重试 |
| 公式显示空白或未更新 | 在 Excel/WPS 中重算并保存，确认输入文件也有缓存值；不是通过安装 Word 解决 |

日志和排障截图可能包含项目内容，分享前检查。当前交付不提供自动更新、开机启动、数据迁移或远程访问。

## 六、验收清单

部署人员逐项记录实际 OS、Python、浏览器、Excel/WPS 版本及结果：

- [ ] 完整 Python 安装程序在断网目标 Windows 可安装，版本和位数正确。
- [ ] 新建 venv、无 pip 缓存，用 --no-index 和 --require-hashes 安装成功；pip check 通过。
- [ ] 新目录首次启动、10 张表、phase_wbs、空项目列表、停止重启正常。
- [ ] 清空浏览器缓存/新配置，阻断公网但保留本机 loopback，5 Tab 可用且无公网资源请求。
- [ ] 导入样例副本、分类、生成估算、编辑、下载、删除 Session、相同编号重导入均符合预期。
- [ ] 10 Sheet 输出在选定 Excel/WPS 中无修复提示，公式重算和关键汇总/打印布局正确。
- [ ] 中文/空格路径、中文下载名、重复导入下载、同名文件占用行为可接受。
- [ ] 仅监听 127.0.0.1:8001，另一台局域网电脑不能访问，不开入站规则。
- [ ] 测试业务数据不放入正式交付 ZIP，不覆盖原有数据。

可选运行 ` .\.venv\Scripts\python.exe scripts\verify_offline.py`：脚本在系统临时目录复制代码、创建临时数据库、监听随机 loopback 端口验证 HTTP 流程，结束后移除临时目录。它不替代 Windows 双击启动、浏览器交互和 Office 重算验收；它需要当前样例的必要数据结构保持完整。

## 七、Mac 兼容与源码交付

Mac 继续 `python3 run.py`，原 `requirements.txt` 的 5 个版本未变，不要在 Mac 安装 Windows 锁文件。Mac 同样仅本机监听、使用本地 Alpine.js，且不再开放两个 API 开发页面。Windows venv 与锁文件不改变 Mac 的 Python 安装环境。

项目方生成源码包（仅标准库，无需安装打包组件）：

```text
python3 scripts/build_delivery.py --output dist/EstimateM-windows-source.zip
```

Windows 可改用 `py -3.13`。脚本严格按允许列表打包，包含文件 SHA-256、Git 提交与工作区状态，不包含 `.git`、业务 DB、历史导出、wheelhouse、venv、日志或凭据。输出路径已存在会报错，避免覆盖以前交付包。正式交付应从已提交且干净的代码生成，并记录打印的 ZIP SHA-256；这个源码包仍需部署人员加入安装程序和完整 wheelhouse 才能离线安装。

源码继续提交到已有仓库 `AriaNova315/EstimateM`。当前保持私有，不新增 Actions。未来公开前需要另行检查原样例、文档及历史中的内容，单独确认公开权限；本次部署不改写 Git 历史。
