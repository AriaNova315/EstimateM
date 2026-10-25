# 离线交付实现与验证记录

实现日期：2026-10-24—26。源码基线：`e46fb5514cf6c09c8010f4ca5b69d17d09bdee79`，最终提交号以 Git 和交付包 `DELIVERY_MANIFEST.json` 为准。

**状态：最小离线交付实现已完成；Windows 实机与 Office 验收尚未完成。** 不将交叉解析/下载或 Mac 回归等同于 Windows 安装运行通过。正式步骤见 `DEPLOY_WINDOWS_OFFLINE.md`。

## 已实现范围

- Alpine.js 3.14.1 改为本地静态文件，保留原版本，带官方 MIT 许可证、来源与 SHA-256。
- `run.py` 限制为 `127.0.0.1:8001`；单进程、不启用 reload。
- 关闭 FastAPI `/docs`、`/redoc`，业务 API、健康检查和 OpenAPI JSON 保留。
- Windows CRLF/ASCII 启动脚本，显式切换目录、解释器路径加引号、UTF-8 环境，错误保留在窗口中。
- 16 项 Windows Python 3.13 x64 候选依赖全部固定，附目标 wheel 哈希；原 5 项直接依赖版本未改。
- 允许列表源码 ZIP 打包工具，包内文件哈希/提交信息清单；不覆盖已有输出文件。
- 使用原输入样例，不另做空白模板；未改动两个原有 Excel 文件、估算算法、业务模型或导入导出规则。
- 无新 Actions、数据库迁移、自动下载、自动升级、Windows 服务、额外业务功能或 Git 历史改写。

## 已通过的核实

### 依赖与静态检查

- 本机 macOS + Python 3.14.3，未升级系统或项目当前依赖。
- 用现有 uv 针对 `x86_64-pc-windows-msvc`、Python 3.13 的环境解析，结果 16 项；同时检查各 wheel METADATA 的直接与条件依赖。
- pip 以 `--platform win_amd64 --python-version 3.13 --implementation cp --abi cp313 --only-binary=:all:` 下载全部目标 wheels 到临时目录；包括 `pydantic_core-2.50.1-cp313-cp313-win_amd64.whl`。
- 以完整锁文件执行本地 `--no-index --no-cache-dir --require-hashes --only-binary` 安装 dry-run，16 项目标 wheels 和 SHA-256 全部通过。此步骤没有执行 Windows 二进制或在 Windows 创建 venv。
- Alpine.js 44,659 bytes，SHA-256：`358d9afbb1ab5befa2f48061a30776e5bcd7707f410a606ba985f98bc3b1c034`，Node 语法检查通过；运行不要求安装 Node。

### Mac 原环境和候选依赖版本回归

`scripts/verify_offline.py` 在系统临时目录复制应用代码，使用含中文及空格的目录，在随机 loopback 端口启动独立服务。不连接现有数据库，结束时停止进程并移除临时测试数据。

在原有 Mac 环境和临时 Mac venv 均通过以下流程。临时环境安装与候选锁相同的 16 项版本（使用 Mac 原生 wheel，不使用 Windows wheel），依赖兼容检查也通过：

- 全新数据库、空项目列表、10 张表及 `phase_wbs` 字段。
- `/health`、主页、本地 JS 返回；`/docs`、`/redoc` 返回 404。
- 现有样例副本导入、任务列表/分类、创建 Session、编辑估算值。
- 下载可打开的 OOXML ZIP，10 Sheet、含跨表使用的公式；不宣称公式计算或打印布局已验证。
- Session 删除、同编号重新导入后清除旧 Session。
- 正常停止、重启、测试数据保留。
- 原目录数据库 SHA-256 与检查前一致：`530d002992b5d311e99e8f83a19df4f9323fc17e8e9d3a6839a9876b4806478d`。

### 浏览器验证边界

独立临时服务使用 18017 端口，避免使用真实业务实例或 8001 端口。浏览器加载业务页面和本地 JS 成功，Alpine 版本为 3.14.1；点击“+ 手动新建”后表单正常显示，没有控制台错误。加载请求记录仅包含本机页面、JS 和 API，源 HTML 没有公网引用。

浏览器工具不允许 `Network.setBlockedURLs` 或 `Network.clearBrowserCache`，因此没有声称完成“强制断网+清缓存”的浏览器验收，该项仍由 Windows 部署人员实测。

### 源码包检查

允许列表构建了 32 个文件加 `DELIVERY_MANIFEST.json` 的源码 ZIP；逐项重新读取并核对文件 SHA-256 全部通过。包内未发现数据库、日志、pyc、wheel、.git 或 generated 内容。正式交付包在提交完成后重新生成，清单应显示已提交且无未提交修改。

## 仍需部署人员执行

- [ ] Windows 11 x64、标准 CPython 3.13.17 完整安装程序在断网环境安装。
- [ ] Windows 重新下载匹配 lock 的 wheelhouse，干净 venv、无缓存、--no-index 与哈希校验安装成功，pip check 通过。
- [ ] 双击 CMD 启动、中文/空格路径、端口冲突处理、停止/重启、文件锁和中文下载名。
- [ ] 新浏览器配置/无缓存、禁止公网但保留本机 loopback，5 Tab 和完整流程可用，只有本机资源请求。
- [ ] Windows 确认仅 loopback 监听，其他局域网电脑不可访问；不开放入站规则。
- [ ] 10 Sheet 在现场 Excel（或约定 WPS）无修复提示，重算、汇总、保存缓存、打印布局验收。
- [ ] 记录实际 Windows/Python/浏览器/Office 版本与结果；验收完成前不标记“Windows 离线交付验收通过”。

## 发布内容与仓库

源码包由 `scripts/build_delivery.py` 明确允许列表生成，wheelhouse、Python 安装程序由部署人员准备。数据库、历史导出、测试临时文件、本地 venv、.git 及凭据不入包。ZIP 内的 `DELIVERY_MANIFEST.json` 列出每个源文件 SHA-256，并标明提交及工作区是否有未提交修改。

沿用 `AriaNova315/EstimateM`。本次不修改仓库权限，未来公开前须另行检查样例和文档内容及 Git 历史。不因提供样例给使用人员而视为其历史内容已获公开授权。
