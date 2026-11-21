# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概览

Windows 本地桌面工具（Python 3.13 + Tkinter），用于登记审计底稿中的“未完成事项”，并能从清单定位回原工作簿的工作表和单元格。运行时只读取 Excel / WPS 表格，不写入、不保存、不关闭底稿。当前版本 0.3.0。

## 常用命令

```bat
REM 安装依赖
python -m pip install -r requirements.txt

REM 从源码运行（或双击 run_dev.bat）
pythonw -X utf8 -m workpaper_todo.app

REM 核心单元测试（纯逻辑，不需要 Excel/WPS）
python -m pytest tests/test_core.py -q

REM 单个测试
python -m pytest "tests/test_core.py::<test_name>" -q

REM 全量测试
python -m pytest tests -q

REM 真实 Excel/WPS 联动测试（默认跳过；会打开自己创建的临时工作簿）
set WPT_RUN_INTEGRATION=1
python -m pytest tests -q

REM 使用隔离的数据目录做手动验收
set WORKPAPER_TODO_DATA_DIR=E:\临时测试数据

REM 打包与发布
build_exe.bat        REM PyInstaller -> dist\底稿未完成助手\
build_release.bat    REM 生成 release\底稿未完成助手-<version>.zip
build_inno.bat       REM 生成 Inno Setup 安装包（需已安装 Inno Setup 6）
```

真实联动测试会激活本机 Excel/WPS 窗口，只应在明确需要时运行；它们只关闭自己创建的测试工作簿。

## 整体架构

**启动链路**：`run_app.py` → `workpaper_todo.app.main()`。`main()` 依次做单实例互斥体检查、DPI 感知、日志初始化、创建 `Config(data_dir)` 与 `Store(data.db)`，最后进入 Tk 主循环。所有运行状态落在本地数据目录（默认 `%LOCALAPPDATA%\WorkpaperTodo`，`WORKPAPER_TODO_DATA_DIR` 可覆盖，`location.txt` 记录迁移后的目录）。

**线程模型（关键约束）**：Tk UI 主线程从不直接调用 COM。所有 Excel/WPS 对象模型操作都通过 `App._submit()` 交给 `Worker`——一个长期运行的后台线程，启动时显式 `CoInitializeEx(STA)` + `comtypes.CoInitialize()`，逐条串行执行任务。结果写回 `ui_queue`，由 Tk 的 `after` 轮询、在主线程更新界面。全局热键由独立的 `HotkeyThread` 注册并持有消息循环，热键触发时只把前台窗口句柄投递到 `ui_queue`，等主线程再派发任务。新增任何表格交互都应遵循这条链路，不要从回调里直接碰 COM。

**连接 Excel/WPS（`nativeom.py`）**：Excel 和 WPS 的工作表视图窗口类名都是 `EXCEL7`。通过 `oleacc.AccessibleObjectFromWindow(EXCEL7, OBJID_NATIVEOM)` 取得该窗口的 `Window` 对象，再取 `Window.Parent`（工作簿）和 `Workbook.Application`。这样不依赖 `Excel.Application` / `Ket.Application` 的 COM 注册，避免 WPS 单实例/多标签连到错误实例。登记使用前台窗口自己的 `Window.Selection` 和 `Window.Parent.ActiveSheet`；定位先激活目标工作簿，再从其取得工作表和选区。

**登记流程（`capture.py`）**：前台窗口 → 确认是工作表窗口 → 读取选区 → 生成 `SelectionInfo`（文件路径、文件名、文件指纹 volume/index、工作表、单元格地址、合并区域、内容快照、来源软件）。多选、空单元格、未保存工作簿、云文档链接等不安全情形会拒绝登记。

**定位流程（`jump.py`）**：查找已打开的工作簿，必要时启动 Excel/WPS 打开文件；核对工作簿路径 key；激活并验证 `ActiveWorkbook` 确实是目标；查找工作表；把当前内容与登记快照比对，不一致则要求人工确认；最后选中单元格并置前。任何一步无法确认都抛错，不猜测。

**只读约束**：对底稿只允许 `Activate` / `Select` / 读取属性。禁止调用 `Save` / `Close` / `Quit`，也不写单元格、批注或任何标记。

**存储（`store.py`）**：单个 SQLite 文件 `data.db`，用 `PRAGMA user_version` 管理 schema 版本；升级前自动备份为 `.vN.bak`。核心表为 `items`（事项）、`item_events`（字段变更历史）、`project_meta`（项目归档）。搜索优先用 FTS5（trigram），不可用或结果为空时退回 `LIKE`。状态常量和数据类集中在 `models.py`。

**其他模块**：`tray.py` 原生 Windows 托盘；`win32util.py` 纯 ctypes 的窗口、进程、文件指纹与全局热键；`dialogs.py` 各类确认/详情/设置对话框；`toast.py` 右下角提示条；`secure.py` Windows DPAPI 加密备份；文件改名/移动后的重新关联同样走 `jump.py` 的安全核对逻辑。

## 开发约定

- **失败不猜**：内容变化、文件缺失、工作表缺失、路径不一致都必须停下来要求人工确认；不要为了让流程继续而自动接受不一致。
- **线程边界**：新增的 COM 读取或文件操作放进 `Worker` 任务；UI 更新只能回到主线程。
- **数据兼容**：修改 schema 必须递增 `SCHEMA_VERSION` 并补迁移与备份逻辑；不要直接改用户现有 `data.db`。
- **版本号**：`workpaper_todo/__init__.py` 的 `__version__` 是程序内版本来源；`build_release.bat` 与 `build_release.ps1` 的默认 `-Version` 也需同步更新，发版时一并检查。
- **测试分层**：`tests/test_core.py` 是不依赖 Office 的纯逻辑测试；`test_real_apps.py`、`test_button_real_apps.py` 依赖真实 Excel/WPS，默认跳过，需设置 `WPT_RUN_INTEGRATION=1`；`test_gui_ux.py`、`test_multimonitor_gui.py` 需要真实 Tk 窗口/多显示器环境。
- **打包**：新增运行时依赖要同时更新 `requirements.txt` 与 PyInstaller 的排除/包含策略（见 `build_exe.bat`）；入口固定为 `run_app.py`。
- **文档语言**：用户可见文案与 `docs/` 文档使用中文。
