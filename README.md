# 底稿未完成助手 0.3.0

审计底稿“未完成事项”的本地登记与定位工具。用户在 Excel / WPS 表格中选中审计程序单元格，
按一次热键登记；以后在清单里点击“定位底稿”，自动回到对应文件、工作表和单元格。

- 纯本地、单机单人使用
- 只读取工作簿，不写入、不保存、不关闭底稿
- 同时支持 Microsoft Excel 和 WPS 表格
- 关闭主窗口后默认留在系统托盘，热键继续有效

## 主要功能

- 热键/按钮登记当前单元格，支持普通单元格和一个合并单元格。
- 自动记录文件路径、工作表、位置、内容快照、时间、来源软件和文件指纹。
- 清单支持状态、项目、账龄、归档筛选和全文搜索。
- 多选批量改状态、改项目、删除；详情预览和变更历史。
- 缺失资料负责人、索取日期、收到日期和等待天数。
- 项目自动建议和归档。
- 重复登记可查看、更新现有事项或仍然新增；登记后 5 秒内可撤销。
- 定位底稿不保存、不覆盖、不关闭原文件；文件改名/移动可重新关联并核对文件指纹。
- 每日自动备份、普通备份、Windows DPAPI 加密备份/解密。
- 导出 CSV、带文件链接的 Excel 复核清单和 HTML 复核清单。
- 数据目录可在设置中迁移（复制后重启生效）。
- 运行日志自动轮转并脱敏绝对路径。
- 首次使用引导、热键可用性提示和本机合成练习底稿；练习文件生成在数据目录下，不覆盖旧文件。
- 快捷视图“待办 / 待跟进 / 无人负责 / 等待>7天”，支持一键清除筛选和按等待天数排序。
- 选中事项可快速补录资料、负责人、索取/收到日期和手动状态；收到资料不会自动改变状态。
- 登记提示可撤销或直接查看事项；文件重新关联提供路径/指纹确认，定位成功后才保存新路径。

## 运行环境

- Windows 10 / 11（64 位）
- 发布版不需要安装 Python
- 源码运行：Python 3.13（开发环境使用 Anaconda）
- 运行时依赖：`comtypes`、`openpyxl`（Tkinter、SQLite 为 Python 标准库）
- 测试/构建依赖：`pytest`、`pyinstaller`、`pillow`

## 启动发布版

双击：

```text
dist\底稿未完成助手\底稿未完成助手.exe
```

或先构建/生成发布包：

```bat
build_exe.bat
build_release.bat
```

发布包在 `release\底稿未完成助手-0.3.0.zip`，内含按用户安装程序 `安装程序.exe`、Inno Setup 6 标准安装包 `底稿未完成助手-0.3.0-安装程序.exe`、PowerShell 安装/卸载脚本和 Inno 脚本。两种安装程序都不需要管理员权限，默认安装到 `%LOCALAPPDATA%\Programs\底稿未完成助手`。如事务所 IT 已安装 Inno Setup 6，可运行 `build_inno.bat` 重新生成标准安装包。详见 `docs/发布与签名说明.md`。

数据默认保存在 `%LOCALAPPDATA%\WorkpaperTodo\`，不在程序目录内。程序是本地明文 SQLite；
加密备份只能由同一 Windows 用户在同一台电脑恢复。

## 从源码运行

```bat
python -m pip install -r requirements.txt
pythonw -X utf8 -m workpaper_todo.app
```

或双击 `run_dev.bat`（需要 `pythonw` 在 PATH 中）。

测试数据目录可通过环境变量覆盖：

```bat
set WORKPAPER_TODO_DATA_DIR=E:\临时测试数据
```

## 测试

```bat
python -m pytest tests/test_core.py -q
set WPT_RUN_INTEGRATION=1
python -m pytest tests -q
```

真实联动测试会打开自己创建的临时工作簿，激活窗口，验证登记/定位/合并单元格/内容变化/未保存修改/
文件改名重新关联等场景；结束时只关闭自己创建的测试工作簿，不会关闭或保存用户已打开的其他文档。
当前本机完整套件共 **86 项通过**（核心 76 + 真实 Excel/WPS 4 + 真实 Tk 按钮 2 + 真实多显示器 1 + 首次引导/重新关联/主界面布局 Tk 3），完整运行通过。
最终打包版验收记录见 `testdata/final_acceptance_v030.log`。

## 目录结构

```text
workpaper_todo/
  app.py          Tkinter 主窗口、托盘、任务队列、热键、批量操作等
  capture.py      从当前表格窗口读取选区，生成待登记信息
  jump.py         从清单定位回底稿（打开/切换/核对/选中）
  nativeom.py     通过 EXCEL7 窗口的 NativeOM 访问 Excel/WPS 对象模型
  store.py        SQLite 存储、迁移、FTS5、历史、归档、导出、备份
  secure.py       Windows DPAPI 加密备份
  tray.py         原生 Windows 托盘图标和菜单
  config.py       本地配置、数据目录指针
  models.py       状态常量与数据类
  win32util.py    窗口、进程、文件指纹、全局热键（纯 ctypes）
  dialogs.py      重复登记、内容变化、详情、历史、归档、设置对话框
  toast.py        右下角提示条（支持操作按钮）
tests/
  test_core.py              纯逻辑单元测试
  test_real_apps.py         真实 Excel/WPS 联动测试（默认跳过）
  test_button_real_apps.py  真实 Tk 登记按钮联动测试（默认跳过）
tools/                      安装/发布/签名/验收/探测脚本（不属于运行时）
  installer_app.py          按用户安装程序源码
  install_per_user.ps1      PowerShell 安装脚本
  uninstall_per_user.ps1    PowerShell 卸载脚本
  sign_exe.ps1              使用 PFX 证书签名
  test_signing.ps1          自签名流程验证（不用于正式分发）
installer.iss               Inno Setup 6 脚本
build_inno.bat              生成 Inno Setup 标准安装包
docs/使用说明.md            面向使用者的说明
docs/发布与签名说明.md      构建、安装、签名和回滚
docs/兼容性验证记录.md      实际验证边界
docs/加载项方案评估.md      Excel/WPS 加载项方案比较
```

## 实现要点

- **如何连接到正在编辑的窗口**：Excel 和 WPS 的工作表视图窗口类名都是 `EXCEL7`。
  通过 `oleacc.AccessibleObjectFromWindow(EXCEL7, OBJID_NATIVEOM)` 取得该窗口的
  `Window` 对象，再取 `Window.Parent`（工作簿）和 `Workbook.Application`。
  这样不依赖 `Excel.Application` / `Ket.Application` 的 COM 注册，也不会因为
  WPS 的 COM 单实例/多标签问题而连到错误的实例。
- **多窗口/多标签**：登记时使用前台窗口自己的 `Window.Selection` 和
  `Window.Parent.ActiveSheet`；定位时先激活目标工作簿，再从工作簿取工作表和选区，
  避免 WPS 多标签下“工作表对象属于另一个标签”的问题。
- **不写底稿**：所有操作只调用 `Activate` / `Select` / 读取属性，从不调用
  `Save` / `Close` / `Quit`，也不写单元格、批注或标记。
- **线程与 COM**：UI 主线程不碰 COM；所有 NativeOM 调用在单独的工作线程中执行，
  该线程启动时显式 `CoInitialize()`。
- **失败不猜**：多选、空单元格、未保存工作簿、加载项、云文档链接都会拒绝登记并提示；
  内容变化、文件缺失、工作表缺失都会要求人工确认。
- **本地优先**：数据库、日志、备份和加密密钥都在本机；程序没有上传路径，也不依赖云服务。

## 已知限制

详见 `docs/使用说明.md` 第 8 节和 `docs/兼容性验证记录.md`。主要包括：仅验证 64 位 Windows 11、
仅单单元格/合并单元格、不支持云文档链接、不自动跟随被移动到其他单元格的程序、单机单人、
代码签名未完成、Windows 10/32 位、真实混合 DPI 缩放仍需按事务所环境补测；当前主显示器
150% 缩放和多显示器窗口/对话框已用本机 3 屏环境验证。
