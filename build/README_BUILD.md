# Windows 构建

采用 Python 3.13 + PyInstaller；依赖版本固定于 requirements.txt。

```powershell
py -3.13 -m venv .venv
.venv\Scripts\python -m pip install -r requirements-dev.txt
.venv\Scripts\python -m pytest -q
build\build_gui.bat
build\build_cli.bat
```

GUI 脚本依次构建文件夹版和单文件版；CLI 独立构建。脚本优先使用仓库 .venv，否则使用 py -3.13。失败时返回非零退出码，不继续打印成功。

构建入口 build.py 使用当前解释器，路径相对于仓库计算；不依赖开发者机器上的绝对路径。生成的 spec 和中间文件位于 .artifacts。不会删除用户的 comtypes 缓存或自动升级全局 pip。

```powershell
py -3.13 build\build.py --kind all --output .artifacts\release-2.0.0
```

输出文件：
- GalBGMController/GalBGMController.exe（需保留同目录 _internal）
- GalBGMController-single.exe
- GalBGMController-CLI.exe

发布前运行两个 GUI 产物的 --diagnostics <json> 和隔离配置 --smoke-test 3；检查退出码和日志。先构建到独立目录，验证后复制至版本化安装目录，再备份并更新快捷方式，避免覆盖正在使用的 exe。
