# Galgame BGM Controller 2.0

让音乐跟随游戏。Windows 下的 Galgame 音频控制器：切到后台或最小化时自动静音，切回时恢复接管前的静音状态。

## 使用

1. 启动 `GalBGMController.exe`（文件夹版）或 `GalBGMController-single.exe`（单文件版）。
2. 点击 **添加游戏**，在图标网格中勾选正在运行的游戏。
3. 关闭主窗口后，控制器继续驻留托盘。
4. 已添加的游戏以后启动时会自动接管，无需重复启动控制器。

选择器只枚举具有任务栏窗口的应用，包含最小化窗口；没有推荐、评分、前台优先或后台进程列表。同一 exe 的多个窗口合并为一张卡片，显示图标和名称，不展示完整路径。支持搜索、多选、自动刷新；特殊启动器或没有普通任务栏窗口的游戏可通过 **从文件添加** 选择实际发声的 exe。

## 多游戏与托盘

- 每个运行中的已登记 exe 有自己的游戏图标；游戏退出后只移除它自己的图标。
- 每个图标可以独立暂停/继续、设置静音规则，以及 **停止此游戏控制（本次）**。
- 停止本次控制会恢复音频并撤掉该游戏的图标，不影响其他游戏。库中记录保留，再次启动游戏后自动接管；也可在游戏库选中它并点击继续。
- 没有游戏受控时显示控制器图标，方便重新打开游戏库。
- **暂停全部** 与 **退出全部控制器** 才作用于全部游戏。
- 同一配置只允许一个控制器实例；重复双击会唤回现有主窗口。

Windows 决定通知区域图标的排列和折叠位置；新图标可能出现在右下角的折叠菜单中。按 exe 路径区分游戏，同一个 exe 的多个进程共用一个控制图标。

## 音频与恢复

默认规则可选“切到后台时静音”或“仅最小化时静音”；每款游戏可以单独覆盖。前台检测按实际进程路径匹配，最小化规则要求该 exe 的任务栏窗口全部最小化。

- 精确匹配完整 exe 路径，避免两个目录中的 `Game.exe` 互相影响。
- 每约 0.5 秒重新枚举所有可用播放设备和音频会话，处理默认输出切换、非默认输出、新音频会话和设备重连。
- 静音前保存原状态；暂停、移除、停止本次控制、正常退出时恢复，不改变主音量和游戏音量数值。
- 用户原先已静音的会话会保持静音。“允许播放”是控制策略状态，不保证游戏正在发声。
- 单个设备失败不停止其他游戏的生命周期处理；设备恢复后重试。
- 音频恢复记录会在更改静音前写入磁盘。设备暂时离线时保留记录；重连或控制器下次启动时，对仍存在的同一会话恢复。
- 强制结束进程、断电或 Windows 注销不能保证即时执行清理。恢复记录只能恢复可重新找到的会话，不能恢复已永久消失的设备或会话。

使用 helper 子进程发声的游戏，需要将实际发声 exe 加入库；不会凭同名文件、进程树或猜测去控制其他程序。任务栏枚举采用 Win32 顶层窗口规则，特殊 Shell/UWP 托管窗口可能需要从文件添加。

## 配置和日志

沿用旧版的实际配置位置，原游戏清单和默认规则无需重新录入：

```text
%USERPROFILE%\gal_audio_controller_config.json
%USERPROFILE%\gal_audio_controller_config.audio-state.json
%USERPROFILE%\gal_audio_controller_config.log
```

旧的 `selected_procs` / `mute_mode` 兼容；名称和单游戏规则保存在可选的 `apps` 字段。配置只在修改时原子写入；文件损坏时明确报错，不用空配置覆盖。运行中的外部文件修改不会热加载，请退出后编辑再重启。日志按 1 MB 轮转，保留两个历史文件。

## 架构

```text
Win32 任务栏窗口 → windows.py → 图标选择器 → Config
进程 + 前台窗口 → WinFocus → Service → 每游戏 GameState → Queue → Tk / Tray
                                    ↓
                          PycawAudio → 所有播放端点
                                    ↓
                          原静音状态 + 原子恢复记录
```

- `core/windows.py`：窗口筛选和按 exe 聚合。
- `core/focus.py`：每轮一次进程遍历、一次窗口遍历，产生活跃目标快照。
- `core/service.py`：每游戏状态、暂停/停止/重启生命周期；工作线程持有音频 COM 对象。
- `core/audio.py`：设备/会话重新枚举、精确路径控制、原状态恢复。
- `core/config.py`：旧格式兼容、线程锁、原子保存。
- `core/instance.py`：Windows 命名互斥量和唤醒事件。
- `ui/picker.py` / `ui/tray.py`：图标选择、多托盘图标；通过队列进入 Tk 主线程。
- `app_gui.py` / `app_cli.py`：入口；GUI 与 CLI 共用同一个控制服务。

## 开发、测试与编译

Windows + Python 3.13。保留原来的 PyInstaller 文件夹版、单文件 GUI 和 CLI 打包方式，脚本不再自动升级全局环境或删除 comtypes 缓存。

```powershell
py -3.13 -m venv .venv
.venv\Scripts\python -m pip install -r requirements-dev.txt
.venv\Scripts\python -m pytest -q
build\build_gui.bat
build\build_cli.bat
```

输出：`dist\GalBGMController\GalBGMController.exe`、`dist\GalBGMController-single.exe`、`dist\GalBGMController-CLI.exe`。正在运行的版本不要原地覆盖；可以指定独立目录：

```powershell
py -3.13 build\build.py --kind all --output .artifacts\release-2.0.0
```

真实 Windows 无声双进程测试（只改测试程序的会话）：

```powershell
New-Item -ItemType Directory -Force .artifacts
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:winexe /out:.artifacts\GameA.exe /reference:System.Windows.Forms.dll /reference:System.Drawing.dll tests\fixtures\AudioFixture.cs
Copy-Item .artifacts\GameA.exe .artifacts\GameB.exe
py -3.13 tests\windows_audio_smoke.py
```

GUI 支持 `--config <path>` 隔离配置、`--diagnostics <json>` 只读打包检查、`--smoke-test <seconds>` 正常启动后自动恢复并退出。不要将测试配置指向日常配置。

[详细重构和验证报告](docs/REFACTOR-2.0.zh-CN.md) · [构建说明](build/README_BUILD.md)
