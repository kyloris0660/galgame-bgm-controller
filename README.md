# Galgame BGM Controller

自动控制 Galgame 在后台/最小化时的 BGM 播放。切换窗口状态时自动静音/恢复音量。支持同时监控多个游戏进程。

## 快速开始

1. 前往 [Releases](https://github.com/kyloris0660/galgame-bgm-controller/releases/latest) 下载最新版本
2. 解压到任意目录（建议不要解压到 Program Files 等需要管理员权限的目录）
3. 运行 `GalgameBGMController.exe`（将自动请求管理员权限）
4. 选择需要控制的游戏进程（已选择过的进程会显示为绿色背景）
5. 程序会自动最小化到系统托盘运行

## 功能说明

### 🎮 进程控制
- 支持同时监控多个游戏进程
- 记录历史游戏进程，支持下次启动时自动匹配
- 可随时添加新的监控进程
- 支持管理和移除已监控的进程
- 可设置游戏退出时自动关闭程序

### 🔇 音频控制
- 仅最小化时静音（默认）
- 非前台时静音（可选）
- 多进程支持：同时运行多个游戏时，只保留前台的游戏音频

### 🔧 系统托盘
- 显示当前监控状态
- 暂停/继续监控
- 切换静音模式
- 重新选择进程
- 清空历史记录

## ⚙️ 系统要求

- Windows 7/8/10/11
- 管理员权限（用于控制音频）

## 📝 注意事项

- 程序运行时生成的文件位于 `GalgameBGMController/_internal` 目录：
  - `gal_audio_controller_config.json`：配置文件
  - `bgm_controller.log`：日志文件

## 💻 开发相关

### 从源码运行

```bash
git clone https://github.com/kyloris0660/galgame-bgm-controller.git
cd galgame-bgm-controller
pip install pycaw psutil pywin32 pillow pytk pystray
python start.pyw
```

### 构建可执行文件

```bash
pip install pyinstaller
python build.py
```