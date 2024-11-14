# Galgame BGM Controller

一个简单的Python脚本，用于自动控制Galgame在后台运行时/最小化时的BGM播放。当游戏窗口不在前台时自动静音，切换回前台时自动恢复音量。

## 功能描述

- 🎮 支持选择特定的Galgame进程进行音频监控
  - 支持重新选择进程
  - 默认当游戏退出时同时自动退出脚本（可关闭）
  - 支持系统托盘调整可选项
  - 支持临时暂停监控
- 🔇 支持两种自动静音控制
  - 默认模式：仅在游戏最小化时静音（可切换）
  - 可选模式：非前台时自动静音
- 💾 支持设置保存
  - 记录历史选择过的游戏进程
  - 下次启动时自动匹配已知进程（可关闭）

## 系统要求

- Windows 7/8/10/11
- Python 3.6+
- 管理员权限（用于控制音频）

## 从源码运行

1. 克隆仓库：

```bash
git clone https://github.com/kyloris0660/galgame-bgm-controller.git
cd galgame-bgm-controller
```

2. 安装依赖：

```bash
pip install pycaw psutil pywin32 pillow pytk pystray
```

3. 运行程序：

```bash
python start.pyw
```

## 从源码构建

1. 安装额外的构建依赖：

```bash
pip install pyinstaller
```

2. 运行构建脚本：

```bash
python build.py
```

3. 构建完成后，可执行文件将位于 `dist` 目录中

## 直接使用可执行文件（未测试）

1. 下载最新版本的 `GalgameBGMController.zip`
2. 解压到任意目录
3. 运行 `GalgameBGMController.exe`（需要管理员权限）
4. 在弹出的窗口中选择需要控制的Galgame进程（历史记录将显示为绿色背景）
5. 程序会自动监控选定游戏的音频状态，可在托盘中找到

## 系统托盘功能

右键点击系统托盘图标可以：

- 查看当前监控状态
- 暂停/继续监控
- 切换静音模式（仅最小化/非前台）
- 开关进程结束自动关闭
- 开关自动匹配历史进程
- 重新选择进程
- 清空历史记录
- 退出程序

## 注意事项

- 建议在运行游戏前先启动此程序
- 某些使用特殊音频引擎的游戏可能不受支持（未经测试）
- 配置文件保存在程序目录下的 `gal_audio_controller_config.json` 中
- 程序运行日志保存在程序目录下的 `bgm_controller.log` 文件中