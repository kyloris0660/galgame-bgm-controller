# Galgame BGM Controller

一个简单的Python脚本，用于自动控制Galgame在后台运行时的BGM播放。当游戏窗口不在前台时自动静音，切换回前台时自动恢复音量。

## 功能描述

- 🎮 支持选择特定的Galgame进程进行音频监控
- 🔇 当游戏窗口切换到后台时自动静音
- 🔊 当游戏窗口切换回前台时自动恢复音量

## 系统要求

- Windows 7/8/10/11
- Python 3.6+
- 管理员权限（用于控制音频）

## 使用方法

1. 克隆仓库：
```bash
git clone https://github.com/kyloris0660/galgame-bgm-controller.git
cd galgame-bgm-controller
```

2. 安装依赖：
```bash
pip install pycaw psutil pywin32 pytk pillow pystray
```

3. 运行程序：
```bash
python admin_launcher.py
```

4. 在弹出的窗口中选择需要控制的Galgame进程
5. 程序会自动监控选定游戏的音频状态
6. 按 `Ctrl+C` 退出程序

## 注意事项

- 建议在运行游戏前先启动此程序
- 某些使用特殊音频引擎的游戏可能不受支持（未经测试）