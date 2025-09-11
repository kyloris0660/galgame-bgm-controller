# Galgame BGM Controller

一个基于 **Python + pycaw** 的轻量级工具，用于在 Windows 下自动静音 Galgame BGM。  
当游戏窗口不在前台或被最小化时，它会静音指定进程的音频会话，并在右下角托盘显示当前被静音应用的图标。

---

## ✨ 功能特性

- 🎮 **自动静音**：支持两种模式  
  - **不在前台时静音**（默认）  
  - **仅最小化时静音**  
- 🖼 **进程选择器升级**：  
  - 显示每个进程的应用图标  
  - 前台进程淡黄高亮，便于快速定位  
  - 支持多选 & 右侧大图标预览  
- 🗑 **管理监听列表**：  
  - 支持删除选中的监听进程  
  - 监听列表实时保存到配置文件  
- 🪟 **托盘图标**：  
  - 正常状态使用默认图标  
  - 被静音时，托盘图标变为对应应用的图标  
- 🐛 **调试改进**：  
  - 控制台日志仅在状态变化时打印  
  - 支持在界面右上角勾选/取消调试输出  
- 💾 **配置文件**：  
  - 默认路径：`~\gal_audio_controller_config.json`  
  - 示例：  
    ```json
    {
      "selected_procs": [
        "C:\\Galgame\\悠刻のファムファタル_DL版\\femme_fatale.exe"
      ],
      "mute_mode": "not_foreground"
    }
    ```

---

## 📥 安装依赖

确保已安装 **Python 3.13 (64-bit)** 和 **git**。

```bash
pip install -U pip
pip install --user pycaw==20240210 "comtypes>=1.4.8" pywin32==311 psutil Pillow pystray
