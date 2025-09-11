# Galgame BGM Controller

Galgame BGM Controller 可以在 Windows 上自动静音后台 Galgame 进程并在切回时恢复音量。  
支持系统托盘图标动态显示当前被静音应用的图标，并提供多种静音策略。

---

## ✨ 功能
- 监听多个游戏进程，前台优先或仅最小化时静音。
- 托盘右键菜单：暂停/继续、切换静音规则、退出。
- 托盘图标随被静音的应用改变。
- 无需 Python 环境：提供打包好的单文件 exe。
- 自动生成用户配置（`config.json`），不污染仓库。

---

## 📦 下载与使用
1. 前往 [Releases](../../releases) 页面，下载最新的 **`GalBGMController-single.exe`**。
2. 双击运行，第一次会生成默认配置文件 `config.json`。
3. 点击「选择监听软件」添加需要静音的游戏进程。
4. 切出游戏时会自动静音，切回时恢复。

---

## 🛠️ 本地构建
如需自行构建：
```bash
# 安装依赖
py -3.13 -m pip install --user -r requirements.txt

# 打包 GUI（单文件）
build\build_gui.bat
