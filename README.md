
# Galgame BGM Controller — v5 重写版（可测、可扩展）

## 亮点
- **分层解耦**：Policy/Config/Focus/Audio/Service 清晰职责
- **可测试**：自带 pytest 测试（含集成回路），核心无需 UI 也能验证逻辑
- **双入口**：`app_cli.py`（无 UI、用于验证后台静音）和 `app_gui.py`（最小 GUI + 托盘）
- **安全降级**：托盘失败不影响后台静音；pycaw 不可用也可用 DummyAudio 做回归测试

## 快速开始
```powershell
pip install -r requirements.txt
# 先用 CLI 验证后台静音逻辑（推荐）
python app_cli.py
# 或者 GUI
python app_gui.py
```

## 运行测试
```powershell
pytest -q
```

## 配置文件
保存在 `~\\gal_audio_controller_config.json`：
```json
{
  "selected_procs": ["xxx\\Game.exe", "xxx\\Player.exe"],
  "mute_mode": "not_foreground"  // or "minimized_only"
}
```
