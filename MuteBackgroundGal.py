import time
import tkinter as tk
from tkinter import ttk, messagebox
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import win32gui
import win32process
import pystray
from PIL import Image, ImageDraw
import threading
import sys
import json
import os
from pathlib import Path

class AudioController:
    def __init__(self):
        self.target_pid = None
        self.target_name = None
        self.running = True
        self.monitoring_thread = None
        self.tray_icon = None
        self.paused = False
        self.minimize_only = True
        self.auto_close = True
        self.auto_match = True  # 默认开启自动匹配
        self.last_muted_state = {}
        self.history_processes = set()  # 先初始化为空集合
        # 修改配置文件路径到当前目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(current_dir, 'gal_audio_controller_config.json')
        
        # 加载配置
        config = self.load_config()
        self.history_processes = set(config.get('history_processes', []))  # 更新历史进程记录

    def load_config(self):
        """加载配置文件"""
        default_config = {'history_processes': [], 'auto_match': True}
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config
            else:
                # 如果配置文件不存在，创建默认配置
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    json.dump(default_config, f, ensure_ascii=False, indent=2)
                return default_config
        except Exception as e:
            print(f"加载配置文件失败: {e}")
            return default_config

    def save_config(self):
        """保存配置文件"""
        try:
            config = {
                'history_processes': list(self.history_processes),  # 转换集合为列表以便JSON序列化
                'auto_match': self.auto_match
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置文件失败: {e}")

    def add_to_history(self, process_name):
        """添加进程到历史记录"""
        if process_name:
            self.history_processes.add(process_name)
            self.save_config()

    def create_icon(self, is_pause=False):
        """创建状态图标"""
        icon_size = 64
        image = Image.new('RGB', (icon_size, icon_size), color='white')
        drawing = ImageDraw.Draw(image)
        
        if is_pause:
            # 暂停状态：红色双竖线
            pause_color = '#D32F2F'
            drawing.rectangle([20, 15, 30, 49], fill=pause_color)
            drawing.rectangle([39, 15, 49, 49], fill=pause_color)
        else:
            # 播放状态：绿色三角形
            play_color = '#00C853'
            points = [(20, 15), (20, 49), (49, 32)]
            drawing.polygon(points, fill=play_color)
        
        return image

    def update_icon_and_menu(self):
        """更新图标和菜单状态"""
        if self.tray_icon:
            self.tray_icon.icon = self.create_icon(self.paused)
            self.tray_icon.menu = self.create_menu()

    def create_menu(self):
        """创建托盘菜单"""
        return pystray.Menu(
            pystray.MenuItem("Galgame音频控制器", None, enabled=False),
            pystray.MenuItem(
                f"当前进程: {self.target_name if self.target_name else '未选择'}", 
                None, 
                enabled=False
            ),
            pystray.MenuItem(
                f"状态: {'已暂停' if self.paused else '监控中'}", 
                None, 
                enabled=False
            ),
            pystray.MenuItem(
                f"{'继续监控' if self.paused else '暂停监控'}", 
                self.toggle_pause
            ),
            pystray.MenuItem(
                "仅最小化时静音",
                self.toggle_minimize_only,
                checked=lambda item: self.minimize_only
            ),
            pystray.MenuItem(
                "进程结束时自动关闭",
                self.toggle_auto_close,
                checked=lambda item: self.auto_close
            ),
            pystray.MenuItem(
                "自动匹配历史进程",
                self.toggle_auto_match,
                checked=lambda item: self.auto_match
            ),
            pystray.MenuItem(
                "重新选择进程",
                self.reselect_process
            ),
            pystray.MenuItem(
                "清空历史记录",
                self.clear_history
            ),
            pystray.MenuItem(
                "退出程序",
                self.stop_monitoring
            )
        )

    def create_tray_icon(self):
        """创建系统托盘图标"""
        self.tray_icon = pystray.Icon(
            "galgame_audio_controller",
            self.create_icon(),
            "Galgame音频控制器\n右键可暂停/继续",
            menu=self.create_menu()
        )

    def toggle_pause(self):
        """切换暂停状态"""
        self.paused = not self.paused
        if self.paused and self.target_pid:
            self.restore_volume(self.target_pid)
        self.update_icon_and_menu()

    def toggle_minimize_only(self):
        """切换是否仅在最小化时静音"""
        self.minimize_only = not self.minimize_only
        self.update_icon_and_menu()

    def toggle_auto_close(self):
        """切换是否在目标进程结束时自动关闭"""
        self.auto_close = not self.auto_close
        self.update_icon_and_menu()

    def toggle_auto_match(self):
        """切换是否自动匹配历史进程"""
        self.auto_match = not self.auto_match
        self.save_config()
        self.update_icon_and_menu()

    def clear_history(self):
        """清空历史记录"""
        self.history_processes.clear()
        self.save_config()
        messagebox.showinfo("提示", "历史记录已清空")

    def find_matching_processes(self):
        """查找匹配的历史进程"""
        matching_processes = []
        sessions = AudioUtilities.GetAllSessions()
        
        for session in sessions:
            if session.Process and session.Process.name() in self.history_processes:
                matching_processes.append((session.Process.pid, session.Process.name()))
        
        return matching_processes

    def auto_select_process(self):
        """尝试自动选择进程"""
        if not self.auto_match or not self.history_processes:
            return False
            
        matching_processes = self.find_matching_processes()
        
        if len(matching_processes) == 1:
            # 只有一个匹配的进程，自动选择
            self.target_pid, self.target_name = matching_processes[0]
            self.last_muted_state[self.target_pid] = False
            print(f"自动选择进程: {self.target_name} (PID: {self.target_pid})")
            return True
        elif len(matching_processes) > 1:
            # 有多个匹配的进程，让用户手动选择
            print(f"找到多个匹配的进程: {[p[1] for p in matching_processes]}")
            return False
        else:
            print("未找到匹配的历史进程")
            return False

    def restore_volume(self, pid):
        """恢复指定进程的音量"""
        try:
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and session.Process.pid == pid:
                    volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                    volume.SetMute(0, None)
                    break
        except Exception as e:
            print(f"恢复音量失败: {e}")

    def restore_all_volumes(self):
        """恢复所有被跟踪进程的音量"""
        try:
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and session.Process.pid in self.last_muted_state:
                    volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                    volume.SetMute(0, None)
        except Exception as e:
            print(f"恢复所有音量失败: {e}")

    def stop_monitoring(self):
        """停止监控并退出程序"""
        self.running = False
        self.restore_all_volumes()
        if self.tray_icon:
            self.tray_icon.stop()

    def reselect_process(self):
        """重新选择要监控的进程"""
        if self.target_pid:
            self.restore_volume(self.target_pid)
        
        # 强制显示选择窗口，不进行自动匹配
        temp_auto_match = self.auto_match  # 保存当前自动匹配设置
        self.auto_match = False  # 临时关闭自动匹配
        success = self.select_target_process()  # 显示选择窗口
        self.auto_match = temp_auto_match  # 恢复自动匹配设置
        
        if success:
            self.update_icon_and_menu()
        else:
            # 如果没有选择新进程，恢复原来的进程
            print("未选择新进程，保持原有设置")

    def select_target_process(self):
        """创建进程选择窗口"""
        # 先尝试自动匹配
        if self.auto_match and self.auto_select_process():
            return True

        root = tk.Tk()
        root.title("选择要监控的程序")
        root.geometry("400x450")

        # 创建主框架
        main_frame = ttk.Frame(root)
        main_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        # 添加标签
        label_text = "请选择要监控的游戏进程："
        if self.history_processes:
            label_text += "（绿色背景为历史记录）"
        ttk.Label(main_frame, text=label_text).pack(anchor='w')

        # 创建树形视图和滚动条
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(expand=True, fill=tk.BOTH, pady=(5,0))

        tree = ttk.Treeview(tree_frame, columns=('PID', 'Name'), show='headings', height=15)
        tree.heading('PID', text='进程ID')
        tree.heading('Name', text='进程名称')
        tree.column('PID', width=100)
        tree.column('Name', width=280)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        # 填充数据
        matched_items = []  # 用于存储匹配的历史进程项
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process:
                item = tree.insert('', tk.END, values=(session.Process.pid, session.Process.name()))
                if session.Process.name() in self.history_processes:
                    matched_items.append(item)
                    tree.item(item, tags=('history',))

        # 设置历史进程的特殊样式
        tree.tag_configure('history', background='#E8F5E9')  # 浅绿色背景

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 如果有匹配的历史进程，选中第一个
        if matched_items:
            tree.selection_set(matched_items[0])
            tree.see(matched_items[0])

        selected_result = False  # 用于跟踪是否成功选择了进程

        def on_select():
            nonlocal selected_result
            selected_item = tree.selection()
            if selected_item:
                pid, name = tree.item(selected_item[0])['values']
                self.target_pid = pid
                self.target_name = name
                self.last_muted_state[pid] = False
                self.add_to_history(name)  # 添加到历史记录
                selected_result = True
                root.destroy()
            else:
                messagebox.showwarning("提示", "请先选择一个进程")

        # 确认按钮
        ttk.Button(main_frame, text="确定", command=on_select, width=20).pack(pady=10)

        # 窗口居中显示
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        root.minsize(400, 450)

        # 双击选择功能
        tree.bind('<Double-1>', lambda e: on_select())
        
        # 处理窗口关闭
        def on_closing():
            nonlocal selected_result
            if not selected_result:
                self.target_pid = None
                self.target_name = None
            root.destroy()
            
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        root.mainloop()
        return selected_result  # 返回是否选择了进程

    def is_window_minimized(self, pid):
        """检查指定进程的窗口是否最小化"""
        def callback(hwnd, hwnds):
            try:
                _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                if window_pid == pid and win32gui.IsWindowVisible(hwnd):
                    hwnds.append(hwnd)
            except:
                pass
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        
        return all(win32gui.IsIconic(hwnd) for hwnd in hwnds) if hwnds else False

    def get_foreground_window_pid(self):
        """获取当前前台窗口的进程ID"""
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return pid
        except:
            return None

    def monitor_target_app(self):
        """监控目标应用的音频状态"""
        if not self.target_pid:
            print("请先选择要监控的程序！")
            return

        while self.running:
            try:
                # 检查目标进程是否仍在运行
                target_running = False
                sessions = AudioUtilities.GetAllSessions()
                
                # 首先检查目标进程是否存在
                for session in sessions:
                    if session.Process and session.Process.pid == self.target_pid:
                        target_running = True
                        break
                
                # 如果启用了自动关闭且目标进程不存在
                if self.auto_close and not target_running:
                    print(f"目标进程 {self.target_name} (PID: {self.target_pid}) 已结束，程序自动关闭")
                    self.stop_monitoring()
                    break

                if self.paused:
                    time.sleep(1)
                    continue

                foreground_pid = self.get_foreground_window_pid()
                if not foreground_pid:
                    time.sleep(1)
                    continue

                for session in sessions:
                    if not session.Process:
                        continue

                    volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                    pid = session.Process.pid

                    if pid == self.target_pid:
                        # 确定是否应该静音
                        should_mute = (self.is_window_minimized(pid) if self.minimize_only 
                                     else foreground_pid != pid)

                        # 更新静音状态
                        if should_mute != self.last_muted_state.get(pid, False):
                            volume.SetMute(should_mute, None)
                            self.last_muted_state[pid] = should_mute
                    
                    # 恢复非目标进程的音量
                    elif pid in self.last_muted_state and self.last_muted_state[pid]:
                        volume.SetMute(0, None)
                        del self.last_muted_state[pid]

            except Exception as e:
                print(f"监控过程中出现错误: {e}")
                time.sleep(1)

            time.sleep(1)

    def start(self):
        """启动程序"""
        # 先选择进程
        self.select_target_process()
        if not self.target_pid:
            return
            
        # 再创建托盘图标
        self.create_tray_icon()
        
        # 启动监控线程
        self.monitoring_thread = threading.Thread(target=self.monitor_target_app)
        self.monitoring_thread.daemon = True
        self.monitoring_thread.start()
        
        # 运行托盘图标
        self.tray_icon.run()

def main():
    controller = AudioController()
    controller.start()

if __name__ == "__main__":
    main()