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
import logging
from pathlib import Path

class AudioController:
    def __init__(self):
        self.target_processes = {}  # 改为字典，存储 {pid: name} 的映射
        self.running = True
        self.monitoring_thread = None
        self.tray_icon = None
        self.paused = False
        self.history_processes = set()  # 先初始化为空集合
        
        # 修改配置文件路径到当前目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(current_dir, 'gal_audio_controller_config.json')
        
        # 加载配置
        config = self.load_config()
        self.history_processes = set(config.get('history_processes', []))
        self.minimize_only = config.get('minimize_only', True)
        self.auto_close = config.get('auto_close', True)
        self.auto_match = config.get('auto_match', True)
        self.last_muted_state = {}

    def load_config(self):
        """加载配置文件"""
        default_config = {
            'history_processes': [],
            'auto_match': True,
            'minimize_only': True,
            'auto_close': False
        }
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
            logging.info(f"加载配置文件失败: {e}")
            return default_config

    def save_config(self):
        """保存配置文件"""
        try:
            config = {
                'history_processes': list(self.history_processes),
                'auto_match': self.auto_match,
                'minimize_only': self.minimize_only,
                'auto_close': self.auto_close
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.info(f"保存配置文件失败: {e}")

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
        def add_process_callback(icon, item):
            self.add_process()
            
        def manage_processes_callback(icon, item):
            self.manage_processes()
            
        def toggle_pause_callback(icon, item):
            self.toggle_pause()
            
        def toggle_minimize_callback(icon, item):
            self.toggle_minimize_only()
            
        def toggle_auto_close_callback(icon, item):
            self.toggle_auto_close()
            
        def toggle_auto_match_callback(icon, item):
            self.toggle_auto_match()
            
        def clear_history_callback(icon, item):
            self.clear_history()
            
        def stop_monitoring_callback(icon, item):
            self.stop_monitoring()

        return pystray.Menu(
            pystray.MenuItem("Galgame音频控制器", None, enabled=False),
            pystray.MenuItem(
                f"状态: {'已暂停' if self.paused else '监控中'}", 
                None, 
                enabled=False
            ),
            pystray.MenuItem(
                f"{'继续监控' if self.paused else '暂停监控'}", 
                toggle_pause_callback
            ),
            pystray.MenuItem(
                "仅最小化时静音",
                toggle_minimize_callback,
                checked=lambda item: self.minimize_only
            ),
            pystray.MenuItem(
                "进程结束时自动关闭",
                toggle_auto_close_callback,
                checked=lambda item: self.auto_close
            ),
            pystray.MenuItem(
                "自动匹配历史进程",
                toggle_auto_match_callback,
                checked=lambda item: self.auto_match
            ),
            pystray.MenuItem(
                "添加监控进程",
                add_process_callback
            ),
            pystray.MenuItem(
                "管理监控进程",
                manage_processes_callback
            ),
            pystray.MenuItem(
                "清空历史记录",
                clear_history_callback
            ),
            pystray.MenuItem(
                "退出程序",
                stop_monitoring_callback
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
        if self.paused and self.target_processes:
            self.restore_volume(self.target_processes.keys())
        self.update_icon_and_menu()

    def toggle_minimize_only(self):
        """切换是否仅在最小化时静音"""
        self.minimize_only = not self.minimize_only
        self.save_config()  # 保存设置
        self.update_icon_and_menu()

    def toggle_auto_close(self):
        """切换是否在目标进程结束时自动关闭"""
        self.auto_close = not self.auto_close
        self.save_config()  # 保存设置
        self.update_icon_and_menu()

    def toggle_auto_match(self):
        """切换是否自动匹配历史进程"""
        self.auto_match = not self.auto_match
        self.save_config()  # 保存设置
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
            
        found_new_process = False
        sessions = AudioUtilities.GetAllSessions()
        
        for session in sessions:
            if session.Process and session.Process.name() in self.history_processes:
                pid = session.Process.pid
                # 只添加还未在监控列表中的进程
                if pid not in self.target_processes:
                    self.target_processes[pid] = session.Process.name()
                    self.last_muted_state[pid] = False
                    logging.info(f"自动添加进程: {session.Process.name()} (PID: {pid})")
                    found_new_process = True
        
        return found_new_process

    def restore_volume(self, pids):
        """恢复指定进程的音量"""
        if not isinstance(pids, (list, set)):
            pids = [pids]
            
        try:
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and session.Process.pid in pids:
                    volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                    volume.SetMute(0, None)
        except Exception as e:
            logging.info(f"恢复音量失败: {e}")

    def restore_all_volumes(self):
        """恢复所有被跟踪进程的音量"""
        try:
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and session.Process.pid in self.last_muted_state:
                    volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                    volume.SetMute(0, None)
        except Exception as e:
            logging.info(f"恢复所有音量失败: {e}")

    def stop_monitoring(self):
        """停止监控并退出程序"""
        self.running = False
        self.restore_all_volumes()
        # 确保在退出前保存配置
        self.save_config()
        if self.tray_icon:
            self.tray_icon.stop()

    def add_process(self, icon=None, item=None):
        """添加新的监控进程"""
        # 直接显示进程选择窗口，不进行自动匹配
        self.select_target_process(skip_auto_match=True)
        self.update_icon_and_menu()

    def manage_processes(self):
        """管理当前监控的进程"""
        if not self.target_processes:
            messagebox.showinfo("提示", "当前没有监控的进程")
            return
            
        root = tk.Tk()
        root.title("管理监控进程")
        root.geometry("400x300")
        
        # 创建主框架
        main_frame = ttk.Frame(root)
        main_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        
        # 添加标签
        ttk.Label(main_frame, text="当前监控的进程:").pack(anchor='w')
        
        # 创建列表框和滚动条
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(expand=True, fill=tk.BOTH, pady=(5,0))
        
        listbox = tk.Listbox(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        
        # 填充数据
        for pid, name in self.target_processes.items():
            listbox.insert(tk.END, f"{name} (PID: {pid})")
        
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        def remove_selected():
            selected = listbox.curselection()
            if not selected:
                messagebox.showwarning("提示", "请先选择一个进程")
                return
                
            # 从列表中获取PID
            item_text = listbox.get(selected[0])
            pid = int(item_text.split("PID: ")[1].strip(")"))
            
            # 恢复音量并移除进程
            self.restore_volume(pid)
            if pid in self.target_processes:
                del self.target_processes[pid]
            if pid in self.last_muted_state:
                del self.last_muted_state[pid]
                
            # 更新列表
            listbox.delete(selected[0])
        
        ttk.Button(button_frame, text="移除选中进程", command=remove_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="关闭", command=root.destroy).pack(side=tk.RIGHT, padx=5)
        
        # 窗口居中显示
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        root.mainloop()

    def select_target_process(self, skip_auto_match=False):
        """创建进程选择窗口"""
        # 只在启动时自动匹配，手动添加进程时跳过自动匹配
        if not skip_auto_match and self.auto_match and self.auto_select_process():
            return True

        root = tk.Tk()
        root.title("选择要监控的程序")
        root.geometry("400x450")
        root.attributes('-topmost', True)  # 窗口置顶

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
                pid = session.Process.pid
                # 跳过已经在监控列表中的进程
                if pid in self.target_processes:
                    continue
                    
                item = tree.insert('', tk.END, values=(pid, session.Process.name()))
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

        def on_select():
            selected_item = tree.selection()
            if selected_item:
                pid, name = tree.item(selected_item[0])['values']
                pid = int(pid)  # 确保 pid 是整数
                self.target_processes[pid] = name
                self.last_muted_state[pid] = False
                self.add_to_history(name)  # 添加到历史记录
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
        root.protocol("WM_DELETE_WINDOW", root.destroy)
        
        root.mainloop()

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
        while self.running:
            try:
                # 不管是否有目标进程，都尝试自动匹配新进程
                if self.auto_match:
                    self.auto_select_process()
                    
                if not self.target_processes:
                    time.sleep(1)
                    continue

                if self.paused:
                    time.sleep(1)
                    continue

                # 获取当前所有音频会话
                sessions = AudioUtilities.GetAllSessions()
                
                # 检查目标进程是否仍在运行，移除已结束的进程
                active_pids = set()
                for session in sessions:
                    if session.Process:
                        active_pids.add(session.Process.pid)
                
                # 移除已结束的进程
                ended_processes = []
                for pid in list(self.target_processes.keys()):
                    if pid not in active_pids:
                        ended_processes.append((pid, self.target_processes[pid]))
                        del self.target_processes[pid]
                        if pid in self.last_muted_state:
                            del self.last_muted_state[pid]
                
                # 如果所有进程都结束且设置了自动关闭
                if ended_processes and not self.target_processes and self.auto_close:
                    process_names = ", ".join([name for _, name in ended_processes])
                    logging.info(f"所有监控进程已结束: {process_names}，程序自动关闭")
                    self.save_config()
                    self.stop_monitoring()
                    break
                
                foreground_pid = self.get_foreground_window_pid()
                if not foreground_pid:
                    time.sleep(1)
                    continue

                # 处理每个音频会话
                for session in sessions:
                    if not session.Process:
                        continue

                    pid = session.Process.pid
                    
                    # 只处理我们监控的进程
                    if pid in self.target_processes:
                        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                        
                        # 确定是否应该静音
                        should_mute = (self.is_window_minimized(pid) if self.minimize_only 
                                    else foreground_pid != pid)

                        # 更新静音状态
                        if should_mute != self.last_muted_state.get(pid, False):
                            volume.SetMute(should_mute, None)
                            self.last_muted_state[pid] = should_mute

                time.sleep(1)
                
            except Exception as e:
                logging.info(f"监控过程中出现错误: {e}")
                time.sleep(1)

    def start(self):
        """启动程序"""
        # 先尝试自动匹配进程
        if not self.auto_select_process():
            # 如果没有自动匹配到，则显示选择窗口
            self.select_target_process()
            
        # 创建托盘图标
        self.create_tray_icon()
        
        # 启动监控线程
        self.monitoring_thread = threading.Thread(target=self.monitor_target_app)
        self.monitoring_thread.daemon = True
        self.monitoring_thread.start()
        
        # 运行托盘图标
        self.tray_icon.run()

def setup_logging():
    """设置日志输出"""
    log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bgm_controller.log')
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

def main():
    setup_logging()  # 设置日志输出
    logging.info("="*50)
    logging.info("程序启动")

    controller = AudioController()
    controller.start()

if __name__ == "__main__":
    main()