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
        self.last_muted_state = {}

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
                "重新选择进程",
                self.reselect_process
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
        self.select_target_process()
        self.update_icon_and_menu()

    def select_target_process(self):
        """创建进程选择窗口"""
        root = tk.Tk()
        root.title("选择要监控的程序")
        root.geometry("400x450")

        # 创建主框架
        main_frame = ttk.Frame(root)
        main_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        # 添加标签
        ttk.Label(main_frame, text="请选择要监控的游戏进程：").pack(anchor='w')

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
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process:
                tree.insert('', tk.END, values=(session.Process.pid, session.Process.name()))

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def on_select():
            selected_item = tree.selection()
            if selected_item:
                pid, name = tree.item(selected_item[0])['values']
                self.target_pid = pid
                self.target_name = name
                self.last_muted_state[pid] = False
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