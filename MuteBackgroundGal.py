import time
import tkinter as tk
from tkinter import ttk
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import psutil
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
        self.minimize_only = True  # 默认设置为仅最小化时静音
        self.last_muted_state = {}  # 用于跟踪每个进程的静音状态

    def create_play_icon(self):
        """创建播放状态的图标（绿色三角形）"""
        icon_size = 64
        image = Image.new('RGB', (icon_size, icon_size), color='white')
        drawing = ImageDraw.Draw(image)
        
        # 绘制绿色三角形（播放图标）
        play_color = '#00C853'  # 使用鲜艳的绿色
        points = [
            (20, 15),  # 左上角
            (20, 49),  # 左下角
            (49, 32)   # 右中点
        ]
        drawing.polygon(points, fill=play_color)
        return image
        
    def create_pause_icon(self):
        """创建暂停状态的图标（红色双竖线）"""
        icon_size = 64
        image = Image.new('RGB', (icon_size, icon_size), color='white')
        drawing = ImageDraw.Draw(image)
        
        # 绘制红色双竖线（暂停图标）
        pause_color = '#D32F2F'  # 使用鲜艳的红色
        # 左边竖线
        drawing.rectangle([20, 15, 30, 49], fill=pause_color)
        # 右边竖线
        drawing.rectangle([39, 15, 49, 49], fill=pause_color)
        return image

    def toggle_minimize_only(self):
        """切换是否仅在最小化时静音"""
        self.minimize_only = not self.minimize_only
        self.update_icon_and_menu()

    def get_menu(self):
        """获取当前状态的菜单"""
        return pystray.Menu(
            pystray.MenuItem(
                "Galgame音频控制器",
                None,
                enabled=False
            ),
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
        # 初始使用播放图标
        initial_icon = self.create_play_icon()
        
        # 创建托盘图标
        self.tray_icon = pystray.Icon(
            "galgame_audio_controller",
            initial_icon,
            "Galgame音频控制器\n右键可暂停/继续",
            menu=self.get_menu()
        )

    def update_icon_and_menu(self):
        """更新图标和菜单状态"""
        if self.tray_icon:
            new_icon = self.create_pause_icon() if self.paused else self.create_play_icon()
            self.tray_icon.icon = new_icon
            self.tray_icon.menu = self.get_menu()

    def toggle_pause(self):
        """切换暂停状态"""
        self.paused = not self.paused
        if self.paused:
            # 暂停时确保取消静音
            try:
                sessions = AudioUtilities.GetAllSessions()
                for session in sessions:
                    if (session.Process and 
                        session.Process.pid == self.target_pid):
                        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                        volume.SetMute(0, None)  # 取消静音
                        break
            except Exception as e:
                print(f"暂停时取消静音失败: {e}")
        
        # 更新图标和菜单
        self.update_icon_and_menu()

    def stop_monitoring(self):
        """停止监控并退出程序"""
        try:
            # 在退出前确保取消所有进程的静音
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and session.Process.pid in self.last_muted_state:
                    volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                    volume.SetMute(0, None)  # 强制取消静音
        except Exception as e:
            print(f"退出时取消静音失败: {e}")
            
        self.running = False
        if self.tray_icon:
            self.tray_icon.stop()

    def reselect_process(self):
        """重新选择要监控的进程"""
        # 如果正在监控，先取消当前进程的静音状态
        if self.target_pid:
            try:
                sessions = AudioUtilities.GetAllSessions()
                for session in sessions:
                    if (session.Process and 
                        session.Process.pid == self.target_pid):
                        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                        volume.SetMute(0, None)  # 取消静音
                        break
            except Exception as e:
                print(f"取消静音失败: {e}")

        # 创建新的选择窗口
        self.select_target_process()
        
        # 更新图标和菜单
        self.update_icon_and_menu()

    def select_target_process(self):
        """创建进程选择窗口"""
        root = tk.Tk()
        root.title("选择要监控的程序")
        root.geometry("400x450")  # 增加窗口默认高度
        
        # 添加标签
        label = ttk.Label(root, text="请选择要监控的游戏进程：")
        label.pack(pady=(10,0), padx=10, anchor='w')
        
        # 创建主框架
        frame = ttk.Frame(root)
        frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        # 创建树形视图和滚动条
        tree = ttk.Treeview(frame, columns=('PID', 'Name'), show='headings', height=15)  # 设置固定高度
        tree.heading('PID', text='进程ID')
        tree.heading('Name', text='进程名称')
        tree.column('PID', width=100)  # 设置PID列宽
        tree.column('Name', width=280)  # 设置Name列宽
        
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # 填充数据
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process:
                tree.insert('', tk.END, values=(session.Process.pid, session.Process.name()))
        
        # 布局树形视图和滚动条
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建按钮框架
        button_frame = ttk.Frame(root)
        button_frame.pack(pady=10, fill=tk.X)
        
        def on_select():
            selected_item = tree.selection()
            if selected_item:
                pid, name = tree.item(selected_item[0])['values']
                self.target_pid = pid
                self.target_name = name
                # 初始化选定进程的静音状态
                self.last_muted_state[pid] = False
                root.destroy()
                # 立即更新托盘图标的菜单显示
                if self.tray_icon:
                    self.update_icon_and_menu()
            else:
                tk.messagebox.showwarning("提示", "请先选择一个进程")
        
        # 确认按钮
        select_btn = ttk.Button(button_frame, text="确定", command=on_select, width=20)
        select_btn.pack(pady=5)
        
        # 窗口居中显示
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        # 设置最小窗口大小
        root.minsize(400, 450)
        
        # 双击选择功能
        def on_double_click(event):
            on_select()
            
        tree.bind('<Double-1>', on_double_click)
        
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
        
        # 检查所有找到的窗口
        for hwnd in hwnds:
            if not win32gui.IsIconic(hwnd):  # 如果有任何一个窗口不是最小化的
                return False
        return len(hwnds) > 0  # 所有窗口都是最小化的（且至少找到了一个窗口）

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
                # 如果处于暂停状态，不进行音频控制
                if self.paused:
                    time.sleep(1)
                    continue
                    
                foreground_pid = self.get_foreground_window_pid()
                if not foreground_pid:
                    time.sleep(1)
                    continue
                
                sessions = AudioUtilities.GetAllSessions()
                for session in sessions:
                    if session.Process:
                        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                        pid = session.Process.pid
                        
                        # 如果是目标进程
                        if pid == self.target_pid:
                            # 检查窗口状态
                            is_minimized = self.is_window_minimized(pid)
                            is_foreground = (foreground_pid == pid)
                            
                            # 根据设置决定是否静音
                            should_mute = False
                            if self.minimize_only:
                                # 仅在最小化时静音
                                should_mute = is_minimized
                            else:
                                # 在非前台时静音
                                should_mute = not is_foreground
                            
                            # 更新静音状态
                            if should_mute != self.last_muted_state.get(pid, False):
                                volume.SetMute(should_mute, None)
                                self.last_muted_state[pid] = should_mute
                        
                        # 如果不是目标进程但在last_muted_state中（说明之前是目标进程）
                        elif pid in self.last_muted_state and self.last_muted_state[pid]:
                            # 恢复非目标进程的音量
                            volume.SetMute(0, None)
                            del self.last_muted_state[pid]  # 从跟踪列表中移除
                
                time.sleep(1)
                
            except Exception as e:
                print(f"监控过程中出现错误: {e}")
                time.sleep(1)

    def start(self):
        """启动程序"""
        # 创建托盘图标
        self.create_tray_icon()
        
        # 选择目标进程
        self.select_target_process()
        
        if self.target_pid:
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