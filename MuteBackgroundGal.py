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
        
    def create_tray_icon(self):
        """创建系统托盘图标"""
        # 创建一个简单的图标
        icon_size = 64
        image = Image.new('RGB', (icon_size, icon_size), color='white')
        drawing = ImageDraw.Draw(image)
        drawing.rectangle([20, 20, 44, 44], fill='black')
        
        def exit_with_confirmation():
            self.stop_monitoring()
            
        # 创建托盘图标菜单
        menu = (
            pystray.MenuItem("Galgame音频控制器", lambda: None, enabled=False),
            pystray.MenuItem("状态: 运行中", lambda: None, enabled=False),
            pystray.MenuItem("说明: 右键点击退出即可", lambda: None, enabled=False),
            pystray.MenuItem("退出", exit_with_confirmation)
        )
        
        # 创建托盘图标
        self.tray_icon = pystray.Icon(
            "galgame_audio_controller",
            image,
            "Galgame音频控制器\n右键点击退出",
            menu
        )

    def stop_monitoring(self):
        """停止监控并退出程序"""
        try:
            # 在退出前确保取消静音
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if (session.Process and 
                    session.Process.pid == self.target_pid):
                    volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                    # 强制取消静音
                    volume.SetMute(0, None)
                    break
        except Exception as e:
            print(f"退出时取消静音失败: {e}")
            
        self.running = False
        if self.tray_icon:
            self.tray_icon.stop()
        
    def select_target_process(self):
        """创建进程选择窗口"""
        root = tk.Tk()
        root.title("选择要监控的程序")
        root.geometry("400x300")
        
        frame = ttk.Frame(root)
        frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        tree = ttk.Treeview(frame, columns=('PID', 'Name'), show='headings')
        tree.heading('PID', text='进程ID')
        tree.heading('Name', text='进程名称')
        
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
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
                root.destroy()
                # 更新托盘图标状态文本
                if self.tray_icon:
                    new_menu = (
                        pystray.MenuItem(f"状态: 正在监控 {name}", lambda: None, enabled=False),
                        pystray.MenuItem("退出", self.stop_monitoring)
                    )
                    self.tray_icon.menu = new_menu
        
        select_btn = ttk.Button(root, text="确定", command=on_select)
        select_btn.pack(pady=10)
        
        # 显示使用提示
        tips_text = "提示：\n1. 选择要监控的游戏进程后点击确定\n2. 程序会最小化到系统托盘"
        tips_label = ttk.Label(root, text=tips_text, justify=tk.LEFT, wraplength=380)
        tips_label.pack(pady=10, padx=10)
        
        root.mainloop()
    
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
                foreground_pid = self.get_foreground_window_pid()
                if not foreground_pid:
                    time.sleep(1)
                    continue
                
                sessions = AudioUtilities.GetAllSessions()
                for session in sessions:
                    if (session.Process and 
                        session.Process.pid == self.target_pid):
                        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                        is_foreground = (foreground_pid == self.target_pid)
                        volume.SetMute(not is_foreground, None)
                        break
                
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