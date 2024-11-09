import time
import tkinter as tk
from tkinter import ttk
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import psutil
import win32gui
import win32process

class AudioController:
    def __init__(self):
        self.target_pid = None
        self.target_name = None
        
    def select_target_process(self):
        """创建进程选择窗口"""
        # 创建主窗口
        root = tk.Tk()
        root.title("选择要监控的程序")
        root.geometry("400x300")
        
        # 创建进程列表
        frame = ttk.Frame(root)
        frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        # 创建树形视图
        tree = ttk.Treeview(frame, columns=('PID', 'Name'), show='headings')
        tree.heading('PID', text='进程ID')
        tree.heading('Name', text='进程名称')
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # 获取所有音频会话的进程
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process:
                tree.insert('', tk.END, values=(session.Process.pid, session.Process.name()))
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        def on_select():
            """当用户选择了一个进程"""
            selected_item = tree.selection()
            if selected_item:
                pid, name = tree.item(selected_item[0])['values']
                self.target_pid = pid
                self.target_name = name
                root.destroy()
        
        # 添加选择按钮
        select_btn = ttk.Button(root, text="确定", command=on_select)
        select_btn.pack(pady=10)
        
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
        
        print(f"开始监控程序: {self.target_name} (PID: {self.target_pid})")
        print("提示: 按Ctrl+C退出程序")
        
        try:
            while True:
                # 获取前台进程ID
                foreground_pid = self.get_foreground_window_pid()
                if not foreground_pid:
                    time.sleep(1)
                    continue
                
                # 获取所有音频会话
                sessions = AudioUtilities.GetAllSessions()
                
                # 只处理目标程序的音频
                for session in sessions:
                    if (session.Process and 
                        session.Process.pid == self.target_pid):
                        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                        
                        # 判断目标程序是否在前台
                        is_foreground = (foreground_pid == self.target_pid)
                        
                        # 设置音量状态
                        if is_foreground:
                            # 前台时恢复声音
                            volume.SetMute(0, None)
                            print(f"\r{self.target_name} 已恢复声音", end="")
                        else:
                            # 后台时静音
                            volume.SetMute(1, None)
                            print(f"\r{self.target_name} 已静音", end="")
                        
                        break
                
                time.sleep(1)
                
        except KeyboardInterrupt:
            print("\n程序已退出")

def main():
    controller = AudioController()
    # 先选择要监控的程序
    controller.select_target_process()
    # 开始监控
    controller.monitor_target_app()

if __name__ == "__main__":
    main()