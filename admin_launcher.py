import sys
import os
import ctypes
import subprocess
import MuteBackgroundGal

def is_admin():
    """检查是否具有管理员权限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """以管理员权限重新运行脚本"""
    if is_admin():
        # 如果已经是管理员权限，则直接运行主程序
        MuteBackgroundGal.main()
    else:
        # 如果不是管理员权限，则请求提升
        script = os.path.abspath(sys.argv[0])
        params = ' '.join(sys.argv[1:])
        
        # 请求UAC提权
        ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            sys.executable,
            f'"{script}" {params}',
            None,
            1
        )

if __name__ == '__main__':
    run_as_admin()