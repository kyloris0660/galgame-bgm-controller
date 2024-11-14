import os
import shutil
import subprocess
import glob
import sys

def clean_build():
    """清理build和dist目录"""
    dirs_to_clean = ['build', 'dist']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
    # 清理所有.spec文件
    for file in os.listdir('.'):
        if file.endswith('.spec'):
            os.remove(file)

def post_build_cleanup():
    """构建后清理，保留完整的dist目录"""
    print("\nPerforming post-build cleanup...")
    
    # 只清理build目录和spec文件
    if os.path.exists('build'):
        shutil.rmtree('build')
    for file in os.listdir('.'):
        if file.endswith('.spec'):
            os.remove(file)
    
    # 清理临时文件
    temp_files = [
        'entry.py',
        'admin.manifest',
        '__pycache__',
        '*.pyc',
        '*.pyo',
        '*.pyd',
        '*.log'
    ]
    
    for pattern in temp_files:
        for file_path in glob.glob(pattern):
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Warning: Could not remove {file_path}: {e}")
    
    print("Cleanup completed. Final executable is in the 'dist/GalgameBGMController' directory.")

def install_requirements():
    """安装必要的依赖"""
    print("Installing required packages...")
    packages = [
        'pyinstaller',
        'pycaw',
        'psutil',
        'pywin32',
        'pillow',
        'pystray'
    ]
    for package in packages:
        subprocess.run(['pip', 'install', package], check=True)

def create_manifest():
    """创建管理员权限manifest文件"""
    manifest_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
    version="1.0.0.0"
    processorArchitecture="X86"
    name="GalgameBGMController"
    type="win32"
  />
  <description>Galgame BGM Controller</description>
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
    <security>
      <requestedPrivileges>
        <requestedExecutionLevel level="requireAdministrator" uiAccess="false"/>
      </requestedPrivileges>
    </security>
  </trustInfo>
</assembly>'''
    
    with open('admin.manifest', 'w', encoding='utf-8') as f:
        f.write(manifest_content)

def create_entry_script():
    """创建入口脚本"""
    entry_script = '''import MuteBackgroundGal

if __name__ == '__main__':
    MuteBackgroundGal.main()
'''
    with open('entry.py', 'w', encoding='utf-8') as f:
        f.write(entry_script)

def build_exe():
    """使用PyInstaller构建可执行文件"""
    print("Building executable...")
    
    # 创建入口脚本和manifest文件
    create_entry_script()
    create_manifest()
    
    # PyInstaller命令行参数
    pyinstaller_args = [
        'pyinstaller',
        '--noconfirm',
        '--clean',
        '--name=GalgameBGMController',
        '--icon=app.ico' if os.path.exists('app.ico') else None,  # 如果有图标则使用
        '--noconsole',  # 不显示控制台窗口
        '--uac-admin',  # 请求管理员权限
        '--hidden-import=pystray._win32',
        '--hidden-import=pkg_resources.py2_warn',
        '--hidden-import=win32api',
        '--hidden-import=win32gui',
        '--hidden-import=win32con',
        '--hidden-import=PIL._tkinter_finder',
        '--collect-all=pycaw',
        '--collect-all=pystray',
        '--collect-all=PIL',
        '--add-data=gal_audio_controller_config.json;.',  # 添加配置文件
        '--add-data=MuteBackgroundGal.py;.',  # 添加主程序文件
        '--add-binary=%s;.' % os.path.join(sys.prefix, 'python3.dll'),  # 添加Python DLL
        '--add-binary=%s;.' % os.path.join(sys.prefix, 'python39.dll' if sys.version_info.minor == 9 else f'python3{sys.version_info.minor}.dll'),  # 添加版本特定的Python DLL
        '--manifest=admin.manifest',  # 使用管理员权限manifest
        '--onedir',  # 生成目录形式的输出
        'entry.py'  # 使用新的入口脚本
    ]
    
    # 移除None值
    pyinstaller_args = [arg for arg in pyinstaller_args if arg is not None]
    
    # 执行PyInstaller
    subprocess.run(pyinstaller_args, check=True)

def create_default_config():
    """创建默认配置文件"""
    if not os.path.exists('gal_audio_controller_config.json'):
        default_config = '''{
    "history_processes": [],
    "auto_match": true,
    "minimize_only": true,
    "auto_close": true
}'''
        with open('gal_audio_controller_config.json', 'w', encoding='utf-8') as f:
            f.write(default_config)

def main():
    try:
        # 清理旧的构建文件
        clean_build()
        
        # 安装必要的依赖
        install_requirements()
        
        # 创建默认配置文件
        create_default_config()
        
        # 构建可执行文件
        build_exe()
        
        # 执行构建后清理
        post_build_cleanup()
        
        print("\nBuild completed successfully!")
        print("Final executable and dependencies are in the 'dist/GalgameBGMController' directory")
        
    except subprocess.CalledProcessError as e:
        print(f"Error during build process: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == '__main__':
    main()