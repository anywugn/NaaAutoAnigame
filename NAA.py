import os
import time
import subprocess
from datetime import datetime, timedelta
import tkinter as tk
import sys
import psutil
import json
import threading
import pystray
from PIL import Image
import ctypes
import ctypes.wintypes
from ctypes import POINTER, cast
from tkinter import messagebox
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# 检测并请求管理员权限（如果需要）
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def resource_path(relative_path):
    """获取资源文件的绝对路径，兼容 PyInstaller 打包后的路径"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后的临时文件夹路径
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_config(config_file):
    """从配置文件中加载配置，如果不存在则创建默认配置"""
    default_config = {
        "program_list": [
            "D:\\ProgramPortable\\ZenlessZoneZero-OneDragon\\OneDragon Scheduler.exe",
            "D:\\ProgramPortable\\March7thAssistant\\March7th Assistant.exe"
        ],
        "process_names": [
            "ZenlessZoneZero.exe",
            "StarRail.exe"
        ],
        "run_hour": 12,
        "run_minute": 5,
        "countdown_duration": 10,
        "shutdown_delay": 600,
        "auto_shutdown": False,
        "auto_startup": False
    }

    if not os.path.exists(config_file):
        print(f"配置文件 {config_file} 不存在，正在创建默认配置文件。")
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, ensure_ascii=False, indent=4)
        print(f"已创建默认配置文件 {config_file}，请根据需要修改。")
        # 返回默认配置
        return default_config
    else:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config

def manage_startup(auto_startup):
    """管理开机自启动（使用任务计划程序）"""
    import subprocess

    # 任务名称，可以根据需要修改
    TASK_NAME = "NAAAutoStart"

    # 获取当前脚本的绝对路径
    current_script = os.path.abspath(sys.argv[0])

    if auto_startup:
        # 创建计划任务
        cmd = [
            "schtasks", "/Create", "/F",
            "/SC", "ONLOGON",
            "/RL", "HIGHEST",
            "/TN", TASK_NAME,
            "/TR", f'"{current_script}"'
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                print("已添加开机自启动（任务计划程序）。")
            else:
                print("添加开机自启动失败：", result.stderr)
        except Exception as e:
            print("添加开机自启动时发生异常：", e)
    else:
        # 删除计划任务
        cmd = ["schtasks", "/Delete", "/F", "/TN", TASK_NAME]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                print("已移除开机自启动（任务计划程序）。")
            else:
                print("移除开机自启动失败：", result.stderr)
        except Exception as e:
            print("移除开机自启动时发生异常：", e)

def show_countdown(countdown):
    root = tk.Tk()
    root.title("倒计时")
    root.geometry("600x200")
    # 使窗口不在任务栏显示
    root.attributes('-toolwindow', True)
    root.attributes('-topmost', True)
    root.resizable(False, False)
    label = tk.Label(root, font=("Arial", 24))
    label.pack(expand=True)

    def update_label():
        nonlocal countdown
        if countdown >= 0:
            minutes, seconds = divmod(countdown, 60)
            label.config(text=f"NAA将在 {minutes}分{seconds}秒 后开始运行")
            countdown -= 1
            root.after(1000, update_label)
        else:
            root.destroy()

    update_label()
    root.mainloop()

def is_process_running(process_name):
    """检查进程是否正在运行"""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == process_name:
            return True
    return False

def wait_for_process_start(process_name, timeout=30):
    """等待进程启动"""
    print(f"等待进程 {process_name} 启动...")
    start_time = time.time()
    while True:
        if is_process_running(process_name):
            print(f"进程 {process_name} 已启动。")
            break
        if time.time() - start_time > timeout:
            print(f"等待进程 {process_name} 启动超时。")
            break
        time.sleep(1)

def wait_for_process_exit(process_name):
    """等待进程退出"""
    print(f"等待进程 {process_name} 退出...")
    while is_process_running(process_name):
        time.sleep(1)
    print(f"进程 {process_name} 已退出。")

def run_programs(config):

    mute_system()

    program_list = config['program_list']
    process_names = config['process_names']
    shutdown_delay = config['shutdown_delay']
    auto_shutdown = config['auto_shutdown']

    for program, process_name in zip(program_list, process_names):
        try:
            print(f"正在启动程序：{program}")
            # 启动程序
            subprocess.Popen([program])
            # 等待进程启动
            wait_for_process_start(process_name)
            # 等待指定的进程退出
            wait_for_process_exit(process_name)
            print(f"程序 {program} 已完成。")
        except Exception as e:
            print(f"程序 {program} 执行出错：", e)

    unmute_system()
    if auto_shutdown:
        # 设置自动关机
        os.system(f"shutdown -s -t {shutdown_delay}")
        print(f"系统将在 {shutdown_delay} 秒后自动关机。")

def task_thread(config, run_now_event):
    """任务线程"""
    run_hour = config['run_hour']
    run_minute = config['run_minute']
    countdown_duration = config['countdown_duration']

    while True:
        now = datetime.now()
        run_time = now.replace(
            hour=run_hour, minute=run_minute, second=0, microsecond=0)
        if now >= run_time:
            # 如果当前时间已经超过指定时间，跳到下一天
            run_time += timedelta(days=1)
        time_to_wait = (run_time - now).total_seconds()

        # 等待直到指定时间前倒计时的秒数，或者等待事件被设置
        wait_time = max(0, time_to_wait - countdown_duration)
        if run_now_event.wait(timeout=wait_time):
            # 事件被设置，立即运行程序序列
            print("接收到立即运行程序序列的指令。")
        else:
            # 时间到了，继续执行
            pass

        if run_now_event.is_set():
            run_now_event.clear()
            show_countdown(countdown=0)
            run_programs(config)
            #break  # 如果需要每天运行，注释掉此行
        else:
            # 开始倒计时提醒
            show_countdown(countdown=countdown_duration)
            # 等待剩余的时间，或者等待事件被设置
            now = datetime.now()
            remaining_seconds = (run_time - now).total_seconds()
            if remaining_seconds > 0:
                if run_now_event.wait(timeout=remaining_seconds):
                    print("接收到立即运行程序序列的指令。")
                    run_now_event.clear()
            # 运行程序序列
            run_programs(config)
            break  # 如果需要每天运行，注释掉此行

def create_tray_icon(run_now_event):
    """创建系统托盘图标"""

    # 加载图标
    icon_path = resource_path('icon.ico')  # 图标文件路径
    if not os.path.exists(icon_path):
        print("未找到图标文件 icon.ico，请确保图标文件存在。")
        sys.exit(1)
    else:
        icon_image = Image.open(icon_path)

    def on_ignore_timer(icon, item):
        """忽略定时器，立即运行程序序列"""
        run_now_event.set()
        print("已触发立即运行程序序列的指令。")

    def on_exit(icon, item):
        """退出程序"""
        icon.stop()
        sys.exit()

    # 创建托盘图标
    menu = pystray.Menu(
        pystray.MenuItem('忽略定时，立即运行', on_ignore_timer),
        pystray.MenuItem('退出', on_exit)
    )

    tray_icon = pystray.Icon('MyProgram', icon_image, 'NAA正在后台待命', menu)

    # 启动托盘图标
    tray_icon.run()

def mute_system():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMute(1, None)  # 设置静音

def unmute_system():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMute(0, None)  # 取消静音

def main():

    # 获取当前脚本所在的目录
    current_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    # 将工作目录设置为脚本所在目录
    os.chdir(current_dir)

    if not is_admin():
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        messagebox.showerror("需要管理员权限", "NAA需要以管理员权限运行。")
        sys.exit()    

    global config
    # 加载配置
    config = load_config('config.json')

    # 管理开机自启动
    auto_startup = config.get('auto_startup', False)
    manage_startup(auto_startup)

    # 创建线程事件
    run_now_event = threading.Event()

    # 启动任务线程
    t = threading.Thread(target=task_thread, args=(config, run_now_event))
    t.daemon = True
    t.start()

    # 创建系统托盘图标
    create_tray_icon(run_now_event)

if __name__ == "__main__":
    main()
