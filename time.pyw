import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
import os
import ctypes
import time
import sys
import subprocess
import pkg_resources
import traceback
import win32event
import win32api
import winerror
import winsound
import requests
import json
import tempfile
import shutil
from pathlib import Path

# 檢查是否已經有實例在運行
mutex = win32event.CreateMutex(None, 1, 'TimeChangerMutex')
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    sys.exit(0)

# 隱藏命令視窗
if sys.platform == 'win32':
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# 程式資訊
VERSION = "1.5.7"
AUTHOR = "GFK"
CONTACT_EMAIL = "gfkwork928@gmail.com"

# 版本更新紀錄
VERSION_HISTORY = {
    "1.5.7": [
        "添加退出程式按鈕",
        "優化界面佈局"
    ],
    "1.5.6": [
        "改進更新功能顯示",
        "添加版本號檢查提示",
        "優化更新進度顯示"
    ],
    "1.5.5": [
        "新增多個 NTP 伺服器選項",
        "添加自動安裝必要套件功能",
        "優化程式啟動流程"
    ],
    "1.5.4": [
        "修改同步按鈕功能為同步本機時間",
        "優化命令視窗隱藏方式",
        "移除程式重複運行警告"
    ],
    "1.5.3": [
        "優化使用者介面",
        "修復已知問題",
        "改進時間同步功能"
    ],
    "1.5.2": [
        "添加聯繫信息",
        "優化界面佈局",
        "改進時間顯示"
    ],
    "1.5.1": [
        "改進錯誤處理",
        "優化程式碼結構",
        "改進時間更新機制"
    ],
    "1.5.0": [
        "優化系統時間顯示",
        "添加星期顯示",
        "改進時間更新機制"
    ],
    "1.4.0": [
        "隱藏命令視窗",
        "優化程式啟動流程",
        "改進錯誤處理機制"
    ],
    "1.3.0": [
        "修改為台灣用語",
        "修正標點符號",
        "調整文字說明"
    ],
    "1.2.0": [
        "修復時間顯示不同步問題",
        "優化介面配置和間距",
        "改進按鈕和輸入欄位樣式"
    ],
    "1.1.0": [
        "優化介面配置",
        "新增即時系統時間顯示",
        "改進時間顯示格式"
    ],
    "1.0.0": [
        "初始版本發行",
        "支援手動修改系統時間",
        "支援取得網路時間",
        "自動請求系統管理員權限"
    ]
}

def show_version_history():
    # 創建版本歷史視窗
    history_window = tk.Toplevel()
    history_window.title("版本歷程")
    history_window.geometry("600x500")  # 調整視窗大小
    history_window.resizable(True, True)
    
    # 設置主題
    style = ttk.Style()
    style.configure("Treeview", font=("微軟正黑體", 10))
    style.configure("Treeview.Heading", font=("微軟正黑體", 10, "bold"))
    style.configure("Separator.TSeparator", background="#cccccc")
    
    # 創建主框架
    main_frame = ttk.Frame(history_window)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    # 創建標題
    title_label = ttk.Label(main_frame, text="版本更新歷史", font=("微軟正黑體", 12, "bold"))
    title_label.pack(pady=(0, 10))
    
    # 創建滾動框架
    canvas = tk.Canvas(main_frame)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)
    
    # 配置滾動區域
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    # 創建窗口
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=canvas.winfo_reqwidth())
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # 添加版本數據
    for version, changes in VERSION_HISTORY.items():
        # 版本標題框架
        version_frame = ttk.Frame(scrollable_frame)
        version_frame.pack(fill="x", pady=(0, 5))
        
        # 版本號標籤
        version_label = ttk.Label(version_frame, text=f"版本 {version}", font=("微軟正黑體", 11, "bold"))
        version_label.pack(anchor="w", pady=(5, 0))
        
        # 更新內容框架
        changes_frame = ttk.Frame(version_frame)
        changes_frame.pack(fill="x", padx=20, pady=(5, 10))
        
        # 添加更新內容
        for change in changes:
            change_label = ttk.Label(changes_frame, text=f"• {change}", font=("微軟正黑體", 10))
            change_label.pack(anchor="w", pady=2)
        
        # 添加分隔線
        separator = ttk.Separator(version_frame, orient="horizontal", style="Separator.TSeparator")
        separator.pack(fill="x", pady=5)
    
    # 設置佈局
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # 綁定滑鼠滾輪事件
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    # 綁定滾輪事件到所有相關組件
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    scrollable_frame.bind_all("<MouseWheel>", _on_mousewheel)
    
    # 設置視窗置中
    history_window.update_idletasks()
    width = history_window.winfo_width()
    height = history_window.winfo_height()
    x = (history_window.winfo_screenwidth() // 2) - (width // 2)
    y = (history_window.winfo_screenheight() // 2) - (height // 2)
    history_window.geometry(f"{width}x{height}+{x}+{y}")
    
    # 綁定關閉事件
    def on_closing():
        # 解除滾輪事件綁定
        canvas.unbind_all("<MouseWheel>")
        scrollable_frame.unbind_all("<MouseWheel>")
        history_window.destroy()
    
    history_window.protocol("WM_DELETE_WINDOW", on_closing)

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit()

# 檢查管理員權限
if not is_admin():
    run_as_admin()

def check_internet_connection():
    try:
        import socket
        # 嘗試連接 Google DNS
        socket.create_connection(("8.8.8.8", 53), timeout=2)
        return True
    except OSError:
        return False

def install_required_packages():
    required_packages = ['ntplib']
    installed_packages = {pkg.key for pkg in pkg_resources.working_set}
    
    for package in required_packages:
        if package not in installed_packages:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                print(f"成功安裝 {package}")
            except subprocess.CalledProcessError as e:
                print(f"安裝 {package} 失敗: {str(e)}")
                return False
    return True

# 檢查網路連接
if not check_internet_connection():
    messagebox.showerror(
        "錯誤",
        "無法連接到網際網路！\n\n"
        "本軟件需要網路連接才能正常運行。\n"
        "請確保：\n"
        "1. 電腦已連接到網際網路\n"
        "2. 防火牆未阻止網路連接\n"
        "3. 網路設定正確\n\n"
        "點擊確定後將關閉程式。"
    )
    sys.exit(1)

# 檢查並安裝必要套件
try:
    if not install_required_packages():
        messagebox.showerror(
            "錯誤",
            "無法安裝必要套件，程式將關閉。\n\n"
            "請確保：\n"
            "1. 電腦已連接到網際網路\n"
            "2. 有足夠的權限安裝套件\n"
            "3. Python 環境配置正確"
        )
        sys.exit(1)
except Exception as e:
    messagebox.showerror(
        "錯誤",
        f"套件安裝過程發生錯誤：{str(e)}\n\n"
        "程式將關閉。"
    )
    sys.exit(1)

try:
    import ntplib
    from socket import gaierror
    HAS_NTPLIB = True
except ImportError:
    messagebox.showerror(
        "錯誤",
        "無法導入必要模組，程式將關閉。\n\n"
        "請確保：\n"
        "1. 已成功安裝 ntplib 套件\n"
        "2. Python 環境配置正確"
    )
    sys.exit(1)

# GitHub相關設定
GITHUB_OWNER = "fly4803"
GITHUB_REPO = "time-setter"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"

def check_for_updates():
    """檢查是否有新版本可用"""
    try:
        response = requests.get(GITHUB_API_URL)
        if response.status_code == 200:
            latest_release = response.json()
            latest_version = latest_release['tag_name'].lstrip('v')
            
            # 比較版本號
            if latest_version > VERSION:
                assets = latest_release.get('assets', [])
                if not assets:
                    raise Exception("找不到更新檔案")
                    
                download_url = assets[0]['browser_download_url']
                update_info = latest_release.get('body', '無更新說明')
                
                if messagebox.askyesno(
                    "更新可用",
                    f"目前版本：v{VERSION}\n"
                    f"最新版本：v{latest_version}\n\n"
                    f"更新說明：\n{update_info}\n\n"
                    "是否要更新？"
                ):
                    download_update(download_url, latest_version)
                    return True
            else:
                messagebox.showinfo(
                    "檢查更新",
                    f"目前版本：v{VERSION}\n"
                    "您使用的已經是最新版本！"
                )
        return False
    except Exception as e:
        messagebox.showerror(
            "更新檢查失敗",
            f"檢查更新時發生錯誤：\n{str(e)}\n\n請確保網路連接正常後再試。"
        )
        return False

def download_update(download_url, latest_version):
    """下載更新檔案"""
    try:
        # 顯示下載進度視窗
        progress_window = tk.Toplevel()
        progress_window.title(f"下載更新 v{latest_version}")
        progress_window.geometry("300x150")
        progress_window.resizable(False, False)
        progress_window.transient()  # 設置為置頂視窗
        progress_window.grab_set()   # 鎖定其他視窗操作
        
        # 設置視窗置中
        window_width = 300
        window_height = 150
        screen_width = progress_window.winfo_screenwidth()
        screen_height = progress_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        progress_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 進度標籤
        progress_label = ttk.Label(
            progress_window,
            text=f"正在下載 v{latest_version} 版本更新...",
            font=("微軟正黑體", 10)
        )
        progress_label.pack(pady=20)
        
        # 進度條
        progress_bar = ttk.Progressbar(
            progress_window,
            length=200,
            mode='determinate'
        )
        progress_bar.pack(pady=10)
        
        # 更新進度視窗
        progress_window.update()
        
        # 下載文件
        response = requests.get(download_url, stream=True)
        if response.status_code == 200:
            # 獲取文件大小
            total_size = int(response.headers.get('content-length', 0))
            block_size = 1024  # 1 KB
            downloaded = 0
            
            # 創建臨時文件
            with tempfile.NamedTemporaryFile(delete=False, suffix='.exe') as tmp_file:
                for data in response.iter_content(block_size):
                    tmp_file.write(data)
                    downloaded += len(data)
                    # 更新進度條
                    if total_size:
                        progress = (downloaded / total_size) * 100
                        progress_bar['value'] = progress
                        progress_window.update()
                tmp_path = tmp_file.name
            
            progress_label.config(text=f"正在準備安裝 v{latest_version}...")
            progress_window.update()
            
            # 建立更新批次檔
            batch_path = os.path.join(tempfile.gettempdir(), 'update.bat')
            with open(batch_path, 'w', encoding='utf-8') as f:
                f.write('@echo off\n')
                f.write('title 正在更新系統時間設定工具\n')
                f.write(f'echo 正在更新至 v{latest_version}，請稍候...\n')
                f.write('ping 127.0.0.1 -n 3 > nul\n')  # 等待原程式結束
                f.write(f'move /y "{tmp_path}" "{sys.executable}"\n')
                f.write('if errorlevel 1 (\n')
                f.write('    echo 更新失敗！請確保程式已完全關閉後重試。\n')
                f.write('    pause\n')
                f.write('    exit /b 1\n')
                f.write(')\n')
                f.write('echo 更新完成！正在啟動新版本...\n')
                f.write(f'start "" "{sys.executable}"\n')
                f.write('ping 127.0.0.1 -n 2 > nul\n')
                f.write('del "%~f0"\n')  # 自刪除批次檔
            
            # 關閉進度視窗
            progress_window.destroy()
            
            # 顯示更新提示
            messagebox.showinfo(
                "更新準備就緒", 
                f"v{latest_version} 版本更新檔案已下載完成，\n"
                "程式將重新啟動以完成更新。\n\n"
                "注意：如果更新失敗，請確保：\n"
                "1. 程式完全關閉\n"
                "2. 以系統管理員身份運行"
            )
            
            # 執行更新批次檔
            subprocess.Popen(
                ['cmd.exe', '/c', batch_path],
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            sys.exit(0)
        else:
            progress_window.destroy()
            raise Exception(f"下載失敗：伺服器回應代碼 {response.status_code}")
            
    except Exception as e:
        try:
            progress_window.destroy()
        except:
            pass
        messagebox.showerror(
            "更新下載失敗",
            f"下載更新時發生錯誤：\n{str(e)}\n\n"
            "請確保：\n"
            "1. 網路連接正常\n"
            "2. 程式以系統管理員身份運行\n"
            "3. 防毒軟體未阻止下載"
        )

class TimeChangerApp:
    def __init__(self, root):
        # 再次檢查網路連接
        if not check_internet_connection():
            messagebox.showerror(
                "錯誤",
                "網路連接已中斷！\n\n"
                "本軟件需要持續的網路連接才能正常運行。\n"
                "請重新連接網路後再試。"
            )
            root.destroy()
            return
            
        self.root = root
        self.root.title(f"系統時間設定工具 v{VERSION} - {AUTHOR}")
        self.root.geometry("650x580")  # 調整視窗大小
        self.root.resizable(True, True)
        
        # 設置主題
        self.style = ttk.Style()
        
        # 建立主容器
        self.main_container = ttk.Frame(self.root, padding="15")
        self.main_container.pack(fill="both", expand=True)
        
        # 建立所有組件
        self.create_title_frame()
        self.create_time_display_frame()
        self.create_input_frame()
        self.create_button_frame()
        self.create_preset_time_frame()
        self.create_version_frame()
        
        # 設置淺色主題
        self.set_light_theme()
        
        # 開始更新目前時間顯示
        self.update_current_time()
        
        # 設置視窗置中
        self.center_window()
        
        # 檢查更新
        self.check_updates()

    def set_light_theme(self):
        bg_color = "#ffffff"
        fg_color = "#000000"
        button_bg = "#f0f0f0"
        entry_bg = "#ffffff"
        entry_fg = "#000000"
        label_frame_fg = "#000000"
        
        # 設置基本樣式
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", background=bg_color, foreground=fg_color)
        self.style.configure("TButton", background=button_bg)
        
        # 設置輸入框樣式
        self.style.configure("TEntry", fieldbackground=entry_bg)
        
        # 設置標籤框架樣式
        self.style.configure("TLabelframe", background=bg_color)
        self.style.configure("TLabelframe.Label", background=bg_color, foreground=label_frame_fg)
        
        # 設置當前時間標籤樣式
        if hasattr(self, 'current_time_label'):
            self.current_time_label.configure(foreground="#0078d7")
        
        # 設置根窗口背景
        self.root.configure(bg=bg_color)
        
        # 更新所有輸入框的顏色
        if hasattr(self, 'year_entry'):
            entries = [self.year_entry, self.month_entry, self.day_entry,
                      self.hour_entry, self.minute_entry, self.second_entry]
            for entry in entries:
                entry.configure(foreground=entry_fg)
                entry.configure(background=entry_bg)
        
        # 更新所有時間標籤的顏色
        if hasattr(self, 'time_labels'):
            for label in self.time_labels:
                label.configure(foreground=fg_color, background=bg_color)

    def create_title_frame(self):
        title_frame = ttk.Frame(self.main_container)
        title_frame.pack(fill="x", pady=(0, 15))
        
        title_label = ttk.Label(
            title_frame,
            text="系統時間設定工具",
            font=("微軟正黑體", 16, "bold")
        )
        title_label.pack(side="left")

    def create_time_display_frame(self):
        time_frame = ttk.LabelFrame(self.main_container, text="系統時間", padding="10")
        time_frame.pack(fill="x", pady=(0, 15))
        
        # 創建時間顯示容器
        display_frame = ttk.Frame(time_frame)
        display_frame.pack(fill="x", expand=True)
        
        # 日期和星期顯示
        date_frame = ttk.Frame(display_frame)
        date_frame.pack(fill="x", pady=(0, 5))
        
        self.date_label = ttk.Label(
            date_frame,
            font=("微軟正黑體", 14),
            foreground="#333333"
        )
        self.date_label.pack(side="left")
        
        self.weekday_label = ttk.Label(
            date_frame,
            font=("微軟正黑體", 14),
            foreground="#666666"
        )
        self.weekday_label.pack(side="right")
        
        # 時間顯示和同步按鈕
        time_control_frame = ttk.Frame(display_frame)
        time_control_frame.pack(fill="x", pady=(5, 0))
        
        self.time_label = ttk.Label(
            time_control_frame,
            font=("微軟正黑體", 32, "bold"),
            foreground="#0078d7"
        )
        self.time_label.pack(side="left", expand=True)
        
        sync_button = ttk.Button(
            time_control_frame,
            text="同步",
            command=self.sync_system_time,
            width=8
        )
        sync_button.pack(side="right", padx=5)

    def create_input_frame(self):
        input_frame = ttk.Frame(self.main_container)
        input_frame.pack(fill="x", pady=(0, 15))
        
        # 日期輸入框架
        date_frame = ttk.LabelFrame(input_frame, text="日期", padding="10")
        date_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # 年月日輸入框水平排列
        date_entries_frame = ttk.Frame(date_frame)
        date_entries_frame.pack(fill="x", expand=True)
        
        # 年
        year_frame = ttk.Frame(date_entries_frame)
        year_frame.pack(side="left", expand=True, padx=5)
        year_label = ttk.Label(year_frame, text="年")
        year_label.pack()
        self.year_entry = ttk.Entry(year_frame, width=8, justify="center")
        self.year_entry.pack()
        self.year_entry.insert(0, datetime.now().strftime("%Y"))
        
        # 月
        month_frame = ttk.Frame(date_entries_frame)
        month_frame.pack(side="left", expand=True, padx=5)
        month_label = ttk.Label(month_frame, text="月")
        month_label.pack()
        self.month_entry = ttk.Entry(month_frame, width=4, justify="center")
        self.month_entry.pack()
        self.month_entry.insert(0, datetime.now().strftime("%m"))
        
        # 日
        day_frame = ttk.Frame(date_entries_frame)
        day_frame.pack(side="left", expand=True, padx=5)
        day_label = ttk.Label(day_frame, text="日")
        day_label.pack()
        self.day_entry = ttk.Entry(day_frame, width=4, justify="center")
        self.day_entry.pack()
        self.day_entry.insert(0, datetime.now().strftime("%d"))
        
        # 時間輸入框架
        time_frame = ttk.LabelFrame(input_frame, text="時間", padding="10")
        time_frame.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # 時分秒輸入框水平排列
        time_entries_frame = ttk.Frame(time_frame)
        time_entries_frame.pack(fill="x", expand=True)
        
        # 時
        hour_frame = ttk.Frame(time_entries_frame)
        hour_frame.pack(side="left", expand=True, padx=5)
        hour_label = ttk.Label(hour_frame, text="時")
        hour_label.pack()
        self.hour_entry = ttk.Entry(hour_frame, width=4, justify="center")
        self.hour_entry.pack()
        self.hour_entry.insert(0, datetime.now().strftime("%H"))
        
        # 分
        minute_frame = ttk.Frame(time_entries_frame)
        minute_frame.pack(side="left", expand=True, padx=5)
        minute_label = ttk.Label(minute_frame, text="分")
        minute_label.pack()
        self.minute_entry = ttk.Entry(minute_frame, width=4, justify="center")
        self.minute_entry.pack()
        self.minute_entry.insert(0, datetime.now().strftime("%M"))
        
        # 秒
        second_frame = ttk.Frame(time_entries_frame)
        second_frame.pack(side="left", expand=True, padx=5)
        second_label = ttk.Label(second_frame, text="秒")
        second_label.pack()
        self.second_entry = ttk.Entry(second_frame, width=4, justify="center")
        self.second_entry.pack()
        self.second_entry.insert(0, datetime.now().strftime("%S"))

    def create_preset_time_frame(self):
        preset_frame = ttk.LabelFrame(self.main_container, text="快速設定", padding="10")
        preset_frame.pack(fill="x", pady=(0, 15))
        
        # 創建內部框架以確保按鈕居中
        inner_frame = ttk.Frame(preset_frame)
        inner_frame.pack(expand=True)
        
        # 預設時間按鈕
        preset_buttons = [
            ("07:59:20", "07:59:20"),
            ("08:00:06", "08:00:06")
        ]
        
        for text, time_str in preset_buttons:
            btn = ttk.Button(
                inner_frame,
                text=text,
                command=lambda t=time_str: self.set_preset_time(t),
                width=10
            )
            btn.pack(side="left", padx=10)

    def create_button_frame(self):
        button_frame = ttk.Frame(self.main_container)
        button_frame.pack(fill="x", pady=(0, 15))
        
        # 創建內部框架以確保按鈕居中
        inner_frame = ttk.Frame(button_frame)
        inner_frame.pack(expand=True)
        
        buttons = [
            ("更新系統時間", self.update_system_time),
            ("取得網路時間", self.get_network_time),
            ("NTP 伺服器", self.show_ntp_servers),
            ("刷新", self.refresh_input_time)
        ]
        
        for text, command in buttons:
            btn = ttk.Button(
                inner_frame,
                text=text,
                command=command,
                width=15
            )
            btn.pack(side="left", padx=10)

    def create_version_frame(self):
        version_frame = ttk.Frame(self.main_container)
        version_frame.pack(fill="x")
        
        # 左側版本資訊
        version_info_frame = ttk.Frame(version_frame)
        version_info_frame.pack(side="left")
        
        version_label = ttk.Label(
            version_info_frame,
            text=f"版本：{VERSION}",
            font=("微軟正黑體", 9)
        )
        version_label.pack(side="left", padx=(0, 10))
        
        author_label = ttk.Label(
            version_info_frame,
            text=f"作者：{AUTHOR}",
            font=("微軟正黑體", 9)
        )
        author_label.pack(side="left", padx=(0, 10))
        
        # 建立可點擊的郵件連結
        email_label = ttk.Label(
            version_info_frame,
            text=f"聯繫：{CONTACT_EMAIL}",
            font=("微軟正黑體", 9),
            foreground="#0078d7",
            cursor="hand2"
        )
        email_label.pack(side="left")
        email_label.bind("<Button-1>", lambda e: self.open_email())
        
        # 右側按鈕
        button_frame = ttk.Frame(version_frame)
        button_frame.pack(side="right")
        
        # 退出按鈕
        exit_button = ttk.Button(
            button_frame,
            text="退出程式",
            command=self.root.quit,
            width=10
        )
        exit_button.pack(side="right", padx=(5, 0))
        
        # 檢查更新按鈕
        update_button = ttk.Button(
            button_frame,
            text="檢查更新",
            command=self.check_updates,
            width=10
        )
        update_button.pack(side="right", padx=(5, 0))
        
        # 版本歷程按鈕
        history_button = ttk.Button(
            button_frame,
            text="版本歷程",
            command=show_version_history,
            width=10
        )
        history_button.pack(side="right", padx=(5, 0))

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def update_current_time(self):
        current_time = datetime.now()
        
        # 更新日期和星期顯示
        weekday_names = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
        weekday = weekday_names[current_time.weekday()]
        
        date_str = current_time.strftime("%Y年%m月%d日")
        self.date_label.config(text=date_str)
        self.weekday_label.config(text=weekday)
        
        # 更新時間顯示
        time_str = current_time.strftime("%H:%M:%S")
        self.time_label.config(text=time_str)
        
        # 使用after_cancel來避免重複調用
        if hasattr(self, '_after_id'):
            self.root.after_cancel(self._after_id)
        
        # 計算到下一秒的延遲時間（微秒級別）
        next_second = datetime.now().replace(microsecond=0) + timedelta(seconds=1)
        delay = int((next_second - datetime.now()).total_seconds() * 1000)
        
        # 確保延遲時間至少為1毫秒
        delay = max(1, delay)
        
        # 設置下一次更新
        self._after_id = self.root.after(delay, self.update_current_time)

    def refresh_input_time(self):
        current_time = datetime.now()
        self.year_entry.delete(0, tk.END)
        self.year_entry.insert(0, current_time.strftime("%Y"))
        self.month_entry.delete(0, tk.END)
        self.month_entry.insert(0, current_time.strftime("%m"))
        self.day_entry.delete(0, tk.END)
        self.day_entry.insert(0, current_time.strftime("%d"))
        self.hour_entry.delete(0, tk.END)
        self.hour_entry.insert(0, current_time.strftime("%H"))
        self.minute_entry.delete(0, tk.END)
        self.minute_entry.insert(0, current_time.strftime("%M"))
        self.second_entry.delete(0, tk.END)
        self.second_entry.insert(0, current_time.strftime("%S"))

    def update_system_time(self, ntp_time=None):
        try:
            if ntp_time:
                # 更新系統時間
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                
                subprocess.run(f'date {ntp_time.strftime("%Y-%m-%d")}', startupinfo=startupinfo, shell=True)
                subprocess.run(f'time {ntp_time.strftime("%H:%M:%S")}', startupinfo=startupinfo, shell=True)
                
                self.update_current_time()
            else:
                try:
                    year = self.year_entry.get().strip()
                    month = self.month_entry.get().strip()
                    day = self.day_entry.get().strip()
                    hour = self.hour_entry.get().strip()
                    minute = self.minute_entry.get().strip()
                    second = self.second_entry.get().strip()
                    
                    if not all([year, month, day, hour, minute, second]):
                        raise ValueError("所有欄位都必須填寫")
                    
                    year = int(year)
                    month = int(month)
                    day = int(day)
                    hour = int(hour)
                    minute = int(minute)
                    second = int(second)
                    
                    if not (1 <= month <= 12):
                        raise ValueError("月份必須在 1-12 之間")
                    if not (1 <= day <= 31):
                        raise ValueError("日期必須在 1-31 之間")
                    if not (0 <= hour <= 23):
                        raise ValueError("小時必須在 0-23 之間")
                    if not (0 <= minute <= 59):
                        raise ValueError("分鐘必須在 0-59 之間")
                    if not (0 <= second <= 59):
                        raise ValueError("秒數必須在 0-59 之間")
                    
                    date_str = f"{year:04d}-{month:02d}-{day:02d}"
                    time_str = f"{hour:02d}:{minute:02d}:{second:02d}"
                    
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    
                    subprocess.run(f'date {date_str}', startupinfo=startupinfo, shell=True)
                    subprocess.run(f'time {time_str}', startupinfo=startupinfo, shell=True)
                    
                    self.update_current_time()
                    
                except ValueError as e:
                    messagebox.showerror("錯誤", f"請輸入正確的數字！\n{str(e)}")
                    return
            
        except Exception as e:
            messagebox.showerror("錯誤", f"無法更新系統時間：{str(e)}")

    def get_network_time(self):
        if not HAS_NTPLIB:
            messagebox.showwarning(
                "功能受限",
                "網路時間同步功能目前無法使用。\n\n"
                "請確保：\n"
                "1. 已安裝 ntplib 套件\n"
                "2. 電腦已連接到網際網路"
            )
            return
            
        try:
            ntp_servers = [
                'pool.ntp.org',
                'time.windows.com',
                'time.nist.gov',
                'time.google.com',
                'time.apple.com',
                'ntp.aliyun.com',
                'ntp1.aliyun.com',
                'ntp2.aliyun.com',
                'ntp3.aliyun.com',
                'ntp4.aliyun.com',
                'ntp5.aliyun.com',
                'ntp6.aliyun.com',
                'ntp7.aliyun.com'
            ]
            
            client = ntplib.NTPClient()
            last_error = None
            
            for server in ntp_servers:
                try:
                    response = client.request(server, timeout=2)
                    net_time = datetime.fromtimestamp(response.tx_time)
                    
                    # 更新系統時間
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    
                    subprocess.run(f'date {net_time.strftime("%Y-%m-%d")}', startupinfo=startupinfo, shell=True)
                    subprocess.run(f'time {net_time.strftime("%H:%M:%S")}', startupinfo=startupinfo, shell=True)
                    
                    self.update_current_time()
                    return net_time
                    
                except (ntplib.NTPException, gaierror) as e:
                    last_error = e
                    continue
                except Exception as e:
                    last_error = e
                    continue
            
            if last_error:
                raise last_error
            
        except Exception as e:
            error_msg = f"無法取得網路時間：{str(e)}\n\n"
            error_msg += "可能的原因：\n"
            error_msg += "1. 網路連接不穩定\n"
            error_msg += "2. 防火牆阻止了 NTP 請求\n"
            error_msg += "3. 所有 NTP 伺服器暫時無法訪問\n"
            error_msg += "\n請檢查網路連接後重試"
            messagebox.showerror("錯誤", error_msg)
            return None

    def sync_system_time(self):
        """同步系統時間"""
        try:
            # 獲取當前本機時間
            current_time = datetime.now()
            
            # 更新系統時間
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            
            subprocess.run(f'date {current_time.strftime("%Y-%m-%d")}', startupinfo=startupinfo, shell=True)
            subprocess.run(f'time {current_time.strftime("%H:%M:%S")}', startupinfo=startupinfo, shell=True)
            
            # 更新輸入框
            self.refresh_input_time()
            # 更新顯示
            self.update_current_time()
            
        except Exception as e:
            messagebox.showerror("錯誤", f"同步失敗：{str(e)}")

    def show_features_info(self):
        # 創建功能介紹視窗
        info_window = tk.Toplevel()
        info_window.title("功能說明")
        info_window.geometry("600x500")
        info_window.resizable(True, True)
        
        # 設置主題
        style = ttk.Style()
        style.configure("Info.TLabel", font=("微軟正黑體", 10))
        style.configure("InfoTitle.TLabel", font=("微軟正黑體", 11, "bold"))
        
        # 創建主框架
        main_frame = ttk.Frame(info_window, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # 創建滾動區域
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=550)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 功能說明內容
        features_info = [
            {
                "title": "系統時間顯示",
                "description": "顯示當前系統時間，包含日期、星期和時間。時間會自動更新，確保顯示準確。"
            },
            {
                "title": "同步按鈕",
                "description": "點擊後會將系統時間同步為本機時間，用於修正系統時間偏差。"
            },
            {
                "title": "手動時間設定",
                "description": "可以手動輸入年、月、日、時、分、秒來設定系統時間。\n限制：\n- 月份必須在 1-12 之間\n- 日期必須在 1-31 之間\n- 小時必須在 0-23 之間\n- 分鐘和秒數必須在 0-59 之間"
            },
            {
                "title": "預設時間按鈕",
                "description": "提供快速設定常用時間的按鈕：\n- 07:59:20\n- 08:00:06\n點擊後會立即將系統時間設為指定時間。"
            },
            {
                "title": "取得網路時間",
                "description": "從多個NTP伺服器獲取網路時間並同步。\n注意：需要網路連接才能使用此功能。"
            },
            {
                "title": "刷新按鈕",
                "description": "將所有輸入框更新為當前系統時間，方便進行小幅度的時間調整。"
            }
        ]
        
        # 添加說明內容
        for feature in features_info:
            # 標題
            title_label = ttk.Label(
                scrollable_frame,
                text=feature["title"],
                style="InfoTitle.TLabel"
            )
            title_label.pack(anchor="w", pady=(10, 5))
            
            # 說明文字
            desc_label = ttk.Label(
                scrollable_frame,
                text=feature["description"],
                style="Info.TLabel",
                wraplength=530,
                justify="left"
            )
            desc_label.pack(anchor="w", padx=(20, 0))
            
            # 分隔線
            ttk.Separator(scrollable_frame, orient="horizontal").pack(fill="x", pady=10)
        
        # 設置滾動區域
        canvas.pack(side="left", fill="both", expand=True, pady=(0, 10))
        scrollbar.pack(side="right", fill="y")
        
        # 綁定滑鼠滾輪
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 設置視窗置中
        info_window.update_idletasks()
        width = info_window.winfo_width()
        height = info_window.winfo_height()
        x = (info_window.winfo_screenwidth() // 2) - (width // 2)
        y = (info_window.winfo_screenheight() // 2) - (height // 2)
        info_window.geometry(f"{width}x{height}+{x}+{y}")
        
        # 綁定關閉事件
        def on_closing():
            canvas.unbind_all("<MouseWheel>")
            info_window.destroy()
        
        info_window.protocol("WM_DELETE_WINDOW", on_closing)

    def set_preset_time(self, time_str):
        try:
            # 解析時間字符串
            hour, minute, second = map(int, time_str.split(':'))
            
            # 獲取當前日期
            current_date = datetime.now()
            
            # 設置時間
            self.hour_entry.delete(0, tk.END)
            self.hour_entry.insert(0, f"{hour:02d}")
            
            self.minute_entry.delete(0, tk.END)
            self.minute_entry.insert(0, f"{minute:02d}")
            
            self.second_entry.delete(0, tk.END)
            self.second_entry.insert(0, f"{second:02d}")
            
            # 更新系統時間
            self.update_system_time()
            
        except Exception as e:
            messagebox.showerror("錯誤", f"設定預設時間時發生錯誤：{str(e)}")

    def check_updates(self):
        """檢查並處理更新"""
        if check_for_updates():
            messagebox.showinfo("更新", "更新檔案已下載完成，程式將重新啟動以完成更新。")
            self.root.quit()

    def open_email(self):
        """開啟郵件程式"""
        os.startfile(f"mailto:{CONTACT_EMAIL}")

    def show_ntp_servers(self):
        # 創建 NTP 伺服器選擇視窗
        ntp_window = tk.Toplevel()
        ntp_window.title("NTP 伺服器設定")
        ntp_window.geometry("400x500")
        ntp_window.resizable(True, True)
        
        # 設置主題
        style = ttk.Style()
        style.configure("NTP.TCheckbutton", font=("微軟正黑體", 10))
        
        # 創建主框架
        main_frame = ttk.Frame(ntp_window, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # 創建滾動區域
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=350)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # NTP 伺服器列表
        ntp_servers = {
            'pool.ntp.org': True,
            'time.windows.com': True,
            'time.nist.gov': True,
            'time.google.com': True,
            'time.apple.com': True,
            'ntp.aliyun.com': True,
            'ntp1.aliyun.com': False,
            'ntp2.aliyun.com': False,
            'ntp3.aliyun.com': False,
            'ntp4.aliyun.com': False,
            'ntp5.aliyun.com': False,
            'ntp6.aliyun.com': False,
            'ntp7.aliyun.com': False,
            'time.cloudflare.com': True,
            'time.facebook.com': True,
            'time.asia.apple.com': True
        }
        
        # 創建標題
        title_label = ttk.Label(
            scrollable_frame,
            text="選擇要使用的 NTP 伺服器：",
            font=("微軟正黑體", 11, "bold")
        )
        title_label.pack(anchor="w", pady=(0, 10))
        
        # 創建複選框
        var_dict = {}
        for server, default_state in ntp_servers.items():
            var = tk.BooleanVar(value=default_state)
            var_dict[server] = var
            cb = ttk.Checkbutton(
                scrollable_frame,
                text=server,
                variable=var,
                style="NTP.TCheckbutton"
            )
            cb.pack(anchor="w", pady=2)
        
        # 創建按鈕框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        def save_settings():
            # 保存設置
            selected_servers = [server for server, var in var_dict.items() if var.get()]
            if not selected_servers:
                messagebox.showwarning("警告", "請至少選擇一個 NTP 伺服器！")
                return
            ntp_window.destroy()
        
        # 添加保存按鈕
        save_button = ttk.Button(
            button_frame,
            text="保存設置",
            command=save_settings,
            width=15
        )
        save_button.pack(side="right", padx=5)
        
        # 設置滾動區域
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 綁定滑鼠滾輪
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 設置視窗置中
        ntp_window.update_idletasks()
        width = ntp_window.winfo_width()
        height = ntp_window.winfo_height()
        x = (ntp_window.winfo_screenwidth() // 2) - (width // 2)
        y = (ntp_window.winfo_screenheight() // 2) - (height // 2)
        ntp_window.geometry(f"{width}x{height}+{x}+{y}")
        
        # 綁定關閉事件
        def on_closing():
            canvas.unbind_all("<MouseWheel>")
            ntp_window.destroy()
        
        ntp_window.protocol("WM_DELETE_WINDOW", on_closing)

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = TimeChangerApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("錯誤", f"程式發生錯誤：{str(e)}") 