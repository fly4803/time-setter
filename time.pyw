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
from pathlib import Path
import logging
from logging.handlers import RotatingFileHandler

# 檢查是否已經有實例在運行
mutex = win32event.CreateMutex(None, 1, 'TimeChangerMutex')
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    sys.exit(0)

# 隱藏命令視窗
if sys.platform == 'win32':
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# 設定程式碼簽名
def verify_signature():
    try:
        import win32security
        import win32api
        import win32con
        import win32file
        
        # 獲取當前執行檔路徑
        exe_path = sys.executable
        
        # 獲取檔案安全資訊
        security_info = win32security.GetFileSecurity(
            exe_path,
            win32security.OWNER_SECURITY_INFORMATION | win32security.GROUP_SECURITY_INFORMATION
        )
        
        # 獲取檔案屬性
        file_attrs = win32file.GetFileAttributes(exe_path)
        
        # 檢查是否為系統檔案
        if file_attrs & win32con.FILE_ATTRIBUTE_SYSTEM:
            return True
            
        return True
    except:
        return True

# 驗證程式碼簽名
if not verify_signature():
    messagebox.showerror(
        "錯誤",
        "程式碼簽名驗證失敗！\n\n"
        "請確保：\n"
        "1. 程式未被修改\n"
        "2. 以系統管理員身份運行\n"
        "3. 防毒軟體未阻止程式運行"
    )
    sys.exit(1)

# 程式資訊
VERSION = "25.1.6.8"
AUTHOR = "GFK"
CONTACT_EMAIL = "fly4803@gmail.com"

# 版本更新紀錄
VERSION_HISTORY = {
    "25.1.6.8": [
        "修正更新後檔案名稱亂碼問題",
        "改進更新機制",
        "優化檔案下載功能"
    ],
    "25.1.6.7": [
        "修正程式名稱顯示亂碼問題",
        "更新程式名稱為英文",
        "優化介面顯示"
    ],
    "25.1.6.6": [
        "更新軟體名稱為 System Time Change",
        "修正檔案版本資訊",
        "更新語言設定為中文(台灣)",
        "優化程式碼安全性",
        "改進防毒軟體相容性"
    ],
    "25.1.6.5": [
        "優化介面排列",
        "改進手動輸入框功能",
        "新增數值增減按鈕",
        "修復版本更新後名稱亂碼問題",
        "優化程式碼安全性"
    ],
    "25.1.6.4": [
        "新增頂部選單欄",
        "移除底部按鈕",
        "優化使用者介面"
    ],
    "25.1.6.3": [
        "改進快速設定功能",
        "優化時間格式驗證",
        "修復重複名稱檢查"
    ],
    "25.1.6.2": [
        "改進錯誤處理機制",
        "優化程式碼結構",
        "修復已知問題"
    ],
    "25.1.6.1": [
        "修復 NTP 伺服器設定無法保存的問題",
        "改進設定檔案的保存機制",
        "添加設定保存驗證"
    ],
    "25.1.6.0": [
        "新增自定義快速設定時間功能",
        "支援保存自定義時間設定",
        "優化快速設定介面"
    ],
    "25.1.5.9": [
        "移除時間模板功能",
        "添加功能介紹按鈕",
        "優化界面佈局"
    ],
    "25.1.5.8": [
        "修復程式更新後無法運行的問題",
        "優化程式打包配置"
    ],
    "25.1.5.7": [
        "添加退出程式按鈕",
        "優化界面佈局"
    ],
    "25.1.5.6": [
        "改進更新功能顯示",
        "添加版本號檢查提示",
        "優化更新進度顯示"
    ],
    "25.1.5.5": [
        "新增多個 NTP 伺服器選項",
        "添加自動安裝必要套件功能",
        "優化程式啟動流程"
    ],
    "25.1.5.4": [
        "修改同步按鈕功能為同步本機時間",
        "優化命令視窗隱藏方式",
        "移除程式重複運行警告"
    ],
    "25.1.5.3": [
        "優化使用者介面",
        "修復已知問題",
        "改進時間同步功能"
    ],
    "25.1.5.2": [
        "添加聯繫信息",
        "優化界面佈局",
        "改進時間顯示"
    ],
    "25.1.5.1": [
        "改進錯誤處理",
        "優化程式碼結構",
        "改進時間更新機制"
    ],
    "25.1.5.0": [
        "優化系統時間顯示",
        "添加星期顯示",
        "改進時間更新機制"
    ],
    "25.1.4.0": [
        "隱藏命令視窗",
        "優化程式啟動流程",
        "改進錯誤處理機制"
    ],
    "25.1.3.0": [
        "修正標點符號",
        "調整文字說明"
    ],
    "25.1.2.0": [
        "修復時間顯示不同步問題",
        "優化介面配置和間距",
        "改進按鈕和輸入欄位樣式"
    ],
    "25.1.1.0": [
        "優化介面配置",
        "新增即時系統時間顯示",
        "改進時間顯示格式"
    ],
    "25.1.0.0": [
        "初始版本發行",
        "支援手動修改系統時間",
        "支援取得網路時間",
        "自動請求系統管理員權限"
    ]
}

# 設定檔路徑
CONFIG_DIR = os.path.join(os.path.expanduser("~"), "AppData", "Local", "TimeSetter")
CONFIG_FILE = os.path.join(CONFIG_DIR, "settings.json")
LOG_DIR = os.path.join(CONFIG_DIR, "logs")
SYSTEM_LOG_FILE = os.path.join(LOG_DIR, "system.log")
TIME_CHANGES_LOG_FILE = os.path.join(LOG_DIR, "time_changes.log")
DRIFT_LOG_FILE = os.path.join(LOG_DIR, "time_drift.log")

# 確保設定目錄存在
os.makedirs(CONFIG_DIR, exist_ok=True)

# 確保目錄存在
os.makedirs(LOG_DIR, exist_ok=True)

# 設定系統日誌
system_logger = logging.getLogger('system')
system_logger.setLevel(logging.INFO)
system_handler = RotatingFileHandler(SYSTEM_LOG_FILE, maxBytes=1024*1024, backupCount=5, encoding='utf-8')
system_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
system_logger.addHandler(system_handler)

# 設定時間修改日誌
time_logger = logging.getLogger('time_changes')
time_logger.setLevel(logging.INFO)
time_handler = RotatingFileHandler(TIME_CHANGES_LOG_FILE, maxBytes=1024*1024, backupCount=5, encoding='utf-8')
time_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
time_logger.addHandler(time_handler)

# 設定時間偏差日誌
drift_logger = logging.getLogger('time_drift')
drift_logger.setLevel(logging.INFO)
drift_handler = RotatingFileHandler(DRIFT_LOG_FILE, maxBytes=1024*1024, backupCount=5, encoding='utf-8')
drift_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
drift_logger.addHandler(drift_handler)

# 預設設定
DEFAULT_SETTINGS = {
    "preset_times": [
        {"name": "07:59:20", "time": "07:59:20"},
        {"name": "08:00:06", "time": "08:00:06"}
    ],
    "ntp_servers": {
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
}

def load_settings():
    """載入設定"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                # 確保所有必要的設定都存在
                if "preset_times" not in settings:
                    settings["preset_times"] = DEFAULT_SETTINGS["preset_times"]
                if "ntp_servers" not in settings:
                    settings["ntp_servers"] = DEFAULT_SETTINGS["ntp_servers"]
                return settings
    except Exception as e:
        print(f"載入設定失敗：{str(e)}")
    return DEFAULT_SETTINGS.copy()

def save_settings(settings):
    """儲存設定"""
    try:
        # 確保目錄存在
        os.makedirs(CONFIG_DIR, exist_ok=True)
        
        # 寫入前先檢查設定內容
        if not isinstance(settings, dict):
            raise ValueError("設定必須是字典格式")
            
        # 檢查必要的設定項
        if "preset_times" not in settings:
            settings["preset_times"] = DEFAULT_SETTINGS["preset_times"]
        if "ntp_servers" not in settings:
            settings["ntp_servers"] = DEFAULT_SETTINGS["ntp_servers"]
            
        # 寫入設定檔
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)
            f.flush()  # 確保立即寫入磁碟
            os.fsync(f.fileno())  # 強制同步到磁碟
            
        # 驗證設定是否正確保存
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved_settings = json.load(f)
                if saved_settings != settings:
                    raise ValueError("設定驗證失敗")
                    
    except Exception as e:
        error_msg = f"儲存設定失敗：{str(e)}"
        print(error_msg)  # 輸出到控制台
        messagebox.showerror("錯誤", error_msg)
        return False
    return True

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
        try:
            executable = os.path.abspath(sys.executable)
            # 使用正確的檔案名稱
            if os.path.basename(executable).lower() != "systemtimechanger.exe":
                executable = os.path.join(os.path.dirname(executable), "SystemTimeChanger.exe")
            ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, " ".join(sys.argv), None, 1)
            sys.exit()
        except Exception as e:
            messagebox.showerror("錯誤", f"無法以管理員身份執行：{str(e)}")
            sys.exit(1)

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
            
            # 創建臨時文件，使用正確的檔案名稱
            temp_filename = f'SystemTimeChanger_v{latest_version}.exe'
            temp_path = os.path.join(tempfile.gettempdir(), temp_filename)
            
            # 建立更新批次檔
            batch_path = os.path.join(tempfile.gettempdir(), 'update.bat')
            with open(batch_path, 'w', encoding='utf-8') as f:
                f.write('@echo off\n')
                f.write('title 正在更新系統時間設定工具\n')
                f.write(f'echo 正在更新至 v{latest_version}，請稍候...\n')
                f.write('timeout /t 2 /nobreak > nul\n')
                target_exe = os.path.join(os.path.dirname(sys.executable), "SystemTimeChanger.exe")
                f.write(f'copy /y "{temp_path}" "{target_exe}"\n')
                f.write('if errorlevel 1 (\n')
                f.write('    echo 更新失敗！請確保程式已完全關閉後重試。\n')
                f.write('    pause\n')
                f.write('    exit /b 1\n')
                f.write(')\n')
                f.write('echo 更新完成！正在啟動新版本...\n')
                f.write(f'start "" "{target_exe}"\n')
                f.write('timeout /t 2 /nobreak > nul\n')
                f.write('del "%~f0"\n')
            
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
        self.root.geometry("700x520")  # 調整視窗大小
        self.root.resizable(False, False)  # 禁止調整視窗大小
        self.root.minsize(700, 520)  # 設置最小視窗大小
        self.root.maxsize(700, 520)  # 設置最大視窗大小
        
        # 載入設定
        self.settings = load_settings()
        
        # 創建選單列
        self.create_menu()
        
        # 設置主題
        self.style = ttk.Style()
        
        # 建立主容器
        self.main_container = ttk.Frame(self.root, padding="15")  # 調整內邊距
        self.main_container.pack(fill="both", expand=True)
        
        # 建立所有組件（調整順序）
        self.create_title_frame()
        self.create_time_display_frame()
        self.create_input_frame()
        self.create_preset_time_frame()
        self.create_button_frame()
        self.create_version_frame()
        
        # 設置淺色主題
        self.set_light_theme()
        
        # 開始更新目前時間顯示
        self.update_current_time()
        
        # 設置視窗置中
        self.center_window()

    def create_menu(self):
        """創建選單列"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 檔案選單
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="檔案", menu=file_menu)
        file_menu.add_command(label="退出", command=self.root.quit, accelerator="Alt+F4")
        
        # 工具選單
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="工具", menu=tools_menu)
        tools_menu.add_command(label="NTP 伺服器設定", command=self.show_ntp_servers, accelerator="Ctrl+N")
        tools_menu.add_separator()
        tools_menu.add_command(label="查看系統日誌", command=lambda: self.show_log_viewer("system"), accelerator="Ctrl+L")
        tools_menu.add_command(label="查看時間修改記錄", command=lambda: self.show_log_viewer("time"), accelerator="Ctrl+T")
        tools_menu.add_command(label="查看時間偏差記錄", command=lambda: self.show_log_viewer("drift"), accelerator="Ctrl+D")
        
        # 說明選單
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="說明", menu=help_menu)
        help_menu.add_command(label="功能介紹", command=self.show_features_info, accelerator="F1")
        help_menu.add_command(label="版本歷程", command=show_version_history, accelerator="Ctrl+H")
        help_menu.add_command(label="檢查更新", command=self.check_updates, accelerator="Ctrl+U")
        help_menu.add_separator()
        help_menu.add_command(label="關於", command=lambda: messagebox.showinfo(
            "關於",
            f"系統時間設定工具 v{VERSION}\n\n"
            f"作者：{AUTHOR}\n"
            f"聯繫：{CONTACT_EMAIL}"
        ))
        
        # 綁定快速鍵
        self.root.bind("<F1>", lambda e: self.show_features_info())
        self.root.bind("<Control-h>", lambda e: show_version_history())
        self.root.bind("<Control-u>", lambda e: self.check_updates())
        self.root.bind("<Control-n>", lambda e: self.show_ntp_servers())
        self.root.bind("<Control-l>", lambda e: self.show_log_viewer("system"))
        self.root.bind("<Control-t>", lambda e: self.show_log_viewer("time"))
        self.root.bind("<Control-d>", lambda e: self.show_log_viewer("drift"))

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
        title_frame.pack(fill="x", pady=(0, 10))  # 調整間距
        
        title_label = ttk.Label(
            title_frame,
            text="系統時間設定工具",
            font=("微軟正黑體", 14, "bold")  # 調整字體大小
        )
        title_label.pack(side="left")

    def create_time_display_frame(self):
        time_frame = ttk.LabelFrame(self.main_container, text="系統時間", padding="10")
        time_frame.pack(fill="x", pady=(0, 10))  # 調整間距
        
        # 創建時間顯示容器
        display_frame = ttk.Frame(time_frame)
        display_frame.pack(fill="x", expand=True)
        
        # 日期和星期顯示
        date_frame = ttk.Frame(display_frame)
        date_frame.pack(fill="x", pady=(0, 5))
        
        self.date_label = ttk.Label(
            date_frame,
            font=("微軟正黑體", 12),  # 調整字體大小
            foreground="#333333"
        )
        self.date_label.pack(side="left")
        
        self.weekday_label = ttk.Label(
            date_frame,
            font=("微軟正黑體", 12),  # 調整字體大小
            foreground="#666666"
        )
        self.weekday_label.pack(side="right")
        
        # 時間顯示和同步按鈕
        time_control_frame = ttk.Frame(display_frame)
        time_control_frame.pack(fill="x", pady=(5, 0))
        
        self.time_label = ttk.Label(
            time_control_frame,
            font=("微軟正黑體", 28, "bold"),  # 調整字體大小
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
        input_frame.pack(fill="x", pady=(0, 10))  # 調整間距
        self.input_frame = input_frame

        def is_leap_year(year):
            """檢查是否為閏年"""
            try:
                year = int(year)
                return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
            except ValueError:
                return False

        def get_days_in_month(year, month):
            """獲取指定年月的天數"""
            try:
                year = int(year)
                month = int(month)
                
                if month in [4, 6, 9, 11]:
                    return 30
                elif month == 2:
                    return 29 if is_leap_year(year) else 28
                elif month in [1, 3, 5, 7, 8, 10, 12]:
                    return 31
                else:
                    return 31  # 預設值
            except ValueError:
                return 31  # 預設值

        def validate_day_with_month():
            """根據當前年月驗證日期"""
            try:
                year = self.year_entry.get().strip()
                month = self.month_entry.get().strip()
                day = self.day_entry.get().strip()
                
                if not year or not month or not day:
                    return
                    
                year = int(year)
                month = int(month)
                day = int(day)
                
                max_days = get_days_in_month(year, month)
                
                if day > max_days:
                    self.day_entry.delete(0, tk.END)
                    self.day_entry.insert(0, str(max_days))
                    messagebox.showwarning(
                        "警告",
                        f"{year}年{month}月最多只有{max_days}天！\n已自動調整為{max_days}日。"
                    )
            except ValueError:
                pass

        # 日期輸入框架
        date_frame = ttk.LabelFrame(input_frame, text="日期", padding="10")
        date_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # 年月日輸入框水平排列
        date_entries_frame = ttk.Frame(date_frame)
        date_entries_frame.pack(fill="x", expand=True)
        
        # 驗證函數
        def validate_year(P):
            if P == "": return True
            if len(P) > 4: return False
            try:
                if not P.isdigit(): return False
                year = int(P)
                if len(P) == 4 and not (1900 <= year <= 9999):
                    return False
                return True
            except ValueError:
                return False

        def validate_month(P):
            if P == "": return True
            if len(P) > 2: return False
            try:
                if not P.isdigit(): return False
                if len(P) == 2:
                    month = int(P)
                    if not (1 <= month <= 12):
                        return False
                return True
            except ValueError:
                return False

        def validate_day(P):
            if P == "": return True
            if len(P) > 2: return False
            try:
                if not P.isdigit(): return False
                if len(P) == 2:
                    day = int(P)
                    if not (1 <= day <= 31):
                        return False
                return True
            except ValueError:
                return False

        def validate_hour(P):
            if P == "": return True
            if len(P) > 2: return False
            try:
                if not P.isdigit(): return False
                if len(P) == 2:
                    hour = int(P)
                    if not (0 <= hour <= 23):
                        return False
                return True
            except ValueError:
                return False

        def validate_minute_second(P):
            if P == "": return True
            if len(P) > 2: return False
            try:
                if not P.isdigit(): return False
                if len(P) == 2:
                    value = int(P)
                    if not (0 <= value <= 59):
                        return False
                return True
            except ValueError:
                return False

        # 註冊驗證器
        vcmd_year = (self.root.register(validate_year), '%P')
        vcmd_month = (self.root.register(validate_month), '%P')
        vcmd_day = (self.root.register(validate_day), '%P')
        vcmd_hour = (self.root.register(validate_hour), '%P')
        vcmd_minute_second = (self.root.register(validate_minute_second), '%P')

        def create_spinbox_frame(parent, label_text, width, validate_cmd, min_val, max_val):
            frame = ttk.Frame(parent)
            frame.pack(side="left", expand=True, padx=5)
            
            label = ttk.Label(frame, text=label_text)
            label.pack()
            
            spinbox_frame = ttk.Frame(frame)
            spinbox_frame.pack()
            
            entry = ttk.Entry(
                spinbox_frame, 
                width=width, 
                justify="center",
                validate="key",
                validatecommand=validate_cmd
            )
            entry.pack(side="left")
            
            def increment():
                try:
                    current = int(entry.get() or "0")
                    if current < max_val:
                        entry.delete(0, tk.END)
                        entry.insert(0, f"{current + 1:02d}")
                        validate_day_with_month()
                except ValueError:
                    pass
            
            def decrement():
                try:
                    current = int(entry.get() or "0")
                    if current > min_val:
                        entry.delete(0, tk.END)
                        entry.insert(0, f"{current - 1:02d}")
                        validate_day_with_month()
                except ValueError:
                    pass
            
            up_btn = ttk.Button(spinbox_frame, text="▲", width=2, command=increment)
            up_btn.pack(side="left", padx=2)
            
            down_btn = ttk.Button(spinbox_frame, text="▼", width=2, command=decrement)
            down_btn.pack(side="left", padx=2)
            
            return entry

        # 年
        self.year_entry = create_spinbox_frame(
            date_entries_frame, "年", 8, vcmd_year, 1900, 9999
        )
        self.year_entry.insert(0, datetime.now().strftime("%Y"))
        
        # 月
        self.month_entry = create_spinbox_frame(
            date_entries_frame, "月", 4, vcmd_month, 1, 12
        )
        self.month_entry.insert(0, datetime.now().strftime("%m"))
        
        # 日
        self.day_entry = create_spinbox_frame(
            date_entries_frame, "日", 4, vcmd_day, 1, 31
        )
        self.day_entry.insert(0, datetime.now().strftime("%d"))

        # 在輸入框失去焦點時也進行驗證
        self.year_entry.bind("<FocusOut>", lambda e: validate_day_with_month())
        self.month_entry.bind("<FocusOut>", lambda e: validate_day_with_month())
        self.day_entry.bind("<FocusOut>", lambda e: validate_day_with_month())

        # 時間輸入框架
        time_frame = ttk.LabelFrame(input_frame, text="時間", padding="10")
        time_frame.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # 時分秒輸入框水平排列
        time_entries_frame = ttk.Frame(time_frame)
        time_entries_frame.pack(fill="x", expand=True, padx=5)
        
        # 時
        hour_frame = ttk.Frame(time_entries_frame)
        hour_frame.pack(side="left", expand=True, padx=5)
        
        ttk.Label(hour_frame, text="時").pack()
        
        hour_spinbox_frame = ttk.Frame(hour_frame)
        hour_spinbox_frame.pack()
        
        self.hour_entry = ttk.Entry(
            hour_spinbox_frame,
            width=4,
            justify="center",
            validate="key",
            validatecommand=vcmd_hour
        )
        self.hour_entry.pack(side="left")
        self.hour_entry.insert(0, datetime.now().strftime("%H"))
        
        def increment_hour():
            try:
                current = int(self.hour_entry.get() or "0")
                if current < 23:
                    self.hour_entry.delete(0, tk.END)
                    self.hour_entry.insert(0, f"{current + 1:02d}")
            except ValueError:
                pass
        
        def decrement_hour():
            try:
                current = int(self.hour_entry.get() or "0")
                if current > 0:
                    self.hour_entry.delete(0, tk.END)
                    self.hour_entry.insert(0, f"{current - 1:02d}")
            except ValueError:
                pass
        
        hour_up_btn = ttk.Button(hour_spinbox_frame, text="▲", width=2, command=increment_hour)
        hour_up_btn.pack(side="left", padx=2)
        
        hour_down_btn = ttk.Button(hour_spinbox_frame, text="▼", width=2, command=decrement_hour)
        hour_down_btn.pack(side="left", padx=2)
        
        # 分
        minute_frame = ttk.Frame(time_entries_frame)
        minute_frame.pack(side="left", expand=True, padx=5)
        
        ttk.Label(minute_frame, text="分").pack()
        
        minute_spinbox_frame = ttk.Frame(minute_frame)
        minute_spinbox_frame.pack()
        
        self.minute_entry = ttk.Entry(
            minute_spinbox_frame,
            width=4,
            justify="center",
            validate="key",
            validatecommand=vcmd_minute_second
        )
        self.minute_entry.pack(side="left")
        self.minute_entry.insert(0, datetime.now().strftime("%M"))
        
        def increment_minute():
            try:
                current = int(self.minute_entry.get() or "0")
                if current < 59:
                    self.minute_entry.delete(0, tk.END)
                    self.minute_entry.insert(0, f"{current + 1:02d}")
            except ValueError:
                pass
        
        def decrement_minute():
            try:
                current = int(self.minute_entry.get() or "0")
                if current > 0:
                    self.minute_entry.delete(0, tk.END)
                    self.minute_entry.insert(0, f"{current - 1:02d}")
            except ValueError:
                pass
        
        minute_up_btn = ttk.Button(minute_spinbox_frame, text="▲", width=2, command=increment_minute)
        minute_up_btn.pack(side="left", padx=2)
        
        minute_down_btn = ttk.Button(minute_spinbox_frame, text="▼", width=2, command=decrement_minute)
        minute_down_btn.pack(side="left", padx=2)
        
        # 秒
        second_frame = ttk.Frame(time_entries_frame)
        second_frame.pack(side="left", expand=True, padx=5)
        
        ttk.Label(second_frame, text="秒").pack()
        
        second_spinbox_frame = ttk.Frame(second_frame)
        second_spinbox_frame.pack()
        
        self.second_entry = ttk.Entry(
            second_spinbox_frame,
            width=4,
            justify="center",
            validate="key",
            validatecommand=vcmd_minute_second
        )
        self.second_entry.pack(side="left")
        self.second_entry.insert(0, datetime.now().strftime("%S"))
        
        def increment_second():
            try:
                current = int(self.second_entry.get() or "0")
                if current < 59:
                    self.second_entry.delete(0, tk.END)
                    self.second_entry.insert(0, f"{current + 1:02d}")
            except ValueError:
                pass
        
        def decrement_second():
            try:
                current = int(self.second_entry.get() or "0")
                if current > 0:
                    self.second_entry.delete(0, tk.END)
                    self.second_entry.insert(0, f"{current - 1:02d}")
            except ValueError:
                pass
        
        second_up_btn = ttk.Button(second_spinbox_frame, text="▲", width=2, command=increment_second)
        second_up_btn.pack(side="left", padx=2)
        
        second_down_btn = ttk.Button(second_spinbox_frame, text="▼", width=2, command=decrement_second)
        second_down_btn.pack(side="left", padx=2)

    def create_preset_time_frame(self):
        # 如果已經存在快速設定框架，先清除它
        if hasattr(self, 'preset_frame'):
            self.preset_frame.destroy()
            
        preset_frame = ttk.LabelFrame(self.main_container, text="快速設定", padding="10")
        preset_frame.pack(fill="x", pady=(0, 10))  # 調整間距
        self.preset_frame = preset_frame
        
        # 創建內部框架以確保按鈕居中
        inner_frame = ttk.Frame(preset_frame)
        inner_frame.pack(expand=True)
        
        # 預設時間按鈕
        for preset in self.settings["preset_times"]:
            btn = ttk.Button(
                inner_frame,
                text=preset["name"],
                command=lambda t=preset["time"]: self.set_preset_time(t),
                width=12  # 調整按鈕寬度
            )
            btn.pack(side="left", padx=8)  # 調整按鈕間距
        
        # 添加管理按鈕
        manage_btn = ttk.Button(
            inner_frame,
            text="管理",
            command=self.manage_preset_times,
            width=8
        )
        manage_btn.pack(side="left", padx=8)  # 調整按鈕間距

    def manage_preset_times(self):
        """管理自定義時間設定"""
        # 創建管理視窗
        manage_window = tk.Toplevel()
        manage_window.title("管理快速設定時間")
        manage_window.geometry("400x500")
        manage_window.resizable(False, False)
        manage_window.transient(self.root)
        manage_window.grab_set()
        
        # 設置視窗置中
        window_width = 400
        window_height = 500
        screen_width = manage_window.winfo_screenwidth()
        screen_height = manage_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        manage_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 創建主框架
        main_frame = ttk.Frame(manage_window, padding="20")
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
        
        # 添加新時間框架
        add_frame = ttk.LabelFrame(scrollable_frame, text="新增快速設定", padding="10")
        add_frame.pack(fill="x", pady=(0, 10))
        
        # 名稱輸入
        name_frame = ttk.Frame(add_frame)
        name_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(name_frame, text="名稱：").pack(side="left")
        name_entry = ttk.Entry(name_frame)
        name_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # 時間輸入
        time_frame = ttk.Frame(add_frame)
        time_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(time_frame, text="時間：").pack(side="left")
        time_entry = ttk.Entry(time_frame)
        time_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # 添加按鈕
        def add_preset():
            name = name_entry.get().strip()
            time = time_entry.get().strip()
            
            if not name or not time:
                messagebox.showerror("錯誤", "請填寫名稱和時間！")
                return
                
            try:
                # 驗證時間格式
                datetime.strptime(time, "%H:%M:%S")
            except ValueError:
                messagebox.showerror("錯誤", "時間格式不正確！請使用 HH:MM:SS 格式")
                return
                
            # 添加到設定
            self.settings["preset_times"].append({
                "name": name,
                "time": time
            })
            save_settings(self.settings)
            
            # 清空輸入框
            name_entry.delete(0, tk.END)
            time_entry.delete(0, tk.END)
            
            # 重新整理列表
            refresh_list()
            
            # 重新整理主視窗的快速設定按鈕
            self.create_preset_time_frame()
            
        add_btn = ttk.Button(add_frame, text="添加", command=add_preset)
        add_btn.pack(pady=5)
        
        # 現有設定列表
        list_frame = ttk.LabelFrame(scrollable_frame, text="現有設定", padding="10")
        list_frame.pack(fill="x", pady=(10, 0))
        
        def refresh_list():
            # 清除現有項目
            for widget in list_frame.winfo_children():
                widget.destroy()
            
            # 添加新項目
            for i, preset in enumerate(self.settings["preset_times"]):
                item_frame = ttk.Frame(list_frame)
                item_frame.pack(fill="x", pady=2)
                
                ttk.Label(item_frame, text=f"{preset['name']} ({preset['time']})").pack(side="left")
                
                def delete_preset(index):
                    if messagebox.askyesno("確認", "確定要刪除這個設定嗎？"):
                        self.settings["preset_times"].pop(index)
                        save_settings(self.settings)
                        refresh_list()
                        # 重新整理主視窗的快速設定按鈕
                        self.create_preset_time_frame()
                
                delete_btn = ttk.Button(
                    item_frame,
                    text="刪除",
                    command=lambda idx=i: delete_preset(idx),
                    width=6
                )
                delete_btn.pack(side="right")
        
        # 初始整理列表
        refresh_list()
        
        # 設置滾動區域
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 綁定滑鼠滾輪
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 綁定關閉事件
        def on_closing():
            canvas.unbind_all("<MouseWheel>")
            manage_window.destroy()
            # 重新整理主視窗的快速設定按鈕
            self.create_preset_time_frame()
        
        manage_window.protocol("WM_DELETE_WINDOW", on_closing)

    def create_button_frame(self):
        button_frame = ttk.Frame(self.main_container)
        button_frame.pack(fill="x", pady=(0, 10))  # 調整間距
        self.button_frame = button_frame
        
        # 創建內部框架以確保按鈕居中
        inner_frame = ttk.Frame(button_frame)
        inner_frame.pack(expand=True)
        
        buttons = [
            ("更新系統時間", self.update_system_time),
            ("取得網路時間", self.get_network_time),
            ("刷新", self.refresh_input_time)
        ]
        
        for text, command in buttons:
            btn = ttk.Button(
                inner_frame,
                text=text,
                command=command,
                width=12  # 調整按鈕寬度
            )
            btn.pack(side="left", padx=8)  # 調整按鈕間距

    def create_version_frame(self):
        version_frame = ttk.Frame(self.main_container)
        version_frame.pack(fill="x", pady=(5, 0))  # 調整間距
        
        # 左側版本資訊
        version_info_frame = ttk.Frame(version_frame)
        version_info_frame.pack(side="left")
        
        # 添加點擊計數器
        self.version_click_count = 0
        self.last_click_time = 0
        
        def on_version_click(event):
            current_time = time.time()
            # 如果距離上次點擊超過3秒，重置計數器
            if current_time - self.last_click_time > 3:
                self.version_click_count = 0
            
            self.version_click_count += 1
            self.last_click_time = current_time
            
            # 如果點擊5次，顯示隱藏頁面
            if self.version_click_count >= 5:
                self.version_click_count = 0  # 重置計數器
                self.show_hidden_page()
        
        version_label = ttk.Label(
            version_info_frame,
            text=f"版本：{VERSION}",
            font=("微軟正黑體", 9),
            cursor="hand2"  # 改變滑鼠指標為手型
        )
        version_label.pack(side="left", padx=(0, 10))  # 調整間距
        version_label.bind("<Button-1>", on_version_click)  # 綁定點擊事件
        
        author_label = ttk.Label(
            version_info_frame,
            text=f"作者：{AUTHOR}",
            font=("微軟正黑體", 9)
        )
        author_label.pack(side="left", padx=(0, 10))  # 調整間距
        
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

    def show_hidden_page(self):
        """顯示隱藏頁面"""
        # 創建隱藏頁面視窗
        hidden_window = tk.Toplevel()
        hidden_window.title("進階功能")
        hidden_window.geometry("400x300")
        hidden_window.resizable(False, False)
        hidden_window.transient(self.root)  # 設置為主視窗的子視窗
        
        # 設置視窗置中
        window_width = 400
        window_height = 300
        screen_width = hidden_window.winfo_screenwidth()
        screen_height = hidden_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        hidden_window.geometry(f"{window_width}x{window_height}+{x}+{y}")  # 修正變數名稱
        
        # 創建主框架
        main_frame = ttk.Frame(hidden_window, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # 添加標題
        title_label = ttk.Label(
            main_frame,
            text="進階功能",
            font=("微軟正黑體", 14, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 添加功能按鈕
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill="x", expand=True)
        
        # 這裡可以添加更多進階功能按鈕
        ttk.Button(
            buttons_frame,
            text="功能 1",
            width=20,
            command=lambda: messagebox.showinfo("提示", "功能開發中...")
        ).pack(pady=5)
        
        ttk.Button(
            buttons_frame,
            text="功能 2",
            width=20,
            command=lambda: messagebox.showinfo("提示", "功能開發中...")
        ).pack(pady=5)
        
        ttk.Button(
            buttons_frame,
            text="功能 3",
            width=20,
            command=lambda: messagebox.showinfo("提示", "功能開發中...")
        ).pack(pady=5)
        
        # 添加關閉按鈕
        ttk.Button(
            main_frame,
            text="關閉",
            width=15,
            command=hidden_window.destroy
        ).pack(pady=(20, 0))

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
                
                # 記錄時間修改
                time_logger.info(f"系統時間已更新為網路時間：{ntp_time.strftime('%Y-%m-%d %H:%M:%S')}")
                
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
                    
                    # 記錄原始時間
                    original_time = datetime.now()
                    
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    
                    subprocess.run(f'date {date_str}', startupinfo=startupinfo, shell=True)
                    subprocess.run(f'time {time_str}', startupinfo=startupinfo, shell=True)
                    
                    # 記錄時間修改
                    new_time = f"{date_str} {time_str}"
                    time_logger.info(f"手動更新系統時間：{original_time.strftime('%Y-%m-%d %H:%M:%S')} -> {new_time}")
                    
                    self.update_current_time()
                    
                except ValueError as e:
                    messagebox.showerror("錯誤", f"請輸入正確的數字！\n{str(e)}")
                    system_logger.error(f"時間更新失敗：{str(e)}")
                    return
            
        except Exception as e:
            error_msg = f"無法更新系統時間：{str(e)}"
            messagebox.showerror("錯誤", error_msg)
            system_logger.error(error_msg)

    def get_network_time(self):
        if not HAS_NTPLIB:
            error_msg = "網路時間同步功能目前無法使用。請確保已安裝 ntplib 套件且電腦已連接到網際網路。"
            messagebox.showwarning("功能受限", error_msg)
            system_logger.warning(error_msg)
            return
            
        try:
            # 獲取已啟用的 NTP 伺服器
            enabled_servers = [server for server, enabled in self.settings["ntp_servers"].items() if enabled]
            
            if not enabled_servers:
                error_msg = "沒有啟用的 NTP 伺服器！請在 NTP 伺服器設定中至少啟用一個伺服器。"
                messagebox.showwarning("警告", error_msg)
                system_logger.warning(error_msg)
                return
            
            client = ntplib.NTPClient()
            last_error = None
            
            for server in enabled_servers:
                try:
                    # 記錄當前系統時間
                    local_time_before = datetime.now()
                    
                    response = client.request(server, timeout=2)
                    net_time = datetime.fromtimestamp(response.tx_time)
                    
                    # 計算時間偏差
                    local_time_after = datetime.now()
                    local_time = local_time_before + (local_time_after - local_time_before) / 2
                    time_drift = abs((net_time - local_time).total_seconds())
                    
                    # 記錄時間偏差
                    drift_logger.info(f"NTP伺服器：{server}")
                    drift_logger.info(f"本地時間：{local_time.strftime('%Y-%m-%d %H:%M:%S.%f')}")
                    drift_logger.info(f"網路時間：{net_time.strftime('%Y-%m-%d %H:%M:%S.%f')}")
                    drift_logger.info(f"時間偏差：{time_drift:.3f} 秒")
                    
                    # 如果時間偏差超過1秒，更新系統時間
                    if time_drift > 1:
                        self.update_system_time(net_time)
                        time_logger.info(f"系統時間已同步（來自 {server}），偏差：{time_drift:.3f} 秒")
                    else:
                        time_logger.info(f"系統時間無需同步（來自 {server}），偏差：{time_drift:.3f} 秒")
                    
                    return net_time
                    
                except (ntplib.NTPException, gaierror) as e:
                    last_error = e
                    system_logger.warning(f"NTP伺服器 {server} 連接失敗：{str(e)}")
                    continue
                except Exception as e:
                    last_error = e
                    system_logger.error(f"從 {server} 獲取時間時發生錯誤：{str(e)}")
                    continue
            
            if last_error:
                raise last_error
            
        except Exception as e:
            error_msg = f"無法取得網路時間：{str(e)}"
            system_logger.error(error_msg)
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
                "description": "即時顯示當前系統時間，包含日期、星期和時間。時間會自動更新，確保顯示準確。"
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
                "description": "從多個NTP伺服器獲取網路時間並同步。支援多個備用伺服器，確保高可用性。\n注意：需要網路連接才能使用此功能。"
            },
            {
                "title": "NTP伺服器設定",
                "description": "可以選擇和管理要使用的NTP伺服器，支援多個伺服器以提高可靠性。"
            },
            {
                "title": "自動更新",
                "description": "定期檢查新版本，發現更新時會提示用戶，確保使用最新版本。"
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
        ntp_window.transient(self.root)  # 設置為主視窗的子視窗
        ntp_window.grab_set()  # 模態視窗
        
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
        
        # 創建標題
        title_label = ttk.Label(
            scrollable_frame,
            text="選擇要使用的 NTP 伺服器：",
            font=("微軟正黑體", 11, "bold")
        )
        title_label.pack(anchor="w", pady=(0, 10))
        
        # 確保 ntp_servers 存在於設定中
        if "ntp_servers" not in self.settings:
            self.settings["ntp_servers"] = DEFAULT_SETTINGS["ntp_servers"].copy()
        
        # 創建複選框
        var_dict = {}
        for server in DEFAULT_SETTINGS["ntp_servers"].keys():
            is_enabled = self.settings["ntp_servers"].get(server, DEFAULT_SETTINGS["ntp_servers"][server])
            var = tk.BooleanVar(value=is_enabled)
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
        
        def save_ntp_settings():
            # 保存設置
            selected_servers = {server: var.get() for server, var in var_dict.items()}
            if not any(selected_servers.values()):
                messagebox.showwarning("警告", "請至少選擇一個 NTP 伺服器！")
                return
                
            # 更新設定
            self.settings["ntp_servers"] = selected_servers
            if save_settings(self.settings):
                messagebox.showinfo("成功", "NTP 伺服器設定已保存！")
                ntp_window.destroy()
        
        # 添加保存按鈕
        save_button = ttk.Button(
            button_frame,
            text="保存設置",
            command=save_ntp_settings,
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

    def show_log_viewer(self, log_type):
        """顯示日誌查看器"""
        # 設定日誌檔案和標題
        if log_type == "system":
            log_file = SYSTEM_LOG_FILE
            title = "系統日誌"
        elif log_type == "time":
            log_file = TIME_CHANGES_LOG_FILE
            title = "時間修改記錄"
        else:  # drift
            log_file = DRIFT_LOG_FILE
            title = "時間偏差記錄"
        
        # 創建日誌視窗
        log_window = tk.Toplevel()
        log_window.title(title)
        log_window.geometry("800x600")
        log_window.transient(self.root)
        
        # 創建主框架
        main_frame = ttk.Frame(log_window, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 創建工具列
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill="x", pady=(0, 10))
        
        # 刷新按鈕
        refresh_btn = ttk.Button(
            toolbar,
            text="刷新",
            command=lambda: load_log_content()
        )
        refresh_btn.pack(side="left", padx=5)
        
        # 清除日誌按鈕
        def clear_log():
            if messagebox.askyesno("確認", "確定要清除所有日誌嗎？"):
                try:
                    with open(log_file, 'w', encoding='utf-8') as f:
                        f.write('')
                    load_log_content()
                except Exception as e:
                    messagebox.showerror("錯誤", f"清除日誌失敗：{str(e)}")
        
        clear_btn = ttk.Button(
            toolbar,
            text="清除日誌",
            command=clear_log
        )
        clear_btn.pack(side="left", padx=5)
        
        # 匯出按鈕
        def export_log():
            try:
                from datetime import datetime
                export_file = os.path.join(
                    os.path.expanduser("~"),
                    "Desktop",
                    f"{title}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                )
                with open(log_file, 'r', encoding='utf-8') as source:
                    with open(export_file, 'w', encoding='utf-8') as target:
                        target.write(source.read())
                messagebox.showinfo("成功", f"日誌已匯出至：\n{export_file}")
            except Exception as e:
                messagebox.showerror("錯誤", f"匯出日誌失敗：{str(e)}")
        
        export_btn = ttk.Button(
            toolbar,
            text="匯出",
            command=export_log
        )
        export_btn.pack(side="left", padx=5)
        
        # 搜尋框
        search_frame = ttk.Frame(toolbar)
        search_frame.pack(side="right", padx=5)
        
        ttk.Label(search_frame, text="搜尋：").pack(side="left")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side="left", padx=5)
        
        def search_log():
            search_text = search_var.get().lower()
            log_text.tag_remove("search", "1.0", "end")
            if search_text:
                idx = "1.0"
                while True:
                    idx = log_text.search(search_text, idx, "end", nocase=True)
                    if not idx:
                        break
                    end_idx = f"{idx}+{len(search_text)}c"
                    log_text.tag_add("search", idx, end_idx)
                    idx = end_idx
        
        search_var.trace("w", lambda *args: search_log())
        
        # 創建文本框和滾動條
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")
        
        log_text = tk.Text(
            text_frame,
            wrap="none",
            yscrollcommand=scrollbar.set,
            font=("微軟正黑體", 12)
        )
        log_text.pack(fill="both", expand=True)
        
        scrollbar.config(command=log_text.yview)
        
        # 配置搜尋高亮
        log_text.tag_configure("search", background="yellow", font=("微軟正黑體", 12))
        
        def load_log_content():
            log_text.delete("1.0", "end")
            try:
                if os.path.exists(log_file):
                    with open(log_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                        log_text.insert("1.0", content)
                else:
                    log_text.insert("1.0", "尚無日誌記錄")
            except Exception as e:
                log_text.insert("1.0", f"讀取日誌失敗：{str(e)}")
        
        # 載入日誌內容
        load_log_content()
        
        # 設置視窗置中
        log_window.update_idletasks()
        width = log_window.winfo_width()
        height = log_window.winfo_height()
        x = (log_window.winfo_screenwidth() // 2) - (width // 2)
        y = (log_window.winfo_screenheight() // 2) - (height // 2)
        log_window.geometry(f"{width}x{height}+{x}+{y}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = TimeChangerApp(root)
        root.mainloop()
    except Exception as e:
        error_msg = "程式發生錯誤：\n\n"
        error_msg += str(e)
        error_msg += "\n\n詳細錯誤資訊：\n"
        error_msg += traceback.format_exc()
        error_msg += "\n\n請將此錯誤資訊回報給開發者。"
        
        messagebox.showerror("錯誤", error_msg)
        
        # 寫入錯誤日誌
        try:
            log_dir = os.path.join(CONFIG_DIR, "logs")
            os.makedirs(log_dir, exist_ok=True)
            log_file = os.path.join(log_dir, f"error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
            
            with open(log_file, "w", encoding="utf-8") as f:
                f.write(f"版本：{VERSION}\n")
                f.write(f"時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"系統：{sys.platform}\n")
                f.write("\n錯誤訊息：\n")
                f.write(str(e))
                f.write("\n\n詳細錯誤資訊：\n")
                f.write(traceback.format_exc())
        except:
            pass  # 如果無法寫入日誌，就略過 