import tkinter as tk
from tkinter import messagebox, simpledialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import winreg
import ctypes
from ctypes import windll
import os
import subprocess
import json

DEFAULT_PAC_URL = "http://127.0.0.1:8123/proxy.pac"
CONFIG_FILE = "config.json"
PROCESS_NAME = "PrincessConnectReDive"
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    ctypes.windll.user32.SetProcessDPIAware()


class SystemControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("公主连结小工具")
        self.root.geometry("420x480")
        self.root.resizable(False, False)
        self.pac_url = self.load_config()
        container = ttk.Frame(root, padding=15)
        container.pack(fill=BOTH, expand=YES)
        self.frame_proxy = ttk.Labelframe(
            container, text="acgp代理DMM网页", padding=10, bootstyle="primary"
        )
        self.frame_proxy.pack(fill=X, pady=5)
        status_frame = ttk.Frame(self.frame_proxy)
        status_frame.pack(fill=X, pady=(0, 5))
        self.proxy_status_var = tk.StringVar(value="检测中...")
        self.lbl_proxy_status = ttk.Label(
            status_frame, textvariable=self.proxy_status_var, font=("微软雅黑", 9)
        )
        self.lbl_proxy_status.pack(side=LEFT)
        btn_frame = ttk.Frame(self.frame_proxy)
        btn_frame.pack(fill=X)
        ttk.Button(
            btn_frame,
            text="⚙️ 设置地址",
            command=self.change_pac_url,
            width=10,
            bootstyle="secondary-outline",
        ).pack(side=LEFT, padx=(0, 5))
        self.btn_proxy_toggle = ttk.Button(
            btn_frame,
            text="切换状态",
            command=self.toggle_proxy,
            width=10,
            bootstyle="primary",
        )
        self.btn_proxy_toggle.pack(side=RIGHT)
        self.lbl_current_pac = ttk.Label(
            self.frame_proxy,
            text=f"当前默认: {self.pac_url}",
            font=("Consolas", 10),
            bootstyle="secondary",
        )
        self.lbl_current_pac.pack(anchor="w", pady=(5, 0))
        self.frame_audio = ttk.Labelframe(
            container, text="音效设置", padding=10, bootstyle="info"
        )
        self.frame_audio.pack(fill=X, pady=10)
        ttk.Button(
            self.frame_audio,
            text="↗ 打开声音设置页 ",
            command=self.open_sound_settings,
            width=30,
            bootstyle="info-outline",
        ).pack()
        ttk.Label(
            self.frame_audio,
            text="手动打开或关闭空间音效",
            font=("微软雅黑", 8),
            bootstyle="secondary",
        ).pack(pady=(5, 0))
        self.frame_window = ttk.Labelframe(
            container, text="游戏窗口控制", padding=10, bootstyle="warning"
        )
        self.frame_window.pack(fill=X, pady=5)
        ttk.Label(
            self.frame_window,
            text="⚠️ 必须以管理员身份运行本工具",
            font=("微软雅黑", 9, "bold"),
            bootstyle="danger",
        ).pack(anchor="w", pady=(0, 10))
        win_btn_frame = ttk.Frame(self.frame_window)
        win_btn_frame.pack(fill=X)
        ttk.Button(
            win_btn_frame,
            text="⛶ 强制全屏",
            command=self.force_fullscreen,
            width=12,
            bootstyle="warning",
        ).pack(side=LEFT, padx=5, expand=YES, fill=X)
        ttk.Button(
            win_btn_frame,
            text="↩ 恢复窗口",
            command=self.restore_window,
            width=12,
            bootstyle="secondary",
        ).pack(side=RIGHT, padx=5, expand=YES, fill=X)
        self.setup_theme_switcher(container)
        self.check_proxy_status()

    def load_config(self):
        """读取本地配置文件，没有则用默认"""
        config_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), CONFIG_FILE
        )
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    return data.get("pac_url", DEFAULT_PAC_URL)
            except:
                pass
        return DEFAULT_PAC_URL

    def save_config(self, url):
        """保存配置到本地 json"""
        config_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), CONFIG_FILE
        )
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump({"pac_url": url}, f, indent=4)
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存配置: {e}")

    def change_pac_url(self):
        """弹出输入框让用户修改地址"""
        new_url = simpledialog.askstring(
            "设置代理地址",
            "请输入 PAC 地址:",
            initialvalue=self.pac_url,
            parent=self.root,
        )
        if new_url:
            self.pac_url = new_url.strip()
            self.save_config(self.pac_url)
            self.lbl_current_pac.configure(text=f"当前: {self.pac_url}")
            messagebox.showinfo("成功", "代理地址已更新并保存！\n(下次开启代理时生效)")

    def setup_theme_switcher(self, parent):
        footer = ttk.Frame(parent)
        footer.pack(fill=X, side=BOTTOM, pady=10)
        ttk.Label(
            footer, text="🎨 主题切换:", bootstyle="secondary", font=("微软雅黑", 9)
        ).pack(side=LEFT, padx=(0, 5))
        theme_names = self.root.style.theme_names()
        self.theme_cbo = ttk.Combobox(
            footer, values=theme_names, state="readonly", width=15
        )
        self.theme_cbo.pack(side=LEFT)
        current_theme = self.root.style.theme.name
        self.theme_cbo.set(current_theme)
        self.theme_cbo.bind(
            "<<ComboboxSelected>>",
            lambda e: self.root.style.theme_use(self.theme_cbo.get()),
        )

    def find_window_by_process(self):
        try:
            cmd = f'powershell "(Get-Process {PROCESS_NAME} -ErrorAction SilentlyContinue).MainWindowHandle"'
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                startupinfo=startupinfo,
            )
            out, err = process.communicate()
            output = out.decode().strip()
            if not output or output == "0":
                messagebox.showwarning(
                    "未找到",
                    f"未找到进程: {PROCESS_NAME}.exe\n请先启动游戏，并确保已管理员身份运行。",
                )
                return None
            hwnd = int(output.splitlines()[0])
            if hwnd <= 0:
                return None
            return hwnd
        except Exception as e:
            messagebox.showerror("系统错误", str(e))
            return None

    def force_fullscreen(self):
        hwnd = self.find_window_by_process()
        if not hwnd:
            return
        user32 = windll.user32
        screen_w = user32.GetSystemMetrics(0)
        screen_h = user32.GetSystemMetrics(1)
        old_style = user32.GetWindowLongW(hwnd, -16)
        new_style = old_style & ~0x00C00000 & ~0x00040000
        user32.SetWindowLongW(hwnd, -16, new_style)
        ret = user32.SetWindowPos(hwnd, 0, 0, 0, screen_w, screen_h, 0x0020 | 0x0040)
        if ret == 0:
            messagebox.showerror("权限不足", "修改失败！\n请务必以管理员身份运行。")

    def restore_window(self):
        hwnd = self.find_window_by_process()
        if not hwnd:
            return
        user32 = windll.user32
        old_style = user32.GetWindowLongW(hwnd, -16)
        new_style = old_style | 0x00C00000 | 0x00040000
        user32.SetWindowLongW(hwnd, -16, new_style)
        user32.SetWindowPos(hwnd, 0, 100, 100, 1296, 759, 0x0020 | 0x0040)

    def open_sound_settings(self):
        os.system("start ms-settings:sound")

    def check_proxy_status(self):
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
                0,
                winreg.KEY_READ,
            )
            try:
                winreg.QueryValueEx(key, "AutoConfigURL")
                self.is_proxy_on = True
                self.proxy_status_var.set("✅ 代理已开启")
                self.lbl_proxy_status.configure(bootstyle="success")
            except FileNotFoundError:
                self.is_proxy_on = False
                self.proxy_status_var.set("⬜ 代理已关闭")
                self.lbl_proxy_status.configure(bootstyle="default")
            finally:
                winreg.CloseKey(key)
        except Exception:
            pass

    def toggle_proxy(self):
        try:
            REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_WRITE
            )
            target_on = not self.is_proxy_on
            if target_on:
                winreg.SetValueEx(key, "AutoConfigURL", 0, winreg.REG_SZ, self.pac_url)
                winreg.SetValueEx(key, "ProxyEnable", 0, winreg.REG_DWORD, 0)
            else:
                try:
                    winreg.DeleteValue(key, "AutoConfigURL")
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
            try:
                ctypes.windll.Wininet.InternetSetOptionW(0, 39, 0, 0)
                ctypes.windll.Wininet.InternetSetOptionW(0, 37, 0, 0)
            except:
                pass
            self.check_proxy_status()
        except Exception as e:
            messagebox.showerror("错误", str(e))


if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = SystemControlApp(root)
    root.mainloop()
