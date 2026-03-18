# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器
作者：laozi4vip
GitHub：https://github.com/laozi4vip/QuickLauncher
"""

__author__ = "laozi4vip"
__github__ = "https://github.com/laozi4vip/QuickLauncher"
__app_name__ = "QuickLauncher"
__version__ = "2.1"
__description__ = "Windows 任务栏快捷启动器"

import wx
import wx.adv
import os
import json
import subprocess
import psutil
import ctypes
import sys
import winreg
import shlex
import re
import time

# ---------------------------
# Windows API & 常量
# ---------------------------
user32 = ctypes.windll.user32
SW_MINIMIZE = 6
SW_RESTORE = 9

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008

GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080

RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_RUN_NAME = "QuickLauncher"

BROWSER_SET = {"chrome", "msedge", "chromium", "brave", "firefox"}

# ---------------------------
# 配置路径
# ---------------------------
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

LAST_ACTIVE_PROFILE_HWND = {}

# ★ 浏览器主进程映射缓存
_browser_main_map_cache = {}
_browser_main_map_ts = 0.0


# ---------------------------
# 配置 / 自启
# ---------------------------
def get_launch_command():
    if getattr(sys, 'frozen', False):
        return f'"{sys.executable}"'
    return f'"{sys.executable}" "{os.path.abspath(__file__)}"'


def is_autostart_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_READ) as key:
            val, _ = winreg.QueryValueEx(key, APP_RUN_NAME)
            return bool(val)
    except Exception:
        return False


def set_autostart(enable: bool):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_SET_VALUE) as key:
            if enable:
                winreg.SetValueEx(key, APP_RUN_NAME, 0, winreg.REG_SZ, get_launch_command())
            else:
                try:
                    winreg.DeleteValue(key, APP_RUN_NAME)
                except FileNotFoundError:
                    pass
        return True
    except Exception as e:
        print("set_autostart error:", e)
        return False


def load_config():
    default_data = {"programs": [], "autostart": is_autostart_enabled()}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            programs = data.get("programs", [])
            autostart = data.get("autostart", is_autostart_enabled())
            for p in programs:
                p.setdefault("name", "")
                p.setdefault("path", "")
                p.setdefault("args", "")
                p.setdefault("hotkey", "")
                p.setdefault("window_keyword", "")
                p.setdefault("match_mode", "title")
                p.setdefault("bind_hwnd", 0)
                p.setdefault("profile_name", "")
                p.setdefault("title_sig", "")
            return {"programs": programs, "autostart": autostart}
        except Exception as e:
            print("load_config error:", e)
    return default_data


def save_config(programs, autostart):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"programs": programs, "autostart": autostart}, f, ensure_ascii=False, indent=4)


# ---------------------------
# 热键工具
# ---------------------------
def normalize_hotkey(hotkey: str) -> str:
    return (hotkey or "").strip().lower().replace(" ", "")


def hotkey_to_mod_vk(hotkey: str):
    hotkey = normalize_hotkey(hotkey)
    if not hotkey:
        raise ValueError("快捷键不能为空")
    parts = hotkey.split("+")
    if any(p == "" for p in parts):
        raise ValueError("快捷键格式错误")

    key_map = {str(i): 0x30 + i for i in range(10)}
    for i in range(26):
        key_map[chr(ord("a") + i)] = 0x41 + i
    for i in range(1, 13):
        key_map[f"f{i}"] = 0x70 + (i - 1)

    mods = 0
    main_key = None
    for p in parts:
        if p == "ctrl":
            mods |= MOD_CONTROL
        elif p == "alt":
            mods |= MOD_ALT
        elif p == "shift":
            mods |= MOD_SHIFT
        elif p == "win":
            mods |= MOD_WIN
        else:
            if main_key is not None:
                raise ValueError("只能有一个主键")
            if p not in key_map:
                raise ValueError("不支持的主键")
            main_key = p

    if main_key is None:
        raise ValueError("缺少主键")
    if mods == 0 and not main_key.startswith("f"):
        raise ValueError("无修饰键仅支持 F1-F12")

    ordered = []
    if mods & MOD_CONTROL:
        ordered.append("ctrl")
    if mods & MOD_SHIFT:
        ordered.append("shift")
    if mods & MOD_ALT:
        ordered.append("alt")
    if mods & MOD_WIN:
        ordered.append("win")
    normalized = "+".join(ordered + [main_key]) if ordered else main_key
    return mods, key_map[main_key], normalized


def wx_event_to_hotkey(event: wx.KeyEvent) -> str:
    key = event.GetKeyCode()
    main_key = None
    if wx.WXK_F1 <= key <= wx.WXK_F12:
        main_key = f"f{key - wx.WXK_F1 + 1}"
    elif ord('0') <= key <= ord('9'):
        main_key = chr(key).lower()
    elif ord('A') <= key <= ord('Z'):
        main_key = chr(key).lower()
    elif ord('a') <= key <= ord('z'):
        main_key = chr(key).lower()
    else:
        return ""
    mods = []
    if event.ControlDown():
        mods.append("ctrl")
    if event.ShiftDown():
        mods.append("shift")
    if event.AltDown():
        mods.append("alt")
    if event.MetaDown():
        mods.append("win")
    if not mods and not main_key.startswith("f"):
        return ""
    return "+".join(mods + [main_key]) if mods else main_key


# ---------------------------
# 窗口 / 进程工具
# ---------------------------
def get_window_title(hwnd):
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value or ""


def get_pid_from_hwnd(hwnd):
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def is_alt_tab_window(hwnd):
    if not user32.IsWindowVisible(hwnd):
        return False
    title = get_window_title(hwnd)
    if not title.strip():
        return False
    exstyle = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    if exstyle & WS_EX_TOOLWINDOW:
        return False
    return True


def is_hwnd_valid(hwnd: int):
    try:
        return bool(hwnd) and bool(user32.IsWindow(hwnd)) and bool(user32.IsWindowVisible(hwnd))
    except Exception:
        return False


def get_proc_path_name_cmdline(pid):
    try:
        p = psutil.Process(pid)
        path = p.exe() or ""
        name = p.name() or ""
        try:
            cmdline_list = p.cmdline()
        except Exception:
            cmdline_list = []
        return path, name, cmdline_list
    except Exception:
        return "", "", []


def parse_profile_from_cmdline(proc_name: str, cmdline_list):
    name = (proc_name or "").lower().replace(".exe", "")
    args = cmdline_list or []

    def value_after(prefix):
        for i, a in enumerate(args):
            la = a.lower()
            if la.startswith(prefix + "="):
                return a.split("=", 1)[1].strip().strip('"')
            if la == prefix and i + 1 < len(args):
                return (args[i + 1] or "").strip().strip('"')
        return ""

    if name in ("chrome", "msedge", "chromium", "brave"):
        v = value_after("--profile-directory")
        if v:
            return v
        u = value_after("--user-data-dir")
        if u:
            return os.path.basename(u.rstrip("\\/"))
        return ""

    if name == "firefox":
        for i, a in enumerate(args):
            la = a.lower()
            if la in ("-p", "--profile") and i + 1 < len(args):
                return (args[i + 1] or "").strip().strip('"')
            if la == "-profile" and i + 1 < len(args):
                p = (args[i + 1] or "").strip().strip('"')
                return os.path.basename(p.rstrip("\\/"))
    return ""


def parse_profile_from_title(proc_name: str, title: str):
    name = (proc_name or "").lower().replace(".exe", "")
    t = (title or "").strip()
    if not t:
        return ""
    m = re.search(r"\b(Profile\s*\d+|Default|Personal|Work|Guest)\b", t, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    if name == "firefox":
        m2 = re.search(r"\b(Profile\s*\d+|Default)\b", t, re.IGNORECASE)
        if m2:
            return m2.group(1).strip()
    return ""


def guess_profile(proc_name: str, cmdline_list, title: str):
    p1 = parse_profile_from_cmdline(proc_name, cmdline_list)
    if p1:
        return p1
    return parse_profile_from_title(proc_name, title)


def make_title_signature(title: str):
    t = (title or "").strip().lower()
    if not t:
        return ""
    t = re.sub(r"\s+", " ", t)
    parts = [x.strip() for x in t.split("-") if x.strip()]
    if len(parts) >= 2:
        return " - ".join(parts[-2:])[:120]
    return t[:120]


# ---------------------------
# ★ 判断程序是否属于浏览器
# ---------------------------
def is_browser_program(program):
    """根据程序路径判断是否属于浏览器"""
    path = program.get("path", "") or ""
    exe_name = os.path.basename(path).lower().replace(".exe", "")
    return exe_name in BROWSER_SET


# ---------------------------
# ★ 浏览器主进程扫描 & 进程树回溯
# ---------------------------
def build_browser_main_proc_map():
    """
    扫描所有浏览器进程，找出主进程（命令行不含 --type= 的）。
    返回 {pid: profile_directory}
    无显式 --profile-directory 的主进程视为 "Default"。
    """
    main_map = {}
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            pname = (proc.name() or "").lower().replace(".exe", "")
            if pname not in BROWSER_SET:
                continue
            try:
                cmdline = proc.cmdline()
            except (psutil.AccessDenied, psutil.ZombieProcess):
                continue
            # 主浏览器进程不含 --type= 参数
            if any(a.lower().startswith('--type=') for a in cmdline):
                continue
            profile = parse_profile_from_cmdline(proc.name(), cmdline)
            if not profile:
                profile = "Default"
            main_map[proc.pid] = profile
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return main_map


def get_cached_browser_main_map(max_age=3.0):
    """带缓存的浏览器主进程映射（避免高频调用时重复扫描）"""
    global _browser_main_map_cache, _browser_main_map_ts
    now = time.time()
    if now - _browser_main_map_ts < max_age and _browser_main_map_cache:
        return _browser_main_map_cache
    _browser_main_map_cache = build_browser_main_proc_map()
    _browser_main_map_ts = now
    return _browser_main_map_cache


def get_profile_by_pid_tree(pid, main_map):
    """
    从窗口的 PID 向上遍历进程树，找到它属于哪个主浏览器进程，返回 profile。
    """
    if pid in main_map:
        return main_map[pid]
    try:
        visited = {pid}
        current = psutil.Process(pid)
        for _ in range(15):
            try:
                parent = current.parent()
                if parent is None or parent.pid == 0 or parent.pid in visited:
                    break
                visited.add(parent.pid)
                if parent.pid in main_map:
                    return main_map[parent.pid]
                current = parent
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                break
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass
    return ""


# ---------------------------
# 窗口枚举（★ 使用进程树回溯检测 Profile）
# ---------------------------
def enum_visible_app_windows():
    windows = []

    # ★ 预先构建浏览器主进程映射
    browser_main_map = get_cached_browser_main_map(max_age=2.0)

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def callback(hwnd, _):
        try:
            if not user32.IsWindow(hwnd):
                return True
            if not is_alt_tab_window(hwnd):
                return True

            pid = get_pid_from_hwnd(hwnd)
            title = get_window_title(hwnd).strip()
            path, proc_name, cmdline = get_proc_path_name_cmdline(pid)

            pn = (proc_name or "").lower().replace(".exe", "")
            if pn in BROWSER_SET:
                # ★ 优先通过进程树回溯获取 profile
                profile = get_profile_by_pid_tree(pid, browser_main_map)
                if not profile:
                    profile = guess_profile(proc_name, cmdline, title)
            else:
                profile = guess_profile(proc_name, cmdline, title)

            windows.append({
                "hwnd": int(hwnd),
                "pid": int(pid),
                "title": title,
                "title_sig": make_title_signature(title),
                "path": path,
                "proc_name": proc_name,
                "cmdline": cmdline,
                "profile": profile
            })
        except Exception:
            pass
        return True

    user32.EnumWindows(callback, 0)
    return windows


def enum_windows_for_program(path):
    target = os.path.normcase(path or "")
    result = []
    all_ws = enum_visible_app_windows()
    for w in all_ws:
        wpath = os.path.normcase(w.get("path", "") or "")
        if wpath and target and wpath == target:
            result.append(w)
    if not result and path:
        exe_name = os.path.basename(path).lower().replace(".exe", "")
        for w in all_ws:
            pn = (w.get("proc_name", "") or "").lower().replace(".exe", "")
            if pn == exe_name:
                result.append(w)
    return result


def update_last_active_cache():
    hwnd = user32.GetForegroundWindow()
    if not is_hwnd_valid(hwnd):
        return
    try:
        pid = get_pid_from_hwnd(hwnd)
        path, proc_name, cmdline = get_proc_path_name_cmdline(pid)
        title = get_window_title(hwnd).strip()

        exe = (os.path.basename(path) if path else proc_name or "").lower().replace(".exe", "")
        if exe in BROWSER_SET:
            # ★ 用进程树回溯获取 profile
            browser_main_map = get_cached_browser_main_map(max_age=5.0)
            profile = get_profile_by_pid_tree(pid, browser_main_map)
            if not profile:
                profile = guess_profile(proc_name, cmdline, title)
            if profile:
                LAST_ACTIVE_PROFILE_HWND[(exe, profile.strip().lower())] = (int(hwnd), time.time())
    except Exception:
        pass


def build_profile_args(proc_name: str, profile: str):
    pn = (proc_name or "").lower().replace(".exe", "")
    pf = (profile or "").strip()
    if not pf:
        return ""
    if pn in ("chrome", "msedge", "chromium", "brave"):
        return f'--profile-directory="{pf}"'
    if pn == "firefox":
        return f'-P "{pf}"'
    return ""


def score_window_for_program(program, w):
    score = 0
    keyword = (program.get("window_keyword", "") or "").strip().lower()
    profile_name = (program.get("profile_name", "") or "").strip().lower()
    title_sig = (program.get("title_sig", "") or "").strip().lower()
    match_mode = (program.get("match_mode", "title") or "title").strip().lower()

    w_title = (w.get("title", "") or "").lower()
    w_sig = (w.get("title_sig", "") or "").lower()
    w_prof = (w.get("profile", "") or "").strip().lower()

    if keyword and w_prof:
        if w_prof == keyword:
            score += 120
        elif keyword in w_prof:
            score += 90

    if profile_name and w_prof:
        if w_prof == profile_name:
            score += 100
        elif profile_name in w_prof:
            score += 70

    if title_sig and w_sig:
        if w_sig == title_sig:
            score += 60
        elif title_sig in w_sig or w_sig in title_sig:
            score += 35

    if match_mode == "title" and keyword and keyword in w_title:
        score += 80

    hwnd = int(w.get("hwnd", 0) or 0)
    if hwnd == user32.GetForegroundWindow():
        score += 20
    if hwnd and not user32.IsIconic(hwnd):
        score += 10

    return score


# ---------------------------
# 核心查找（含 HWND 自动重绑）
# ---------------------------
def find_window_for_program(program):
    path = program.get("path", "")
    if not path:
        return None

    match_mode = (program.get("match_mode", "title") or "title").strip().lower()
    keyword = (program.get("window_keyword", "") or "").strip().lower()
    bind_hwnd = int(program.get("bind_hwnd", 0) or 0)
    title_sig = (program.get("title_sig", "") or "").strip().lower()
    profile_name = (program.get("profile_name", "") or "").strip().lower()

    if match_mode == "hwnd" and bind_hwnd and is_hwnd_valid(bind_hwnd):
        pid = get_pid_from_hwnd(bind_hwnd)
        ppath, _, _ = get_proc_path_name_cmdline(pid)
        if os.path.normcase(ppath or "") == os.path.normcase(path or ""):
            return bind_hwnd

    candidates = enum_windows_for_program(path)
    if not candidates:
        return None

    scored = []
    for w in candidates:
        s = score_window_for_program(program, w)
        scored.append((s, int(w["hwnd"]), w))
    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best_hwnd, _ = scored[0]

    # ★ 浏览器类：必须有正向匹配分数，否则不返回任何窗口（不兜底）
    if is_browser_program(program):
        if best_score <= 0:
            return None
        return best_hwnd

    # ★ 非浏览器类：保留原有逻辑，有关键词/profile/title_sig 但分数<=0 则不匹配
    if (keyword or profile_name or title_sig) and best_score <= 0:
        return None
    return best_hwnd


def toggle_program(program):
    """
    ★ 修改版：
    - 浏览器类：找不到匹配窗口时不启动，直接返回 (None, False)
    - 非浏览器类：找不到匹配窗口时按路径启动程序
    """
    path = program.get("path", "")
    args = program.get("args", "")
    if not path:
        return None, False

    hwnd = find_window_for_program(program)
    if hwnd:
        fg = user32.GetForegroundWindow()
        if hwnd == fg:
            user32.ShowWindow(hwnd, SW_MINIMIZE)
        else:
            if user32.IsIconic(hwnd):
                user32.ShowWindow(hwnd, SW_RESTORE)
            user32.SetForegroundWindow(hwnd)
        return hwnd, True
    else:
        # ★ 浏览器类：没匹配上就不做任何事
        if is_browser_program(program):
            return None, False

        # ★ 非浏览器类：按路径启动
        if not os.path.exists(path):
            wx.MessageBox(f"程序不存在：\n{path}", "错误", wx.OK | wx.ICON_ERROR)
            return None, False
        cmd = [path]
        if args.strip():
            try:
                cmd.extend(shlex.split(args.strip(), posix=False))
            except Exception:
                cmd.extend(args.strip().split())
        subprocess.Popen(cmd)
        return None, False


def ask_profile_input(parent, default_profile=""):
    dlg = wx.TextEntryDialog(
        parent,
        "请输入浏览器 Profile（如 Profile 1 / Default / Work）",
        "确认 Profile",
        value=default_profile or ""
    )
    ret = dlg.ShowModal()
    val = dlg.GetValue().strip() if ret == wx.ID_OK else ""
    dlg.Destroy()
    return ret == wx.ID_OK, val


# ---------------------------
# 交互对话
# ---------------------------
class HotkeyCaptureDialog(wx.Dialog):
    def __init__(self, parent, current_hotkey=""):
        super().__init__(parent, title="设置快捷键", size=(460, 220))
        self.captured_hotkey = normalize_hotkey(current_hotkey)
        panel = wx.Panel(self)
        s = wx.BoxSizer(wx.VERTICAL)
        s.Add(wx.StaticText(panel, label="直接按组合键（Esc取消，Backspace清空）"), 0, wx.ALL, 10)
        self.show = wx.TextCtrl(panel, value=self.captured_hotkey, style=wx.TE_READONLY)
        s.Add(self.show, 0, wx.ALL | wx.EXPAND, 10)
        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(wx.Button(panel, wx.ID_OK, "确定"), 0, wx.ALL, 5)
        row.Add(wx.Button(panel, wx.ID_CANCEL, "取消"), 0, wx.ALL, 5)
        clear = wx.Button(panel, wx.ID_ANY, "清空")
        row.Add(clear, 0, wx.ALL, 5)
        s.Add(row, 0, wx.ALIGN_CENTER)
        panel.SetSizer(s)
        self.Bind(wx.EVT_CHAR_HOOK, self.on_char)
        clear.Bind(wx.EVT_BUTTON, self.on_clear)

    def on_clear(self, _):
        self.captured_hotkey = ""
        self.show.SetValue("")

    def on_char(self, event):
        key = event.GetKeyCode()
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CANCEL)
            return
        if key in (wx.WXK_BACK, wx.WXK_DELETE):
            self.on_clear(None)
            return
        hk = wx_event_to_hotkey(event)
        if not hk:
            wx.Bell()
            return
        try:
            _, _, n = hotkey_to_mod_vk(hk)
            self.captured_hotkey = n
            self.show.SetValue(n)
        except Exception:
            wx.Bell()


def get_icon_path():
    """获取图标路径，优先从exe同目录查找"""
    # 打包后从exe所在目录查找
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = BASE_DIR
    
    # 优先查找 ICO，其次 PNG
    for ext in ['icon.ico', 'icon.png']:
        icon_path = os.path.join(base, ext)
        if os.path.exists(icon_path):
            return icon_path
        # 备选：从当前目录查找
        alt_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ext)
        if os.path.exists(alt_path):
            return alt_path
    return None

class QuickLauncherTaskBar(wx.adv.TaskBarIcon):
    def __init__(self, frame):
        super().__init__()
        self.frame = frame
        # 加载图标文件
        icon_path = get_icon_path()
        if icon_path and os.path.exists(icon_path):
            if icon_path.endswith('.png'):
                bmp = wx.Bitmap(icon_path, wx.BITMAP_TYPE_PNG)
                icon = wx.Icon()
                icon.CopyFromBitmap(bmp)
            else:
                icon = wx.Icon(icon_path, wx.BITMAP_TYPE_ICO)
        else:
            bmp = wx.ArtProvider.GetBitmap(wx.ART_EXECUTABLE_FILE, wx.ART_OTHER, (16, 16))
            icon = wx.Icon()
            icon.CopyFromBitmap(bmp)
        self.SetIcon(icon, f"{__app_name__} v{__version__}")
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK, lambda e: self.frame.show_from_tray())

    def CreatePopupMenu(self):
        menu = wx.Menu()
        s = menu.Append(wx.ID_ANY, "显示主窗口")
        h = menu.Append(wx.ID_ANY, "隐藏到托盘")
        menu.AppendSeparator()
        about = menu.Append(wx.ID_ABOUT, f"关于 QuickLauncher v{__version__}")
        menu.AppendSeparator()
        x = menu.Append(wx.ID_EXIT, "退出")
        self.Bind(wx.EVT_MENU, lambda e: self.frame.show_from_tray(), s)
        self.Bind(wx.EVT_MENU, lambda e: self.frame.hide_to_tray(), h)
        self.Bind(wx.EVT_MENU, lambda e: self.frame.on_about(None), about)
        self.Bind(wx.EVT_MENU, lambda e: self.frame.exit_app(), x)
        return menu


# ---------------------------
# 主窗口
# ---------------------------
class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title=f"{__app_name__} v{__version__}", size=(980, 580))
        
        # 设置窗口图标
        icon_path = get_icon_path()
        if icon_path and os.path.exists(icon_path):
            if icon_path.endswith('.png'):
                bmp = wx.Bitmap(icon_path, wx.BITMAP_TYPE_PNG)
                icon = wx.Icon()
                icon.CopyFromBitmap(bmp)
                self.SetIcon(icon)
            else:
                self.SetIcon(wx.Icon(icon_path, wx.BITMAP_TYPE_ICO))
        
        cfg = load_config()
        self.programs = cfg["programs"]
        self.autostart = cfg["autostart"]
        self.hotkey_id_to_index = {}
        self.registered_hotkey_ids = []
        self.exiting = False

        self.init_ui()
        self.Centre()

        self.taskbar = QuickLauncherTaskBar(self)
        self.refresh_list()
        self.register_all_hotkeys()

        self.Bind(wx.EVT_CLOSE, self.on_close)

        self.fg_timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.on_timer, self.fg_timer)
        self.fg_timer.Start(1200)

    def on_timer(self, _):
        update_last_active_cache()

    def init_ui(self):
        # 创建菜单栏
        menubar = wx.MenuBar()
        
        # 帮助菜单
        help_menu = wx.Menu()
        check_update_item = help_menu.Append(wx.ID_ANY, "检查更新", "检查是否有新版本")
        self.Bind(wx.EVT_MENU, self.check_for_updates, check_update_item)
        about_item = help_menu.Append(wx.ID_ABOUT, f"关于 {__app_name__}", "关于本软件")
        self.Bind(wx.EVT_MENU, self.on_about, about_item)
        help_menu.AppendSeparator()
        exit_item = help_menu.Append(wx.ID_EXIT, "退出", "退出程序")
        self.Bind(wx.EVT_MENU, self.exit_app, exit_item)
        
        menubar.Append(help_menu, "帮助")
        self.SetMenuBar(menubar)
        
        panel = wx.Panel(self)
        root = wx.BoxSizer(wx.VERTICAL)

        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "名称", width=120)
        self.list_ctrl.InsertColumn(1, "快捷键", width=120)
        self.list_ctrl.InsertColumn(2, "模式", width=80)
        self.list_ctrl.InsertColumn(3, "关键词", width=160)
        self.list_ctrl.InsertColumn(4, "HWND", width=90)
        self.list_ctrl.InsertColumn(5, "Profile", width=100)
        self.list_ctrl.InsertColumn(6, "TitleSig", width=180)
        self.list_ctrl.InsertColumn(7, "路径", width=240)
        root.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        row = wx.BoxSizer(wx.HORIZONTAL)
        for label, fn in [
            ("手动添加", self.on_manual_add),
            ("从运行窗口添加", self.on_add_from_running),
            ("删除", self.on_delete),
            ("设置快捷键", self.on_set_hotkey),
            ("设置匹配", self.on_set_match),
            ("最小化到托盘", lambda e: self.hide_to_tray()),
        ]:
            b = wx.Button(panel, label=label)
            b.Bind(wx.EVT_BUTTON, fn)
            row.Add(b, 0, wx.ALL, 4)
        root.Add(row, 0, wx.ALIGN_CENTER)

        row2 = wx.BoxSizer(wx.HORIZONTAL)
        self.autostart_cb = wx.CheckBox(panel, label="开机自启")
        self.autostart_cb.SetValue(self.autostart)
        row2.Add(self.autostart_cb, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 6)
        save_btn = wx.Button(panel, label="保存设置")
        save_btn.Bind(wx.EVT_BUTTON, self.on_apply_settings)
        row2.Add(save_btn, 0, wx.ALL, 6)
        root.Add(row2, 0, wx.ALIGN_CENTER)

        self.status = wx.StaticText(panel, label="就绪")
        root.Add(self.status, 0, wx.ALL | wx.ALIGN_CENTER, 6)
        panel.SetSizer(root)

    def persist(self):
        save_config(self.programs, self.autostart_cb.GetValue())

    def refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for p in self.programs:
            self.list_ctrl.Append([
                p.get("name", ""),
                p.get("hotkey", ""),
                p.get("match_mode", "title"),
                p.get("window_keyword", ""),
                str(int(p.get("bind_hwnd", 0) or 0)),
                p.get("profile_name", ""),
                p.get("title_sig", ""),
                p.get("path", "")
            ])

    def update_status(self, s):
        self.status.SetLabel(s)

    def hide_to_tray(self):
        self.Hide()
        self.update_status("已最小化到托盘")

    def show_from_tray(self):
        self.Show()
        self.Raise()
        self.Iconize(False)
        self.update_status("窗口已恢复")

    def on_close(self, event):
        if self.exiting:
            self.unregister_all_hotkeys()
            if self.taskbar:
                self.taskbar.RemoveIcon()
                self.taskbar.Destroy()
            event.Skip()
        else:
            event.Veto()
            self.hide_to_tray()

    def on_about(self, _):
        """显示关于对话框"""
        info = wx.adv.AboutDialogInfo()
        info.SetName(__app_name__)
        info.SetVersion(__version__)
        info.SetDescription(f"{__description__}\n\n作者：{__author__}\nGitHub：{__github__}")
        info.SetWebSite(__github__)
        info.AddDeveloper(__author__)
        wx.adv.AboutBox(info)

    def check_for_updates(self):
        """检查更新"""
        import urllib.request
        import json
        try:
            url = f"https://api.github.com/repos/laozi4vip/QuickLauncher/releases/latest"
            req = urllib.request.Request(url, headers={'User-Agent': 'QuickLauncher'})
            with urllib.request.urlopen(req, timeout=10) as response:
                data = json.loads(response.read().decode())
                latest_version = data.get('tag_name', 'v1.0.0').lstrip('v')
                current_version = __version__
                
                # 比较版本
                def parse_version(v):
                    parts = v.split('.')
                    return [int(p) for p in parts] + [0] * (3 - len(parts))
                
                latest = parse_version(latest_version)
                current = parse_version(current_version)
                
                if latest > current:
                    dlg = wx.Dialog(self, title="检查更新", size=(400, 180))
                    panel = wx.Panel(dlg)
                    sizer = wx.BoxSizer(wx.VERTICAL)
                    sizer.Add(wx.StaticText(panel, label=f"当前版本：v{current_version}"), 0, wx.ALL, 10)
                    sizer.Add(wx.StaticText(panel, label=f"最新版本：v{latest_version}"), 0, wx.ALL, 10)
                    sizer.Add(wx.StaticText(panel, label="发现新版本！"), 0, wx.ALL, 10)
                    
                    btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
                    ok_btn = wx.Button(panel, label="前往下载")
                    ok_btn.Bind(lambda e: wx.LaunchBrowserInDefaultBrowser(__github__ + "/releases"))
                    btn_sizer.Add(ok_btn, 0, wx.ALL, 5)
                    cancel_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
                    btn_sizer.Add(cancel_btn, 0, wx.ALL, 5)
                    sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER|wx.ALL, 10)
                    
                    panel.SetSizer(sizer)
                    dlg.ShowModal()
                    dlg.Destroy()
                else:
                    wx.MessageBox(f"当前已是最新版本 (v{current_version})", "检查更新", wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"检查更新失败：{str(e)}", "错误", wx.OK | wx.ICON_ERROR)

    def exit_app(self):
        self.exiting = True
        self.Close()

    def on_apply_settings(self, _):
        ok = set_autostart(self.autostart_cb.GetValue())
        if not ok:
            wx.MessageBox("开机自启设置失败", "错误", wx.OK | wx.ICON_ERROR)
            self.autostart_cb.SetValue(is_autostart_enabled())
            return
        self.persist()
        wx.MessageBox("设置已保存", "成功")

    # ---- 自动补绑逻辑 ----
    def get_used_hwnds_by_same_path(self, path, exclude_index=None):
        used = set()
        tpath = os.path.normcase(path or "")
        for i, p in enumerate(self.programs):
            if exclude_index is not None and i == exclude_index:
                continue
            ppath = os.path.normcase(p.get("path", "") or "")
            if ppath != tpath:
                continue
            h = int(p.get("bind_hwnd", 0) or 0)
            if h and is_hwnd_valid(h):
                used.add(h)
        return used

    def auto_bind_program_if_needed(self, idx, windows_cache=None, save=False):
        if idx < 0 or idx >= len(self.programs):
            return False
        p = self.programs[idx]
        if (p.get("match_mode", "title") or "title").lower() != "hwnd":
            return False

        current = int(p.get("bind_hwnd", 0) or 0)
        path = p.get("path", "")
        if not path:
            return False

        if current and is_hwnd_valid(current):
            pid = get_pid_from_hwnd(current)
            cpath, _, _ = get_proc_path_name_cmdline(pid)
            if os.path.normcase(cpath or "") == os.path.normcase(path or ""):
                return False

        cands = windows_cache if windows_cache is not None else enum_windows_for_program(path)
        if not cands:
            return False

        used = self.get_used_hwnds_by_same_path(path, exclude_index=idx)
        scored = []
        for w in cands:
            hwnd = int(w.get("hwnd", 0) or 0)
            if not hwnd or hwnd in used or not is_hwnd_valid(hwnd):
                continue
            s = score_window_for_program(p, w)
            scored.append((s, hwnd, w))

        if not scored:
            return False

        scored.sort(key=lambda x: x[0], reverse=True)
        best_score, best_hwnd, best_w = scored[0]

        kw = (p.get("window_keyword", "") or "").strip()
        pf = (p.get("profile_name", "") or "").strip()
        ts = (p.get("title_sig", "") or "").strip()
        if (kw or pf or ts) and best_score <= 0:
            return False

        p["bind_hwnd"] = int(best_hwnd)
        if not (p.get("title_sig", "") or "").strip():
            p["title_sig"] = best_w.get("title_sig", "") or make_title_signature(best_w.get("title", ""))
        if save:
            self.persist()
        return True

    def auto_bind_unbound_same_browser(self, base_index):
        if base_index < 0 or base_index >= len(self.programs):
            return 0
        base = self.programs[base_index]
        path = base.get("path", "")
        if not path:
            return 0
        cands = enum_windows_for_program(path)
        if not cands:
            return 0
        changed = 0
        for i, p in enumerate(self.programs):
            if os.path.normcase(p.get("path", "") or "") != os.path.normcase(path or ""):
                continue
            if (p.get("match_mode", "title") or "title").lower() != "hwnd":
                continue
            if self.auto_bind_program_if_needed(i, windows_cache=cands, save=False):
                changed += 1
        if changed:
            self.persist()
            self.refresh_list()
        return changed

    # ---- 热键 ----
    def unregister_all_hotkeys(self):
        for i in self.registered_hotkey_ids:
            try:
                self.UnregisterHotKey(i)
            except Exception:
                pass
        self.registered_hotkey_ids.clear()
        self.hotkey_id_to_index.clear()

    def register_all_hotkeys(self):
        self.unregister_all_hotkeys()
        used = set()
        fail = []
        base_id = 1000
        for idx, p in enumerate(self.programs):
            hk = normalize_hotkey(p.get("hotkey", ""))
            if not hk:
                continue
            try:
                mods, vk, n = hotkey_to_mod_vk(hk)
                self.programs[idx]["hotkey"] = n
            except Exception as e:
                fail.append(f"{p.get('name', '')}：{hk} ({e})")
                continue
            if n in used:
                fail.append(f"{p.get('name', '')}：{n}（重复）")
                continue
            used.add(n)
            hotkey_id = base_id + idx
            if self.RegisterHotKey(hotkey_id, mods, vk):
                self.registered_hotkey_ids.append(hotkey_id)
                self.hotkey_id_to_index[hotkey_id] = idx
                self.Bind(wx.EVT_HOTKEY, self.on_hotkey, id=hotkey_id)
            else:
                fail.append(f"{p.get('name', '')}：{n}（被占用）")
        self.persist()
        self.refresh_list()
        if fail:
            wx.MessageBox("以下快捷键注册失败：\n" + "\n".join(fail), "提示")

    def on_hotkey(self, event):
        idx = self.hotkey_id_to_index.get(event.GetId())
        if idx is None or idx >= len(self.programs):
            return
        p = self.programs[idx]
        mode = (p.get("match_mode", "title") or "title").lower()
        old_hwnd = int(p.get("bind_hwnd", 0) or 0)

        if mode == "hwnd":
            self.auto_bind_program_if_needed(idx, save=True)
            old_hwnd = int(p.get("bind_hwnd", 0) or 0)

        hwnd, _focused = toggle_program(p)

        # ★ 浏览器类：没匹配到窗口（hwnd 为 None 且未聚焦），直接忽略，不做任何事
        if is_browser_program(p) and hwnd is None and not _focused:
            return

        changed = False

        if mode == "hwnd":
            keep_old = False
            if old_hwnd and is_hwnd_valid(old_hwnd):
                pid = get_pid_from_hwnd(old_hwnd)
                cpath, _, _ = get_proc_path_name_cmdline(pid)
                if os.path.normcase(cpath or "") == os.path.normcase(p.get("path", "") or ""):
                    keep_old = True
            if hwnd and not keep_old and hwnd != old_hwnd:
                p["bind_hwnd"] = int(hwnd)
                p["title_sig"] = make_title_signature(get_window_title(hwnd))
                changed = True
            auto_cnt = self.auto_bind_unbound_same_browser(idx)
            if auto_cnt > 0:
                changed = True

        if changed:
            self.persist()
            self.refresh_list()
        self.update_status(f"已切换：{p.get('name', '')}")

    # ---- 添加/编辑 ----
    def on_manual_add(self, _):
        dlg = wx.Dialog(self, title="手动添加", size=(700, 360))
        panel = wx.Panel(dlg)
        s = wx.BoxSizer(wx.VERTICAL)

        def row(label, ctrl):
            r = wx.BoxSizer(wx.HORIZONTAL)
            r.Add(wx.StaticText(panel, label=label), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 6)
            r.Add(ctrl, 1, wx.ALL | wx.EXPAND, 6)
            s.Add(r, 0, wx.EXPAND)

        name = wx.TextCtrl(panel)
        path = wx.TextCtrl(panel)
        args = wx.TextCtrl(panel)
        mode = wx.Choice(panel, choices=["title", "profile", "hwnd"])
        mode.SetStringSelection("title")
        kw = wx.TextCtrl(panel)
        hwnd = wx.TextCtrl(panel, value="0")

        row("名称:", name)
        row("路径:", path)
        row("启动参数:", args)
        row("模式:", mode)
        row("关键词:", kw)
        row("HWND:", hwnd)

        b = wx.Button(panel, label="浏览exe")

        def browse(_e):
            fd = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if fd.ShowModal() == wx.ID_OK:
                pth = fd.GetPath()
                path.SetValue(pth)
                if not name.GetValue().strip():
                    name.SetValue(os.path.splitext(os.path.basename(pth))[0])
            fd.Destroy()

        b.Bind(wx.EVT_BUTTON, browse)
        s.Add(b, 0, wx.ALL | wx.ALIGN_CENTER, 6)

        btns = wx.StdDialogButtonSizer()
        okb = wx.Button(panel, wx.ID_OK)
        cb = wx.Button(panel, wx.ID_CANCEL)
        btns.AddButton(okb)
        btns.AddButton(cb)
        btns.Realize()
        s.Add(btns, 0, wx.ALL | wx.ALIGN_CENTER, 8)
        panel.SetSizer(s)

        if dlg.ShowModal() == wx.ID_OK:
            p = {
                "name": name.GetValue().strip(),
                "path": path.GetValue().strip(),
                "args": args.GetValue().strip(),
                "hotkey": "",
                "window_keyword": kw.GetValue().strip(),
                "match_mode": mode.GetStringSelection() or "title",
                "bind_hwnd": int(hwnd.GetValue().strip() or "0") if hwnd.GetValue().strip().isdigit() else 0,
                "profile_name": kw.GetValue().strip() if (mode.GetStringSelection() == "profile") else "",
                "title_sig": ""
            }
            if not p["name"] or not p["path"]:
                wx.MessageBox("名称和路径必填", "提示")
            elif not os.path.exists(p["path"]):
                wx.MessageBox("路径不存在", "错误")
            else:
                self.programs.append(p)
                self.persist()
                self.refresh_list()
                self.register_all_hotkeys()
        dlg.Destroy()

    def on_add_from_running(self, _):
        ws = enum_visible_app_windows()
        if not ws:
            wx.MessageBox("未找到运行窗口", "提示")
            return
        ws = [w for w in ws if w.get("path", "").lower().endswith(".exe")]
        ws.sort(key=lambda x: (x.get("title", "") or "").lower())

        dlg = wx.Dialog(self, title="从运行窗口添加", size=(1020, 520))
        panel = wx.Panel(dlg)
        s = wx.BoxSizer(wx.VERTICAL)
        tip = wx.StaticText(panel, label="双击添加：浏览器默认用 hwnd 模式 + 自动重绑（Profile 列现在会正确区分）")
        s.Add(tip, 0, wx.ALL, 8)

        lc = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        lc.InsertColumn(0, "标题", width=330)
        lc.InsertColumn(1, "程序", width=110)
        lc.InsertColumn(2, "Profile", width=120)
        lc.InsertColumn(3, "HWND", width=90)
        lc.InsertColumn(4, "路径", width=330)

        rows = []
        for w in ws:
            path = w["path"]
            name = os.path.splitext(os.path.basename(path))[0]
            row = {
                "name": name,
                "path": path,
                "title": w.get("title", ""),
                "profile": w.get("profile", ""),
                "proc_name": w.get("proc_name", ""),
                "hwnd": int(w.get("hwnd", 0)),
                "title_sig": w.get("title_sig", "")
            }
            rows.append(row)
            i = lc.InsertItem(lc.GetItemCount(), row["title"] or "(无标题)")
            lc.SetItem(i, 1, row["name"])
            lc.SetItem(i, 2, row["profile"])
            lc.SetItem(i, 3, str(row["hwnd"]))
            lc.SetItem(i, 4, row["path"])

        def on_dbl(_e):
            i = lc.GetFirstSelected()
            if i < 0:
                return
            it = rows[i]
            proc = (it.get("proc_name", "") or "").lower().replace(".exe", "")
            profile = (it.get("profile", "") or "").strip()

            mode = "hwnd" if proc in BROWSER_SET else "title"
            kw = profile if (mode == "hwnd" and profile) else (it.get("title", "")[:80])
            args = ""

            if proc in BROWSER_SET:
                ok, m_profile = ask_profile_input(self, default_profile=profile)
                if ok and m_profile:
                    args = build_profile_args(proc, m_profile)
                    kw = m_profile
                    profile = m_profile

            self.programs.append({
                "name": it["name"],
                "path": it["path"],
                "args": args,
                "hotkey": "",
                "window_keyword": kw,
                "match_mode": mode,
                "bind_hwnd": it["hwnd"] if mode == "hwnd" else 0,
                "profile_name": profile,
                "title_sig": it.get("title_sig", "")
            })
            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()
            dlg.Destroy()

        lc.Bind(wx.EVT_LIST_ITEM_ACTIVATED, on_dbl)
        s.Add(lc, 1, wx.ALL | wx.EXPAND, 8)

        close_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dlg.Destroy())
        s.Add(close_btn, 0, wx.ALL | wx.ALIGN_CENTER, 6)
        panel.SetSizer(s)
        dlg.ShowModal()
        dlg.Destroy()

    def on_delete(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择", "提示")
            return
        self.programs.pop(idx)
        self.persist()
        self.refresh_list()
        self.register_all_hotkeys()

    def on_set_hotkey(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择程序", "提示")
            return
        dlg = HotkeyCaptureDialog(self, self.programs[idx].get("hotkey", ""))
        if dlg.ShowModal() == wx.ID_OK:
            hk = normalize_hotkey(dlg.captured_hotkey)
            if not hk:
                self.programs[idx]["hotkey"] = ""
            else:
                try:
                    _, _, n = hotkey_to_mod_vk(hk)
                except Exception as e:
                    wx.MessageBox(f"快捷键无效: {e}", "错误")
                    dlg.Destroy()
                    return
                for i, p in enumerate(self.programs):
                    if i != idx and normalize_hotkey(p.get("hotkey", "")) == n:
                        wx.MessageBox("快捷键重复", "错误")
                        dlg.Destroy()
                        return
                self.programs[idx]["hotkey"] = n
            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()
        dlg.Destroy()

    def on_set_match(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择程序", "提示")
            return
        p = self.programs[idx]
        dlg = wx.Dialog(self, title="设置匹配", size=(560, 300))
        panel = wx.Panel(dlg)
        s = wx.BoxSizer(wx.VERTICAL)

        mode = wx.Choice(panel, choices=["title", "profile", "hwnd"])
        mode.SetStringSelection(p.get("match_mode", "title"))
        kw = wx.TextCtrl(panel, value=p.get("window_keyword", ""))
        hwnd = wx.TextCtrl(panel, value=str(int(p.get("bind_hwnd", 0) or 0)))
        prof = wx.TextCtrl(panel, value=p.get("profile_name", ""))
        tsig = wx.TextCtrl(panel, value=p.get("title_sig", ""))

        for lab, ctrl in [("模式:", mode), ("关键词:", kw), ("绑定HWND:", hwnd), ("Profile名:", prof), ("TitleSig:", tsig)]:
            r = wx.BoxSizer(wx.HORIZONTAL)
            r.Add(wx.StaticText(panel, label=lab), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 6)
            r.Add(ctrl, 1, wx.ALL | wx.EXPAND, 6)
            s.Add(r, 0, wx.EXPAND)

        tip = wx.StaticText(panel, label="浏览器推荐：mode=hwnd，profile_name 填 Profile 目录名（如 Profile 1）")
        tip.SetForegroundColour(wx.Colour(100, 100, 100))
        s.Add(tip, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        btns = wx.StdDialogButtonSizer()
        okb = wx.Button(panel, wx.ID_OK)
        cb = wx.Button(panel, wx.ID_CANCEL)
        btns.AddButton(okb)
        btns.AddButton(cb)
        btns.Realize()
        s.Add(btns, 0, wx.ALL | wx.ALIGN_CENTER, 8)
        panel.SetSizer(s)

        if dlg.ShowModal() == wx.ID_OK:
            p["match_mode"] = mode.GetStringSelection() or "title"
            p["window_keyword"] = kw.GetValue().strip()
            p["profile_name"] = prof.GetValue().strip()
            p["title_sig"] = tsig.GetValue().strip()
            try:
                p["bind_hwnd"] = int(hwnd.GetValue().strip() or "0")
            except Exception:
                p["bind_hwnd"] = 0
            self.persist()
            self.refresh_list()
            wx.MessageBox("已保存", "成功")
        dlg.Destroy()


class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True


if __name__ == "__main__":
    app = QuickLauncherApp()
    app.MainLoop()
