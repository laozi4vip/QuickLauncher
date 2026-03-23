# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器
作者：laozi4vip
GitHub：https://github.com/laozi4vip/QuickLauncher
"""

__author__ = "laozi4vip"
__github__ = "https://github.com/laozi4vip/QuickLauncher"
__app_name__ = "QuickLauncher"
__version__ = "3.6"
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
SW_HIDE = 0
SW_SHOW = 5
SW_MINIMIZE = 6
SW_RESTORE = 9

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008

GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000

SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SWP_FRAMECHANGED = 0x0020

RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_RUN_NAME = "QuickLauncher"

BROWSER_SET = {"chrome", "msedge", "chromium", "brave", "firefox"}

# ---------------------------
# 配置路径
# ---------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

LAST_ACTIVE_PROFILE_HWND = {}

_browser_main_map_cache = {}
_browser_main_map_ts = 0.0


# ---------------------------
# 配置 / 自启
# ---------------------------
def get_launch_command(start_in_tray=False):
    tray_arg = " --tray" if start_in_tray else ""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"{tray_arg}'
    return f'"{sys.executable}" "{os.path.abspath(__file__)}"{tray_arg}'


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
                winreg.SetValueEx(key, APP_RUN_NAME, 0, winreg.REG_SZ, get_launch_command(start_in_tray=True))
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
                p.setdefault("browser_fallback_exe", False)
                p.setdefault("browser_group_toggle", True)
                p.setdefault("hotkey_action", "toggle")
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
    hk = (hotkey or "").strip().lower().replace(" ", "")
    hk = hk.replace("grave", "`").replace("tilde", "`")
    return hk

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

    key_map["`"] = 0xC0
    key_map["grave"] = 0xC0
    key_map["tilde"] = 0xC0
    
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
    elif ord("0") <= key <= ord("9"):
        main_key = chr(key).lower()
    elif ord("A") <= key <= ord("Z"):
        main_key = chr(key).lower()
    elif ord("a") <= key <= ord("z"):
        main_key = chr(key).lower()
    elif key in (ord("`"), ord("~"), 0xC0):
        main_key = "`"
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
        return bool(hwnd) and bool(user32.IsWindow(hwnd))
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

def parse_profile_from_profile_path_text(s: str):
    t = (s or "").strip().replace("/", "\\").rstrip("\\")
    if not t:
        return ""
    base = os.path.basename(t)
    if re.match(r"(?i)^profile\s*\d+$", base):
        n = re.findall(r"\d+", base)[0]
        return f"Profile {n}"
    if base.lower() in ("default", "guest", "personal", "work"):
        return base.capitalize()
    return ""


def parse_profile_from_cwd(pid):
    try:
        p = psutil.Process(pid)
        cwd = p.cwd() or ""
    except Exception:
        return ""

    if not cwd:
        return ""

    c = cwd.replace("/", "\\").rstrip("\\")
    base = os.path.basename(c)
    if re.match(r"(?i)^profile\s*\d+$", base):
        n = re.findall(r"\d+", base)[0]
        return f"Profile {n}"
    if base.lower() in ("default", "guest", "personal", "work"):
        return base.capitalize()

    return ""


def parse_profile_from_title(proc_name: str, title: str):
    t = (title or "").strip()
    if not t:
        return ""

    m = re.search(r"\((profile\s*\d+|default|personal|work|guest)\)", t, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    m2 = re.search(r"\b(profile\s*\d+|default|personal|work|guest)\b", t, re.IGNORECASE)
    if m2:
        return m2.group(1).strip()

    return ""


def guess_profile(proc_name: str, cmdline_list, title: str, pid: int = 0):
    if pid:
        p0 = parse_profile_from_cwd(pid)
        if p0:
            return p0

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


def set_window_exstyle(hwnd: int, exstyle: int):
    try:
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, exstyle)
        user32.SetWindowPos(
            hwnd,
            0,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    except Exception:
        pass


# ---------------------------
# 判断程序是否属于浏览器
# ---------------------------
def is_browser_program(program):
    path = program.get("path", "") or ""
    exe_name = os.path.basename(path).lower().replace(".exe", "")
    return exe_name in BROWSER_SET


def browser_fallback_enabled(program):
    return bool(program.get("browser_fallback_exe", False))


def browser_group_toggle_enabled(program):
    return bool(program.get("browser_group_toggle", True))


def is_hide_action(program):
    return (program.get("hotkey_action", "toggle") or "toggle").strip().lower() == "hide"


# ---------------------------
# 浏览器主进程扫描 & 进程树回溯
# ---------------------------
def build_browser_main_proc_map():
    main_map = {}
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            pname = (proc.name() or "").lower().replace(".exe", "")
            if pname not in BROWSER_SET:
                continue
            try:
                cmdline = proc.cmdline()
            except (psutil.AccessDenied, psutil.ZombieProcess):
                continue
            if any(a.lower().startswith("--type=") for a in cmdline):
                continue
            profile = parse_profile_from_cwd(proc.pid)
            if not profile:
                profile = parse_profile_from_cmdline(proc.name(), cmdline)
            if not profile:
                profile = "Default"
            main_map[proc.pid] = profile
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return main_map


def get_cached_browser_main_map(max_age=3.0):
    global _browser_main_map_cache, _browser_main_map_ts
    now = time.time()
    if now - _browser_main_map_ts < max_age and _browser_main_map_cache:
        return _browser_main_map_cache
    _browser_main_map_cache = build_browser_main_proc_map()
    _browser_main_map_ts = now
    return _browser_main_map_cache


def get_profile_by_pid_tree(pid, main_map):
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
# 窗口枚举
# ---------------------------
def enum_visible_app_windows():
    windows = []
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
                profile = get_profile_by_pid_tree(pid, browser_main_map)
                if not profile:
                    profile = guess_profile(proc_name, cmdline, title, pid)
            else:
                profile = guess_profile(proc_name, cmdline, title, pid)

            windows.append(
                {
                    "hwnd": int(hwnd),
                    "pid": int(pid),
                    "title": title,
                    "title_sig": make_title_signature(title),
                    "path": path,
                    "proc_name": proc_name,
                    "cmdline": cmdline,
                    "profile": profile,
                    "profile_norm": _normalize_profile_text(profile),
                    "profile_raw": profile,
                }
            )
            
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
            browser_main_map = get_cached_browser_main_map(max_age=5.0)
            profile = get_profile_by_pid_tree(pid, browser_main_map)
            if not profile:
                profile = guess_profile(proc_name, cmdline, title, pid)
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


def _normalize_profile_text(s: str):
    t = (s or "").strip().lower()
    if not t:
        return ""
    t = t.replace("　", " ")
    t = re.sub(r"\s+", " ", t)
    t = t.replace("-", " ")
    t = re.sub(r"^profile\s*(\d+)$", r"profile \1", t)
    if t == "default profile":
        t = "default"
    return t

def _profile_match(target_profile: str, w_profile: str):
    tp = _normalize_profile_text(target_profile)
    wp = _normalize_profile_text(w_profile)
    if not tp or not wp:
        return False
    return wp == tp or (tp in wp) or (wp in tp)


# 新增：严格匹配要求（设置了哪些条件，就必须全部命中）
def browser_window_matches_all_configured(program, w):
    match_mode = (program.get("match_mode", "title") or "title").strip().lower()
    keyword = (program.get("window_keyword", "") or "").strip()
    profile_name = (program.get("profile_name", "") or "").strip()
    title_sig = (program.get("title_sig", "") or "").strip()
    bind_hwnd = int(program.get("bind_hwnd", 0) or 0)

    w_hwnd = int(w.get("hwnd", 0) or 0)
    w_title = (w.get("title", "") or "").lower()
    w_sig = (w.get("title_sig", "") or "").lower()

    w_pf = _normalize_profile_text(w.get("profile", ""))
    w_cmd_pf = _normalize_profile_text(parse_profile_from_cmdline(w.get("proc_name", ""), w.get("cmdline", [])))
    w_t_pf = _normalize_profile_text(parse_profile_from_title(w.get("proc_name", ""), w.get("title", "")))

    checks = []

    # mode 对应条件
    if match_mode == "hwnd":
        checks.append(bind_hwnd > 0 and w_hwnd == bind_hwnd)
    elif match_mode == "title":
        checks.append(bool(keyword) and (keyword.lower() in w_title))
    elif match_mode == "profile":
        tp = _normalize_profile_text(profile_name or keyword)
        checks.append(bool(tp) and any(_profile_match(tp, x) for x in (w_pf, w_cmd_pf, w_t_pf) if x))

    # 额外配置：只要填了，也必须匹配
    if keyword:
        if match_mode == "profile":
            tp_kw = _normalize_profile_text(keyword)
            checks.append(any(_profile_match(tp_kw, x) for x in (w_pf, w_cmd_pf, w_t_pf) if x))
        elif match_mode == "hwnd":
            checks.append(True)  # hwnd 模式下 keyword 仅视作备注，不强制
        else:
            checks.append(keyword.lower() in w_title)

    if profile_name:
        tp_pf = _normalize_profile_text(profile_name)
        checks.append(any(_profile_match(tp_pf, x) for x in (w_pf, w_cmd_pf, w_t_pf) if x))

    if title_sig:
        ts = title_sig.lower()
        checks.append(bool(w_sig) and (w_sig == ts or ts in w_sig or w_sig in ts))

    if not checks:
        return False
    return all(checks)


def score_window_for_program(program, w):
    score = 0
    keyword = (program.get("window_keyword", "") or "").strip().lower()
    profile_name = _normalize_profile_text(program.get("profile_name", ""))
    title_sig = (program.get("title_sig", "") or "").strip().lower()
    match_mode = (program.get("match_mode", "title") or "title").strip().lower()

    w_title = (w.get("title", "") or "").lower()
    w_sig = (w.get("title_sig", "") or "").lower()
    w_prof = _normalize_profile_text(w.get("profile", ""))
    w_cmd_prof = _normalize_profile_text(
        parse_profile_from_cmdline(w.get("proc_name", ""), w.get("cmdline", []))
    )
    if not w_cmd_prof:
        w_cmd_prof = _normalize_profile_text(
            parse_profile_from_title(w.get("proc_name", ""), w.get("title", ""))
        )
    
    if profile_name and w_cmd_prof:
        if w_cmd_prof == profile_name:
            score += 140
        elif profile_name in w_cmd_prof or w_cmd_prof in profile_name:
            score += 90
            
    if w_prof and w_cmd_prof and w_prof == w_cmd_prof:
        score += 20

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


def find_window_for_program(program):
    path = program.get("path", "")
    if not path:
        return None

    match_mode = (program.get("match_mode", "title") or "title").strip().lower()
    bind_hwnd = int(program.get("bind_hwnd", 0) or 0)

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
    
    if is_browser_program(program):
        # 浏览器改为严格：不使用“有分就算匹配”
        strict = [x for x in candidates if browser_window_matches_all_configured(program, x)]
        if not strict:
            return None
        strict.sort(key=lambda w: score_window_for_program(program, w), reverse=True)
        return int(strict[0].get("hwnd", 0) or 0) or None

    if best_score <= 0:
        return None
    return best_hwnd


def find_browser_group_windows(program):
    path = program.get("path", "")
    if not path:
        return []

    cands = enum_windows_for_program(path)
    if not cands:
        return []

    # 浏览器组联动改为严格：配置了哪些条件就全部命中
    matched = []
    for w in cands:
        hwnd = int(w.get("hwnd", 0) or 0)
        if not hwnd or not is_hwnd_valid(hwnd):
            continue
        if browser_window_matches_all_configured(program, w):
            matched.append(w)

    if not matched:
        return []

    fg = user32.GetForegroundWindow()
    scored = []
    for w in matched:
        hwnd = int(w.get("hwnd", 0) or 0)
        s = score_window_for_program(program, w)
        if hwnd == fg:
            s += 50
        if hwnd and not user32.IsIconic(hwnd):
            s += 15
        scored.append((s, hwnd))

    scored.sort(key=lambda x: x[0], reverse=True)

    seen = set()
    result = []
    for _, h in scored:
        if h not in seen:
            seen.add(h)
            result.append(h)
    return result


def minimize_windows(hwnds):
    for h in hwnds:
        try:
            if is_hwnd_valid(h):
                user32.ShowWindow(h, SW_MINIMIZE)
        except Exception:
            pass


def restore_windows(hwnds):
    for h in hwnds:
        try:
            if is_hwnd_valid(h) and user32.IsIconic(h):
                user32.ShowWindow(h, SW_RESTORE)
        except Exception:
            pass


def launch_program_by_path(path, args):
    if not os.path.exists(path):
        wx.MessageBox(f"程序不存在：\n{path}", "错误", wx.OK | wx.ICON_ERROR)
        return False
    cmd = [path]
    if (args or "").strip():
        try:
            cmd.extend(shlex.split(args.strip(), posix=False))
        except Exception:
            cmd.extend(args.strip().split())
    subprocess.Popen(cmd)
    return True


def toggle_program(program):
    path = program.get("path", "")
    args = program.get("args", "")
    if not path:
        return None, False

    if is_browser_program(program) and browser_group_toggle_enabled(program):
        group_hwnds = find_browser_group_windows(program)
        if group_hwnds:
            fg = user32.GetForegroundWindow()

            if fg in group_hwnds:
                minimize_windows(group_hwnds)
                return group_hwnds[0], True

            restore_windows(group_hwnds)
            try:
                user32.SetForegroundWindow(group_hwnds[0])
            except Exception:
                pass
            return group_hwnds[0], True

        if browser_fallback_enabled(program):
            ok = launch_program_by_path(path, args)
            return None, ok
        return None, False

    hwnd = find_window_for_program(program)
    if hwnd:
        fg = user32.GetForegroundWindow()
        if hwnd == fg:
            user32.ShowWindow(hwnd, SW_MINIMIZE)
        else:
            if user32.IsIconic(hwnd):
                user32.ShowWindow(hwnd, SW_RESTORE)
            try:
                user32.SetForegroundWindow(hwnd)
            except Exception:
                pass
        return hwnd, True

    if is_browser_program(program):
        if browser_fallback_enabled(program):
            ok = launch_program_by_path(path, args)
            return None, ok
        return None, False

    ok = launch_program_by_path(path, args)
    return None, ok


def ask_profile_input(parent, default_profile=""):
    dlg = wx.TextEntryDialog(
        parent,
        "请输入浏览器 Profile（如 Profile 1 / Default / Work）",
        "确认 Profile",
        value=default_profile or "",
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
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = BASE_DIR

    for ext in ["icon.ico", "icon.png"]:
        icon_path = os.path.join(base, ext)
        if os.path.exists(icon_path):
            return icon_path
        alt_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ext)
        if os.path.exists(alt_path):
            return alt_path
    return None


class QuickLauncherTaskBar(wx.adv.TaskBarIcon):
    def __init__(self, frame):
        super().__init__()
        self.frame = frame
        icon_path = get_icon_path()
        if icon_path and os.path.exists(icon_path):
            if icon_path.endswith(".png"):
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


class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title=f"{__app_name__} v{__version__}", size=(1220, 650))

        icon_path = get_icon_path()
        if icon_path and os.path.exists(icon_path):
            if icon_path.endswith(".png"):
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
        self.hidden_states = {}

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
        menubar = wx.MenuBar()

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
        self.list_ctrl.InsertColumn(2, "动作", width=90)
        self.list_ctrl.InsertColumn(3, "模式", width=80)
        self.list_ctrl.InsertColumn(4, "关键词", width=150)
        self.list_ctrl.InsertColumn(5, "HWND", width=90)
        self.list_ctrl.InsertColumn(6, "Profile", width=100)
        self.list_ctrl.InsertColumn(7, "TitleSig", width=160)
        self.list_ctrl.InsertColumn(8, "浏览器兜底", width=90)
        self.list_ctrl.InsertColumn(9, "组联动", width=80)
        self.list_ctrl.InsertColumn(10, "路径", width=280)
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
            fb = "是" if p.get("browser_fallback_exe", False) else "否"
            gt = "是" if p.get("browser_group_toggle", True) else "否"
            act = "隐藏/恢复" if is_hide_action(p) else "切换/启动"
            self.list_ctrl.Append(
                [
                    p.get("name", ""),
                    p.get("hotkey", ""),
                    act,
                    p.get("match_mode", "title"),
                    p.get("window_keyword", ""),
                    str(int(p.get("bind_hwnd", 0) or 0)),
                    p.get("profile_name", ""),
                    p.get("title_sig", ""),
                    fb,
                    gt,
                    p.get("path", ""),
                ]
            )

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
            self.restore_all_hidden_windows()
            self.unregister_all_hotkeys()
            if self.taskbar:
                self.taskbar.RemoveIcon()
                self.taskbar.Destroy()
            event.Skip()
        else:
            event.Veto()
            self.hide_to_tray()

    def on_about(self, _):
        info = wx.adv.AboutDialogInfo()
        info.SetName(__app_name__)
        info.SetVersion(__version__)
        info.SetDescription(f"{__description__}\n\n作者：{__author__}\nGitHub：{__github__}")
        info.SetWebSite(__github__)
        info.AddDeveloper(__author__)
        wx.adv.AboutBox(info)

    def check_for_updates(self, _):
        import urllib.request
        import json as _json

        try:
            url = "https://api.github.com/repos/laozi4vip/QuickLauncher/releases/latest"
            req = urllib.request.Request(url, headers={"User-Agent": "QuickLauncher"})
            with urllib.request.urlopen(req, timeout=10) as response:
                data = _json.loads(response.read().decode())
                latest_version = data.get("tag_name", "v1.0.0").lstrip("v")
                current_version = __version__

                def parse_version(v):
                    parts = v.split(".")
                    return [int(p) for p in parts] + [0] * (3 - len(parts))

                latest = parse_version(latest_version)
                current = parse_version(current_version)

                if latest > current:
                    dlg = wx.Dialog(self, title="检查更新", size=(420, 190))
                    panel = wx.Panel(dlg)
                    sizer = wx.BoxSizer(wx.VERTICAL)
                    sizer.Add(wx.StaticText(panel, label=f"当前版本：v{current_version}"), 0, wx.ALL, 10)
                    sizer.Add(wx.StaticText(panel, label=f"最新版本：v{latest_version}"), 0, wx.ALL, 10)
                    sizer.Add(wx.StaticText(panel, label="发现新版本！"), 0, wx.ALL, 10)

                    btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
                    ok_btn = wx.Button(panel, label="前往下载")
                    ok_btn.Bind(wx.EVT_BUTTON, lambda e: wx.LaunchDefaultBrowser(__github__ + "/releases"))
                    btn_sizer.Add(ok_btn, 0, wx.ALL, 5)
                    cancel_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
                    btn_sizer.Add(cancel_btn, 0, wx.ALL, 5)
                    sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 10)

                    panel.SetSizer(sizer)
                    dlg.ShowModal()
                    dlg.Destroy()
                else:
                    wx.MessageBox(f"当前已是最新版本 (v{current_version})", "检查更新", wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"检查更新失败：{str(e)}", "错误", wx.OK | wx.ICON_ERROR)

    def exit_app(self, _=None):
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

    def _hide_window_and_taskbar(self, hwnd: int):
        if not is_hwnd_valid(hwnd):
            return None
        old_ex = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        new_ex = (old_ex | WS_EX_TOOLWINDOW) & (~WS_EX_APPWINDOW)
        set_window_exstyle(hwnd, new_ex)
        user32.ShowWindow(hwnd, SW_HIDE)
        return {"hwnd": int(hwnd), "exstyle": int(old_ex)}

    def _restore_window_and_taskbar(self, item):
        hwnd = int(item.get("hwnd", 0) or 0)
        if not is_hwnd_valid(hwnd):
            return False
        old_ex = int(item.get("exstyle", 0))
        set_window_exstyle(hwnd, old_ex)
        user32.ShowWindow(hwnd, SW_SHOW)
        user32.ShowWindow(hwnd, SW_RESTORE)
        return True

    def restore_all_hidden_windows(self):
        keys = list(self.hidden_states.keys())
        for k in keys:
            state = self.hidden_states.get(k, {})
            items = state.get("items", [])
            for it in items:
                self._restore_window_and_taskbar(it)
        self.hidden_states.clear()

    def toggle_hide_for_program(self, idx, p):
        key = int(idx)
        state = self.hidden_states.get(key)

        if state:
            items = state.get("items", [])
            restored_any = False
            for it in items:
                if self._restore_window_and_taskbar(it):
                    restored_any = True
            if restored_any and items:
                try:
                    user32.SetForegroundWindow(int(items[0].get("hwnd", 0) or 0))
                except Exception:
                    pass
            self.hidden_states.pop(key, None)
            return None, restored_any

        hwnds = []
        if is_browser_program(p) and browser_group_toggle_enabled(p):
            hwnds = find_browser_group_windows(p)
        else:
            h = find_window_for_program(p)
            if h:
                hwnds = [h]

        if not hwnds:
            return None, False

        hidden_items = []
        for h in hwnds:
            it = self._hide_window_and_taskbar(h)
            if it:
                hidden_items.append(it)

        if not hidden_items:
            return None, False

        self.hidden_states[key] = {"items": hidden_items}
        return hidden_items[0]["hwnd"], True

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
    
        target_pf = _normalize_profile_text(p.get("profile_name", ""))
    
        for w in cands:
            hwnd = int(w.get("hwnd", 0) or 0)
            if not hwnd or hwnd in used or not is_hwnd_valid(hwnd):
                continue
    
            if target_pf:
                w_pf = _normalize_profile_text(w.get("profile", ""))
                w_cmd_pf = _normalize_profile_text(
                    parse_profile_from_cmdline(w.get("proc_name", ""), w.get("cmdline", []))
                )
                w_t_pf = _normalize_profile_text(
                    parse_profile_from_title(w.get("proc_name", ""), w.get("title", ""))
                )
                if not any(_profile_match(target_pf, x) for x in (w_pf, w_cmd_pf, w_t_pf) if x):
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

        if is_hide_action(p):
            hwnd, acted = self.toggle_hide_for_program(idx, p)
            if acted:
                if idx in self.hidden_states:
                    self.update_status(f"已隐藏：{p.get('name', '')}")
                else:
                    self.update_status(f"已恢复：{p.get('name', '')}")
            return

        mode = (p.get("match_mode", "title") or "title").lower()
        old_hwnd = int(p.get("bind_hwnd", 0) or 0)

        if mode == "hwnd":
            self.auto_bind_program_if_needed(idx, save=True)
            old_hwnd = int(p.get("bind_hwnd", 0) or 0)

        hwnd, acted = toggle_program(p)

        if is_browser_program(p) and not acted:
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

    def on_manual_add(self, _):
        dlg = wx.Dialog(self, title="手动添加", size=(800, 560))
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
        action = wx.Choice(panel, choices=["切换/启动", "隐藏/恢复（窗口+任务栏图标）"])
        action.SetStringSelection("切换/启动")
        mode = wx.Choice(panel, choices=["title", "profile", "hwnd"])
        mode.SetStringSelection("title")
        kw = wx.TextCtrl(panel)
        hwnd = wx.TextCtrl(panel, value="0")

        fallback_cb = wx.CheckBox(panel, label="浏览器找不到窗口时，允许按 EXE 兜底启动")
        fallback_cb.SetValue(False)

        group_toggle_cb = wx.CheckBox(panel, label="同 Profile 多窗口联动（含隐私窗口）")
        group_toggle_cb.SetValue(True)

        row("名称:", name)
        row("路径:", path)
        row("启动参数:", args)
        row("热键动作:", action)
        row("模式:", mode)
        row("关键词:", kw)
        row("HWND:", hwnd)
        s.Add(fallback_cb, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)
        s.Add(group_toggle_cb, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)

        b = wx.Button(panel, label="浏览exe")

        def refresh_browser_options():
            fake_program = {"path": path.GetValue().strip()}
            is_b = is_browser_program(fake_program)
            fallback_cb.Enable(is_b)
            group_toggle_cb.Enable(is_b)
            if not is_b:
                fallback_cb.SetValue(False)
                group_toggle_cb.SetValue(False)

        def browse(_e):
            fd = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if fd.ShowModal() == wx.ID_OK:
                pth = fd.GetPath()
                path.SetValue(pth)
                if not name.GetValue().strip():
                    name.SetValue(os.path.splitext(os.path.basename(pth))[0])
                refresh_browser_options()
            fd.Destroy()

        path.Bind(wx.EVT_TEXT, lambda e: (refresh_browser_options(), e.Skip()))
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

        refresh_browser_options()

        if dlg.ShowModal() == wx.ID_OK:
            hwnd_val = 0
            try:
                hwnd_val = int((hwnd.GetValue() or "").strip() or "0")
            except Exception:
                hwnd_val = 0

            p = {
                "name": name.GetValue().strip(),
                "path": path.GetValue().strip(),
                "args": args.GetValue().strip(),
                "hotkey": "",
                "window_keyword": kw.GetValue().strip(),
                "match_mode": mode.GetStringSelection() or "title",
                "bind_hwnd": hwnd_val,
                "profile_name": kw.GetValue().strip() if (mode.GetStringSelection() == "profile") else "",
                "title_sig": "",
                "browser_fallback_exe": bool(fallback_cb.GetValue()),
                "browser_group_toggle": bool(group_toggle_cb.GetValue()),
                "hotkey_action": "hide" if action.GetStringSelection().startswith("隐藏/恢复") else "toggle",
            }

            if not p["name"] or not p["path"]:
                wx.MessageBox("名称和路径必填", "提示")
            elif not os.path.exists(p["path"]):
                wx.MessageBox("路径不存在", "错误")
            else:
                if not is_browser_program(p):
                    p["browser_fallback_exe"] = False
                    p["browser_group_toggle"] = False
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

        dlg = wx.Dialog(self, title="从运行窗口添加", size=(1040, 540))
        panel = wx.Panel(dlg)
        s = wx.BoxSizer(wx.VERTICAL)
        tip = wx.StaticText(panel, label="双击添加：浏览器默认用 hwnd 模式（可选 EXE 兜底 + 组联动）")
        s.Add(tip, 0, wx.ALL, 8)

        lc = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        lc.InsertColumn(0, "标题", width=330)
        lc.InsertColumn(1, "程序", width=110)
        lc.InsertColumn(2, "Profile", width=120)
        lc.InsertColumn(3, "HWND", width=90)
        lc.InsertColumn(4, "路径", width=330)

        rows = []
        for w in ws:
            pth = w["path"]
            name = os.path.splitext(os.path.basename(pth))[0]
            row = {
                "name": name,
                "path": pth,
                "title": w.get("title", ""),
                "profile": w.get("profile", ""),
                "proc_name": w.get("proc_name", ""),
                "hwnd": int(w.get("hwnd", 0)),
                "title_sig": w.get("title_sig", ""),
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
            fallback = False
            group_toggle = False

            if proc in BROWSER_SET:
                ok, m_profile = ask_profile_input(self, default_profile=profile)
                if ok and m_profile:
                    args = build_profile_args(proc, m_profile)
                    kw = m_profile
                    profile = m_profile

                ask = wx.MessageBox(
                    "浏览器找不到匹配窗口时，是否允许按 EXE 启动？",
                    "浏览器 EXE 兜底",
                    wx.YES_NO | wx.ICON_QUESTION,
                )
                fallback = ask == wx.YES

                ask2 = wx.MessageBox(
                    "是否启用“同 Profile 多窗口联动（含隐私窗口）”？",
                    "浏览器组联动",
                    wx.YES_NO | wx.ICON_QUESTION,
                )
                group_toggle = ask2 == wx.YES

            self.programs.append(
                {
                    "name": it["name"],
                    "path": it["path"],
                    "args": args,
                    "hotkey": "",
                    "window_keyword": kw,
                    "match_mode": mode,
                    "bind_hwnd": it["hwnd"] if mode == "hwnd" else 0,
                    "profile_name": profile,
                    "title_sig": it.get("title_sig", ""),
                    "browser_fallback_exe": fallback,
                    "browser_group_toggle": group_toggle,
                    "hotkey_action": "toggle",
                }
            )
            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()
            dlg.EndModal(wx.ID_OK)

        lc.Bind(wx.EVT_LIST_ITEM_ACTIVATED, on_dbl)
        s.Add(lc, 1, wx.ALL | wx.EXPAND, 8)

        close_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dlg.EndModal(wx.ID_CANCEL))
        s.Add(close_btn, 0, wx.ALL | wx.ALIGN_CENTER, 6)
        panel.SetSizer(s)

        dlg.ShowModal()
        dlg.Destroy()

    def on_delete(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择", "提示")
            return
        self.hidden_states.pop(int(idx), None)
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
        dlg = wx.Dialog(self, title="设置匹配", size=(680, 520))
        panel = wx.Panel(dlg)
        s = wx.BoxSizer(wx.VERTICAL)

        action = wx.Choice(panel, choices=["切换/启动", "隐藏/恢复（窗口+任务栏图标）"])
        action.SetStringSelection("隐藏/恢复（窗口+任务栏图标）" if is_hide_action(p) else "切换/启动")

        mode = wx.Choice(panel, choices=["title", "profile", "hwnd"])
        mode.SetStringSelection(p.get("match_mode", "title"))
        kw = wx.TextCtrl(panel, value=p.get("window_keyword", ""))
        hwnd = wx.TextCtrl(panel, value=str(int(p.get("bind_hwnd", 0) or 0)))
        prof = wx.TextCtrl(panel, value=p.get("profile_name", ""))
        tsig = wx.TextCtrl(panel, value=p.get("title_sig", ""))

        fallback_cb = wx.CheckBox(panel, label="浏览器找不到窗口时，允许 EXE 兜底启动")
        fallback_cb.SetValue(bool(p.get("browser_fallback_exe", False)))

        group_toggle_cb = wx.CheckBox(panel, label="同 Profile 多窗口联动（含隐私窗口）")
        group_toggle_cb.SetValue(bool(p.get("browser_group_toggle", True)))

        is_b = is_browser_program(p)
        fallback_cb.Enable(is_b)
        group_toggle_cb.Enable(is_b)

        for lab, ctrl in [
            ("热键动作:", action),
            ("模式:", mode),
            ("关键词:", kw),
            ("绑定HWND:", hwnd),
            ("Profile名:", prof),
            ("TitleSig:", tsig),
        ]:
            r = wx.BoxSizer(wx.HORIZONTAL)
            r.Add(wx.StaticText(panel, label=lab), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 6)
            r.Add(ctrl, 1, wx.ALL | wx.EXPAND, 6)
            s.Add(r, 0, wx.EXPAND)

        s.Add(fallback_cb, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)
        s.Add(group_toggle_cb, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)

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
            p["hotkey_action"] = "hide" if action.GetStringSelection().startswith("隐藏/恢复") else "toggle"
            p["match_mode"] = mode.GetStringSelection() or "title"
            p["window_keyword"] = kw.GetValue().strip()
            p["profile_name"] = prof.GetValue().strip()
            p["title_sig"] = tsig.GetValue().strip()
            try:
                p["bind_hwnd"] = int(hwnd.GetValue().strip() or "0")
            except Exception:
                p["bind_hwnd"] = 0

            if is_browser_program(p):
                p["browser_fallback_exe"] = bool(fallback_cb.GetValue())
                p["browser_group_toggle"] = bool(group_toggle_cb.GetValue())
            else:
                p["browser_fallback_exe"] = False
                p["browser_group_toggle"] = False

            self.persist()
            self.refresh_list()
            wx.MessageBox("已保存", "成功")
        dlg.Destroy()


class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
    
        start_in_tray = any(arg.lower() in ("--tray", "/tray") for arg in sys.argv[1:])
        if start_in_tray:
            self.frame.hide_to_tray()
        else:
            self.frame.Show()
    
        return True


if __name__ == "__main__":
    app = QuickLauncherApp()
    app.MainLoop()
