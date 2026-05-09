# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器
作者：laozi4vip
GitHub：https://github.com/laozi4vip/QuickLauncher
"""

__author__ = "laozi4vip"
__github__ = "https://github.com/laozi4vip/QuickLauncher"
__app_name__ = "QuickLauncher"
__version__ = "4.0"
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
import threading

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
MATCH_MODES = ["title", "profile", "hwnd", "program"]

# ---------------------------
# 配置路径
# ---------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

LAST_ACTIVE_PROFILE_HWND = {}
LAST_ACTIVE_MAX = 50  # [FIX] 限制最大条目数，防止长期运行泄漏

_browser_main_map_cache = {}
_browser_main_map_ts = 0.0
_browser_main_map_lock = threading.Lock()  # [FIX] 线程安全
CACHE_TTL = 300.0

# [FIX] 新增：枚举窗口结果缓存，防止单次热键处理中重复全量扫描
_enum_windows_cache = []
_enum_windows_cache_ts = 0.0
ENUM_WINDOWS_CACHE_TTL = 1.5  # 1.5秒内复用


def invalidate_enum_cache():
    """[FIX] 在隐藏/恢复窗口后主动失效缓存"""
    global _enum_windows_cache_ts
    _enum_windows_cache_ts = 0.0


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
                if p["match_mode"] not in MATCH_MODES:
                    p["match_mode"] = "title"
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
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"programs": programs, "autostart": autostart}, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print("save_config error:", e)


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
    try:
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        return buf.value or ""
    except Exception:
        return ""


def get_pid_from_hwnd(hwnd):
    try:
        pid = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value
    except Exception:
        return 0


def is_alt_tab_window(hwnd):
    try:
        if not user32.IsWindowVisible(hwnd):
            return False
        title = get_window_title(hwnd)
        if not title.strip():
            return False
        exstyle = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        if exstyle & WS_EX_TOOLWINDOW:
            return False
        return True
    except Exception:
        return False


def is_hwnd_valid(hwnd: int):
    try:
        return bool(hwnd) and bool(user32.IsWindow(hwnd))
    except Exception:
        return False


def get_proc_path_name_cmdline(pid):
    if not pid:
        return "", "", []
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
    """[FIX] 增加完整异常保护和有效性检查"""
    try:
        if not is_hwnd_valid(hwnd):
            return
        cur = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        if cur == exstyle:
            return
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, exstyle)
        user32.SetWindowPos(
            hwnd, 0, 0, 0, 0, 0,
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
    try:
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
    except Exception:
        pass
    return main_map


def get_cached_browser_main_map(max_age=CACHE_TTL):
    global _browser_main_map_cache, _browser_main_map_ts
    now = time.time()
    with _browser_main_map_lock:
        if now - _browser_main_map_ts < max_age and _browser_main_map_cache:
            return dict(_browser_main_map_cache)  # [FIX] 返回副本
    new_map = build_browser_main_proc_map()
    with _browser_main_map_lock:
        _browser_main_map_cache = new_map
        _browser_main_map_ts = time.time()
        return dict(new_map)


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
# 窗口枚举 [FIX] 增加结果缓存
# ---------------------------
def enum_visible_app_windows(max_cache_age=ENUM_WINDOWS_CACHE_TTL):
    global _enum_windows_cache, _enum_windows_cache_ts

    now = time.time()
    if max_cache_age > 0 and (now - _enum_windows_cache_ts) < max_cache_age and _enum_windows_cache:
        return list(_enum_windows_cache)

    windows = []
    # [FIX] 统一使用较长缓存，避免每次热键都扫描全部进程
    browser_main_map = get_cached_browser_main_map(max_age=30.0)

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def callback(hwnd, _):
        try:
            if not user32.IsWindow(hwnd):
                return True
            if not is_alt_tab_window(hwnd):
                return True

            pid = get_pid_from_hwnd(hwnd)
            if not pid:
                return True
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

    try:
        user32.EnumWindows(callback, 0)
    except Exception:
        pass

    _enum_windows_cache = windows
    _enum_windows_cache_ts = time.time()
    return list(windows)


def enum_windows_for_program(path, all_windows=None):
    """[FIX] 支持传入预扫描结果，避免重复枚举"""
    target = os.path.normcase(path or "")
    result = []
    all_ws = all_windows if all_windows is not None else enum_visible_app_windows()
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
        if not pid:
            return

        path, proc_name, cmdline = get_proc_path_name_cmdline(pid)
        exe = (os.path.basename(path) if path else proc_name or "").lower().replace(".exe", "")

        if exe not in BROWSER_SET:
            return

        title = get_window_title(hwnd).strip()
        browser_main_map = get_cached_browser_main_map(max_age=60.0)  # [FIX] 延长缓存

        profile = get_profile_by_pid_tree(pid, browser_main_map)
        if not profile:
            profile = guess_profile(proc_name, cmdline, title, pid)

        if profile:
            now = time.time()
            key = (exe, profile.strip().lower())
            LAST_ACTIVE_PROFILE_HWND[key] = (int(hwnd), now)

            # [FIX] 更积极的清理策略，限制最大50条
            if len(LAST_ACTIVE_PROFILE_HWND) > LAST_ACTIVE_MAX:
                sorted_keys = sorted(
                    LAST_ACTIVE_PROFILE_HWND.keys(),
                    key=lambda k: LAST_ACTIVE_PROFILE_HWND[k][1]
                )
                for k in sorted_keys[: len(sorted_keys) - LAST_ACTIVE_MAX // 2]:
                    LAST_ACTIVE_PROFILE_HWND.pop(k, None)

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


def browser_window_matches_all_configured(program, w):
    match_mode = (program.get("match_mode", "title") or "title").strip().lower()
    keyword = (program.get("window_keyword", "") or "").strip()
    profile_name = (program.get("profile_name", "") or "").strip()
    title_sig = (program.get("title_sig", "") or "").strip()
    bind_hwnd = int(program.get("bind_hwnd", 0) or 0)

    if match_mode == "program":
        return True

    w_hwnd = int(w.get("hwnd", 0) or 0)
    w_title = (w.get("title", "") or "").lower()
    w_sig = (w.get("title_sig", "") or "").lower()

    w_pf = _normalize_profile_text(w.get("profile", ""))
    w_cmd_pf = _normalize_profile_text(parse_profile_from_cmdline(w.get("proc_name", ""), w.get("cmdline", [])))
    w_t_pf = _normalize_profile_text(parse_profile_from_title(w.get("proc_name", ""), w.get("title", "")))

    checks = []

    if match_mode == "hwnd" and bind_hwnd > 0:
        checks.append(w_hwnd == bind_hwnd)
    elif match_mode == "title" and keyword:
        checks.append(keyword.lower() in w_title)
    elif match_mode == "profile":
        tp = _normalize_profile_text(profile_name or keyword)
        if tp:
            checks.append(any(_profile_match(tp, x) for x in (w_pf, w_cmd_pf, w_t_pf) if x))

    if keyword:
        if match_mode == "profile":
            tp_kw = _normalize_profile_text(keyword)
            checks.append(any(_profile_match(tp_kw, x) for x in (w_pf, w_cmd_pf, w_t_pf) if x))
        elif match_mode == "hwnd":
            checks.append(True)
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

    if match_mode == "program":
        hwnd = int(w.get("hwnd", 0) or 0)
        try:
            if hwnd == user32.GetForegroundWindow():
                score += 200
        except Exception:
            pass
        if hwnd and is_hwnd_valid(hwnd):
            try:
                if not user32.IsIconic(hwnd):
                    score += 120
            except Exception:
                pass
        score += 50
        return score

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
    try:
        if hwnd == user32.GetForegroundWindow():
            score += 20
    except Exception:
        pass
    if hwnd and is_hwnd_valid(hwnd):
        try:
            if not user32.IsIconic(hwnd):
                score += 10
        except Exception:
            pass

    return score


def find_window_for_program(program):
    path = program.get("path", "")
    if not path:
        return None

    match_mode = (program.get("match_mode", "title") or "title").strip().lower()

    if match_mode == "program":
        cands = enum_windows_for_program(path)
        if not cands:
            return None
        scored = []
        for w in cands:
            hwnd = int(w.get("hwnd", 0) or 0)
            if hwnd and is_hwnd_valid(hwnd):
                scored.append((score_window_for_program(program, w), hwnd))
        if not scored:
            return None
        scored.sort(key=lambda x: x[0], reverse=True)
        return scored[0][1]

    if is_browser_program(program) and (browser_group_toggle_enabled(program) or browser_fallback_enabled(program)):
        cands = enum_windows_for_program(path)
        valid = [int(w.get("hwnd", 0) or 0) for w in cands if is_hwnd_valid(int(w.get("hwnd", 0) or 0))]
        return valid[0] if valid else None

    bind_hwnd = int(program.get("bind_hwnd", 0) or 0)

    if match_mode == "hwnd" and bind_hwnd and is_hwnd_valid(bind_hwnd):
        pid = get_pid_from_hwnd(bind_hwnd)
        ppath, _, _ = get_proc_path_name_cmdline(pid)
        if os.path.normcase(ppath or "") == os.path.normcase(path or ""):
            return bind_hwnd

    candidates = enum_windows_for_program(path)
    if not candidates:
        return None

    if is_browser_program(program):
        strict = [x for x in candidates if browser_window_matches_all_configured(program, x)]
        if not strict:
            return None
        strict.sort(key=lambda w: score_window_for_program(program, w), reverse=True)
        return int(strict[0].get("hwnd", 0) or 0) or None

    scored = []
    for w in candidates:
        s = score_window_for_program(program, w)
        scored.append((s, int(w["hwnd"]), w))
    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best_hwnd, _ = scored[0]

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

    match_mode = (program.get("match_mode", "title") or "title").strip().lower()

    if match_mode == "program":
        matched = []
        for w in cands:
            hwnd = int(w.get("hwnd", 0) or 0)
            if hwnd and is_hwnd_valid(hwnd):
                matched.append(w)
    elif browser_group_toggle_enabled(program) or browser_fallback_enabled(program):
        matched = []
        for w in cands:
            hwnd = int(w.get("hwnd", 0) or 0)
            if hwnd and is_hwnd_valid(hwnd):
                matched.append(w)
    else:
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
        if hwnd and is_hwnd_valid(hwnd):
            try:
                if not user32.IsIconic(hwnd):
                    s += 15
            except Exception:
                pass
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
    try:
        subprocess.Popen(cmd)
    except Exception as e:
        wx.MessageBox(f"启动失败：{e}", "错误", wx.OK | wx.ICON_ERROR)
        return False
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


# ======================================================
# 第二部分（GUI 类 + 主入口）—— 续
# ======================================================

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


# ---------------------------
# 托盘图标
# ---------------------------
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
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK, self.on_left_dclick)

    def on_left_dclick(self, _):
        self.frame.show_main_window()

    def CreatePopupMenu(self):
        menu = wx.Menu()
        item_show = menu.Append(wx.ID_ANY, "打开主窗口")
        item_quit = menu.Append(wx.ID_ANY, "退出")
        self.Bind(wx.EVT_MENU, lambda e: self.frame.show_main_window(), item_show)
        self.Bind(wx.EVT_MENU, lambda e: self.frame.on_quit(e), item_quit)
        return menu


# ---------------------------
# 隐藏窗口管理器（hide 动作）
# ---------------------------
class HiddenWindowManager:
    """管理被 hide 动作隐藏的窗口，支持恢复"""

    def __init__(self):
        self._hidden = {}  # key -> list of (hwnd, original_exstyle)

    def hide_windows(self, key: str, hwnds: list):
        records = []
        for h in hwnds:
            if not is_hwnd_valid(h):
                continue
            try:
                orig_ex = user32.GetWindowLongW(h, GWL_EXSTYLE)
                user32.ShowWindow(h, SW_HIDE)
                # [FIX] 不批量发 SWP_FRAMECHANGED，仅在必要时修改 exstyle
                new_ex = (orig_ex | WS_EX_TOOLWINDOW) & ~WS_EX_APPWINDOW
                if new_ex != orig_ex:
                    user32.SetWindowLongW(h, GWL_EXSTYLE, new_ex)
                records.append((h, orig_ex))
            except Exception:
                pass
        if records:
            self._hidden[key] = records
            invalidate_enum_cache()
        return len(records) > 0

    def restore_windows(self, key: str):
        records = self._hidden.pop(key, [])
        if not records:
            return False
        last_valid = None
        for h, orig_ex in records:
            if not is_hwnd_valid(h):
                continue
            try:
                user32.SetWindowLongW(h, GWL_EXSTYLE, orig_ex)
                user32.ShowWindow(h, SW_SHOW)
                if user32.IsIconic(h):
                    user32.ShowWindow(h, SW_RESTORE)
                last_valid = h
            except Exception:
                pass
        # [FIX] 只对最后一个窗口发一次 FRAMECHANGED 通知
        if last_valid and is_hwnd_valid(last_valid):
            try:
                user32.SetWindowPos(
                    last_valid, 0, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
                )
                user32.SetForegroundWindow(last_valid)
            except Exception:
                pass
        invalidate_enum_cache()
        return True

    def is_hidden(self, key: str) -> bool:
        records = self._hidden.get(key, [])
        return any(is_hwnd_valid(h) for h, _ in records)

    def cleanup_stale(self):
        """清理已失效的句柄"""
        stale_keys = []
        for key, records in self._hidden.items():
            if not any(is_hwnd_valid(h) for h, _ in records):
                stale_keys.append(key)
        for k in stale_keys:
            self._hidden.pop(k, None)

    def restore_all(self):
        for key in list(self._hidden.keys()):
            self.restore_windows(key)


# ---------------------------
# 主窗口
# ---------------------------
class QuickLauncherFrame(wx.Frame):

    # [FIX] 热键防抖间隔（秒）
    HOTKEY_DEBOUNCE = 0.35

    def __init__(self, start_hidden=False):
        super().__init__(
            None,
            title=f"{__app_name__} v{__version__} - {__description__}",
            size=(900, 520),
            style=wx.DEFAULT_FRAME_STYLE,
        )
        self.SetMinSize((720, 400))
        self.Centre()

        # 数据
        config = load_config()
        self.programs = config["programs"]
        self.autostart = config["autostart"]

        # 热键管理
        self._hotkey_ids = {}       # id -> (mod, vk, normalized)
        self._hotkey_id_counter = 0x7000

        # [FIX] 热键防抖时间戳
        self._last_hotkey_ts = {}   # hotkey_id -> float

        # 隐藏窗口管理器
        self.hidden_mgr = HiddenWindowManager()

        # [FIX] 定时器 ID
        self._profile_timer_id = wx.NewIdRef()
        self._cleanup_timer_id = wx.NewIdRef()
        self._timers_started = False

        # 托盘
        self.taskbar = QuickLauncherTaskBar(self)

        # 构建 UI
        self._build_ui()

        # 注册热键
        self._register_all_hotkeys()

        # [FIX] 使用 wx.Timer 代替 threading.Timer / time.sleep
        self._start_timers()

        # 绑定事件
        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.Bind(wx.EVT_HOTKEY, self.on_hotkey)

        if start_hidden:
            self.Hide()
        else:
            self.Show()

    # ---------- UI 构建 ----------

    def _build_ui(self):
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # 工具栏按钮
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_add = wx.Button(panel, label="➕ 添加")
        self.btn_edit = wx.Button(panel, label="✏️ 编辑")
        self.btn_del = wx.Button(panel, label="🗑️ 删除")
        self.btn_up = wx.Button(panel, label="⬆ 上移")
        self.btn_down = wx.Button(panel, label="⬇ 下移")
        self.cb_autostart = wx.CheckBox(panel, label="开机自启")
        self.cb_autostart.SetValue(self.autostart)

        for b in (self.btn_add, self.btn_edit, self.btn_del, self.btn_up, self.btn_down):
            btn_sizer.Add(b, 0, wx.ALL, 3)
        btn_sizer.AddStretchSpacer()
        btn_sizer.Add(self.cb_autostart, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 3)

        main_sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 列表
        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        cols = [
            ("名称", 120), ("程序路径", 200), ("参数", 100),
            ("快捷键", 100), ("匹配模式", 80), ("窗口关键字", 100),
            ("Profile", 80), ("动作", 60),
        ]
        for i, (name, w) in enumerate(cols):
            self.list_ctrl.InsertColumn(i, name, width=w)
        main_sizer.Add(self.list_ctrl, 1, wx.EXPAND | wx.ALL, 5)

        panel.SetSizer(main_sizer)

        # 事件绑定
        self.btn_add.Bind(wx.EVT_BUTTON, self.on_add)
        self.btn_edit.Bind(wx.EVT_BUTTON, self.on_edit)
        self.btn_del.Bind(wx.EVT_BUTTON, self.on_delete)
        self.btn_up.Bind(wx.EVT_BUTTON, self.on_move_up)
        self.btn_down.Bind(wx.EVT_BUTTON, self.on_move_down)
        self.cb_autostart.Bind(wx.EVT_CHECKBOX, self.on_autostart_toggle)
        self.list_ctrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.on_edit)

        self._refresh_list()

    def _refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for i, p in enumerate(self.programs):
            idx = self.list_ctrl.InsertItem(i, p.get("name", ""))
            self.list_ctrl.SetItem(idx, 1, p.get("path", ""))
            self.list_ctrl.SetItem(idx, 2, p.get("args", ""))
            self.list_ctrl.SetItem(idx, 3, p.get("hotkey", ""))
            self.list_ctrl.SetItem(idx, 4, p.get("match_mode", "title"))
            self.list_ctrl.SetItem(idx, 5, p.get("window_keyword", ""))
            self.list_ctrl.SetItem(idx, 6, p.get("profile_name", ""))
            action = p.get("hotkey_action", "toggle")
            self.list_ctrl.SetItem(idx, 7, action)

    # ---------- 定时器 [FIX] ----------

    def _start_timers(self):
        """使用 wx.Timer 替代阻塞式 sleep"""
        if self._timers_started:
            return

        # Profile 活跃窗口缓存更新 - 每 3 秒
        self._profile_timer = wx.Timer(self, self._profile_timer_id)
        self.Bind(wx.EVT_TIMER, self._on_profile_timer, id=self._profile_timer_id)
        self._profile_timer.Start(3000)

        # 失效隐藏句柄清理 - 每 30 秒
        self._cleanup_timer = wx.Timer(self, self._cleanup_timer_id)
        self.Bind(wx.EVT_TIMER, self._on_cleanup_timer, id=self._cleanup_timer_id)
        self._cleanup_timer.Start(30000)

        self._timers_started = True

    def _stop_timers(self):
        """[FIX] 退出时必须停止定时器"""
        if hasattr(self, "_profile_timer"):
            self._profile_timer.Stop()
        if hasattr(self, "_cleanup_timer"):
            self._cleanup_timer.Stop()
        self._timers_started = False

    def _on_profile_timer(self, _):
        try:
            update_last_active_cache()
        except Exception:
            pass

    def _on_cleanup_timer(self, _):
        try:
            self.hidden_mgr.cleanup_stale()
        except Exception:
            pass

    # ---------- 热键注册 ----------

    def _next_hotkey_id(self):
        self._hotkey_id_counter += 1
        return self._hotkey_id_counter

    def _register_all_hotkeys(self):
        self._unregister_all_hotkeys()
        for i, p in enumerate(self.programs):
            hk = p.get("hotkey", "")
            if not hk:
                continue
            try:
                mod, vk, norm = hotkey_to_mod_vk(hk)
                hid = self._next_hotkey_id()
                ok = user32.RegisterHotKey(self.GetHandle(), hid, mod, vk)
                if ok:
                    self._hotkey_ids[hid] = (mod, vk, norm, i)
                else:
                    print(f"[WARN] 热键注册失败: {norm} (可能被占用)")
            except Exception as e:
                print(f"[WARN] 热键解析失败: {hk} -> {e}")

    def _unregister_all_hotkeys(self):
        for hid in list(self._hotkey_ids.keys()):
            try:
                user32.UnregisterHotKey(self.GetHandle(), hid)
            except Exception:
                pass
        self._hotkey_ids.clear()

    # ---------- 热键处理 [FIX: 防抖 + 无阻塞] ----------

    def on_hotkey(self, event):
        hid = event.GetId()
        info = self._hotkey_ids.get(hid)
        if not info:
            return
        _, _, norm, prog_idx = info

        # [FIX] 防抖：忽略过快的重复触发
        now = time.time()
        last = self._last_hotkey_ts.get(hid, 0.0)
        if now - last < self.HOTKEY_DEBOUNCE:
            return
        self._last_hotkey_ts[hid] = now

        if prog_idx < 0 or prog_idx >= len(self.programs):
            return

        program = self.programs[prog_idx]

        # [FIX] 用 wx.CallAfter 确保不阻塞消息泵
        wx.CallAfter(self._handle_hotkey_action, program)

    def _handle_hotkey_action(self, program):
        """实际执行热键动作，在主线程事件循环中安全执行"""
        try:
            action = (program.get("hotkey_action", "toggle") or "toggle").strip().lower()

            if action == "hide":
                self._do_hide_toggle(program)
            else:
                self._do_toggle(program)
        except Exception as e:
            print(f"[ERROR] hotkey action: {e}")
        finally:
            # [FIX] 操作完成后失效枚举缓存
            invalidate_enum_cache()

    def _make_hide_key(self, program):
        """为 hide 动作生成唯一 key"""
        path = os.path.normcase(program.get("path", ""))
        profile = (program.get("profile_name", "") or "").strip().lower()
        keyword = (program.get("window_keyword", "") or "").strip().lower()
        return f"{path}|{profile}|{keyword}"

    def _do_hide_toggle(self, program):
        key = self._make_hide_key(program)

        # 如果已隐藏则恢复
        if self.hidden_mgr.is_hidden(key):
            self.hidden_mgr.restore_windows(key)
            return

        # 否则查找并隐藏
        if is_browser_program(program) and browser_group_toggle_enabled(program):
            group = find_browser_group_windows(program)
            if group:
                self.hidden_mgr.hide_windows(key, group)
                return
        else:
            hwnd = find_window_for_program(program)
            if hwnd and is_hwnd_valid(hwnd):
                self.hidden_mgr.hide_windows(key, [hwnd])
                return

        # 没找到窗口，尝试启动
        path = program.get("path", "")
        args = program.get("args", "")
        if path:
            launch_program_by_path(path, args)

    def _do_toggle(self, program):
        toggle_program(program)

    # ---------- 列表操作 ----------

    def _get_selected_index(self):
        return self.list_ctrl.GetFirstSelected()

    def on_add(self, _):
        dlg = ProgramEditDialog(self, program=None)
        if dlg.ShowModal() == wx.ID_OK:
            new_prog = dlg.get_program()
            self.programs.append(new_prog)
            self._save_and_refresh()
        dlg.Destroy()

    def on_edit(self, _):
        idx = self._get_selected_index()
        if idx < 0:
            wx.MessageBox("请先选择一项", "提示", wx.OK | wx.ICON_INFORMATION)
            return
        dlg = ProgramEditDialog(self, program=self.programs[idx])
        if dlg.ShowModal() == wx.ID_OK:
            self.programs[idx] = dlg.get_program()
            self._save_and_refresh()
        dlg.Destroy()

    def on_delete(self, _):
        idx = self._get_selected_index()
        if idx < 0:
            wx.MessageBox("请先选择一项", "提示", wx.OK | wx.ICON_INFORMATION)
            return
        name = self.programs[idx].get("name", "(未命名)")
        if wx.MessageBox(f"确认删除「{name}」？", "确认", wx.YES_NO | wx.ICON_QUESTION) == wx.YES:
            self.programs.pop(idx)
            self._save_and_refresh()

    def on_move_up(self, _):
        idx = self._get_selected_index()
        if idx <= 0:
            return
        self.programs[idx - 1], self.programs[idx] = self.programs[idx], self.programs[idx - 1]
        self._save_and_refresh()
        self.list_ctrl.Select(idx - 1)

    def on_move_down(self, _):
        idx = self._get_selected_index()
        if idx < 0 or idx >= len(self.programs) - 1:
            return
        self.programs[idx], self.programs[idx + 1] = self.programs[idx + 1], self.programs[idx]
        self._save_and_refresh()
        self.list_ctrl.Select(idx + 1)

    def on_autostart_toggle(self, _):
        self.autostart = self.cb_autostart.GetValue()
        set_autostart(self.autostart)
        self._save_config_only()

    def _save_and_refresh(self):
        save_config(self.programs, self.autostart)
        self._refresh_list()
        self._register_all_hotkeys()

    def _save_config_only(self):
        save_config(self.programs, self.autostart)

    # ---------- 窗口显示/隐藏 ----------

    def show_main_window(self):
        if self.IsIconized():
            self.Iconize(False)
        self.Show()
        self.Raise()

    def on_close(self, event):
        """关闭窗口时最小化到托盘"""
        if event.CanVeto():
            event.Veto()
            self.Hide()
        else:
            self._do_quit()

    def on_quit(self, _):
        dlg = wx.MessageDialog(
            self, "确定退出 QuickLauncher？", "退出确认",
            wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION,
        )
        if dlg.ShowModal() == wx.YES:
            self._do_quit()
        dlg.Destroy()

    def _do_quit(self):
        """[FIX] 安全退出：先停定时器、注销热键、恢复隐藏窗口"""
        try:
            self._stop_timers()
        except Exception:
            pass
        try:
            self._unregister_all_hotkeys()
        except Exception:
            pass
        try:
            self.hidden_mgr.restore_all()
        except Exception:
            pass
        try:
            self.taskbar.RemoveIcon()
            self.taskbar.Destroy()
        except Exception:
            pass
        self.Destroy()
        wx.GetApp().ExitMainLoop()


# ---------------------------
# 程序编辑对话框
# ---------------------------
class ProgramEditDialog(wx.Dialog):
    def __init__(self, parent, program=None):
        super().__init__(
            parent,
            title="编辑程序" if program else "添加程序",
            size=(560, 520),
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
        )
        self.program = program or {}
        self.Centre()
        self._build_ui()

    def _build_ui(self):
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)
        grid = wx.FlexGridSizer(cols=3, hgap=8, vgap=8)
        grid.AddGrowableCol(1, 1)

        # 名称
        grid.Add(wx.StaticText(panel, label="名称："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.txt_name = wx.TextCtrl(panel, value=self.program.get("name", ""))
        grid.Add(self.txt_name, 1, wx.EXPAND)
        grid.AddSpacer(0)

        # 路径
        grid.Add(wx.StaticText(panel, label="程序路径："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.txt_path = wx.TextCtrl(panel, value=self.program.get("path", ""))
        grid.Add(self.txt_path, 1, wx.EXPAND)
        btn_browse = wx.Button(panel, label="浏览…")
        btn_browse.Bind(wx.EVT_BUTTON, self.on_browse)
        grid.Add(btn_browse, 0)

        # 参数
        grid.Add(wx.StaticText(panel, label="启动参数："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.txt_args = wx.TextCtrl(panel, value=self.program.get("args", ""))
        grid.Add(self.txt_args, 1, wx.EXPAND)
        grid.AddSpacer(0)

        # 快捷键
        grid.Add(wx.StaticText(panel, label="快捷键："), 0, wx.ALIGN_CENTER_VERTICAL)
        hk_row = wx.BoxSizer(wx.HORIZONTAL)
        self.txt_hotkey = wx.TextCtrl(
            panel, value=self.program.get("hotkey", ""), style=wx.TE_READONLY
        )
        hk_row.Add(self.txt_hotkey, 1, wx.EXPAND)
        btn_capture = wx.Button(panel, label="录入")
        btn_capture.Bind(wx.EVT_BUTTON, self.on_capture_hotkey)
        hk_row.Add(btn_capture, 0, wx.LEFT, 4)
        grid.Add(hk_row, 1, wx.EXPAND)
        grid.AddSpacer(0)

        # 匹配模式
        grid.Add(wx.StaticText(panel, label="匹配模式："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.ch_mode = wx.Choice(panel, choices=MATCH_MODES)
        cur_mode = self.program.get("match_mode", "title")
        if cur_mode in MATCH_MODES:
            self.ch_mode.SetSelection(MATCH_MODES.index(cur_mode))
        else:
            self.ch_mode.SetSelection(0)
        grid.Add(self.ch_mode, 1, wx.EXPAND)
        grid.AddSpacer(0)

        # 窗口关键字
        grid.Add(wx.StaticText(panel, label="窗口关键字："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.txt_keyword = wx.TextCtrl(panel, value=self.program.get("window_keyword", ""))
        grid.Add(self.txt_keyword, 1, wx.EXPAND)
        grid.AddSpacer(0)

        # Profile
        grid.Add(wx.StaticText(panel, label="Profile："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.txt_profile = wx.TextCtrl(panel, value=self.program.get("profile_name", ""))
        grid.Add(self.txt_profile, 1, wx.EXPAND)
        grid.AddSpacer(0)

        # 动作
        grid.Add(wx.StaticText(panel, label="热键动作："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.ch_action = wx.Choice(panel, choices=["toggle", "hide"])
        cur_action = self.program.get("hotkey_action", "toggle")
        self.ch_action.SetSelection(1 if cur_action == "hide" else 0)
        grid.Add(self.ch_action, 1, wx.EXPAND)
        grid.AddSpacer(0)

        sizer.Add(grid, 0, wx.EXPAND | wx.ALL, 12)

        # 浏览器选项
        browser_box = wx.StaticBox(panel, label="浏览器选项")
        bsizer = wx.StaticBoxSizer(browser_box, wx.VERTICAL)
        self.cb_fallback = wx.CheckBox(panel, label="未找到窗口时启动新实例")
        self.cb_fallback.SetValue(self.program.get("browser_fallback_exe", False))
        bsizer.Add(self.cb_fallback, 0, wx.ALL, 4)
        self.cb_group = wx.CheckBox(panel, label="按组切换所有同 exe 窗口")
        self.cb_group.SetValue(self.program.get("browser_group_toggle", True))
        bsizer.Add(self.cb_group, 0, wx.ALL, 4)
        sizer.Add(bsizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

        # 按钮
        btn_sizer = self.CreateStdDialogButtonSizer(wx.OK | wx.CANCEL)
        sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 12)

        panel.SetSizer(sizer)

    def on_browse(self, _):
        dlg = wx.FileDialog(
            self, "选择程序", wildcard="可执行文件 (*.exe)|*.exe|所有文件|*.*",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
        )
        if dlg.ShowModal() == wx.ID_OK:
            self.txt_path.SetValue(dlg.GetPath())
            if not self.txt_name.GetValue().strip():
                base = os.path.splitext(os.path.basename(dlg.GetPath()))[0]
                self.txt_name.SetValue(base)
        dlg.Destroy()

    def on_capture_hotkey(self, _):
        dlg = HotkeyCaptureDialog(self, self.txt_hotkey.GetValue())
        if dlg.ShowModal() == wx.ID_OK:
            self.txt_hotkey.SetValue(dlg.captured_hotkey)
        dlg.Destroy()

    def get_program(self) -> dict:
        mode = MATCH_MODES[self.ch_mode.GetSelection()]
        action = "hide" if self.ch_action.GetSelection() == 1 else "toggle"
        return {
            "name": self.txt_name.GetValue().strip(),
            "path": self.txt_path.GetValue().strip(),
            "args": self.txt_args.GetValue().strip(),
            "hotkey": normalize_hotkey(self.txt_hotkey.GetValue()),
            "window_keyword": self.txt_keyword.GetValue().strip(),
            "match_mode": mode,
            "bind_hwnd": int(self.program.get("bind_hwnd", 0) or 0),
            "profile_name": self.txt_profile.GetValue().strip(),
            "title_sig": self.program.get("title_sig", ""),
            "browser_fallback_exe": self.cb_fallback.GetValue(),
            "browser_group_toggle": self.cb_group.GetValue(),
            "hotkey_action": action,
        }


# ---------------------------
# 主入口
# ---------------------------
def main():
    start_in_tray = "--tray" in sys.argv

    app = wx.App(False)

    # 单实例检测
    instance_name = f"{__app_name__}_SingleInstance_Mutex"
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, instance_name)
    last_err = ctypes.windll.kernel32.GetLastError()
    if last_err == 183:  # ERROR_ALREADY_EXISTS
        wx.MessageBox(
            f"{__app_name__} 已在运行中。",
            "提示", wx.OK | wx.ICON_INFORMATION,
        )
        return

    frame = QuickLauncherFrame(start_hidden=start_in_tray)
    app.MainLoop()


if __name__ == "__main__":
    main()
