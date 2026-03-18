# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器（浏览器多用户稳定匹配增强版）
修复重点：
1) 浏览器窗口匹配支持 profile（cmdline + 标题兜底推断）
2) 从运行窗口添加时，浏览器 profile 可手动确认/修正，避免“都一样”
3) 浏览器新增项自动补启动参数（profile 参数）
4) 兼容旧配置（无 match_mode 时沿用 title contains 逻辑）
"""

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
                p.setdefault("match_mode", "title")  # title / profile
            return {"programs": programs, "autostart": autostart}
        except Exception as e:
            print("load_config error:", e)

    return default_data


def save_config(programs, autostart):
    data = {"programs": programs, "autostart": autostart}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


# ---------------------------
# 热键工具
# ---------------------------
def normalize_hotkey(hotkey: str) -> str:
    if not hotkey:
        return ""
    return hotkey.strip().lower().replace(" ", "")


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
    """返回 profile 名（如 Profile 1、Default、Work），拿不到返回空"""
    name = (proc_name or "").lower().replace(".exe", "")
    args = cmdline_list or []

    def value_after(prefix):
        # 支持 --xxx=yyy 或 --xxx yyy
        for i, a in enumerate(args):
            la = a.lower()
            if la.startswith(prefix + "="):
                return a.split("=", 1)[1].strip().strip('"')
            if la == prefix and i + 1 < len(args):
                return (args[i + 1] or "").strip().strip('"')
        return ""

    # Chrome / Edge / Chromium 系
    if name in ("chrome", "msedge", "chromium", "brave"):
        v = value_after("--profile-directory")
        if v:
            return v
        u = value_after("--user-data-dir")
        if u:
            return os.path.basename(u.rstrip("\\/"))
        return ""

    # Firefox
    if name == "firefox":
        for i, a in enumerate(args):
            la = a.lower()
            if la in ("-p", "--profile") and i + 1 < len(args):
                return (args[i + 1] or "").strip().strip('"')
            if la == "-profile" and i + 1 < len(args):
                p = (args[i + 1] or "").strip().strip('"')
                return os.path.basename(p.rstrip("\\/"))
        return ""

    return ""


def parse_profile_from_title(proc_name: str, title: str):
    """从标题兜底猜 profile（不是100%准确，但能修正很多“都一样”的情况）"""
    name = (proc_name or "").lower().replace(".exe", "")
    t = (title or "").strip()

    if not t:
        return ""

    # 常见：Profile 1 / Profile 2 / Default
    m = re.search(r"\b(Profile\s*\d+|Default)\b", t, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # Edge 常见：Personal / Work
    if name == "msedge":
        m2 = re.search(r"\b(Personal|Work|Guest)\b", t, re.IGNORECASE)
        if m2:
            return m2.group(1).strip()

    # Firefox 也可能在标题出现 profile 词
    if name == "firefox":
        m3 = re.search(r"\b(Profile\s*\d+|Default)\b", t, re.IGNORECASE)
        if m3:
            return m3.group(1).strip()

    return ""


def guess_profile(proc_name: str, cmdline_list, title: str):
    p1 = parse_profile_from_cmdline(proc_name, cmdline_list)
    if p1:
        return p1
    return parse_profile_from_title(proc_name, title)


def enum_visible_app_windows():
    windows = []

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
            profile = guess_profile(proc_name, cmdline, title)

            windows.append({
                "hwnd": int(hwnd),
                "pid": int(pid),
                "title": title,
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
        if not wpath:
            continue
        if wpath == target:
            result.append((w["hwnd"], w["title"], w["profile"], w["pid"]))

    if not result and path:
        exe_name = os.path.basename(path).lower().replace(".exe", "")
        for w in all_ws:
            pn = (w.get("proc_name", "") or "").lower().replace(".exe", "")
            if pn == exe_name:
                result.append((w["hwnd"], w["title"], w["profile"], w["pid"]))

    return result


def find_window_for_program(program):
    path = program.get("path", "")
    if not path:
        return None

    match_mode = (program.get("match_mode", "title") or "title").strip().lower()
    keyword = (program.get("window_keyword", "") or "").strip().lower()

    candidates = enum_windows_for_program(path)

    if keyword:
        if match_mode == "profile":
            filtered = []
            for h, t, prof, pid in candidates:
                p = (prof or "").strip().lower()
                if p == keyword or (keyword in p):
                    filtered.append((h, t, prof, pid))
            candidates = filtered
        else:
            candidates = [(h, t, prof, pid) for (h, t, prof, pid) in candidates if keyword in (t or "").lower()]

    if not candidates:
        return None

    fg = user32.GetForegroundWindow()
    for hwnd, *_ in candidates:
        if hwnd == fg:
            return hwnd

    for hwnd, *_ in candidates:
        if not user32.IsIconic(hwnd):
            return hwnd

    return candidates[0][0]


def toggle_program(program):
    path = program.get("path", "")
    args = program.get("args", "")
    if not path:
        return

    hwnd = find_window_for_program(program)

    if hwnd:
        fg = user32.GetForegroundWindow()
        if hwnd == fg:
            user32.ShowWindow(hwnd, SW_MINIMIZE)
        else:
            if user32.IsIconic(hwnd):
                user32.ShowWindow(hwnd, SW_RESTORE)
            user32.SetForegroundWindow(hwnd)
    else:
        if not os.path.exists(path):
            wx.MessageBox(f"程序不存在：\n{path}", "错误", wx.OK | wx.ICON_ERROR)
            return

        cmd = [path]
        if args.strip():
            try:
                cmd.extend(shlex.split(args.strip(), posix=False))
            except Exception:
                cmd.extend(args.strip().split())

        subprocess.Popen(cmd)


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


def ask_profile_input(parent, default_profile=""):
    dlg = wx.TextEntryDialog(
        parent,
        "请输入浏览器 Profile（例如：Profile 1 / Default / Work）",
        "确认 Profile",
        value=default_profile or ""
    )
    ret = dlg.ShowModal()
    val = dlg.GetValue().strip() if ret == wx.ID_OK else ""
    dlg.Destroy()
    return ret == wx.ID_OK, val


# ---------------------------
# 程序搜索
# ---------------------------
def collect_search_dirs():
    dirs = []
    for env_key in ["ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"]:
        v = os.environ.get(env_key, "")
        if v and os.path.isdir(v):
            dirs.append(v)

    lap = os.environ.get("LOCALAPPDATA", "")
    if lap:
        p = os.path.join(lap, "Programs")
        if os.path.isdir(p):
            dirs.append(p)

    for p in os.environ.get("PATH", "").split(";"):
        p = p.strip().strip('"')
        if p and os.path.isdir(p):
            dirs.append(p)

    seen = set()
    result = []
    for d in dirs:
        ld = d.lower()
        if ld not in seen:
            seen.add(ld)
            result.append(d)
    return result


def search_executables(query: str, max_results=300):
    query = (query or "").strip().lower()
    if not query:
        return []

    search_dirs = collect_search_dirs()
    results = []
    seen_path = set()

    for base in search_dirs:
        base_lower = base.lower()
        if "windows" in base_lower and "system32" in base_lower:
            continue

        for root, dirnames, filenames in os.walk(base):
            rel = os.path.relpath(root, base)
            depth = 0 if rel == "." else rel.count(os.sep) + 1
            if depth > 4:
                dirnames[:] = []
                continue

            for fn in filenames:
                if not fn.lower().endswith(".exe"):
                    continue
                if query not in fn.lower():
                    continue

                full = os.path.join(root, fn)
                lf = full.lower()
                if lf in seen_path:
                    continue
                seen_path.add(lf)
                results.append({"name": os.path.splitext(fn)[0], "path": full})
                if len(results) >= max_results:
                    return results

    results.sort(key=lambda x: (len(x["path"]), x["name"].lower()))
    return results


class ProgramSearchDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title="搜索程序并添加", size=(760, 520))
        self.selected_item = None
        self.results_data = []

        panel = wx.Panel(self)
        root = wx.BoxSizer(wx.VERTICAL)

        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(wx.StaticText(panel, label="搜索关键词（程序名）:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 6)
        self.query_ctrl = wx.TextCtrl(panel, size=(260, -1))
        self.query_ctrl.SetHint("例如：chrome / firefox / code / wechat")
        row.Add(self.query_ctrl, 0, wx.ALL, 6)

        self.search_btn = wx.Button(panel, label="开始搜索")
        row.Add(self.search_btn, 0, wx.ALL, 6)

        tips = wx.StaticText(panel, label="提示：搜索范围为常见安装目录，结果可双击直接添加")
        tips.SetForegroundColour(wx.Colour(100, 100, 100))
        row.Add(tips, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 6)
        root.Add(row, 0, wx.EXPAND)

        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "程序名", width=160)
        self.list_ctrl.InsertColumn(1, "路径", width=560)
        root.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        btn_row.Add(wx.Button(panel, wx.ID_OK, "添加所选"), 0, wx.ALL, 6)
        btn_row.Add(wx.Button(panel, wx.ID_CANCEL, "取消"), 0, wx.ALL, 6)
        root.Add(btn_row, 0, wx.ALIGN_CENTER)

        panel.SetSizer(root)

        self.search_btn.Bind(wx.EVT_BUTTON, self.on_search)
        self.query_ctrl.Bind(wx.EVT_TEXT_ENTER, self.on_search)
        self.list_ctrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.on_dbl_click)

    def on_search(self, _):
        q = self.query_ctrl.GetValue().strip()
        if not q:
            wx.MessageBox("请输入搜索关键词", "提示")
            return

        wx.BeginBusyCursor()
        try:
            results = search_executables(q, max_results=400)
        finally:
            if wx.IsBusy():
                wx.EndBusyCursor()

        self.results_data = results
        self.list_ctrl.DeleteAllItems()
        for item in results:
            self.list_ctrl.Append([item["name"], item["path"]])

        if not results:
            wx.MessageBox("未找到匹配程序，请尝试更短关键词", "提示")

    def on_dbl_click(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if 0 <= idx < len(self.results_data):
            self.selected_item = self.results_data[idx]
            self.EndModal(wx.ID_OK)

    def get_selected(self):
        idx = self.list_ctrl.GetFirstSelected()
        if 0 <= idx < len(self.results_data):
            return self.results_data[idx]
        return self.selected_item


# ---------------------------
# 快捷键捕获对话框
# ---------------------------
class HotkeyCaptureDialog(wx.Dialog):
    def __init__(self, parent, current_hotkey=""):
        super().__init__(parent, title="设置快捷键（按键自动录入）", size=(500, 250))
        self.captured_hotkey = normalize_hotkey(current_hotkey)

        panel = wx.Panel(self)
        s = wx.BoxSizer(wx.VERTICAL)

        tips = wx.StaticText(
            panel,
            label=(
                "请直接按下目标快捷键（例如 Ctrl+1、Alt+Q、Ctrl+Shift+F2）\n"
                "单键仅支持 F1-F12\n"
                "Esc 取消，Backspace 清空"
            ),
        )
        tips.SetForegroundColour(wx.Colour(90, 90, 90))
        s.Add(tips, 0, wx.ALL | wx.EXPAND, 10)

        self.hk_show = wx.TextCtrl(panel, value=self.captured_hotkey, style=wx.TE_READONLY)
        s.Add(self.hk_show, 0, wx.ALL | wx.EXPAND, 10)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        ok_btn = wx.Button(panel, wx.ID_OK, "确定")
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        clear_btn = wx.Button(panel, wx.ID_ANY, "清空")
        btn_row.Add(ok_btn, 0, wx.ALL, 5)
        btn_row.Add(cancel_btn, 0, wx.ALL, 5)
        btn_row.Add(clear_btn, 0, wx.ALL, 5)
        s.Add(btn_row, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        panel.SetSizer(s)

        self.Bind(wx.EVT_CHAR_HOOK, self.on_char_hook)
        clear_btn.Bind(wx.EVT_BUTTON, self.on_clear)

    def on_clear(self, _):
        self.captured_hotkey = ""
        self.hk_show.SetValue("")

    def on_char_hook(self, event):
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
            _, _, normalized = hotkey_to_mod_vk(hk)
            self.captured_hotkey = normalized
            self.hk_show.SetValue(normalized)
        except ValueError:
            wx.Bell()


# ---------------------------
# 托盘
# ---------------------------
class QuickLauncherTaskBar(wx.adv.TaskBarIcon):
    def __init__(self, frame):
        super().__init__()
        self.frame = frame
        bmp = wx.ArtProvider.GetBitmap(wx.ART_EXECUTABLE_FILE, wx.ART_OTHER, (16, 16))
        icon = wx.Icon()
        icon.CopyFromBitmap(bmp)
        self.SetIcon(icon, "QuickLauncher")
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK, self.on_show)

    def CreatePopupMenu(self):
        menu = wx.Menu()
        show_item = menu.Append(wx.ID_ANY, "显示主窗口")
        hide_item = menu.Append(wx.ID_ANY, "隐藏到托盘")
        menu.AppendSeparator()
        exit_item = menu.Append(wx.ID_EXIT, "退出")
        self.Bind(wx.EVT_MENU, self.on_show, show_item)
        self.Bind(wx.EVT_MENU, self.on_hide, hide_item)
        self.Bind(wx.EVT_MENU, self.on_exit, exit_item)
        return menu

    def on_show(self, _):
        self.frame.show_from_tray()

    def on_hide(self, _):
        self.frame.hide_to_tray()

    def on_exit(self, _):
        self.frame.exit_app()


# ---------------------------
# 主窗口
# ---------------------------
class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(980, 580))
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
        self.update_status("运行中... | 全局热键已就绪")

        self.Bind(wx.EVT_CLOSE, self.on_close)

    def init_ui(self):
        panel = wx.Panel(self)
        root = wx.BoxSizer(wx.VERTICAL)

        title = wx.StaticText(panel, label="程序列表")
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        root.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        desc = wx.StaticText(
            panel,
            label=(
                "建议：浏览器多用户请使用“匹配模式=profile”，关键词填 Profile 1 / Profile 2 / Default\n"
                "若运行窗口识别不准，添加时可手动改 profile"
            )
        )
        desc.SetForegroundColour(wx.Colour(100, 100, 100))
        root.Add(desc, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "程序名称", width=120)
        self.list_ctrl.InsertColumn(1, "快捷键", width=120)
        self.list_ctrl.InsertColumn(2, "匹配模式", width=90)
        self.list_ctrl.InsertColumn(3, "关键词", width=180)
        self.list_ctrl.InsertColumn(4, "程序路径", width=320)
        self.list_ctrl.InsertColumn(5, "启动参数", width=140)
        root.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)

        add_btn = wx.Button(panel, label="手动添加")
        add_btn.Bind(wx.EVT_BUTTON, self.on_manual_add)
        btn_row.Add(add_btn, 0, wx.ALL, 5)

        add_running_btn = wx.Button(panel, label="从运行窗口添加")
        add_running_btn.Bind(wx.EVT_BUTTON, self.on_add_from_running)
        btn_row.Add(add_running_btn, 0, wx.ALL, 5)

        del_btn = wx.Button(panel, label="删除")
        del_btn.Bind(wx.EVT_BUTTON, self.on_delete)
        btn_row.Add(del_btn, 0, wx.ALL, 5)

        set_hotkey_btn = wx.Button(panel, label="设置快捷键")
        set_hotkey_btn.Bind(wx.EVT_BUTTON, self.on_set_hotkey)
        btn_row.Add(set_hotkey_btn, 0, wx.ALL, 5)

        set_kw_btn = wx.Button(panel, label="设置关键词/模式")
        set_kw_btn.Bind(wx.EVT_BUTTON, self.on_set_match)
        btn_row.Add(set_kw_btn, 0, wx.ALL, 5)

        hide_btn = wx.Button(panel, label="最小化到托盘")
        hide_btn.Bind(wx.EVT_BUTTON, lambda e: self.hide_to_tray())
        btn_row.Add(hide_btn, 0, wx.ALL, 5)

        root.Add(btn_row, 0, wx.ALIGN_CENTER)

        setting_row = wx.BoxSizer(wx.HORIZONTAL)
        self.autostart_cb = wx.CheckBox(panel, label="开机自启")
        self.autostart_cb.SetValue(self.autostart)
        setting_row.Add(self.autostart_cb, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 8)

        apply_setting_btn = wx.Button(panel, label="保存设置")
        apply_setting_btn.Bind(wx.EVT_BUTTON, self.on_apply_settings)
        setting_row.Add(apply_setting_btn, 0, wx.ALL, 5)
        root.Add(setting_row, 0, wx.ALIGN_CENTER)

        self.status_text = wx.StaticText(panel, label="就绪")
        root.Add(self.status_text, 0, wx.ALL | wx.ALIGN_CENTER, 6)

        panel.SetSizer(root)

    def update_status(self, text):
        self.status_text.SetLabel(text)

    def refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for p in self.programs:
            self.list_ctrl.Append([
                p.get("name", ""),
                p.get("hotkey", ""),
                p.get("match_mode", "title"),
                p.get("window_keyword", ""),
                p.get("path", ""),
                p.get("args", "")
            ])

    def hide_to_tray(self):
        self.Hide()
        self.update_status("已最小化到托盘（双击托盘图标可恢复）")

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

    def exit_app(self):
        self.exiting = True
        self.Close()

    def persist(self):
        save_config(self.programs, self.autostart_cb.GetValue())

    def on_apply_settings(self, _):
        target = self.autostart_cb.GetValue()
        ok = set_autostart(target)
        if not ok:
            wx.MessageBox("开机自启设置失败（可能权限不足）", "错误", wx.OK | wx.ICON_ERROR)
            self.autostart_cb.SetValue(is_autostart_enabled())
            return
        self.autostart = target
        self.persist()
        wx.MessageBox("设置已保存", "成功", wx.OK | wx.ICON_INFORMATION)

    def unregister_all_hotkeys(self):
        for hotkey_id in self.registered_hotkey_ids:
            try:
                self.UnregisterHotKey(hotkey_id)
            except Exception:
                pass
        self.registered_hotkey_ids.clear()
        self.hotkey_id_to_index.clear()

    def register_all_hotkeys(self):
        self.unregister_all_hotkeys()

        used = set()
        base_id = 1000
        fail_msgs = []

        for idx, p in enumerate(self.programs):
            hk = normalize_hotkey(p.get("hotkey", ""))
            if not hk:
                continue
            try:
                mods, vk, normalized = hotkey_to_mod_vk(hk)
                self.programs[idx]["hotkey"] = normalized
                hk = normalized
            except ValueError as e:
                fail_msgs.append(f"{p.get('name','')}：{hk}（{e}）")
                continue

            if hk in used:
                fail_msgs.append(f"{p.get('name','')}：{hk}（与列表内重复）")
                continue
            used.add(hk)

            hotkey_id = base_id + idx
            ok = self.RegisterHotKey(hotkey_id, mods, vk)
            if ok:
                self.registered_hotkey_ids.append(hotkey_id)
                self.hotkey_id_to_index[hotkey_id] = idx
                self.Bind(wx.EVT_HOTKEY, self.on_hotkey, id=hotkey_id)
            else:
                fail_msgs.append(f"{p.get('name','')}：{hk}（注册失败，可能被占用）")

        self.persist()
        self.refresh_list()

        if fail_msgs:
            wx.MessageBox("以下快捷键注册失败：\n\n" + "\n".join(fail_msgs), "热键提示", wx.OK | wx.ICON_WARNING)

    def on_hotkey(self, event):
        idx = self.hotkey_id_to_index.get(event.GetId())
        if idx is None or not (0 <= idx < len(self.programs)):
            return
        program = self.programs[idx]
        try:
            toggle_program(program)
            self.update_status(f"已切换：{program.get('name', '')}")
        except Exception as e:
            self.update_status("切换失败")
            wx.MessageBox(f"切换失败：{e}", "错误", wx.OK | wx.ICON_ERROR)

    def on_manual_add(self, _):
        dialog = wx.Dialog(self, title="添加程序", size=(760, 420))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        row1 = wx.BoxSizer(wx.HORIZONTAL)
        row1.Add(wx.StaticText(panel, label="程序名称:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        name_ctrl = wx.TextCtrl(panel, size=(520, -1))
        row1.Add(name_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(row1, 0, wx.ALL | wx.EXPAND, 5)

        row2 = wx.BoxSizer(wx.HORIZONTAL)
        row2.Add(wx.StaticText(panel, label="程序路径:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        path_ctrl = wx.TextCtrl(panel, size=(420, -1))
        row2.Add(path_ctrl, 1, wx.ALL | wx.EXPAND, 5)

        def on_browse(_evt):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                p = dlg.GetPath()
                path_ctrl.SetValue(p)
                if not name_ctrl.GetValue().strip():
                    name_ctrl.SetValue(os.path.basename(p).replace(".exe", ""))
            dlg.Destroy()

        browse_btn = wx.Button(panel, label="浏览")
        browse_btn.Bind(wx.EVT_BUTTON, on_browse)
        row2.Add(browse_btn, 0, wx.ALL, 5)

        def on_search(_evt):
            sd = ProgramSearchDialog(self)
            if sd.ShowModal() == wx.ID_OK:
                item = sd.get_selected()
                if item:
                    path_ctrl.SetValue(item["path"])
                    if not name_ctrl.GetValue().strip():
                        name_ctrl.SetValue(item["name"])
            sd.Destroy()

        search_btn = wx.Button(panel, label="搜索程序")
        search_btn.Bind(wx.EVT_BUTTON, on_search)
        row2.Add(search_btn, 0, wx.ALL, 5)
        sizer.Add(row2, 0, wx.ALL | wx.EXPAND, 5)

        row3 = wx.BoxSizer(wx.HORIZONTAL)
        row3.Add(wx.StaticText(panel, label="启动参数:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        args_ctrl = wx.TextCtrl(panel, size=(520, -1))
        row3.Add(args_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(row3, 0, wx.ALL | wx.EXPAND, 5)

        row4 = wx.BoxSizer(wx.HORIZONTAL)
        row4.Add(wx.StaticText(panel, label="匹配模式:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        mode_choice = wx.Choice(panel, choices=["title", "profile"])
        mode_choice.SetSelection(0)
        row4.Add(mode_choice, 0, wx.ALL, 5)
        row4.Add(wx.StaticText(panel, label="关键词:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        kw_ctrl = wx.TextCtrl(panel, size=(360, -1))
        kw_ctrl.SetHint("title: 标题包含；profile: 例如 Profile 1 / Default")
        row4.Add(kw_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(row4, 0, wx.ALL | wx.EXPAND, 5)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)

        def on_ok(_evt):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            args = args_ctrl.GetValue().strip()
            kw = kw_ctrl.GetValue().strip()
            mode = mode_choice.GetStringSelection() or "title"

            if not name or not path:
                wx.MessageBox("请填写程序名称和路径", "提示")
                return
            if not os.path.exists(path):
                wx.MessageBox("程序路径不存在，请检查", "错误", wx.OK | wx.ICON_ERROR)
                return

            self.programs.append({
                "name": name,
                "path": path,
                "args": args,
                "hotkey": "",
                "window_keyword": kw,
                "match_mode": mode
            })
            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()
            dialog.Destroy()

        ok_btn = wx.Button(panel, wx.ID_OK, "添加")
        ok_btn.Bind(wx.EVT_BUTTON, on_ok)
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        cancel_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())
        btn_row.Add(ok_btn, 0, wx.ALL, 5)
        btn_row.Add(cancel_btn, 0, wx.ALL, 5)
        sizer.Add(btn_row, 0, wx.ALIGN_CENTER | wx.ALL, 8)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_add_from_running(self, _):
        windows = enum_visible_app_windows()
        running = []

        for w in windows:
            path = w.get("path", "")
            if not path or not path.lower().endswith(".exe"):
                continue
            name = os.path.splitext(os.path.basename(path))[0]
            running.append({
                "name": name,
                "path": path,
                "title": w.get("title", ""),
                "profile": w.get("profile", ""),
                "proc_name": w.get("proc_name", "")
            })

        if not running:
            wx.MessageBox("没有找到可添加的运行窗口", "提示")
            return

        running.sort(key=lambda x: (x["title"] or x["name"]).lower())

        dialog = wx.Dialog(self, title="选择运行中的窗口", size=(980, 520))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer.Add(wx.StaticText(panel, label="双击添加。浏览器会优先 profile 匹配；若不准会让你手动确认"), 0, wx.ALL, 8)

        list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        list_ctrl.InsertColumn(0, "窗口标题", width=340)
        list_ctrl.InsertColumn(1, "程序名", width=120)
        list_ctrl.InsertColumn(2, "Profile(推断)", width=140)
        list_ctrl.InsertColumn(3, "程序路径", width=340)

        for item in running:
            i = list_ctrl.InsertItem(list_ctrl.GetItemCount(), item["title"] or "(无标题)")
            list_ctrl.SetItem(i, 1, item["name"])
            list_ctrl.SetItem(i, 2, item["profile"] or "")
            list_ctrl.SetItem(i, 3, item["path"])

        sizer.Add(list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        def on_double_click(_evt):
            i = list_ctrl.GetFirstSelected()
            if i < 0:
                return
            item = running[i]

            proc = (item.get("proc_name", "") or "").lower().replace(".exe", "")
            profile = (item.get("profile", "") or "").strip()
            title = (item.get("title", "") or "").strip()

            if proc in BROWSER_SET:
                mode = "profile"

                ok, manual_profile = ask_profile_input(self, default_profile=profile)
                if not ok:
                    return
                if not manual_profile:
                    wx.MessageBox("浏览器建议填写 profile（例如 Profile 1 / Default）", "提示")
                    return

                kw = manual_profile
                auto_args = build_profile_args(proc, manual_profile)
            else:
                mode = "title"
                kw = title[:80]
                auto_args = ""

            self.programs.append({
                "name": item["name"],
                "path": item["path"],
                "args": auto_args,
                "hotkey": "",
                "window_keyword": kw,
                "match_mode": mode
            })
            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()
            dialog.Destroy()

        list_ctrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, on_double_click)

        close_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())
        sizer.Add(close_btn, 0, wx.ALL | wx.ALIGN_CENTER, 6)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_delete(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择一项", "提示")
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

        current = self.programs[idx].get("hotkey", "")
        dlg = HotkeyCaptureDialog(self, current_hotkey=current)

        if dlg.ShowModal() == wx.ID_OK:
            hk = normalize_hotkey(dlg.captured_hotkey)
            if not hk:
                self.programs[idx]["hotkey"] = ""
                self.persist()
                self.refresh_list()
                self.register_all_hotkeys()
                wx.MessageBox("已清空快捷键", "提示")
                dlg.Destroy()
                return

            try:
                _, _, normalized = hotkey_to_mod_vk(hk)
            except ValueError as e:
                wx.MessageBox(f"快捷键无效：{e}", "错误", wx.OK | wx.ICON_ERROR)
                dlg.Destroy()
                return

            for i, p in enumerate(self.programs):
                if i != idx and normalize_hotkey(p.get("hotkey", "")) == normalized:
                    wx.MessageBox("该快捷键已被其他程序使用", "错误", wx.OK | wx.ICON_ERROR)
                    dlg.Destroy()
                    return

            self.programs[idx]["hotkey"] = normalized
            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()
            wx.MessageBox(f"已设置快捷键：{normalized}", "成功", wx.OK | wx.ICON_INFORMATION)

        dlg.Destroy()

    def on_set_match(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        p = self.programs[idx]
        cur_mode = p.get("match_mode", "title")
        cur_kw = p.get("window_keyword", "")

        dlg = wx.Dialog(self, title="设置关键词与匹配模式", size=(520, 220))
        panel = wx.Panel(dlg)
        s = wx.BoxSizer(wx.VERTICAL)

        row1 = wx.BoxSizer(wx.HORIZONTAL)
        row1.Add(wx.StaticText(panel, label="匹配模式:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 6)
        mode_choice = wx.Choice(panel, choices=["title", "profile"])
        mode_choice.SetStringSelection(cur_mode if cur_mode in ("title", "profile") else "title")
        row1.Add(mode_choice, 0, wx.ALL, 6)
        s.Add(row1, 0, wx.EXPAND)

        row2 = wx.BoxSizer(wx.HORIZONTAL)
        row2.Add(wx.StaticText(panel, label="关键词:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 6)
        kw_ctrl = wx.TextCtrl(panel, value=cur_kw, size=(360, -1))
        row2.Add(kw_ctrl, 1, wx.ALL | wx.EXPAND, 6)
        s.Add(row2, 0, wx.EXPAND)

        tip = wx.StaticText(panel, label="title=窗口标题包含；profile=浏览器用户目录名（如 Profile 1）")
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
            self.programs[idx]["match_mode"] = mode_choice.GetStringSelection() or "title"
            self.programs[idx]["window_keyword"] = kw_ctrl.GetValue().strip()
            self.persist()
            self.refresh_list()
            wx.MessageBox("匹配规则已保存", "成功", wx.OK | wx.ICON_INFORMATION)

        dlg.Destroy()


class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True


if __name__ == "__main__":
    app = QuickLauncherApp()
    app.MainLoop()
