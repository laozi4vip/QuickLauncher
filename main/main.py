# -*- coding: utf-8 -*-
"""
QuickLauncher - 严格匹配版
特性：
1) 取消所有兜底策略（不按 exe 名模糊匹配，不自动猜测）
2) 热键未命中时：不启动、不切换、无副作用
3) hwnd 失效后仅按稳定特征重绑（profile/title_sig/keyword），避免随机串绑
"""

import wx
import wx.adv
import os
import json
import subprocess
import psutil
import ctypes
import sys
import shlex
import re
import time
import winreg

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

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")


# ---------------------------
# 配置
# ---------------------------
def get_launch_command():
    if getattr(sys, "frozen", False):
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
    if not os.path.exists(CONFIG_FILE):
        return default_data

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        programs = data.get("programs", [])
        for p in programs:
            p.setdefault("name", "")
            p.setdefault("path", "")
            p.setdefault("args", "")
            p.setdefault("hotkey", "")
            p.setdefault("window_keyword", "")
            p.setdefault("match_mode", "title")  # title/profile/hwnd
            p.setdefault("bind_hwnd", 0)
            p.setdefault("profile_name", "")
            p.setdefault("title_sig", "")
        return {"programs": programs, "autostart": data.get("autostart", is_autostart_enabled())}
    except Exception as e:
        print("load_config error:", e)
        return default_data


def save_config(programs, autostart):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"programs": programs, "autostart": autostart}, f, ensure_ascii=False, indent=4)


# ---------------------------
# 热键
# ---------------------------
def normalize_hotkey(hk: str) -> str:
    return (hk or "").strip().lower().replace(" ", "")


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
        elif p == "shift":
            mods |= MOD_SHIFT
        elif p == "alt":
            mods |= MOD_ALT
        elif p == "win":
            mods |= MOD_WIN
        else:
            if main_key is not None:
                raise ValueError("只能有一个主键")
            if p not in key_map:
                raise ValueError(f"不支持的主键: {p}")
            main_key = p

    if main_key is None:
        raise ValueError("缺少主键")
    if mods == 0 and not main_key.startswith("f"):
        raise ValueError("无修饰键仅支持 F1-F12")

    order = []
    if mods & MOD_CONTROL: order.append("ctrl")
    if mods & MOD_SHIFT: order.append("shift")
    if mods & MOD_ALT: order.append("alt")
    if mods & MOD_WIN: order.append("win")
    normalized = "+".join(order + [main_key]) if order else main_key
    return mods, key_map[main_key], normalized


# ---------------------------
# 窗口工具
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
    title = get_window_title(hwnd).strip()
    if not title:
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
            cmd = p.cmdline()
        except Exception:
            cmd = []
        return path, name, cmd
    except Exception:
        return "", "", []


def parse_profile_from_cmdline(proc_name: str, cmdline):
    name = (proc_name or "").lower().replace(".exe", "")
    args = cmdline or []

    def value_after(prefix):
        for i, a in enumerate(args):
            la = (a or "").lower()
            if la.startswith(prefix + "="):
                return a.split("=", 1)[1].strip().strip('"')
            if la == prefix and i + 1 < len(args):
                return (args[i + 1] or "").strip().strip('"')
        return ""

    if name in ("chrome", "msedge", "chromium", "brave"):
        # 优先 profile-directory；其次 user-data-dir basename
        v = value_after("--profile-directory")
        if v:
            return v
        u = value_after("--user-data-dir")
        if u:
            return os.path.basename(u.rstrip("\\/"))
        return ""
    if name == "firefox":
        for i, a in enumerate(args):
            la = (a or "").lower()
            if la in ("-p", "--profile") and i + 1 < len(args):
                return (args[i + 1] or "").strip().strip('"')
            if la == "-profile" and i + 1 < len(args):
                return os.path.basename((args[i + 1] or "").strip().strip('"').rstrip("\\/"))
    return ""


def make_title_signature(title: str):
    t = (title or "").strip().lower()
    if not t:
        return ""
    t = re.sub(r"\s+", " ", t)
    parts = [x.strip() for x in t.split("-") if x.strip()]
    if len(parts) >= 2:
        return " - ".join(parts[-2:])[:120]
    return t[:120]


def enum_visible_app_windows():
    ws = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def cb(hwnd, _):
        try:
            if not user32.IsWindow(hwnd):
                return True
            if not is_alt_tab_window(hwnd):
                return True
            pid = get_pid_from_hwnd(hwnd)
            title = get_window_title(hwnd).strip()
            path, proc_name, cmdline = get_proc_path_name_cmdline(pid)
            profile = parse_profile_from_cmdline(proc_name, cmdline)
            ws.append({
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

    user32.EnumWindows(cb, 0)
    return ws


def enum_windows_for_program(path: str):
    """
    严格模式：只按绝对 path 完全匹配，不再用 exe 名兜底
    """
    target = os.path.normcase(path or "")
    if not target:
        return []
    result = []
    for w in enum_visible_app_windows():
        wpath = os.path.normcase(w.get("path", "") or "")
        if wpath == target:
            result.append(w)
    return result


# ---------------------------
# 严格匹配逻辑（核心）
# ---------------------------
def strict_match_score(program, w):
    mode = (program.get("match_mode", "title") or "title").lower()
    kw = (program.get("window_keyword", "") or "").strip().lower()
    pf = (program.get("profile_name", "") or "").strip().lower()
    ts = (program.get("title_sig", "") or "").strip().lower()

    w_title = (w.get("title", "") or "").lower()
    w_prof = (w.get("profile", "") or "").strip().lower()
    w_sig = (w.get("title_sig", "") or "").strip().lower()

    # 没有任何稳定特征时，不允许匹配（防随机）
    if mode == "hwnd" and not (pf or ts or kw):
        return -1

    if mode == "title":
        if not kw:
            return -1
        return 100 if kw in w_title else -1

    if mode == "profile":
        # profile 模式只认 profile_name/keyword 与窗口 profile 的精确/包含
        target = pf or kw
        if not target or not w_prof:
            return -1
        if w_prof == target:
            return 120
        if target in w_prof:
            return 80
        return -1

    # hwnd 模式：优先 profile/title_sig/keyword 的组合
    score = 0
    if pf:
        if not w_prof:
            return -1
        if w_prof == pf:
            score += 120
        elif pf in w_prof:
            score += 70
        else:
            return -1  # 有 profile_name 就必须命中，否则拒绝

    if ts:
        if w_sig == ts:
            score += 80
        elif ts in w_sig or w_sig in ts:
            score += 30
        else:
            # 有 title_sig 但完全不沾边，拒绝
            return -1

    if kw:
        # hwnd 模式下 keyword 仅作为附加
        if kw in w_title or kw == w_prof or kw in w_prof:
            score += 30

    return score if score > 0 else -1


def find_window_for_program(program):
    path = program.get("path", "")
    if not path:
        return None

    mode = (program.get("match_mode", "title") or "title").lower()
    bind_hwnd = int(program.get("bind_hwnd", 0) or 0)

    # 1) hwnd 已绑定且有效 -> 必须 path 一致才返回
    if mode == "hwnd" and bind_hwnd and is_hwnd_valid(bind_hwnd):
        pid = get_pid_from_hwnd(bind_hwnd)
        ppath, _, _ = get_proc_path_name_cmdline(pid)
        if os.path.normcase(ppath or "") == os.path.normcase(path or ""):
            return bind_hwnd

    # 2) 严格候选匹配
    cands = enum_windows_for_program(path)
    if not cands:
        return None

    scored = []
    for w in cands:
        s = strict_match_score(program, w)
        if s >= 0:
            scored.append((s, int(w["hwnd"]), w))

    if not scored:
        return None

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score = scored[0][0]
    best = [x for x in scored if x[0] == best_score]

    # 分数并列，视为歧义，不动作（防串绑）
    if len(best) > 1:
        return None

    return best[0][1]


def toggle_program(program):
    """
    严格模式：未匹配时不启动程序，直接无动作
    """
    hwnd = find_window_for_program(program)
    if not hwnd:
        return None, False

    fg = user32.GetForegroundWindow()
    if hwnd == fg:
        user32.ShowWindow(hwnd, SW_MINIMIZE)
    else:
        if user32.IsIconic(hwnd):
            user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)
    return hwnd, True


# ---------------------------
# UI
# ---------------------------
class QuickLauncherTaskBar(wx.adv.TaskBarIcon):
    def __init__(self, frame):
        super().__init__()
        self.frame = frame
        bmp = wx.ArtProvider.GetBitmap(wx.ART_EXECUTABLE_FILE, wx.ART_OTHER, (16, 16))
        icon = wx.Icon()
        icon.CopyFromBitmap(bmp)
        self.SetIcon(icon, "QuickLauncher")
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK, lambda e: self.frame.show_from_tray())

    def CreatePopupMenu(self):
        menu = wx.Menu()
        s = menu.Append(wx.ID_ANY, "显示主窗口")
        h = menu.Append(wx.ID_ANY, "隐藏到托盘")
        menu.AppendSeparator()
        x = menu.Append(wx.ID_EXIT, "退出")
        self.Bind(wx.EVT_MENU, lambda e: self.frame.show_from_tray(), s)
        self.Bind(wx.EVT_MENU, lambda e: self.frame.hide_to_tray(), h)
        self.Bind(wx.EVT_MENU, lambda e: self.frame.exit_app(), x)
        return menu


class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher(严格匹配版)", size=(980, 560))
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

    def init_ui(self):
        p = wx.Panel(self)
        root = wx.BoxSizer(wx.VERTICAL)

        self.list_ctrl = wx.ListCtrl(p, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "名称", width=120)
        self.list_ctrl.InsertColumn(1, "快捷键", width=120)
        self.list_ctrl.InsertColumn(2, "模式", width=80)
        self.list_ctrl.InsertColumn(3, "关键词", width=150)
        self.list_ctrl.InsertColumn(4, "HWND", width=100)
        self.list_ctrl.InsertColumn(5, "Profile", width=120)
        self.list_ctrl.InsertColumn(6, "TitleSig", width=180)
        self.list_ctrl.InsertColumn(7, "路径", width=230)
        root.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        for text, fn in [
            ("删除", self.on_delete),
            ("最小化到托盘", lambda e: self.hide_to_tray()),
        ]:
            b = wx.Button(p, label=text)
            b.Bind(wx.EVT_BUTTON, fn)
            btn_row.Add(b, 0, wx.ALL, 4)
        root.Add(btn_row, 0, wx.ALIGN_CENTER)

        row2 = wx.BoxSizer(wx.HORIZONTAL)
        self.autostart_cb = wx.CheckBox(p, label="开机自启")
        self.autostart_cb.SetValue(self.autostart)
        row2.Add(self.autostart_cb, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 6)
        save_btn = wx.Button(p, label="保存设置")
        save_btn.Bind(wx.EVT_BUTTON, self.on_apply_settings)
        row2.Add(save_btn, 0, wx.ALL, 6)
        root.Add(row2, 0, wx.ALIGN_CENTER)

        self.status = wx.StaticText(p, label="就绪（严格模式：未命中不动作）")
        root.Add(self.status, 0, wx.ALL | wx.ALIGN_CENTER, 6)
        p.SetSizer(root)

    def persist(self):
        save_config(self.programs, self.autostart_cb.GetValue())

    def refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for x in self.programs:
            self.list_ctrl.Append([
                x.get("name", ""),
                x.get("hotkey", ""),
                x.get("match_mode", "title"),
                x.get("window_keyword", ""),
                str(int(x.get("bind_hwnd", 0) or 0)),
                x.get("profile_name", ""),
                x.get("title_sig", ""),
                x.get("path", "")
            ])

    def set_status(self, s):
        self.status.SetLabel(s)

    def on_apply_settings(self, _):
        ok = set_autostart(self.autostart_cb.GetValue())
        if not ok:
            wx.MessageBox("开机自启设置失败", "错误", wx.OK | wx.ICON_ERROR)
            self.autostart_cb.SetValue(is_autostart_enabled())
            return
        self.persist()
        wx.MessageBox("已保存", "成功")

    def hide_to_tray(self):
        self.Hide()
        self.set_status("已最小化到托盘")

    def show_from_tray(self):
        self.Show()
        self.Raise()
        self.Iconize(False)

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

    def on_delete(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择一项", "提示")
            return
        self.programs.pop(idx)
        self.persist()
        self.refresh_list()
        self.register_all_hotkeys()

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
        base = 1000

        for idx, prog in enumerate(self.programs):
            hk = normalize_hotkey(prog.get("hotkey", ""))
            if not hk:
                continue
            try:
                mods, vk, n = hotkey_to_mod_vk(hk)
                self.programs[idx]["hotkey"] = n
            except Exception as e:
                fail.append(f"{prog.get('name','')}: {hk} ({e})")
                continue
            if n in used:
                fail.append(f"{prog.get('name','')}: {n}(重复)")
                continue
            used.add(n)

            hid = base + idx
            if self.RegisterHotKey(hid, mods, vk):
                self.registered_hotkey_ids.append(hid)
                self.hotkey_id_to_index[hid] = idx
                self.Bind(wx.EVT_HOTKEY, self.on_hotkey, id=hid)
            else:
                fail.append(f"{prog.get('name','')}: {n}(被占用)")

        self.persist()
        self.refresh_list()
        if fail:
            wx.MessageBox("以下快捷键注册失败：\n" + "\n".join(fail), "提示")

    def on_hotkey(self, event):
        idx = self.hotkey_id_to_index.get(event.GetId())
        if idx is None or idx >= len(self.programs):
            return
        prog = self.programs[idx]

        old = int(prog.get("bind_hwnd", 0) or 0)
        hwnd, ok = toggle_program(prog)

        if not ok:
            # 严格策略：未命中不动作
            self.set_status(f"未命中：{prog.get('name','')}（无任何响应）")
            return

        # 仅在 hwnd 模式且旧绑定无效时，更新绑定
        if (prog.get("match_mode", "title") or "title").lower() == "hwnd":
            keep_old = False
            if old and is_hwnd_valid(old):
                pid = get_pid_from_hwnd(old)
                ppath, _, _ = get_proc_path_name_cmdline(pid)
                keep_old = (os.path.normcase(ppath or "") == os.path.normcase(prog.get("path", "") or ""))
            if not keep_old and hwnd and hwnd != old:
                prog["bind_hwnd"] = int(hwnd)
                prog["title_sig"] = make_title_signature(get_window_title(hwnd))
                self.persist()
                self.refresh_list()

        self.set_status(f"已切换：{prog.get('name','')}")

class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True


if __name__ == "__main__":
    app = QuickLauncherApp()
    app.MainLoop()
