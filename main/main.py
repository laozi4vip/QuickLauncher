# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows 托盘快捷启动器
新增：
1) profile_keyword：用于区分同一 exe 的不同 profile（如 Chrome/Edge）
2) 匹配时同时检查：进程命令行 + 窗口标题
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

# ---------------------------
# 配置路径
# ---------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")


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
                p.setdefault("profile_keyword", "")  # 新增字段兼容

            return {"programs": programs, "autostart": autostart}
        except Exception as e:
            print("load_config error:", e)
    return default_data


def save_config(programs, autostart):
    data = {"programs": programs, "autostart": autostart}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


# ---------------------------
# 热键
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
                raise ValueError(f"不支持的主键：{p}")
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
# 窗口/进程工具
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


def get_proc_path_name_cmdline(pid):
    try:
        p = psutil.Process(pid)
        path = p.exe() or ""
        name = p.name() or ""
        try:
            cmdline = " ".join(p.cmdline()).strip().lower()
        except Exception:
            cmdline = ""
        return path, name, cmdline
    except Exception:
        return "", "", ""


def enum_windows_for_exe(exe_name: str):
    exe_name = exe_name.lower().replace(".exe", "")
    result = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def callback(hwnd, _):
        try:
            if not is_alt_tab_window(hwnd):
                return True

            pid = get_pid_from_hwnd(hwnd)
            path, name, cmdline = get_proc_path_name_cmdline(pid)
            proc_name = (name or "").lower().replace(".exe", "")
            if proc_name != exe_name:
                return True

            result.append({
                "hwnd": int(hwnd),
                "pid": int(pid),
                "title": get_window_title(hwnd).strip(),
                "path": path,
                "cmdline": cmdline
            })
        except Exception:
            pass
        return True

    user32.EnumWindows(callback, 0)
    return result


def find_window_for_program(program):
    path = program.get("path", "")
    if not path:
        return None

    exe_name = os.path.basename(path).lower().replace(".exe", "")
    window_keyword = (program.get("window_keyword", "") or "").strip().lower()
    profile_keyword = (program.get("profile_keyword", "") or "").strip().lower()

    candidates = enum_windows_for_exe(exe_name)
    if not candidates:
        return None

    # 先按 profile_keyword 过滤（命令行 + 标题）
    if profile_keyword:
        filtered = []
        for c in candidates:
            hay = f'{c.get("cmdline","")} {c.get("title","").lower()}'
            if profile_keyword in hay:
                filtered.append(c)
        candidates = filtered if filtered else candidates

    # 再按 window_keyword 过滤（标题）
    if window_keyword:
        filtered = [c for c in candidates if window_keyword in (c.get("title", "").lower())]
        candidates = filtered if filtered else candidates

    if not candidates:
        return None

    fg = user32.GetForegroundWindow()

    # 优先前台窗口命中
    for c in candidates:
        if c["hwnd"] == fg:
            return c["hwnd"]

    # 再选未最小化
    for c in candidates:
        if not user32.IsIconic(c["hwnd"]):
            return c["hwnd"]

    return candidates[0]["hwnd"]


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


# ---------------------------
# 对话框
# ---------------------------
class HotkeyCaptureDialog(wx.Dialog):
    def __init__(self, parent, current_hotkey=""):
        super().__init__(parent, title="设置快捷键（按键自动录入）", size=(500, 230))
        self.captured_hotkey = normalize_hotkey(current_hotkey)

        panel = wx.Panel(self)
        s = wx.BoxSizer(wx.VERTICAL)

        tips = wx.StaticText(panel, label="直接按目标快捷键（Esc取消，Backspace清空）")
        tips.SetForegroundColour(wx.Colour(90, 90, 90))
        s.Add(tips, 0, wx.ALL, 10)

        self.hk_show = wx.TextCtrl(panel, value=self.captured_hotkey, style=wx.TE_READONLY)
        s.Add(self.hk_show, 0, wx.ALL | wx.EXPAND, 10)

        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(wx.Button(panel, wx.ID_OK, "确定"), 0, wx.ALL, 5)
        row.Add(wx.Button(panel, wx.ID_CANCEL, "取消"), 0, wx.ALL, 5)
        clear_btn = wx.Button(panel, wx.ID_ANY, "清空")
        row.Add(clear_btn, 0, wx.ALL, 5)
        s.Add(row, 0, wx.ALIGN_CENTER)

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
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(1020, 580))

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
        self.update_status("运行中...")

        self.Bind(wx.EVT_CLOSE, self.on_close)

    def init_ui(self):
        panel = wx.Panel(self)
        root = wx.BoxSizer(wx.VERTICAL)

        root.Add(wx.StaticText(panel, label="程序列表"), 0, wx.ALL | wx.ALIGN_CENTER, 10)

        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "程序名", width=120)
        self.list_ctrl.InsertColumn(1, "快捷键", width=110)
        self.list_ctrl.InsertColumn(2, "Profile关键词", width=160)
        self.list_ctrl.InsertColumn(3, "窗口关键词", width=180)
        self.list_ctrl.InsertColumn(4, "程序路径", width=320)
        self.list_ctrl.InsertColumn(5, "启动参数", width=160)
        root.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        row = wx.BoxSizer(wx.HORIZONTAL)
        btn_add = wx.Button(panel, label="手动添加")
        btn_del = wx.Button(panel, label="删除")
        btn_hotkey = wx.Button(panel, label="设置快捷键")
        btn_profile = wx.Button(panel, label="设置Profile关键词")
        btn_kw = wx.Button(panel, label="设置窗口关键词")
        btn_hide = wx.Button(panel, label="最小化到托盘")

        btn_add.Bind(wx.EVT_BUTTON, self.on_manual_add)
        btn_del.Bind(wx.EVT_BUTTON, self.on_delete)
        btn_hotkey.Bind(wx.EVT_BUTTON, self.on_set_hotkey)
        btn_profile.Bind(wx.EVT_BUTTON, self.on_set_profile_keyword)
        btn_kw.Bind(wx.EVT_BUTTON, self.on_set_window_keyword)
        btn_hide.Bind(wx.EVT_BUTTON, lambda e: self.hide_to_tray())

        for b in [btn_add, btn_del, btn_hotkey, btn_profile, btn_kw, btn_hide]:
            row.Add(b, 0, wx.ALL, 5)
        root.Add(row, 0, wx.ALIGN_CENTER)

        row2 = wx.BoxSizer(wx.HORIZONTAL)
        self.autostart_cb = wx.CheckBox(panel, label="开机自启")
        self.autostart_cb.SetValue(self.autostart)
        row2.Add(self.autostart_cb, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 8)

        btn_apply = wx.Button(panel, label="保存设置")
        btn_apply.Bind(wx.EVT_BUTTON, self.on_apply_settings)
        row2.Add(btn_apply, 0, wx.ALL, 5)
        root.Add(row2, 0, wx.ALIGN_CENTER)

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
                p.get("profile_keyword", ""),
                p.get("window_keyword", ""),
                p.get("path", ""),
                p.get("args", "")
            ])

    def persist(self):
        save_config(self.programs, self.autostart_cb.GetValue())

    # ----- 托盘 -----
    def hide_to_tray(self):
        self.Hide()
        self.update_status("已隐藏到托盘")

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

    # ----- 设置 -----
    def on_apply_settings(self, _):
        target = self.autostart_cb.GetValue()
        ok = set_autostart(target)
        if not ok:
            wx.MessageBox("开机自启设置失败", "错误", wx.OK | wx.ICON_ERROR)
            self.autostart_cb.SetValue(is_autostart_enabled())
            return
        self.autostart = target
        self.persist()
        wx.MessageBox("设置已保存", "成功", wx.OK | wx.ICON_INFORMATION)

    # ----- 热键注册 -----
    def unregister_all_hotkeys(self):
        for hid in self.registered_hotkey_ids:
            try:
                self.UnregisterHotKey(hid)
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
                fail_msgs.append(f'{p.get("name","")}：{hk}（{e}）')
                continue

            if hk in used:
                fail_msgs.append(f'{p.get("name","")}：{hk}（重复）')
                continue
            used.add(hk)

            hid = base_id + idx
            if self.RegisterHotKey(hid, mods, vk):
                self.registered_hotkey_ids.append(hid)
                self.hotkey_id_to_index[hid] = idx
                self.Bind(wx.EVT_HOTKEY, self.on_hotkey, id=hid)
            else:
                fail_msgs.append(f'{p.get("name","")}：{hk}（注册失败，可能占用）')

        self.persist()
        self.refresh_list()

        if fail_msgs:
            wx.MessageBox("以下快捷键注册失败：\n\n" + "\n".join(fail_msgs),
                          "热键提示", wx.OK | wx.ICON_WARNING)

    def on_hotkey(self, event):
        idx = self.hotkey_id_to_index.get(event.GetId())
        if idx is None or idx < 0 or idx >= len(self.programs):
            return
        try:
            p = self.programs[idx]
            toggle_program(p)
            self.update_status(f'已切换：{p.get("name","")}')
        except Exception as e:
            wx.MessageBox(f"切换失败：{e}", "错误", wx.OK | wx.ICON_ERROR)

    # ----- 列表操作 -----
    def on_manual_add(self, _):
        dialog = wx.Dialog(self, title="添加程序", size=(760, 420))
        panel = wx.Panel(dialog)
        s = wx.BoxSizer(wx.VERTICAL)

        # 名称
        r1 = wx.BoxSizer(wx.HORIZONTAL)
        r1.Add(wx.StaticText(panel, label="程序名称:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        name_ctrl = wx.TextCtrl(panel)
        r1.Add(name_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        s.Add(r1, 0, wx.EXPAND)

        # 路径
        r2 = wx.BoxSizer(wx.HORIZONTAL)
        r2.Add(wx.StaticText(panel, label="程序路径:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        path_ctrl = wx.TextCtrl(panel)
        r2.Add(path_ctrl, 1, wx.ALL | wx.EXPAND, 5)

        def on_browse(_evt):
            fd = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if fd.ShowModal() == wx.ID_OK:
                p = fd.GetPath()
                path_ctrl.SetValue(p)
                if not name_ctrl.GetValue().strip():
                    name_ctrl.SetValue(os.path.basename(p).replace(".exe", ""))
            fd.Destroy()

        btn_browse = wx.Button(panel, label="浏览")
        btn_browse.Bind(wx.EVT_BUTTON, on_browse)
        r2.Add(btn_browse, 0, wx.ALL, 5)
        s.Add(r2, 0, wx.EXPAND)

        # 参数
        r3 = wx.BoxSizer(wx.HORIZONTAL)
        r3.Add(wx.StaticText(panel, label="启动参数:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        args_ctrl = wx.TextCtrl(panel)
        r3.Add(args_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        s.Add(r3, 0, wx.EXPAND)

        # profile关键词
        r4 = wx.BoxSizer(wx.HORIZONTAL)
        r4.Add(wx.StaticText(panel, label="Profile关键词:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        profile_ctrl = wx.TextCtrl(panel)
        profile_ctrl.SetHint("可选：如 profile 1 / person 2 / --profile-directory=Default")
        r4.Add(profile_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        s.Add(r4, 0, wx.EXPAND)

        # 窗口关键词
        r5 = wx.BoxSizer(wx.HORIZONTAL)
        r5.Add(wx.StaticText(panel, label="窗口关键词:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        kw_ctrl = wx.TextCtrl(panel)
        kw_ctrl.SetHint("可选：按窗口标题进一步区分")
        r5.Add(kw_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        s.Add(r5, 0, wx.EXPAND)

        btns = wx.BoxSizer(wx.HORIZONTAL)

        def on_ok(_evt):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            args = args_ctrl.GetValue().strip()
            profile_kw = profile_ctrl.GetValue().strip()
            win_kw = kw_ctrl.GetValue().strip()

            if not name or not path:
                wx.MessageBox("请填写程序名称和路径", "提示")
                return
            if not os.path.exists(path):
                wx.MessageBox("程序路径不存在", "错误", wx.OK | wx.ICON_ERROR)
                return

            self.programs.append({
                "name": name,
                "path": path,
                "args": args,
                "hotkey": "",
                "profile_keyword": profile_kw,
                "window_keyword": win_kw
            })
            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()
            dialog.Destroy()

        btn_ok = wx.Button(panel, wx.ID_OK, "添加")
        btn_ok.Bind(wx.EVT_BUTTON, on_ok)
        btn_cancel = wx.Button(panel, wx.ID_CANCEL, "取消")
        btn_cancel.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())

        btns.Add(btn_ok, 0, wx.ALL, 5)
        btns.Add(btn_cancel, 0, wx.ALL, 5)
        s.Add(btns, 0, wx.ALIGN_CENTER | wx.ALL, 8)

        panel.SetSizer(s)
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

        dlg = HotkeyCaptureDialog(self, current_hotkey=self.programs[idx].get("hotkey", ""))
        if dlg.ShowModal() == wx.ID_OK:
            hk = normalize_hotkey(dlg.captured_hotkey)

            if not hk:
                self.programs[idx]["hotkey"] = ""
            else:
                try:
                    _, _, hk = hotkey_to_mod_vk(hk)
                except ValueError as e:
                    wx.MessageBox(f"快捷键无效：{e}", "错误", wx.OK | wx.ICON_ERROR)
                    dlg.Destroy()
                    return

                for i, p in enumerate(self.programs):
                    if i != idx and normalize_hotkey(p.get("hotkey", "")) == hk:
                        wx.MessageBox("该快捷键已被其他程序使用", "错误", wx.OK | wx.ICON_ERROR)
                        dlg.Destroy()
                        return
                self.programs[idx]["hotkey"] = hk

            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()

        dlg.Destroy()

    def on_set_profile_keyword(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        cur = self.programs[idx].get("profile_keyword", "")
        dlg = wx.TextEntryDialog(
            self,
            "输入 Profile关键词（匹配进程命令行 + 窗口标题）\n"
            "例如：profile 1 / person 2 / default\n"
            "留空表示不按 profile 区分",
            "设置 Profile关键词",
            cur
        )
        if dlg.ShowModal() == wx.ID_OK:
            self.programs[idx]["profile_keyword"] = dlg.GetValue().strip()
            self.persist()
            self.refresh_list()
            wx.MessageBox("Profile关键词已保存", "成功", wx.OK | wx.ICON_INFORMATION)
        dlg.Destroy()

    def on_set_window_keyword(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        cur = self.programs[idx].get("window_keyword", "")
        dlg = wx.TextEntryDialog(
            self,
            "输入窗口关键词（匹配窗口标题）\n留空表示不按标题区分",
            "设置窗口关键词",
            cur
        )
        if dlg.ShowModal() == wx.ID_OK:
            self.programs[idx]["window_keyword"] = dlg.GetValue().strip()
            self.persist()
            self.refresh_list()
            wx.MessageBox("窗口关键词已保存", "成功", wx.OK | wx.ICON_INFORMATION)
        dlg.Destroy()


class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True


if __name__ == "__main__":
    app = QuickLauncherApp()
    app.MainLoop()
