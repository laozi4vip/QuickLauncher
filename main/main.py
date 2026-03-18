# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器（稳定版）
功能：
- 绑定程序到快捷键
- 按快捷键启动或切换窗口
- 窗口在前台时最小化，再按恢复

import wx
import os
import json
import subprocess
import psutil
import ctypes
import sys

# ---------------- Windows API ----------------
user32 = ctypes.windll.user32
SW_MINIMIZE = 6
SW_RESTORE = 9

# 配置路径
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("programs", [])
        except Exception as e:
            print("load_config error:", e)
    return []


def save_config(programs):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"programs": programs}, f, ensure_ascii=False, indent=4)
        print(f"Config saved: {CONFIG_FILE}")
    except Exception as e:
        print("save_config error:", e)


def normalize_exe_name(path_or_name: str) -> str:
    s = os.path.basename(path_or_name).lower()
    if s.endswith(".exe"):
        s = s[:-4]
    return s


def find_window_by_exe(exe_name):
    """按进程名找可见窗口"""
    exe_name = normalize_exe_name(exe_name)
    candidate_hwnds = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def callback(hwnd, _):
        try:
            if user32.IsWindowVisible(hwnd) and user32.IsWindowEnabled(hwnd):
                pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                try:
                    proc = psutil.Process(pid.value)
                    p_name = normalize_exe_name(proc.name())
                    if p_name == exe_name:
                        candidate_hwnds.append(hwnd)
                except Exception:
                    pass
        except Exception:
            pass
        return True

    user32.EnumWindows(callback, 0)

    if not candidate_hwnds:
        return None

    fg = user32.GetForegroundWindow()
    if fg in candidate_hwnds:
        return fg
    return candidate_hwnds[0]


def toggle_program(program):
    path = program.get("path", "").strip()
    if not path:
        return

    exe_name = normalize_exe_name(path)
    hwnd = find_window_by_exe(exe_name)

    if hwnd:
        fg = user32.GetForegroundWindow()
        if hwnd == fg:
            user32.ShowWindow(hwnd, SW_MINIMIZE)
        else:
            if user32.IsIconic(hwnd):
                user32.ShowWindow(hwnd, SW_RESTORE)
            # 有些程序激活受限，先尝试恢复再置前
            user32.ShowWindow(hwnd, SW_RESTORE)
            user32.SetForegroundWindow(hwnd)
    else:
        if os.path.exists(path):
            try:
                subprocess.Popen(path)
            except Exception as e:
                print("Popen error:", e)


def get_process_title(pid):
    titles = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def callback(hwnd, _):
        try:
            if hwnd and user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
                window_pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                if window_pid.value == pid:
                    length = user32.GetWindowTextLengthW(hwnd)
                    if length > 0:
                        buf = ctypes.create_unicode_buffer(length + 1)
                        user32.GetWindowTextW(hwnd, buf, length + 1)
                        t = buf.value.strip()
                        if t:
                            titles.append(t)
        except Exception:
            pass
        return True

    user32.EnumWindows(callback, 0)
    return titles[0] if titles else ""


# ---------------- 热键解析 ----------------
MOD_MAP = {
    "ctrl": wx.MOD_CONTROL,
    "control": wx.MOD_CONTROL,
    "shift": wx.MOD_SHIFT,
    "alt": wx.MOD_ALT,
    "win": wx.MOD_WIN,
    "cmd": wx.MOD_WIN,
}

# 支持 a-z, 0-9, f1-f24
KEY_MAP = {chr(i): ord(chr(i).upper()) for i in range(ord("a"), ord("z") + 1)}
for d in "0123456789":
    KEY_MAP[d] = ord(d)
for i in range(1, 25):
    KEY_MAP[f"f{i}"] = wx.WXK_F1 + (i - 1)


def parse_hotkey(hotkey_str: str):
    """
    输入: "ctrl+alt+1"
    输出: (modifiers, keycode) 或 None
    """
    if not hotkey_str:
        return None
    text = hotkey_str.strip().lower().replace(" ", "")
    parts = [p for p in text.split("+") if p]
    if len(parts) < 2:
        return None

    key_token = parts[-1]
    mod_tokens = parts[:-1]

    mods = 0
    for m in mod_tokens:
        if m not in MOD_MAP:
            return None
        mods |= MOD_MAP[m]

    keycode = KEY_MAP.get(key_token)
    if keycode is None:
        return None

    # 至少一个修饰键，避免劫持普通输入
    if mods == 0:
        return None
    return mods, keycode


def canonical_hotkey(hotkey_str: str):
    """
    规范化显示顺序:"ctrl+shift+alt+win+key"
    """
    p = parse_hotkey(hotkey_str)
    if not p:
        return ""
    mods, keycode = p
    mod_names = []
    if mods & wx.MOD_CONTROL:
        mod_names.append("ctrl")
    if mods & wx.MOD_SHIFT:
        mod_names.append("shift")
    if mods & wx.MOD_ALT:
        mod_names.append("alt")
    if mods & wx.MOD_WIN:
        mod_names.append("win")

    key_name = None
    for k, v in KEY_MAP.items():
        if v == keycode:
            key_name = k
            break
    if not key_name:
        return ""
    return "+".join(mod_names + [key_name])


# ---------------- UI ----------------
class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(700, 470))
        self.programs = load_config()

        # hotkey 注册映射: hotkey_id -> program_index
        self.hotkey_bindings = {}
        self.next_hotkey_id = 1000

        self.init_ui()
        self.Centre()

        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.Bind(wx.EVT_HOTKEY, self.on_hotkey)

        self.register_all_hotkeys()

    def init_ui(self):
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        title = wx.StaticText(panel, label="程序列表", style=wx.ALIGN_CENTER)
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        desc = wx.StaticText(
            panel,
            label="设置快捷键快速启动或切换程序窗口\n窗口在前台时按快捷键会最小化，再按恢复",
        )
        desc.SetForegroundColour(wx.Colour(100, 100, 100))
        sizer.Add(desc, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "程序名称", width=180)
        self.list_ctrl.InsertColumn(1, "快捷键", width=130)
        self.list_ctrl.InsertColumn(2, "程序路径", width=360)
        sizer.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 6)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        add_btn = wx.Button(panel, label="手动添加")
        add_btn.Bind(wx.EVT_BUTTON, self.on_manual_add)
        btn_sizer.Add(add_btn, 0, wx.ALL, 5)

        add_running_btn = wx.Button(panel, label="从运行程序添加")
        add_running_btn.Bind(wx.EVT_BUTTON, self.on_add_from_running)
        btn_sizer.Add(add_running_btn, 0, wx.ALL, 5)

        del_btn = wx.Button(panel, label="删除")
        del_btn.Bind(wx.EVT_BUTTON, self.on_delete)
        btn_sizer.Add(del_btn, 0, wx.ALL, 5)

        set_btn = wx.Button(panel, label="设置快捷键")
        set_btn.Bind(wx.EVT_BUTTON, self.on_set_hotkey)
        btn_sizer.Add(set_btn, 0, wx.ALL, 5)

        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER)

        self.status_text = wx.StaticText(panel, label="运行中... | 等待热键")
        sizer.Add(self.status_text, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        panel.SetSizer(sizer)
        self.refresh_list()

    def refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for p in self.programs:
            self.list_ctrl.Append([
                p.get("name", ""),
                p.get("hotkey", ""),
                p.get("path", "")
            ])

    # -------- hotkey register --------
    def unregister_all_hotkeys(self):
        for hid in list(self.hotkey_bindings.keys()):
            try:
                self.UnregisterHotKey(hid)
            except Exception:
                pass
        self.hotkey_bindings.clear()

    def register_all_hotkeys(self):
        self.unregister_all_hotkeys()
        used = set()  # 检测配置内重复 hotkey（规范化后）

        for idx, p in enumerate(self.programs):
            hk = canonical_hotkey(p.get("hotkey", ""))
            if not hk:
                continue
            if hk in used:
                print(f"[重复跳过] {p.get('name')} hotkey={hk}")
                continue
            used.add(hk)

            parsed = parse_hotkey(hk)
            if not parsed:
                continue
            mods, keycode = parsed

            hid = self.next_hotkey_id
            self.next_hotkey_id += 1

            ok = self.RegisterHotKey(hid, mods, keycode)
            if ok:
                self.hotkey_bindings[hid] = idx
            else:
                print(f"[注册失败] {p.get('name')} hotkey={hk}（可能被系统或其他程序占用）")

        self.status_text.SetLabel(f"已注册热键: {len(self.hotkey_bindings)} 个")

    def on_hotkey(self, event):
        hid = event.GetId()
        idx = self.hotkey_bindings.get(hid)
        if idx is None or idx >= len(self.programs):
            return

        p = self.programs[idx]
        toggle_program(p)
        self.status_text.SetLabel(f"已切换: {p.get('name', '')}")

    # -------- actions --------
    def on_manual_add(self, _):
        dialog = wx.Dialog(self, title="添加程序", size=(560, 200))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # name
        name_sizer = wx.BoxSizer(wx.HORIZONTAL)
        name_sizer.Add(wx.StaticText(panel, label="程序名称:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        name_ctrl = wx.TextCtrl(panel, size=(340, -1))
        name_sizer.Add(name_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(name_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # path
        path_sizer = wx.BoxSizer(wx.HORIZONTAL)
        path_sizer.Add(wx.StaticText(panel, label="程序路径:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        path_ctrl = wx.TextCtrl(panel, size=(300, -1))
        path_sizer.Add(path_ctrl, 1, wx.EXPAND | wx.ALL, 5)

        def on_browse(_e):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                p = dlg.GetPath()
                path_ctrl.SetValue(p)
                if not name_ctrl.GetValue().strip():
                    name_ctrl.SetValue(os.path.basename(p).replace(".exe", ""))
            dlg.Destroy()

        browse_btn = wx.Button(panel, label="浏览")
        browse_btn.Bind(wx.EVT_BUTTON, on_browse)
        path_sizer.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(path_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # buttons
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        def on_ok(_e):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            if not name or not path:
                wx.MessageBox("请填写完整名称和路径", "提示")
                return
            if not os.path.exists(path):
                wx.MessageBox("文件路径不存在", "提示")
                return

            self.programs.append({"name": name, "path": path, "hotkey": ""})
            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()
            dialog.Destroy()

        ok_btn = wx.Button(panel, wx.ID_OK, "添加")
        ok_btn.Bind(wx.EVT_BUTTON, on_ok)
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        cancel_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())

        btn_sizer.Add(ok_btn, 0, wx.ALL, 5)
        btn_sizer.Add(cancel_btn, 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_add_from_running(self, _):
        running = []
        seen = set()

        for proc in psutil.process_iter(["exe", "pid", "name"]):
            try:
                exe = proc.info.get("exe")
                if exe and exe.lower().endswith(".exe"):
                    key = exe.lower()
                    if key in seen:
                        continue
                    seen.add(key)

                    title = get_process_title(proc.info["pid"])
                    running.append({
                        "name": os.path.basename(exe).replace(".exe", ""),
                        "path": exe,
                        "title": title,
                    })
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
            except Exception:
                pass

        if not running:
            wx.MessageBox("没有找到可添加的运行程序", "提示")
            return

        running.sort(key=lambda x: (x["title"] or x["name"]).lower())

        dialog = wx.Dialog(self, title="选择运行程序", size=(560, 430))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer.Add(wx.StaticText(panel, label="双击选择程序添加:"), 0, wx.ALL, 8)
        listbox = wx.ListBox(panel)
        for p in running:
            txt = f"{p['title']} - {p['name']}" if p["title"] else p["name"]
            listbox.Append(txt)
        sizer.Add(listbox, 1, wx.EXPAND | wx.ALL, 8)

        def on_dclick(_e):
            sel = listbox.GetSelection()
            if sel == wx.NOT_FOUND:
                return
            item = running[sel]
            self.programs.append({"name": item["name"], "path": item["path"], "hotkey": ""})
            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()
            dialog.Destroy()

        listbox.Bind(wx.EVT_LISTBOX_DCLICK, on_dclick)

        close_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())
        sizer.Add(close_btn, 0, wx.ALL | wx.ALIGN_CENTER, 8)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_delete(self, _):
        sel = self.list_ctrl.GetFirstSelected()
        if sel < 0:
            wx.MessageBox("请先选择一项", "提示")
            return

        self.programs.pop(sel)
        save_config(self.programs)
        self.refresh_list()
        self.register_all_hotkeys()

    def on_set_hotkey(self, _):
        sel = self.list_ctrl.GetFirstSelected()
        if sel < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        p = self.programs[sel]
        current = p.get("hotkey", "")

        dlg = wx.TextEntryDialog(
            self,
            f"当前快捷键: {current}\n\n输入新快捷键（如 ctrl+alt+1, ctrl+shift+a, alt+f2）\n留空可清除。",
            "设置快捷键",
            current
        )
        if dlg.ShowModal() == wx.ID_OK:
            raw = dlg.GetValue().strip().lower()
            if raw == "":
                self.programs[sel]["hotkey"] = ""
                save_config(self.programs)
                self.refresh_list()
                self.register_all_hotkeys()
                wx.MessageBox("已清除快捷键", "提示")
                dlg.Destroy()
                return

            canon = canonical_hotkey(raw)
            if not canon:
                wx.MessageBox("快捷键格式无效，请使用如 ctrl+1 / alt+f2 / ctrl+shift+a", "错误")
                dlg.Destroy()
                return

            # 检查与其他项冲突
            for i, item in enumerate(self.programs):
                if i == sel:
                    continue
                if canonical_hotkey(item.get("hotkey", "")) == canon:
                    wx.MessageBox(f"快捷键冲突：已被「{item.get('name','')}」使用", "错误")
                    dlg.Destroy()
                    return

            self.programs[sel]["hotkey"] = canon
            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()
            wx.MessageBox(f"已设置快捷键: {canon}", "成功")

        dlg.Destroy()

    def on_close(self, event):
        self.unregister_all_hotkeys()
        event.Skip()


class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True


if __name__ == "__main__":
    app = QuickLauncherApp()
    app.MainLoop()
