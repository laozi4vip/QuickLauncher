"""
QuickLauncher - Windows任务栏快捷启动器（稳定版）
功能：
- 绑定程序到快捷键
- 按快捷键启动或切换窗口
- 窗口在前台时最小化，再按恢复
"""

import wx
import os
import json
import subprocess
import psutil
import ctypes
import sys

# Windows API
user32 = ctypes.windll.user32
SW_MINIMIZE = 6
SW_RESTORE = 9

# 配置路径
if getattr(sys, 'frozen', False):
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
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"programs": programs}, f, ensure_ascii=False, indent=4)
    print("Config saved:", CONFIG_FILE)


def find_window(exe_name):
    exe_name = exe_name.lower().replace(".exe", "")
    candidate_hwnds = []

    def callback(hwnd, _):
        try:
            if user32.IsWindowVisible(hwnd) and user32.IsWindowEnabled(hwnd):
                pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                try:
                    proc = psutil.Process(pid.value)
                    proc_name = proc.name().lower().replace(".exe", "")
                    if exe_name == proc_name:
                        candidate_hwnds.append(hwnd)
                except Exception:
                    pass
        except Exception:
            pass
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    user32.EnumWindows(WNDENUMPROC(callback), 0)

    fg_hwnd = user32.GetForegroundWindow()
    if fg_hwnd in candidate_hwnds:
        return fg_hwnd
    return candidate_hwnds[0] if candidate_hwnds else None


def toggle_program(program):
    path = program.get("path", "")
    if not path:
        return

    exe_name = os.path.basename(path).lower().replace(".exe", "")
    hwnd = find_window(exe_name)

    if hwnd:
        fg_hwnd = user32.GetForegroundWindow()
        if hwnd == fg_hwnd:
            user32.ShowWindow(hwnd, SW_MINIMIZE)
        else:
            if user32.IsIconic(hwnd):
                user32.ShowWindow(hwnd, SW_RESTORE)
            user32.SetForegroundWindow(hwnd)
    else:
        if os.path.exists(path):
            subprocess.Popen(path)


def get_process_title(pid):
    try:
        windows = []

        def callback(hwnd, wins):
            try:
                if hwnd and user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
                    window_pid = ctypes.c_ulong()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                    if window_pid.value == pid:
                        length = user32.GetWindowTextLengthW(hwnd)
                        if length > 0:
                            title = ctypes.create_unicode_buffer(length + 1)
                            user32.GetWindowTextW(hwnd, title, length + 1)
                            wins.append(title.value)
            except Exception:
                pass
            return True

        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)
        user32.EnumWindows(WNDENUMPROC(callback), windows)

        for title in windows:
            if title:
                return title
    except Exception:
        pass
    return ""


def parse_hotkey(hotkey_str):
    """
    把 'ctrl+alt+1' 解析成 (modifiers, keycode)
    """
    if not hotkey_str:
        return None, None

    hotkey = hotkey_str.strip().lower().replace(" ", "")
    parts = hotkey.split("+")
    if len(parts) < 2:
        return None, None

    key_part = parts[-1]
    mod_parts = parts[:-1]

    modifiers = 0
    for m in mod_parts:
        if m == "ctrl":
            modifiers |= wx.MOD_CONTROL
        elif m == "shift":
            modifiers |= wx.MOD_SHIFT
        elif m == "alt":
            modifiers |= wx.MOD_ALT
        else:
            return None, None

    # keycode
    if len(key_part) == 1 and key_part.isalpha():
        keycode = ord(key_part.upper())
    elif len(key_part) == 1 and key_part.isdigit():
        keycode = ord(key_part)
    elif key_part.startswith("f") and key_part[1:].isdigit():
        fn = int(key_part[1:])
        if 1 <= fn <= 12:
            keycode = getattr(wx, f"WXK_F{fn}")
        else:
            return None, None
    else:
        return None, None

    return modifiers, keycode


class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(650, 460))
        self.programs = load_config()

        self.hotkey_id_to_index = {}
        self.registered_hotkey_ids = []

        self.init_ui()
        self.Centre()

        self.bind_hotkeys()
        self.Bind(wx.EVT_HOTKEY, self.on_hotkey_triggered)
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def init_ui(self):
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        title = wx.StaticText(panel, label="程序列表", style=wx.ALIGN_CENTER)
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        desc = wx.StaticText(panel, label="设置快捷键来快速启动或切换程序窗口\n窗口在前台时按快捷键会最小化，再按恢复")
        desc.SetForegroundColour(wx.Colour(100, 100, 100))
        sizer.Add(desc, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "程序名称", width=150)
        self.list_ctrl.InsertColumn(1, "快捷键", width=120)
        self.list_ctrl.InsertColumn(2, "程序路径", width=330)
        sizer.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 5)

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
        self.status_text = wx.StaticText(panel, label="运行中... | 已启用系统级热键监听")
        sizer.Add(self.status_text, 0, wx.ALL | wx.ALIGN_CENTER, 8)

        panel.SetSizer(sizer)
        self.refresh_list()

    def refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for p in self.programs:
            self.list_ctrl.Append([p.get("name", ""), p.get("hotkey", ""), p.get("path", "")])

    def on_manual_add(self, event):
        dialog = wx.Dialog(self, title="添加程序", size=(520, 190))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        name_sizer = wx.BoxSizer(wx.HORIZONTAL)
        name_sizer.Add(wx.StaticText(panel, label="程序名称:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        name_ctrl = wx.TextCtrl(panel, size=(320, -1))
        name_sizer.Add(name_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(name_sizer, 0, wx.EXPAND | wx.ALL, 5)

        path_sizer = wx.BoxSizer(wx.HORIZONTAL)
        path_sizer.Add(wx.StaticText(panel, label="程序路径:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        path_ctrl = wx.TextCtrl(panel, size=(260, -1))
        path_sizer.Add(path_ctrl, 1, wx.EXPAND | wx.ALL, 5)

        def on_browse(evt):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                path = dlg.GetPath()
                path_ctrl.SetValue(path)
                if not name_ctrl.GetValue():
                    name_ctrl.SetValue(os.path.basename(path).replace(".exe", ""))
            dlg.Destroy()

        browse_btn = wx.Button(panel, label="浏览")
        browse_btn.Bind(wx.EVT_BUTTON, on_browse)
        path_sizer.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(path_sizer, 0, wx.EXPAND | wx.ALL, 5)

        def on_ok(evt):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            if not name or not path:
                wx.MessageBox("名称和路径不能为空", "提示")
                return
            self.programs.append({"name": name, "path": path, "hotkey": ""})
            save_config(self.programs)
            self.refresh_list()
            self.bind_hotkeys()
            dialog.Destroy()

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
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

    def on_add_from_running(self, event):
        running = []
        for proc in psutil.process_iter(["exe", "pid", "name"]):
            try:
                exe = proc.info.get("exe")
                if exe and exe.lower().endswith(".exe"):
                    exe_name = os.path.basename(exe)
                    if not any(p["exe"].lower() == exe_name.lower() for p in running):
                        title = get_process_title(proc.info["pid"])
                        running.append({
                            "name": exe_name.replace(".exe", ""),
                            "exe": exe_name,
                            "path": exe,
                            "title": title
                        })
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

        if not running:
            wx.MessageBox("没有找到运行的程序", "提示")
            return

        running.sort(key=lambda x: (x["title"] or x["name"]).lower())

        dialog = wx.Dialog(self, title="选择程序", size=(520, 420))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer.Add(wx.StaticText(panel, label="双击选择程序添加:"), 0, wx.ALL, 5)
        listbox = wx.ListBox(panel, size=(-1, 320))
        for p in running:
            display = f"{p['title']} - {p['name']}" if p["title"] else p["name"]
            listbox.Append(display)
        sizer.Add(listbox, 1, wx.EXPAND | wx.ALL, 5)

        def on_double_click(evt):
            i = listbox.GetSelection()
            if i != wx.NOT_FOUND:
                self.programs.append({
                    "name": running[i]["name"],
                    "path": running[i]["path"],
                    "hotkey": ""
                })
                save_config(self.programs)
                self.refresh_list()
                self.bind_hotkeys()
                dialog.Destroy()

        listbox.Bind(wx.EVT_LISTBOX_DCLICK, on_double_click)

        close_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())
        sizer.Add(close_btn, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_delete(self, event):
        sel = self.list_ctrl.GetFirstSelected()
        if sel >= 0:
            self.programs.pop(sel)
            save_config(self.programs)
            self.refresh_list()
            self.bind_hotkeys()

    def on_set_hotkey(self, event):
        sel = self.list_ctrl.GetFirstSelected()
        if sel < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        current = self.programs[sel].get("hotkey", "")
        dlg = wx.TextEntryDialog(
            self,
            f"当前快捷键: {current}\n\n输入新快捷键（例如 ctrl+1, alt+a, ctrl+shift+f2）:",
            "设置快捷键",
            current
        )

        if dlg.ShowModal() == wx.ID_OK:
            hotkey = dlg.GetValue().strip().lower().replace(" ", "")
            if hotkey:
                # 检查重复
                for i, p in enumerate(self.programs):
                    if i != sel and p.get("hotkey", "").strip().lower().replace(" ", "") == hotkey:
                        wx.MessageBox(f"快捷键已被 [{p.get('name','')}] 占用", "冲突")
                        dlg.Destroy()
                        return

                mod, key = parse_hotkey(hotkey)
                if mod is None or key is None:
                    wx.MessageBox("快捷键格式无效，请使用 ctrl/shift/alt + 字母/数字/F1-F12", "错误")
                    dlg.Destroy()
                    return

                self.programs[sel]["hotkey"] = hotkey
                save_config(self.programs)
                self.refresh_list()
                self.bind_hotkeys()
                wx.MessageBox(f"已设置快捷键: {hotkey}", "成功")

        dlg.Destroy()  # 修复：只销毁 dlg，不要调用不存在的 dialog

    def unbind_hotkeys(self):
        for hotkey_id in self.registered_hotkey_ids:
            try:
                self.UnregisterHotKey(hotkey_id)
            except Exception:
                pass
        self.registered_hotkey_ids.clear()
        self.hotkey_id_to_index.clear()

    def bind_hotkeys(self):
        self.unbind_hotkeys()

        base_id = 1000
        ok_count = 0
        fail_msgs = []

        for i, p in enumerate(self.programs):
            hotkey = p.get("hotkey", "").strip().lower().replace(" ", "")
            if not hotkey:
                continue

            mod, key = parse_hotkey(hotkey)
            if mod is None or key is None:
                fail_msgs.append(f"{p.get('name','')}：格式无效({hotkey})")
                continue

            hotkey_id = base_id + i
            try:
                if self.RegisterHotKey(hotkey_id, mod, key):
                    self.registered_hotkey_ids.append(hotkey_id)
                    self.hotkey_id_to_index[hotkey_id] = i
                    ok_count += 1
                else:
                    fail_msgs.append(f"{p.get('name','')}：注册失败（可能被系统/其他软件占用）{hotkey}")
            except Exception as e:
                fail_msgs.append(f"{p.get('name','')}：{e}")

        status = f"运行中... | 已注册 {ok_count} 个热键"
        if fail_msgs:
            status += f" | 失败 {len(fail_msgs)} 个"
            print("\n".join(fail_msgs))
        self.status_text.SetLabel(status)

    def on_hotkey_triggered(self, event):
        hotkey_id = event.GetId()
        idx = self.hotkey_id_to_index.get(hotkey_id)
        if idx is None or idx >= len(self.programs):
            return

        program = self.programs[idx]
        try:
            toggle_program(program)
            self.status_text.SetLabel(f"已切换: {program.get('name', '')}")
        except Exception as e:
            self.status_text.SetLabel(f"切换失败: {e}")
            print("toggle error:", e)

    def on_close(self, event):
        self.unbind_hotkeys()
        self.Destroy()


class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True


if __name__ == "__main__":
    app = QuickLauncherApp()
    app.MainLoop()
