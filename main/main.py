# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器（稳定版）
功能：
- 绑定程序到全局快捷键
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

# ========== Windows API ==========
user32 = ctypes.windll.user32
SW_MINIMIZE = 6
SW_RESTORE = 9

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008

# ========== 配置路径 ==========
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")


# ========== 配置读写 ==========
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
    print(f"Config saved: {CONFIG_FILE}")


# ========== 窗口查找与切换 ==========
def find_window(exe_name):
    exe_name = exe_name.lower().replace(".exe", "")
    candidate_hwnds = []

    def callback(hwnd, lparam):
        try:
            if user32.IsWindowVisible(hwnd) and user32.IsWindowEnabled(hwnd):
                pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                try:
                    proc = psutil.Process(pid.value)
                    proc_name = proc.name().lower().replace(".exe", "")
                    if proc_name == exe_name:
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
    path = program.get("path", "").strip()
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
            try:
                subprocess.Popen(path)
            except Exception as e:
                print("启动失败:", e)


def get_process_title(pid):
    titles = []

    def callback(hwnd, lparam):
        try:
            if hwnd and user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
                window_pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                if window_pid.value == pid:
                    length = user32.GetWindowTextLengthW(hwnd)
                    if length > 0:
                        buff = ctypes.create_unicode_buffer(length + 1)
                        user32.GetWindowTextW(hwnd, buff, length + 1)
                        text = buff.value.strip()
                        if text:
                            titles.append(text)
        except Exception:
            pass
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    user32.EnumWindows(WNDENUMPROC(callback), 0)

    return titles[0] if titles else ""


# ========== 热键解析 ==========
def parse_hotkey(hotkey: str):
    """
    输入示例: alt+1, ctrl+shift+a, win+f2
    返回: (modifiers, vk) 或 None
    """
    if not hotkey:
        return None

    hk = hotkey.strip().lower().replace(" ", "")
    parts = hk.split("+")
    if len(parts) < 2:
        return None

    key = parts[-1]
    mods = parts[:-1]

    mod_val = 0
    for m in mods:
        if m == "ctrl":
            mod_val |= MOD_CONTROL
        elif m == "alt":
            mod_val |= MOD_ALT
        elif m == "shift":
            mod_val |= MOD_SHIFT
        elif m == "win":
            mod_val |= MOD_WIN
        else:
            return None

    vk = None
    if len(key) == 1 and "a" <= key <= "z":
        vk = ord(key.upper())
    elif len(key) == 1 and "0" <= key <= "9":
        vk = ord(key)
    elif key.startswith("f") and key[1:].isdigit():
        fn = int(key[1:])
        if 1 <= fn <= 12:
            vk = 0x70 + (fn - 1)  # VK_F1 = 0x70

    if vk is None:
        return None
    return mod_val, vk

    return None

class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(700, 480))

        self.programs = load_config()
        self.hotkey_id_to_index = {}   # hotkey_id -> program index
        self.next_hotkey_id = 1000

        self.init_ui()
        self.Centre()

        self.register_all_hotkeys()
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def init_ui(self):
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        title = wx.StaticText(panel, label="程序列表", style=wx.ALIGN_CENTER)
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        desc = wx.StaticText(panel, label="设置快捷键快速启动/切换程序\n窗口在前台时按快捷键会最小化，再按恢复")
        desc.SetForegroundColour(wx.Colour(100, 100, 100))
        sizer.Add(desc, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "程序名称", width=180)
        self.list_ctrl.InsertColumn(1, "快捷键", width=140)
        self.list_ctrl.InsertColumn(2, "程序路径", width=340)
        sizer.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

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

        self.status_text = wx.StaticText(panel, label="运行中... | 全局热键已启用")
        sizer.Add(self.status_text, 0, wx.ALL | wx.ALIGN_CENTER, 8)

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

    def on_manual_add(self, event):
        dialog = wx.Dialog(self, title="添加程序", size=(560, 200))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        name_sizer = wx.BoxSizer(wx.HORIZONTAL)
        name_sizer.Add(wx.StaticText(panel, label="程序名称:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        name_ctrl = wx.TextCtrl(panel, size=(360, -1))
        name_sizer.Add(name_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(name_sizer, 0, wx.EXPAND | wx.ALL, 5)

        path_sizer = wx.BoxSizer(wx.HORIZONTAL)
        path_sizer.Add(wx.StaticText(panel, label="程序路径:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        path_ctrl = wx.TextCtrl(panel, size=(300, -1))
        path_sizer.Add(path_ctrl, 1, wx.EXPAND | wx.ALL, 5)

        def on_browse(_):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="可执行文件 (*.exe)|*.exe", style=wx.FD_OPEN)
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

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        def on_ok(_):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            if not name or not path:
                wx.MessageBox("名称和路径不能为空", "提示")
                return
            self.programs.append({"name": name, "path": path, "hotkey": ""})
            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()
            dialog.EndModal(wx.ID_OK)

        ok_btn = wx.Button(panel, wx.ID_OK, "添加")
        ok_btn.Bind(wx.EVT_BUTTON, on_ok)
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        cancel_btn.Bind(wx.EVT_BUTTON, lambda _: dialog.EndModal(wx.ID_CANCEL))
        btn_sizer.Add(ok_btn, 0, wx.ALL, 5)
        btn_sizer.Add(cancel_btn, 0, wx.ALL, 5)

        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_add_from_running(self, event):
        running = []
        seen_path = set()

        for proc in psutil.process_iter(["exe", "pid", "name"]):
            try:
                exe = proc.info.get("exe")
                if exe and exe.lower().endswith(".exe") and exe not in seen_path:
                    seen_path.add(exe)
                    title = get_process_title(proc.info["pid"])
                    running.append({
                        "name": os.path.basename(exe).replace(".exe", ""),
                        "path": exe,
                        "title": title
                    })
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
            except Exception:
                pass

        if not running:
            wx.MessageBox("没有找到可添加的运行程序", "提示")
            return

        running.sort(key=lambda x: (x["title"] or x["name"]).lower())

        dialog = wx.Dialog(self, title="选择程序", size=(580, 420))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer.Add(wx.StaticText(panel, label="双击选择一个程序添加："), 0, wx.ALL, 8)

        listbox = wx.ListBox(panel, size=(-1, 300))
        for p in running:
            display = f"{p['title']} - {p['name']}" if p["title"] else p["name"]
            listbox.Append(display)
        sizer.Add(listbox, 1, wx.EXPAND | wx.ALL, 8)

        def on_double_click(_):
            idx = listbox.GetSelection()
            if idx == wx.NOT_FOUND:
                return
            self.programs.append({
                "name": running[idx]["name"],
                "path": running[idx]["path"],
                "hotkey": ""
            })
            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()
            dialog.EndModal(wx.ID_OK)

        listbox.Bind(wx.EVT_LISTBOX_DCLICK, on_double_click)

        close_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
        close_btn.Bind(wx.EVT_BUTTON, lambda _: dialog.EndModal(wx.ID_CANCEL))
        sizer.Add(close_btn, 0, wx.ALL | wx.ALIGN_CENTER, 8)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_delete(self, event):
        selection = self.list_ctrl.GetFirstSelected()
        if selection < 0:
            wx.MessageBox("请先选择一项", "提示")
            return
        self.programs.pop(selection)
        save_config(self.programs)
        self.refresh_list()
        self.register_all_hotkeys()

    def on_set_hotkey(self, event):
        selection = self.list_ctrl.GetFirstSelected()
        if selection < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        program = self.programs[selection]
        current_hotkey = program.get("hotkey", "")

        dlg = wx.TextEntryDialog(
            self,
            f"当前快捷键: {current_hotkey}\n\n输入新快捷键（如 alt+1, ctrl+shift+a, win+f2）\n留空可清除：",
            "设置快捷键",
            current_hotkey
        )

        if dlg.ShowModal() == wx.ID_OK:
            hotkey = dlg.GetValue().strip().lower()
            if hotkey:
                if parse_hotkey(hotkey) is None:
                    wx.MessageBox("快捷键格式无效", "错误")
                    dlg.Destroy()
                    return
                # 检查重复
                for i, p in enumerate(self.programs):
                    if i != selection and p.get("hotkey", "").strip().lower() == hotkey:
                        wx.MessageBox("该快捷键已被其他程序占用", "提示")
                        dlg.Destroy()
                        return
                self.programs[selection]["hotkey"] = hotkey
            else:
                self.programs[selection]["hotkey"] = ""  # 清除

            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()
            wx.MessageBox("快捷键已更新", "成功")

        dlg.Destroy()

    # ========== 热键注册 ==========
    def unregister_all_hotkeys(self):
        for hotkey_id in list(self.hotkey_id_to_index.keys()):
            try:
                user32.UnregisterHotKey(int(self.GetHandle()), hotkey_id)
            except Exception:
                pass
        self.hotkey_id_to_index.clear()

    def register_all_hotkeys(self):
        self.unregister_all_hotkeys()
        self.programs = load_config()  # 与文件保持一致

        self.next_hotkey_id = 1000
        failed = []

        for idx, p in enumerate(self.programs):
            hk = p.get("hotkey", "").strip().lower()
            if not hk:
                continue

            parsed = parse_hotkey(hk)
            if parsed is None:
                failed.append(f"{p.get('name', '')}: 格式无效 ({hk})")
                continue

            mod, vk = parsed
            hotkey_id = self.next_hotkey_id
            self.next_hotkey_id += 1

            ok = user32.RegisterHotKey(int(self.GetHandle()), hotkey_id, mod, vk)
            if ok:
                self.hotkey_id_to_index[hotkey_id] = idx
                self.Bind(wx.EVT_HOTKEY, self.on_hotkey, id=hotkey_id)
            else:
                failed.append(f"{p.get('name', '')}: 注册失败（可能与系统/其他软件冲突）({hk})")

        if failed:
            wx.MessageBox("\n".join(failed), "部分热键未注册", wx.OK | wx.ICON_WARNING)

        self.status_text.SetLabel(f"运行中... | 已注册热键: {len(self.hotkey_id_to_index)}")

    def on_hotkey(self, event):
        hotkey_id = event.GetId()
        idx = self.hotkey_id_to_index.get(hotkey_id)
        if idx is None or idx >= len(self.programs):
            return

        program = self.programs[idx]
        toggle_program(program)
        self.status_text.SetLabel(f"已切换: {program.get('name', '')}")

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
