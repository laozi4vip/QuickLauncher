# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器
功能：
- 绑定程序到快捷键
- 按快捷键启动/切换窗口（前台则最小化）
- 托盘常驻
- 开机启动（当前用户）
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

# ---------------- Windows API ----------------
user32 = ctypes.windll.user32
SW_MINIMIZE = 6
SW_RESTORE = 9

# ---------------- 路径与配置 ----------------
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
    APP_CMD = f'"{sys.executable}"'
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    APP_CMD = f'"{sys.executable}" "{os.path.abspath(__file__)}"'

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "QuickLauncher"


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('programs', [])
        except Exception as e:
            print("load_config error:", e)
    return []


def save_config(programs):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump({'programs': programs}, f, ensure_ascii=False, indent=4)


def find_window(exe_name):
    exe_name = exe_name.lower().replace('.exe', '')
    candidate_hwnds = []

    def callback(hwnd, param):
        try:
            if user32.IsWindowVisible(hwnd) and user32.IsWindowEnabled(hwnd):
                pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                try:
                    proc = psutil.Process(pid.value)
                    proc_name = proc.name().lower().replace('.exe', '')
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
    path = program.get('path', '')
    if not path:
        return

    exe_name = os.path.basename(path).lower().replace('.exe', '')
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


# ---------------- 开机启动 ----------------
def is_startup_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_READ) as key:
            val, _ = winreg.QueryValueEx(key, RUN_NAME)
            return val == APP_CMD
    except FileNotFoundError:
        return False
    except Exception as e:
        print("is_startup_enabled error:", e)
        return False


def set_startup(enable: bool):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
            if enable:
                winreg.SetValueEx(key, RUN_NAME, 0, winreg.REG_SZ, APP_CMD)
            else:
                try:
                    winreg.DeleteValue(key, RUN_NAME)
                except FileNotFoundError:
                    pass
        return True
    except Exception as e:
        print("set_startup error:", e)
        return False


# ---------------- 热键解析 ----------------
def parse_hotkey(text: str):
    """
    输入:
      - 单键: f1 ~ f12
      - 组合: alt+1 / ctrl+shift+a / win+f2
    输出: (modifiers, keycode) or (None, None)
    """
    if not text:
        return None, None

    s = text.strip().lower().replace(' ', '')
    parts = s.split('+')

    # 允许单键（仅 F1-F12）
    if len(parts) == 1:
        key_part = parts[0]
        mod_parts = []
    else:
        key_part = parts[-1]
        mod_parts = parts[:-1]

    modifiers = 0
    for m in mod_parts:
        if m == 'ctrl':
            modifiers |= wx.MOD_CONTROL
        elif m == 'shift':
            modifiers |= wx.MOD_SHIFT
        elif m == 'alt':
            modifiers |= wx.MOD_ALT
        elif m in ('win', 'cmd', 'meta'):
            modifiers |= wx.MOD_WIN
        else:
            return None, None

    f_map = {
        'f1': wx.WXK_F1, 'f2': wx.WXK_F2, 'f3': wx.WXK_F3, 'f4': wx.WXK_F4,
        'f5': wx.WXK_F5, 'f6': wx.WXK_F6, 'f7': wx.WXK_F7, 'f8': wx.WXK_F8,
        'f9': wx.WXK_F9, 'f10': wx.WXK_F10, 'f11': wx.WXK_F11, 'f12': wx.WXK_F12
    }

    # 单键只允许 F1-F12；组合可用 F 键或字母数字
    if key_part in f_map:
        keycode = f_map[key_part]
    elif len(mod_parts) > 0 and len(key_part) == 1 and key_part.isalnum():
        keycode = ord(key_part.upper())
    else:
        return None, None

    return modifiers, keycode


# ---------------- 托盘图标 ----------------
class TrayIcon(wx.adv.TaskBarIcon):
    TBMENU_SHOW = wx.NewIdRef()
    TBMENU_STARTUP = wx.NewIdRef()
    TBMENU_EXIT = wx.NewIdRef()

    def __init__(self, frame):
        super().__init__()
        self.frame = frame
        self.set_icon()
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK, self.on_show)

    def set_icon(self):
        bmp = wx.ArtProvider.GetBitmap(wx.ART_EXECUTABLE_FILE, wx.ART_OTHER, (16, 16))
        icon = wx.Icon()
        icon.CopyFromBitmap(bmp)
        self.SetIcon(icon, "QuickLauncher")

    def CreatePopupMenu(self):
        menu = wx.Menu()
        menu.Append(self.TBMENU_SHOW, "显示主窗口")
        startup_text = "关闭开机启动" if is_startup_enabled() else "开启开机启动"
        menu.Append(self.TBMENU_STARTUP, startup_text)
        menu.AppendSeparator()
        menu.Append(self.TBMENU_EXIT, "退出")

        self.Bind(wx.EVT_MENU, self.on_show, id=self.TBMENU_SHOW)
        self.Bind(wx.EVT_MENU, self.on_toggle_startup, id=self.TBMENU_STARTUP)
        self.Bind(wx.EVT_MENU, self.on_exit, id=self.TBMENU_EXIT)
        return menu

    def on_show(self, event):
        self.frame.show_from_tray()

    def on_toggle_startup(self, event):
        enabled = is_startup_enabled()
        ok = set_startup(not enabled)
        if ok:
            self.frame.update_startup_button()
            wx.MessageBox("已更新开机启动状态", "提示")
        else:
            wx.MessageBox("设置失败，请检查权限", "错误")

    def on_exit(self, event):
        self.frame.real_exit()


# ---------------- 主窗口 ----------------
class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(700, 500))
        self.programs = load_config()

        self.hotkey_id_base = 3000
        self.id_to_index = {}
        self.registered_ids = []

        self._really_close = False
        self.tray = TrayIcon(self)

        self.init_ui()
        self.Centre()

        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.Bind(wx.EVT_HOTKEY, self.on_hotkey)

        self.register_hotkeys()

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
        self.list_ctrl.InsertColumn(0, "程序名称", width=150)
        self.list_ctrl.InsertColumn(1, "快捷键", width=120)
        self.list_ctrl.InsertColumn(2, "程序路径", width=380)
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

        self.startup_btn = wx.Button(panel, label="")
        self.startup_btn.Bind(wx.EVT_BUTTON, self.on_toggle_startup)
        btn_sizer.Add(self.startup_btn, 0, wx.ALL, 5)
        self.update_startup_button()

        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER)

        self.status_text = wx.StaticText(panel, label="运行中（已启用托盘）")
        sizer.Add(self.status_text, 0, wx.ALL | wx.ALIGN_CENTER, 6)

        panel.SetSizer(sizer)
        self.refresh_list()

    def update_startup_button(self):
        self.startup_btn.SetLabel("关闭开机启动" if is_startup_enabled() else "开启开机启动")

    def refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for p in self.programs:
            self.list_ctrl.Append([p.get('name', ''), p.get('hotkey', ''), p.get('path', '')])

    # ---------- 热键 ----------
    def unregister_hotkeys(self):
        for hid in self.registered_ids:
            try:
                self.UnregisterHotKey(hid)
            except Exception:
                pass
        self.registered_ids.clear()
        self.id_to_index.clear()

    def register_hotkeys(self):
        self.unregister_hotkeys()
        for i, p in enumerate(self.programs):
            hk = p.get('hotkey', '').strip()
            if not hk:
                continue
            modifiers, keycode = parse_hotkey(hk)
            if modifiers is None:
                print(f"无效热键格式: {hk}")
                continue

            hid = self.hotkey_id_base + i
            ok = self.RegisterHotKey(hid, modifiers, keycode)
            if ok:
                self.registered_ids.append(hid)
                self.id_to_index[hid] = i
            else:
                print(f"注册热键失败（可能冲突）: {hk}")

    def on_hotkey(self, event):
        hid = event.GetId()
        idx = self.id_to_index.get(hid)
        if idx is None or idx >= len(self.programs):
            return
        program = self.programs[idx]
        toggle_program(program)
        self.status_text.SetLabel(f"已切换: {program.get('name', '')}")

    # ---------- 事件 ----------
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
        path_ctrl = wx.TextCtrl(panel, size=(280, -1))
        path_sizer.Add(path_ctrl, 1, wx.EXPAND | wx.ALL, 5)

        def on_browse(evt):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                p = dlg.GetPath()
                path_ctrl.SetValue(p)
                if not name_ctrl.GetValue():
                    name_ctrl.SetValue(os.path.basename(p).replace('.exe', ''))
            dlg.Destroy()

        browse_btn = wx.Button(panel, label="浏览")
        browse_btn.Bind(wx.EVT_BUTTON, on_browse)
        path_sizer.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(path_sizer, 0, wx.EXPAND | wx.ALL, 5)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        def on_ok(evt):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            if name and path:
                self.programs.append({'name': name, 'path': path, 'hotkey': ''})
                save_config(self.programs)
                self.refresh_list()
                self.register_hotkeys()
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

    def on_add_from_running(self, event):
        running = []
        for proc in psutil.process_iter(['exe', 'pid', 'name']):
            try:
                exe = proc.info.get('exe')
                if exe and exe.lower().endswith('.exe'):
                    exe_name = os.path.basename(exe)
                    if not any(p['exe'] == exe_name for p in running):
                        title = get_process_title(proc.info['pid'])
                        running.append({
                            'name': exe_name.replace('.exe', ''),
                            'exe': exe_name,
                            'path': exe,
                            'title': title
                        })
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

        if not running:
            wx.MessageBox("没有找到运行程序", "提示")
            return

        running.sort(key=lambda x: x['title'] or x['name'])

        dialog = wx.Dialog(self, title="选择程序", size=(560, 420))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(wx.StaticText(panel, label="双击选择程序添加:"), 0, wx.ALL, 5)

        listbox = wx.ListBox(panel, size=(-1, 320))
        for p in running:
            display = f"{p['title']} - {p['name']}" if p['title'] else p['name']
            listbox.Append(display)
        sizer.Add(listbox, 1, wx.EXPAND | wx.ALL, 5)

        def on_double_click(evt):
            sel = listbox.GetSelection()
            if sel != wx.NOT_FOUND:
                self.programs.append({
                    'name': running[sel]['name'],
                    'path': running[sel]['path'],
                    'hotkey': ''
                })
                save_config(self.programs)
                self.refresh_list()
                self.register_hotkeys()
                dialog.Destroy()

        listbox.Bind(wx.EVT_LISTBOX_DCLICK, on_double_click)
        close_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())
        sizer.Add(close_btn, 0, wx.ALL | wx.ALIGN_CENTER, 6)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_delete(self, event):
        selection = self.list_ctrl.GetFirstSelected()
        if selection >= 0:
            self.programs.pop(selection)
            save_config(self.programs)
            self.refresh_list()
            self.register_hotkeys()

    def on_set_hotkey(self, event):
        selection = self.list_ctrl.GetFirstSelected()
        if selection < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        current_hotkey = self.programs[selection].get('hotkey', '')
        dlg = wx.TextEntryDialog(
            self,
            f"当前快捷键: {current_hotkey}\n\n输入新快捷键（如 f1, alt+1, ctrl+shift+a, win+f2）:",
            "设置快捷键",
            current_hotkey
        )

        if dlg.ShowModal() == wx.ID_OK:
            hotkey = dlg.GetValue().strip().lower()
            if hotkey:
                m, k = parse_hotkey(hotkey)
                if m is None:
                    wx.MessageBox("热键格式无效", "错误")
                else:
                    self.programs[selection]['hotkey'] = hotkey
                    save_config(self.programs)
                    self.refresh_list()
                    self.register_hotkeys()
                    wx.MessageBox(f"已设置快捷键: {hotkey}", "成功")
        dlg.Destroy()

    def on_toggle_startup(self, event):
        enabled = is_startup_enabled()
        ok = set_startup(not enabled)
        if ok:
            self.update_startup_button()
            wx.MessageBox("已更新开机启动状态", "提示")
        else:
            wx.MessageBox("设置失败，请检查权限", "错误")

    # ---------- 托盘/关闭 ----------
    def on_close(self, event):
        if self._really_close:
            self.unregister_hotkeys()
            if self.tray:
                self.tray.RemoveIcon()
                self.tray.Destroy()
                self.tray = None
            event.Skip()
        else:
            self.Hide()
            self.status_text.SetLabel("已最小化到托盘（右下角图标）")
            event.Veto()

    def show_from_tray(self):
        self.Show()
        self.Raise()
        self.Iconize(False)

    def real_exit(self):
        self._really_close = True
        self.Close()


class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True


if __name__ == '__main__':
    app = QuickLauncherApp()
    app.MainLoop()
