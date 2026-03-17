# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器

功能：
- 绑定程序到快捷键
- 按快捷键启动或切换窗口
- 窗口在前台时按快捷键最小化，再按恢复
- 支持从运行中的程序添加
- 使用 Windows 全局热键（不再轮询键盘，更稳定）

依赖：
pip install wxPython psutil
"""

import wx
import os
import json
import subprocess
import psutil
import ctypes
import sys
import re

# =========================
# Windows API
# =========================
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

SW_MINIMIZE = 6
SW_RESTORE = 9
SW_SHOW = 5

WM_HOTKEY = 0x0312

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008

HWND_TOPMOST = -1
HWND_NOTOPMOST = -2

SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_SHOWWINDOW = 0x0040

# 设置 Windows API 参数/返回值，减少 ctypes 出错概率
user32.EnumWindows.argtypes = [
    ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p),
    ctypes.c_void_p
]
user32.EnumWindows.restype = ctypes.c_bool

user32.IsWindow.argtypes = [ctypes.c_void_p]
user32.IsWindow.restype = ctypes.c_bool

user32.IsWindowVisible.argtypes = [ctypes.c_void_p]
user32.IsWindowVisible.restype = ctypes.c_bool

user32.IsIconic.argtypes = [ctypes.c_void_p]
user32.IsIconic.restype = ctypes.c_bool

user32.GetWindowTextLengthW.argtypes = [ctypes.c_void_p]
user32.GetWindowTextLengthW.restype = ctypes.c_int

user32.GetWindowTextW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int

user32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
user32.GetWindowThreadProcessId.restype = ctypes.c_ulong

user32.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
user32.ShowWindow.restype = ctypes.c_bool

user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
user32.SetForegroundWindow.restype = ctypes.c_bool

user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = ctypes.c_void_p

user32.SetWindowPos.argtypes = [
    ctypes.c_void_p, ctypes.c_void_p,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint
]
user32.SetWindowPos.restype = ctypes.c_bool

user32.RegisterHotKey.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_uint, ctypes.c_uint]
user32.RegisterHotKey.restype = ctypes.c_bool

user32.UnregisterHotKey.argtypes = [ctypes.c_void_p, ctypes.c_int]
user32.UnregisterHotKey.restype = ctypes.c_bool


# =========================
# 配置路径
# =========================
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")


# =========================
# 配置读写
# =========================
def load_config():
    """加载配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                programs = data.get('programs', [])
                if isinstance(programs, list):
                    return programs
        except Exception:
            pass
    return []


def save_config(programs):
    """保存配置"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({'programs': programs}, f, ensure_ascii=False, indent=4)
    except Exception as e:
        wx.MessageBox(f"保存配置失败：{e}", "错误", wx.OK | wx.ICON_ERROR)


# =========================
# 窗口相关工具
# =========================
def get_window_title(hwnd):
    try:
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        return buf.value.strip()
    except Exception:
        return ""


def get_pid_by_hwnd(hwnd):
    try:
        pid = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value
    except Exception:
        return 0


def is_real_window(hwnd):
    """过滤不可见/无标题窗口"""
    try:
        if not hwnd or not user32.IsWindow(hwnd):
            return False
        if not user32.IsWindowVisible(hwnd):
            return False
        title = get_window_title(hwnd)
        if not title:
            return False
        return True
    except Exception:
        return False


def enum_windows():
    """枚举所有可见顶层窗口"""
    result = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def callback(hwnd, lParam):
        try:
            if is_real_window(hwnd):
                result.append(hwnd)
        except Exception:
            pass
        return True

    try:
        user32.EnumWindows(callback, 0)
    except Exception:
        pass

    return result


def find_windows_by_exe(exe_name):
    """根据 exe 名称查找所有窗口"""
    exe_name = exe_name.lower().replace('.exe', '')
    matched = []

    for hwnd in enum_windows():
        try:
            pid = get_pid_by_hwnd(hwnd)
            if not pid:
                continue
            proc = psutil.Process(pid)
            proc_name = proc.name().lower().replace('.exe', '')
            if proc_name == exe_name or exe_name in proc_name or proc_name in exe_name:
                matched.append(hwnd)
        except Exception:
            pass

    return matched


def find_best_window(exe_name):
    """查找最合适的窗口：优先前台匹配窗口，其次第一个匹配窗口"""
    windows = find_windows_by_exe(exe_name)
    if not windows:
        return None

    fg = user32.GetForegroundWindow()
    for hwnd in windows:
        if hwnd == fg:
            return hwnd

    return windows[0]


def is_minimized(hwnd):
    try:
        return bool(user32.IsIconic(hwnd))
    except Exception:
        return False


def restore_window(hwnd):
    try:
        if is_minimized(hwnd):
            user32.ShowWindow(hwnd, SW_RESTORE)
        else:
            user32.ShowWindow(hwnd, SW_SHOW)

        # 有时直接 SetForegroundWindow 会失败，先临时置顶再取消
        user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
        user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
        user32.SetForegroundWindow(hwnd)
    except Exception:
        pass


def minimize_window(hwnd):
    try:
        user32.ShowWindow(hwnd, SW_MINIMIZE)
    except Exception:
        pass


def is_foreground_window(hwnd):
    try:
        return hwnd == user32.GetForegroundWindow()
    except Exception:
        return False


def toggle_program(program):
    """切换程序窗口状态"""
    path = program.get('path', '').strip()
    if not path:
        return False, "程序路径为空"

    exe_name = os.path.basename(path).lower().replace('.exe', '')
    hwnd = find_best_window(exe_name)

    if hwnd:
        if is_foreground_window(hwnd) and not is_minimized(hwnd):
            minimize_window(hwnd)
            return True, f"已最小化：{program.get('name', exe_name)}"
        else:
            restore_window(hwnd)
            return True, f"已切换到：{program.get('name', exe_name)}"
    else:
        if os.path.exists(path):
            try:
                subprocess.Popen(path)
                return True, f"已启动：{program.get('name', exe_name)}"
            except Exception as e:
                return False, f"启动失败：{e}"
        else:
            return False, f"文件不存在：{path}"


def get_process_title(pid):
    """获取进程的窗口标题"""
    titles = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def callback(hwnd, lParam):
        try:
            if hwnd and user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
                window_pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                if window_pid.value == pid:
                    title = get_window_title(hwnd)
                    if title:
                        titles.append(title)
        except Exception:
            pass
        return True

    try:
        user32.EnumWindows(callback, 0)
    except Exception:
        pass

    return titles[0] if titles else ""


# =========================
# 热键处理
# =========================
VK_MAP = {
    '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34,
    '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,
    'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45,
    'f': 0x46, 'g': 0x47, 'h': 0x48, 'i': 0x49, 'j': 0x4A,
    'k': 0x4B, 'l': 0x4C, 'm': 0x4D, 'n': 0x4E, 'o': 0x4F,
    'p': 0x50, 'q': 0x51, 'r': 0x52, 's': 0x53, 't': 0x54,
    'u': 0x55, 'v': 0x56, 'w': 0x57, 'x': 0x58, 'y': 0x59,
    'z': 0x5A,
    'f1': 0x70, 'f2': 0x71, 'f3': 0x72, 'f4': 0x73,
    'f5': 0x74, 'f6': 0x75, 'f7': 0x76, 'f8': 0x77,
    'f9': 0x78, 'f10': 0x79, 'f11': 0x7A, 'f12': 0x7B,
}

VALID_MODIFIERS = {
    'ctrl': MOD_CONTROL,
    'shift': MOD_SHIFT,
    'alt': MOD_ALT,
    'win': MOD_WIN,
}


def normalize_hotkey(hotkey):
    """
    标准化热键字符串
    例如：
    Alt + 1 => alt+1
    shift+ctrl+a => ctrl+shift+a
    """
    if not hotkey:
        return ""

    parts = [p.strip().lower() for p in hotkey.split('+') if p.strip()]
    if len(parts) < 2:
        return ""

    key = parts[-1]
    modifiers = parts[:-1]

    if key not in VK_MAP:
        return ""

    mod_set = []
    seen = set()
    for m in modifiers:
        if m in VALID_MODIFIERS and m not in seen:
            seen.add(m)
            mod_set.append(m)

    if not mod_set:
        return ""

    order = ['ctrl', 'shift', 'alt', 'win']
    mod_set = [m for m in order if m in mod_set]

    return '+'.join(mod_set + [key])


def parse_hotkey(hotkey):
    """把字符串热键转成 (modifiers, vk)"""
    hotkey = normalize_hotkey(hotkey)
    if not hotkey:
        return None

    parts = hotkey.split('+')
    key = parts[-1]
    modifiers = parts[:-1]

    mod_value = 0
    for m in modifiers:
        mod_value |= VALID_MODIFIERS[m]

    vk = VK_MAP.get(key)
    if not vk:
        return None

    return mod_value, vk


# =========================
# 主窗口
# =========================
class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(760, 500))

        self.programs = load_config()
        self.hotkey_map = {}      # hotkey_id -> program_index
        self.registered_ids = []  # 已注册热键 id 列表

        self.init_ui()
        self.Centre()
        self.register_all_hotkeys()

        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.Bind(wx.EVT_HOTKEY, self.on_hotkey)

    def init_ui(self):
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        title = wx.StaticText(panel, label="程序列表", style=wx.ALIGN_CENTER)
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        desc = wx.StaticText(
            panel,
            label="设置全局快捷键来快速启动或切换程序窗口\n窗口在前台时按快捷键会最小化，再按恢复"
        )
        desc.SetForegroundColour(wx.Colour(100, 100, 100))
        sizer.Add(desc, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "程序名称", width=180)
        self.list_ctrl.InsertColumn(1, "快捷键", width=120)
        self.list_ctrl.InsertColumn(2, "程序路径", width=420)
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

        edit_btn = wx.Button(panel, label="编辑路径/名称")
        edit_btn.Bind(wx.EVT_BUTTON, self.on_edit_program)
        btn_sizer.Add(edit_btn, 0, wx.ALL, 5)

        refresh_btn = wx.Button(panel, label="刷新热键")
        refresh_btn.Bind(wx.EVT_BUTTON, self.on_refresh_hotkeys)
        btn_sizer.Add(refresh_btn, 0, wx.ALL, 5)

        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER)

        self.status_text = wx.StaticText(panel, label="运行中... | 已启用全局快捷键")
        sizer.Add(self.status_text, 0, wx.ALL | wx.ALIGN_CENTER, 8)

        panel.SetSizer(sizer)
        self.refresh_list()

    def refresh_list(self):
        self.programs = load_config()
        self.list_ctrl.DeleteAllItems()
        for p in self.programs:
            self.list_ctrl.Append([
                p.get('name', ''),
                p.get('hotkey', ''),
                p.get('path', '')
            ])

    def save_and_refresh(self):
        save_config(self.programs)
        self.refresh_list()
        self.register_all_hotkeys()

    def get_selected_index(self):
        return self.list_ctrl.GetFirstSelected()

    def on_manual_add(self, event):
        dialog = wx.Dialog(self, title="添加程序", size=(560, 220))
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

        def on_browse(evt):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="可执行文件 (*.exe)|*.exe", style=wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                path = dlg.GetPath()
                path_ctrl.SetValue(path)
                if not name_ctrl.GetValue().strip():
                    name_ctrl.SetValue(os.path.basename(path).replace('.exe', ''))
            dlg.Destroy()

        browse_btn = wx.Button(panel, label="浏览")
        browse_btn.Bind(wx.EVT_BUTTON, on_browse)
        path_sizer.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(path_sizer, 0, wx.EXPAND | wx.ALL, 5)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        def on_ok(evt):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()

            if not name:
                wx.MessageBox("请输入程序名称", "提示")
                return
            if not path:
                wx.MessageBox("请选择程序路径", "提示")
                return
            if not os.path.exists(path):
                if wx.MessageBox("文件不存在，仍然要添加吗？", "确认", wx.YES_NO | wx.ICON_QUESTION) != wx.YES:
                    return

            self.programs.append({
                'name': name,
                'path': path,
                'hotkey': ''
            })
            self.save_and_refresh()
            dialog.EndModal(wx.ID_OK)

        ok_btn = wx.Button(panel, wx.ID_OK, "添加")
        ok_btn.Bind(wx.EVT_BUTTON, on_ok)
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        cancel_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.EndModal(wx.ID_CANCEL))
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
                pid = proc.info.get('pid')
                name = proc.info.get('name') or ""

                if not exe or not exe.lower().endswith('.exe'):
                    continue

                title = get_process_title(pid)
                if not title:
                    continue

                exe_name = os.path.basename(exe)
                unique_key = (exe.lower(), title.lower())
                if any(item.get('_key') == unique_key for item in running):
                    continue

                running.append({
                    '_key': unique_key,
                    'name': exe_name.replace('.exe', ''),
                    'exe': exe_name,
                    'path': exe,
                    'title': title
                })
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
            except Exception:
                pass

        if not running:
            wx.MessageBox("没有找到可添加的带窗口程序", "提示")
            return

        running.sort(key=lambda x: (x['name'].lower(), x['title'].lower()))

        dialog = wx.Dialog(self, title="从运行中的程序添加", size=(620, 450))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer.Add(wx.StaticText(panel, label="双击选择程序添加："), 0, wx.ALL, 8)

        listbox = wx.ListBox(panel, size=(-1, 320))
        for p in running:
            display = f"{p['title']}  |  {p['name']}  |  {p['path']}"
            listbox.Append(display)
        sizer.Add(listbox, 1, wx.EXPAND | wx.ALL, 8)

        def do_add(selection):
            item = running[selection]
            path = item['path']

            if any(p.get('path', '').lower() == path.lower() for p in self.programs):
                wx.MessageBox("该程序已存在列表中", "提示")
                return

            self.programs.append({
                'name': item['name'],
                'path': item['path'],
                'hotkey': ''
            })
            self.save_and_refresh()
            dialog.EndModal(wx.ID_OK)

        def on_double_click(evt):
            selection = listbox.GetSelection()
            if selection != wx.NOT_FOUND:
                do_add(selection)

        listbox.Bind(wx.EVT_LISTBOX_DCLICK, on_double_click)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        add_btn = wx.Button(panel, label="添加所选")
        def on_add_click(evt):
            selection = listbox.GetSelection()
            if selection == wx.NOT_FOUND:
                wx.MessageBox("请先选择一个程序", "提示")
                return
            do_add(selection)
        add_btn.Bind(wx.EVT_BUTTON, on_add_click)

        close_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.EndModal(wx.ID_CANCEL))

        btn_sizer.Add(add_btn, 0, wx.ALL, 5)
        btn_sizer.Add(close_btn, 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_delete(self, event):
        selection = self.get_selected_index()
        if selection < 0:
            wx.MessageBox("请先选择要删除的程序", "提示")
            return

        name = self.programs[selection].get('name', '')
        if wx.MessageBox(f"确定删除“{name}”吗？", "确认", wx.YES_NO | wx.ICON_QUESTION) != wx.YES:
            return

        self.programs.pop(selection)
        self.save_and_refresh()
        self.status_text.SetLabel("已删除程序")

    def on_edit_program(self, event):
        selection = self.get_selected_index()
        if selection < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        program = self.programs[selection]

        dialog = wx.Dialog(self, title="编辑程序", size=(560, 220))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        name_sizer = wx.BoxSizer(wx.HORIZONTAL)
        name_sizer.Add(wx.StaticText(panel, label="程序名称:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        name_ctrl = wx.TextCtrl(panel, value=program.get('name', ''), size=(360, -1))
        name_sizer.Add(name_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(name_sizer, 0, wx.EXPAND | wx.ALL, 5)

        path_sizer = wx.BoxSizer(wx.HORIZONTAL)
        path_sizer.Add(wx.StaticText(panel, label="程序路径:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        path_ctrl = wx.TextCtrl(panel, value=program.get('path', ''), size=(300, -1))
        path_sizer.Add(path_ctrl, 1, wx.EXPAND | wx.ALL, 5)

        def on_browse(evt):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="可执行文件 (*.exe)|*.exe", style=wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                path = dlg.GetPath()
                path_ctrl.SetValue(path)
                if not name_ctrl.GetValue().strip():
                    name_ctrl.SetValue(os.path.basename(path).replace('.exe', ''))
            dlg.Destroy()

        browse_btn = wx.Button(panel, label="浏览")
        browse_btn.Bind(wx.EVT_BUTTON, on_browse)
        path_sizer.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(path_sizer, 0, wx.EXPAND | wx.ALL, 5)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        def on_ok(evt):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()

            if not name or not path:
                wx.MessageBox("名称和路径不能为空", "提示")
                return

            self.programs[selection]['name'] = name
            self.programs[selection]['path'] = path
            self.save_and_refresh()
            dialog.EndModal(wx.ID_OK)

        ok_btn = wx.Button(panel, wx.ID_OK, "保存")
        ok_btn.Bind(wx.EVT_BUTTON, on_ok)
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        cancel_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.EndModal(wx.ID_CANCEL))

        btn_sizer.Add(ok_btn, 0, wx.ALL, 5)
        btn_sizer.Add(cancel_btn, 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_set_hotkey(self, event):
        selection = self.get_selected_index()
        if selection < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        program = self.programs[selection]
        current_hotkey = program.get('hotkey', '')

        dlg = wx.TextEntryDialog(
            self,
            "输入新快捷键，例如：\n"
            "alt+1\n"
            "ctrl+shift+a\n"
            "win+f2\n\n"
            "注意：必须至少包含一个修饰键（ctrl/shift/alt/win）",
            "设置快捷键",
            current_hotkey
        )

        if dlg.ShowModal() == wx.ID_OK:
            hotkey = normalize_hotkey(dlg.GetValue().strip())
            if not hotkey:
                wx.MessageBox("快捷键格式无效", "错误", wx.OK | wx.ICON_ERROR)
                dlg.Destroy()
                return

            for i, p in enumerate(self.programs):
                if i != selection and normalize_hotkey(p.get('hotkey', '')) == hotkey:
                    wx.MessageBox(f"快捷键 {hotkey} 已被程序“{p.get('name', '')}”使用", "错误", wx.OK | wx.ICON_ERROR)
                    dlg.Destroy()
                    return

            self.programs[selection]['hotkey'] = hotkey
            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()
            wx.MessageBox(f"已设置快捷键：{hotkey}", "成功")

        dlg.Destroy()

    def on_refresh_hotkeys(self, event):
        self.refresh_list()
        self.register_all_hotkeys()
        self.status_text.SetLabel("热键已刷新")

    def unregister_all_hotkeys(self):
        for hotkey_id in self.registered_ids:
            try:
                self.UnregisterHotKey(hotkey_id)
            except Exception:
                try:
                    user32.UnregisterHotKey(self.GetHandle(), hotkey_id)
                except Exception:
                    pass
        self.registered_ids.clear()
        self.hotkey_map.clear()

    def register_all_hotkeys(self):
        self.unregister_all_hotkeys()
        self.programs = load_config()

        failed = []

        for idx, program in enumerate(self.programs):
            hotkey = normalize_hotkey(program.get('hotkey', ''))
            if not hotkey:
                continue

            parsed = parse_hotkey(hotkey)
            if not parsed:
                failed.append(f"{program.get('name', '')}: {hotkey}")
                continue

            modifiers, vk = parsed
            hotkey_id = 1000 + idx

            try:
                ok = self.RegisterHotKey(hotkey_id, modifiers, vk)
            except Exception:
                ok = user32.RegisterHotKey(self.GetHandle(), hotkey_id, modifiers, vk)

            if ok:
                self.registered_ids.append(hotkey_id)
                self.hotkey_map[hotkey_id] = idx
            else:
                failed.append(f"{program.get('name', '')}: {hotkey}")

        if failed:
            self.status_text.SetLabel("部分热键注册失败，可能已被系统或其他程序占用")
        else:
            self.status_text.SetLabel("运行中... | 全部热键已注册")

    def on_hotkey(self, event):
        hotkey_id = event.GetId()
        idx = self.hotkey_map.get(hotkey_id)

        if idx is None:
            return
        if idx < 0 or idx >= len(self.programs):
            return

        program = self.programs[idx]
        ok, msg = toggle_program(program)
        self.status_text.SetLabel(msg if msg else ("操作成功" if ok else "操作失败"))

    def on_close(self, event):
        self.unregister_all_hotkeys()
        self.Destroy()


# =========================
# App
# =========================
class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True


if __name__ == '__main__':
    app = QuickLauncherApp()
    app.MainLoop()
