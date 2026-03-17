# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器

功能：
- 绑定程序到快捷键
- 按快捷键启动或切换窗口
- 窗口在前台时按快捷键最小化，再按恢复

依赖：
pip install wxPython psutil pynput
"""
import wx
import os
import json
import subprocess
import psutil
import ctypes
import sys
import multiprocessing
import time
from pynput import keyboard

# =========================
# Windows API
# =========================
user32 = ctypes.windll.user32

SW_MINIMIZE = 6
SW_RESTORE = 9
SW_SHOW = 5
HWND_TOPMOST = -1
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_SHOWWINDOW = 0x0040

# 配置路径
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

def load_config():
    """加载配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('programs', [])
        except:
            pass
    return []

def save_config(programs):
    """保存配置"""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump({'programs': programs}, f, ensure_ascii=False, indent=4)

def find_window(exe_name):
    """根据exe名称查找窗口"""
    exe_name = exe_name.lower()
    windows = []
    
    def callback(hwnd, wins):
        try:
            if hwnd and user32.IsWindow(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    wins.append(hwnd)
        except:
            pass
        return True
    
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)
    try:
        user32.EnumWindows(WNDENUMPROC(callback), windows)
    except:
        pass
    
    # 优先查找可见窗口
    for hwnd in windows:
        if user32.IsWindowVisible(hwnd):
            pid = ctypes.c_ulong()
            try:
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                proc = psutil.Process(pid.value)
                if exe_name in proc.name().lower():
                    return hwnd
            except:
                pass
    
    # 查找最小化窗口
    for hwnd in windows:
        if user32.IsIconic(hwnd):
            pid = ctypes.c_ulong()
            try:
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                proc = psutil.Process(pid.value)
                if exe_name in proc.name().lower():
                    return hwnd
            except:
                pass
    
    return None

def is_minimized(hwnd):
    try:
        return user32.IsIconic(hwnd)
    except:
        return False

def restore_window(hwnd):
    try:
        if is_minimized(hwnd):
            user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)
        user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
    except:
        pass

def minimize_window(hwnd):
    try:
        user32.ShowWindow(hwnd, SW_MINIMIZE)
    except:
        pass

def toggle_program(program):
    """切换程序窗口状态"""
    path = program.get('path', '')
    if not path:
        return
    
    exe_name = os.path.basename(path).lower().replace('.exe', '')
    hwnd = find_window(exe_name)
    
    if hwnd:
        if is_minimized(hwnd):
            restore_window(hwnd)
        else:
            minimize_window(hwnd)
    else:
        if os.path.exists(path):
            subprocess.Popen(path)

# =========================
# 热键监听进程
# =========================
def hotkey_listener_process(hotkeys_dict):
    """热键监听进程 - 使用 pynput"""
    hotkeys = {}
    
    # 转换热键格式
    for key_str, program in hotkeys_dict.items():
        try:
            # 转换 alt+1 -> alt+1 格式
            hotkeys[key_str] = program
        except:
            pass
    
    def on_activate(key_str, program):
        """热键激活回调"""
        try:
            toggle_program(program)
        except Exception as e:
            print(f"Error: {e}")
    
    # 创建热键映射
    listener_hotkeys = {}
    for key_str, program in hotkeys.items():
        # 解析热键字符串
        parts = key_str.lower().replace(' ', '').split('+')
        if len(parts) >= 2:
            modifiers = []
            key_part = parts[-1]
            for p in parts[:-1]:
                if p == 'alt':
                    modifiers.append(keyboard.Key.alt)
                elif p == 'ctrl':
                    modifiers.append(keyboard.Key.ctrl_l)
                elif p == 'shift':
                    modifiers.append(keyboard.Key.shift)
            
            # 解析按键
            if key_part.isdigit() or key_part.isalpha():
                key = getattr(keyboard.Key, key_part, key_part)
            elif key_part.startswith('f') and key_part[1:].isdigit():
                key = getattr(keyboard.Key, f'f{key_part[1:]}', None)
            else:
                key = key_part
            
            if key and modifiers:
                combo = tuple(modifiers + [key])
                listener_hotkeys[combo] = lambda p=program: on_activate(key_str, p)
    
    # 使用 pynput 监听
    def for_canonical(f):
        return lambda key: f(listener.canonical(key))
    
    listener = keyboard.Listener(
        on_press=for_canonical(lambda key: None),
        on_release=for_canonical(lambda key: check_hotkey(key))
    )
    
    pressed_keys = set()
    check_hotkey = None
    
    def make_checker(hotkeys_dict):
        def check(key):
            try:
                # 获取修饰键状态
                mods = set()
                if keyboard.Key.alt in [type(key).__dict__.get('_value_', None) for _ in []]:
                    pass
                
                # 简单实现：检查按键组合
                key_repr = str(key).replace("'", "").lower()
                
                for combo, callback in hotkeys_dict.items():
                    combo_str = '+'.join([str(k).lower() for k in combo])
                    if key_repr in combo_str or combo_str in key_repr:
                        callback()
            except:
                pass
        return check
    
    check_hotkey = make_checker(listener_hotkeys)
    listener.start()
    listener.join()

class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(600, 450))
        
        self.programs = load_config()
        self.hotkey_process = None
        
        self.init_ui()
        self.Centre()
        
        # 启动热键监听
        self.start_hotkey_listener()
    
    def init_ui(self):
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        # 标题
        title = wx.StaticText(panel, label="程序列表", style=wx.ALIGN_CENTER)
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(title, 0, wx.ALL|wx.ALIGN_CENTER, 10)
        
        # 说明
        desc = wx.StaticText(panel, label="设置快捷键来快速启动或切换程序窗口\n窗口在前台时按快捷键会最小化，再按恢复")
        desc.SetForegroundColour(wx.Colour(100, 100, 100))
        sizer.Add(desc, 0, wx.ALL|wx.ALIGN_CENTER, 5)
        
        # 列表
        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT|wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "程序名称", width=150)
        self.list_ctrl.InsertColumn(1, "快捷键", width=100)
        self.list_ctrl.InsertColumn(2, "程序路径", width=250)
        
        sizer.Add(self.list_ctrl, 1, wx.ALL|wx.EXPAND, 5)
        
        # 按钮
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
        
        # 状态
        self.status_text = wx.StaticText(panel, label="运行中... | 按快捷键切换程序")
        sizer.Add(self.status_text, 0, wx.ALL|wx.ALIGN_CENTER, 5)
        
        panel.SetSizer(sizer)
        
        self.refresh_list()
        self.Update()
    
    def refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for p in self.programs:
            self.list_ctrl.Append([p.get('name', ''), p.get('hotkey', ''), p.get('path', '')])
    
    def start_hotkey_listener(self):
        """启动热键监听进程"""
        # 停止旧的
        if self.hotkey_process and self.hotkey_process.is_alive():
            self.hotkey_process.terminate()
        
        # 创建热键字典
        hotkeys = {}
        for p in self.programs:
            hotkey = p.get('hotkey', '')
            if hotkey:
                hotkeys[hotkey] = p
        
        if hotkeys:
            # 使用简单的轮询方式（pynput 在某些环境下有问题）
            self.hotkey_process = multiprocessing.Process(
                target=simple_hotkey_listener, 
                args=(hotkeys,),
                daemon=True
            )
            self.hotkey_process.start()
    
    def on_manual_add(self, event):
        """手动添加程序"""
        dialog = wx.Dialog(self, title="添加程序", size=(500, 180))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        name_sizer = wx.BoxSizer(wx.HORIZONTAL)
        name_sizer.Add(wx.StaticText(panel, label="程序名称:"), 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        name_ctrl = wx.TextCtrl(panel, size=(300, -1))
        name_sizer.Add(name_ctrl, 1, wx.EXPAND|wx.ALL, 5)
        sizer.Add(name_sizer, 0, wx.EXPAND|wx.ALL, 5)
        
        path_sizer = wx.BoxSizer(wx.HORIZONTAL)
        path_sizer.Add(wx.StaticText(panel, label="程序路径:"), 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        path_ctrl = wx.TextCtrl(panel, size=(250, -1))
        path_sizer.Add(path_ctrl, 1, wx.EXPAND|wx.ALL, 5)
        
        def on_browse(event):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                path_ctrl.SetValue(dlg.GetPath())
                if not name_ctrl.GetValue():
                    name_ctrl.SetValue(os.path.basename(dlg.GetPath()).replace('.exe', ''))
            dlg.Destroy()
        
        browse_btn = wx.Button(panel, label="浏览")
        browse_btn.Bind(wx.EVT_BUTTON, on_browse)
        path_sizer.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(path_sizer, 0, wx.EXPAND|wx.ALL, 5)
        
        def on_ok(event):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            if name and path:
                self.programs.append({'name': name, 'path': path, 'hotkey': ''})
                save_config(self.programs)
                self.refresh_list()
                self.start_hotkey_listener()
                dialog.Destroy()
        
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        ok_btn = wx.Button(panel, wx.ID_OK, "添加")
        ok_btn.Bind(wx.EVT_BUTTON, on_ok)
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        cancel_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())
        btn_sizer.Add(ok_btn, 0, wx.ALL, 5)
        btn_sizer.Add(cancel_btn, 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER|wx.ALL, 5)
        
        panel.SetSizer(sizer)
        dialog.ShowModal()
    
    def on_add_from_running(self, event):
        """从运行程序添加"""
        running = []
        
        for proc in psutil.process_iter(['exe', 'pid', 'name']):
            try:
                exe = proc.info.get('exe')
                if exe and exe.endswith('.exe'):
                    exe_name = os.path.basename(exe)
                    if not any(p.get('exe') == exe_name for p in running):
                        title = get_process_title(proc.info['pid'])
                        running.append({'name': exe_name.replace('.exe', ''), 'exe': exe_name, 'path': exe, 'title': title})
            except:
                pass
        
        if not running:
            wx.MessageBox("没有找到运行的程序", "提示")
            return
        
        running.sort(key=lambda x: x['title'] or x['name'])
        
        dialog = wx.Dialog(self, title="选择程序", size=(500, 400))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        sizer.Add(wx.StaticText(panel, label="双击选择程序:"), 0, wx.ALL, 5)
        
        listbox = wx.ListBox(panel, size=(-1, 300))
        for p in running:
            display = f"{p['title']} - {p['name']}" if p['title'] else p['name']
            listbox.Append(display)
        sizer.Add(listbox, 1, wx.EXPAND|wx.ALL, 5)
        
        def on_double(event):
            selection = listbox.GetSelection()
            if selection != wx.NOT_FOUND:
                p = running[selection]
                self.programs.append({'name': p['name'], 'path': p['path'], 'hotkey': ''})
                save_config(self.programs)
                self.refresh_list()
                self.start_hotkey_listener()
                dialog.Destroy()
        
        listbox.Bind(wx.EVT_LISTBOX_DCLICK, on_double)
        
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(wx.Button(panel, wx.ID_CANCEL, "关闭"), 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER|wx.ALL, 5)
        
        panel.SetSizer(sizer)
        dialog.ShowModal()
    
    def on_delete(self, event):
        selection = self.list_ctrl.GetFirstSelected()
        if selection >= 0:
            self.programs.pop(selection)
            save_config(self.programs)
            self.refresh_list()
            self.start_hotkey_listener()
    
    def on_set_hotkey(self, event):
        selection = self.list_ctrl.GetFirstSelected()
        if selection < 0:
            wx.MessageBox("请先选择程序", "提示")
            return
        
        dlg = wx.TextEntryDialog(
            self, 
            f"当前快捷键: {self.programs[selection].get('hotkey', '')}\n\n输入新快捷键 (如 alt+1, ctrl+a):",
            "设置快捷键",
            self.programs[selection].get('hotkey', '')
        )
        
        if dlg.ShowModal() == wx.ID_OK:
            hotkey = dlg.GetValue().strip().lower()
            if hotkey:
                self.programs[selection]['hotkey'] = hotkey
                save_config(self.programs)
                self.refresh_list()
                self.start_hotkey_listener()
                wx.MessageBox(f"已设置快捷键: {hotkey}", "成功")
        
        dlg.Destroy()

def get_process_title(pid):
    """获取进程的窗口标题"""
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
            except:
                pass
            return True
        
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)
        user32.EnumWindows(WNDENUMPROC(callback), windows)
        
        for title in windows:
            if title:
                return title
    except:
        pass
    return ""

def simple_hotkey_listener(hotkeys):
    """简单的热键监听 - 使用轮询"""
    last_triggered = {}
    cooldown = 500
    
    # 虚拟键码映射
    vk_map = {
        '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34, '5': 0x35,
        '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39, '0': 0x30,
    }
    
    while True:
        now = time.time() * 1000
        
        for vk, key_name in vk_map.items():
            if user32.GetAsyncKeyState(vk) & 0x8000:
                modifiers = []
                if user32.GetAsyncKeyState(0x11) & 0x8000:  # Ctrl
                    modifiers.append('ctrl')
                if user32.GetAsyncKeyState(0x10) & 0x8000:  # Shift
                    modifiers.append('shift')
                if user32.GetAsyncKeyState(0x12) & 0x8000:  # Alt
                    modifiers.append('alt')
                
                if modifiers:
                    expected = '+'.join(modifiers) + '+' + key_name
                    
                    last_time = last_triggered.get(expected, 0)
                    if now - last_time > cooldown:
                        for program in hotkeys.values():
                            hotkey = program.get('hotkey', '').lower().replace(' ', '')
                            if hotkey == expected:
                                toggle_program(program)
                                last_triggered[expected] = now
                                break
                break
        
        time.sleep(0.02)

class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True

if __name__ == '__main__':
    app = QuickLauncherApp()
    app.MainLoop()
