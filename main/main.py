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
import threading
import time

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
        """启动热键监听线程"""
        # 停止旧的
        if self.hotkey_thread and self.hotkey_thread.is_alive():
            self.running = False
            time.sleep(0.2)
        
        self.running = True
        
        # 创建热键字典（只保存路径和名称）
        self.hotkey_programs = {}
        for p in self.programs:
            hotkey = p.get('hotkey', '')
            if hotkey:
                self.hotkey_programs[hotkey] = {
                    'path': p.get('path', ''),
                    'name': p.get('name', '')
                }
        
        if self.hotkey_programs:
            self.hotkey_thread = threading.Thread(target=self.hotkey_loop, daemon=True)
            self.hotkey_thread.start()
    
    def hotkey_loop(self):
        """热键检测循环"""
        last_triggered = {}
        cooldown = 500
        
        vk_map = {
            '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34, '5': 0x35,
            '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39, '0': 0x30,
        }
        
        while self.running:
            now = time.time() * 1000
            
            for vk, key_name in vk_map.items():
                if user32.GetAsyncKeyState(vk) & 0x8000:
                    modifiers = []
                    if user32.GetAsyncKeyState(0x11) & 0x8000:
                        modifiers.append('ctrl')
                    if user32.GetAsyncKeyState(0x10) & 0x8000:
                        modifiers.append('shift')
                    if user32.GetAsyncKeyState(0x12) & 0x8000:
                        modifiers.append('alt')
                    
                    if modifiers:
                        expected = '+'.join(modifiers) + '+' + key_name
                        
                        last_time = last_triggered.get(expected, 0)
                        if now - last_time > cooldown:
                            program = self.hotkey_programs.get(expected)
                            if program:
                                toggle_program(program)
                                last_triggered[expected] = now
                                wx.CallAfter(self.status_text.SetLabel, f"已切换: {program.get('name', '')}")
                    break
            
            time.sleep(0.02)
    
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

class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True

if __name__ == '__main__':
    app = QuickLauncherApp()
    app.MainLoop()
