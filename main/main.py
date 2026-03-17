"""
QuickLauncher - Windows任务栏快捷启动器
功能：
- 绑定任务栏程序到快捷键
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
import threading
import time

# Windows API
user32 = ctypes.windll.user32
SW_MINIMIZE = 6
SW_RESTORE = 9
SW_SHOW = 5

# 配置路径
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

class Config:
    programs = []
    
    @staticmethod
    def load():
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    Config.programs = data.get('programs', [])
            except:
                Config.programs = []
        else:
            Config.programs = []
    
    @staticmethod
    def save():
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({'programs': Config.programs}, f, ensure_ascii=False, indent=4)

def find_window(exe_name):
    """根据exe名称查找窗口"""
    exe_name = exe_name.lower()
    windows = []
    
    def callback(hwnd, wins):
        try:
            if hwnd and user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
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
    
    for hwnd in windows:
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
        # 置顶
        user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0001 | 0x0002)
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
        return False
    
    exe_name = os.path.basename(path).lower()
    hwnd = find_window(exe_name)
    
    if hwnd:
        if is_minimized(hwnd):
            restore_window(hwnd)
        else:
            minimize_window(hwnd)
        return True
    else:
        # 启动程序
        if os.path.exists(path):
            subprocess.Popen(path)
            return True
    return False

class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(600, 450))
        
        Config.load()
        
        self.init_ui()
        self.Centre()
    
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
        
        add_btn = wx.Button(panel, label="从运行程序添加")
        add_btn.Bind(wx.EVT_BUTTON, self.on_add_from_running)
        btn_sizer.Add(add_btn, 0, wx.ALL, 5)
        
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
        
        # 启动热键监听
        self.start_hotkey_listener()
    
    def refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for i, p in enumerate(Config.programs):
            self.list_ctrl.Append([p.get('name', ''), p.get('hotkey', ''), p.get('path', '')])
    
    def on_add_from_running(self, event):
        """从运行程序添加"""
        running = []
        
        # 获取所有进程
        for proc in psutil.process_iter(['exe', 'pid', 'name']):
            try:
                exe = proc.info.get('exe')
                if exe and exe.endswith('.exe'):
                    exe_name = os.path.basename(exe)
                    # 避免重复
                    if not any(p['exe'] == exe_name for p in running):
                        # 尝试获取窗口标题
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
            wx.MessageBox("没有找到运行的程序", "提示")
            return
        
        # 排序
        running.sort(key=lambda x: x['title'] or x['name'])
        
        # 显示选择对话框
        dialog = wx.Dialog(self, title="选择程序", size=(500, 400))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        sizer.Add(wx.StaticText(panel, label="双击选择程序:"), 0, wx.ALL, 5)
        
        listbox = wx.ListBox(panel, size=(-1, 300))
        for p in running:
            display = f"{p['title']} - {p['name']}" if p['title'] else p['name']
            listbox.Append(display)
        sizer.Add(listbox, 1, wx.EXPAND|wx.ALL, 5)
        
        selected = [None]
        
        def on_select(event):
            selection = listbox.GetSelection()
            if selection != wx.NOT_FOUND:
                selected[0] = running[selection]
        
        listbox.Bind(wx.EVT_LISTBOX_DCLICK, on_select)
        
        def ok():
            if selected[0]:
                Config.programs.append({
                    'name': selected[0]['name'],
                    'path': selected[0]['path'],
                    'hotkey': ''
                })
                Config.save()
                self.refresh_list()
                self.restart_hotkey_listener()
                dialog.Destroy()
        
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(wx.Button(panel, wx.ID_OK, "确定"), 0, wx.ALL, 5)
        btn_sizer.Add(wx.Button(panel, wx.ID_CANCEL, "取消"), 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER|wx.ALL, 5)
        
        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()
    
    def on_delete(self, event):
        selection = self.list_ctrl.GetFirstSelected()
        if selection >= 0:
            Config.programs.pop(selection)
            Config.save()
            self.refresh_list()
            self.restart_hotkey_listener()
    
    def on_set_hotkey(self, event):
        selection = self.list_ctrl.GetFirstSelected()
        if selection < 0:
            wx.MessageBox("请先选择程序", "提示")
            return
        
        dialog = wx.Dialog(self, title="设置快捷键", size=(300, 120))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        sizer.Add(wx.StaticText(panel, label="按下快捷键 (如 Alt+1):"), 0, wx.ALL|wx.ALIGN_CENTER, 10)
        
        hotkey_ctrl = wx.TextCtrl(panel, size=(200, 30))
        sizer.Add(hotkey_ctrl, 0, wx.ALIGN_CENTER|wx.ALL, 10)
        
        def on_key(event):
            keys = []
            key = event.keysym.lower()
            
            # 忽略修饰键
            if key in ('alt_l', 'alt_r', 'shift_l', 'shift_r', 'ctrl_l', 'ctrl_r', 'ctrl', 'shift', 'alt', 'mode'):
                return
            
            # 检测修饰键
            if event.state & 0x1: keys.append('alt')
            if event.state & 0x4: keys.append('shift')
            if event.state & 0x8: keys.append('ctrl')
            
            key = key.upper()
            if key:
                keys.append(key)
            
            if keys:
                hotkey_ctrl.SetValue('+'.join(keys))
        
        hotkey_ctrl.Bind(wx.EVT_KEY_UP, on_key)
        
        def ok():
            hotkey = hotkey_ctrl.GetValue().strip()
            if hotkey and selection < len(Config.programs):
                Config.programs[selection]['hotkey'] = hotkey
                Config.save()
                self.refresh_list()
                self.restart_hotkey_listener()
            dialog.Destroy()
        
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(wx.Button(panel, wx.ID_OK, "确定"), 0, wx.ALL, 5)
        btn_sizer.Add(wx.Button(panel, wx.ID_CANCEL, "取消"), 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER|wx.ALL, 5)
        
        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()
    
    def start_hotkey_listener(self):
        """启动热键监听线程"""
        self.running = True
        self.hotkey_thread = threading.Thread(target=self.hotkey_loop, daemon=True)
        self.hotkey_thread.start()
    
    def restart_hotkey_listener(self):
        """重启热键监听"""
        self.running = False
        time.sleep(0.3)
        self.start_hotkey_listener()
    
    def hotkey_loop(self):
        """热键检测循环"""
        last_triggered = {}
        cooldown = 500  # 500ms冷却
        
        while self.running:
            now = time.time() * 1000
            
            # 检测修饰键
            alt = user32.GetAsyncKeyState(0x12) & 0x8000
            ctrl = user32.GetAsyncKeyState(0x11) & 0x8000
            shift = user32.GetAsyncKeyState(0x10) & 0x8000
            
            # 检测数字键
            for i, vk in enumerate([0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x30]):
                if user32.GetAsyncKeyState(vk) & 0x8000:
                    key = str(i % 10)
                    
                    # 构建期望的快捷键
                    expected = None
                    if alt:
                        expected = f"alt+{key}"
                    elif ctrl:
                        expected = f"ctrl+{key}"
                    elif shift:
                        expected = f"shift+{key}"
                    
                    if expected:
                        # 检查冷却
                        last_time = last_triggered.get(expected, 0)
                        if now - last_time > cooldown:
                            # 匹配并触发
                            for program in Config.programs:
                                hotkey = program.get('hotkey', '').lower().replace(' ', '')
                                if hotkey == expected:
                                    toggle_program(program)
                                    last_triggered[expected] = now
                                    wx.CallAfter(self.status_text.SetLabel, f"已切换: {program.get('name', '')}")
                                    break
                    break
            
            time.sleep(0.05)

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
