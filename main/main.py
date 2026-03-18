# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器（稳定版）
功能：
- 绑定任务栏程序到快捷键
- 按快捷键启动或切换窗口
- 窗口在前台时最小化，再按恢复

改进点：
1) 改为系统级 RegisterHotKey，不再轮询按键（稳定，不会“前两次后失效”）
2) 修复 on_set_hotkey 中错误的 dialog.Destroy() 调用
3) 增加热键冲突/无效提示
4) 统一热键规范化，避免大小写和空格问题
"""

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

# ---------------- 配置路径 ----------------
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
        except Exception as e:
            print("Load config failed:", e)
    return []


def save_config(programs):
    """保存配置"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({'programs': programs}, f, ensure_ascii=False, indent=4)
        print(f"Config saved to: {CONFIG_FILE}")
    except Exception as e:
        print("Save config failed:", e)


def normalize_hotkey(hotkey: str) -> str:
    """热键格式标准化：去空格、小写、修饰键排序"""
    if not hotkey:
        return ""

    hotkey = hotkey.replace(" ", "").lower()
    parts = [p for p in hotkey.split('+') if p]
    if len(parts) < 2:
        return ""

    key = parts[-1]
    mods = set(parts[:-1])

    ordered = []
    for m in ("ctrl", "shift", "alt", "win"):
        if m in mods:
            ordered.append(m)

    if not ordered:
        return ""

    return "+".join(ordered + [key])


def parse_hotkey(hotkey: str):
    """
    将字符串热键转为 wx.RegisterHotKey 需要的 (modifiers, keycode)
    支持：
    - ctrl/shift/alt/win + [a-z, 0-9, f1-f12]
    """
    hk = normalize_hotkey(hotkey)
    if not hk:
        return None, None

    parts = hk.split('+')
    key = parts[-1]
    mods = parts[:-1]

    mod_flags = 0
    for m in mods:
        if m == 'ctrl':
            mod_flags |= wx.MOD_CONTROL
        elif m == 'shift':
            mod_flags |= wx.MOD_SHIFT
        elif m == 'alt':
            mod_flags |= wx.MOD_ALT
        elif m == 'win':
            mod_flags |= wx.MOD_WIN
        else:
            return None, None

    # 主键
    if len(key) == 1 and key.isalpha():
        keycode = ord(key.upper())  # A-Z
    elif len(key) == 1 and key.isdigit():
        keycode = ord(key)  # 0-9
    elif key.startswith('f') and key[1:].isdigit():
        fnum = int(key[1:])
        if 1 <= fnum <= 12:
            keycode = getattr(wx, f"WXK_F{fnum}")
        else:
            return None, None
    else:
        return None, None

    return mod_flags, keycode


def find_window(exe_name):
    """根据exe名称查找窗口"""
    exe_name = exe_name.lower().replace('.exe', '')
    candidate_hwnds = []

    def callback(hwnd, _):
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

    # 优先前台窗口
    fg_hwnd = user32.GetForegroundWindow()
    if fg_hwnd in candidate_hwnds:
        return fg_hwnd

    if candidate_hwnds:
        return candidate_hwnds[0]

    return None


def toggle_program(program):
    """切换程序窗口状态：前台->最小化；非前台->恢复并激活；不存在->启动"""
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
            # 尝试激活
            user32.SetForegroundWindow(hwnd)
    else:
        if os.path.exists(path):
            try:
                subprocess.Popen(path)
            except Exception as e:
                print(f"Launch failed: {path}, {e}")


def get_process_title(pid):
    """获取进程主窗口标题"""
    try:
        titles = []

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
                            text = title.value.strip()
                            if text:
                                wins.append(text)
            except Exception:
                pass
            return True

        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)
        user32.EnumWindows(WNDENUMPROC(callback), titles)

        if titles:
            return titles[0]
    except Exception:
        pass
    return ""


class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(700, 480))
        self.programs = load_config()

        # hotkey相关
        self.hotkey_id_seed = 1000
        self.hotkey_map = {}   # hotkey_id -> program_index
        self.used_hotkeys = set()

        self.init_ui()
        self.Centre()

        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.Bind(wx.EVT_HOTKEY, self.on_hotkey)

        self.register_all_hotkeys()

    # ---------------- UI ----------------
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
        self.list_ctrl.InsertColumn(1, "快捷键", width=120)
        self.list_ctrl.InsertColumn(2, "程序路径", width=360)
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

        self.status_text = wx.StaticText(panel, label="运行中... | 系统级热键已启用")
        sizer.Add(self.status_text, 0, wx.ALL | wx.ALIGN_CENTER, 8)

        panel.SetSizer(sizer)
        self.refresh_list()

    def refresh_list(self):
        self.list_ctrl.DeleteAllItems()
        for p in self.programs:
            self.list_ctrl.Append([
                p.get('name', ''),
                p.get('hotkey', ''),
                p.get('path', '')
            ])

    # ---------------- Hotkey ----------------
    def unregister_all_hotkeys(self):
        for hotkey_id in list(self.hotkey_map.keys()):
            try:
                self.UnregisterHotKey(hotkey_id)
            except Exception:
                pass
        self.hotkey_map.clear()
        self.used_hotkeys.clear()

    def register_all_hotkeys(self):
        self.unregister_all_hotkeys()

        failed = []
        for idx, program in enumerate(self.programs):
            hk = normalize_hotkey(program.get('hotkey', ''))
            if not hk:
                continue

            mods, keycode = parse_hotkey(hk)
            if mods is None:
                failed.append((program.get('name', ''), hk, "格式不支持"))
                continue

            # 同一配置内重复
            if hk in self.used_hotkeys:
                failed.append((program.get('name', ''), hk, "与列表中其他项目重复"))
                continue

            hotkey_id = self.hotkey_id_seed
            self.hotkey_id_seed += 1

            ok = self.RegisterHotKey(hotkey_id, mods, keycode)
            if ok:
                self.hotkey_map[hotkey_id] = idx
                self.used_hotkeys.add(hk)
            else:
                failed.append((program.get('name', ''), hk, "被系统或其他程序占用"))

        if failed:
            msg = "\n".join([f"- {n}: {h} ({reason})" for n, h, reason in failed])
            wx.MessageBox("以下热键注册失败：\n" + msg, "热键提示")

    def on_hotkey(self, event):
        hotkey_id = event.GetId()
        idx = self.hotkey_map.get(hotkey_id)
        if idx is None or idx >= len(self.programs):
            return

        program = self.programs[idx]
        toggle_program(program)
        self.status_text.SetLabel(f"已切换: {program.get('name', '')}")

    # ---------------- 事件 ----------------
    def on_manual_add(self, event):
        dialog = wx.Dialog(self, title="添加程序", size=(520, 200))
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

        def on_browse(_):
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

        def on_ok(_):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            if not (name and path):
                wx.MessageBox("请填写程序名称和路径", "提示")
                return
            if not os.path.exists(path):
                wx.MessageBox("程序路径不存在", "提示")
                return

            self.programs.append({'name': name, 'path': path, 'hotkey': ''})
            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()
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
        seen_path = set()

        for proc in psutil.process_iter(['exe', 'pid', 'name']):
            try:
                exe = proc.info.get('exe')
                if exe and exe.lower().endswith('.exe'):
                    if exe.lower() in seen_path:
                        continue
                    seen_path.add(exe.lower())
                    title = get_process_title(proc.info['pid'])
                    running.append({
                        'name': os.path.basename(exe).replace('.exe', ''),
                        'path': exe,
                        'title': title
                    })
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
            except Exception:
                pass

        if not running:
            wx.MessageBox("没有找到运行的程序", "提示")
            return

        running.sort(key=lambda x: (x['title'] or x['name']).lower())

        dialog = wx.Dialog(self, title="选择程序", size=(560, 420))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer.Add(wx.StaticText(panel, label="双击选择程序添加:"), 0, wx.ALL, 6)

        listbox = wx.ListBox(panel, size=(-1, 320))
        for p in running:
            display = f"{p['title']} - {p['name']}" if p['title'] else p['name']
            listbox.Append(display)
        sizer.Add(listbox, 1, wx.EXPAND | wx.ALL, 6)

        def on_double_click(_):
            sel = listbox.GetSelection()
            if sel == wx.NOT_FOUND:
                return

            item = running[sel]
            # 避免重复添加同一路径
            if any((x.get('path', '').lower() == item['path'].lower()) for x in self.programs):
                wx.MessageBox("该程序已在列表中", "提示")
                return

            self.programs.append({'name': item['name'], 'path': item['path'], 'hotkey': ''})
            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()
            dialog.Destroy()

        listbox.Bind(wx.EVT_LISTBOX_DCLICK, on_double_click)

        close_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())
        sizer.Add(close_btn, 0, wx.ALIGN_CENTER | wx.ALL, 6)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_delete(self, event):
        sel = self.list_ctrl.GetFirstSelected()
        if sel >= 0:
            self.programs.pop(sel)
            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()

    def on_set_hotkey(self, event):
        sel = self.list_ctrl.GetFirstSelected()
        if sel < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        program = self.programs[sel]
        current_hotkey = program.get('hotkey', '')

        dlg = wx.TextEntryDialog(
            self,
            f"当前快捷键: {current_hotkey}\n\n输入新快捷键（如 ctrl+alt+1, alt+f2）：",
            "设置快捷键",
            current_hotkey
        )

        if dlg.ShowModal() == wx.ID_OK:
            hk = normalize_hotkey(dlg.GetValue())
            if not hk:
                wx.MessageBox("热键格式无效，请使用如 ctrl+1 / alt+f2", "提示")
                dlg.Destroy()
                return

            mods, keycode = parse_hotkey(hk)
            if mods is None:
                wx.MessageBox("不支持该热键，请使用字母/数字/F1-F12", "提示")
                dlg.Destroy()
                return

            # 检查配置内冲突
            for i, p in enumerate(self.programs):
                if i != sel and normalize_hotkey(p.get('hotkey', '')) == hk:
                    wx.MessageBox(f"热键冲突：已被「{p.get('name', '')}」使用", "提示")
                    dlg.Destroy()
                    return

            self.programs[sel]['hotkey'] = hk
            save_config(self.programs)
            self.refresh_list()
            self.register_all_hotkeys()
            wx.MessageBox(f"已设置快捷键: {hk}", "成功")

        dlg.Destroy()  # 仅保留这个，修复原来多余 dialog.Destroy() 的 bug

    def on_close(self, event):
        self.unregister_all_hotkeys()
        self.Destroy()


class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True


if __name__ == '__main__':
    app = QuickLauncherApp(False)
    app.MainLoop()
