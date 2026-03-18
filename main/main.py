# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器（增强稳定版）
新增：
1) 设置快捷键时，直接按键自动录入（无需手动输入）
2) 支持“同一浏览器多窗口”分别绑定不同快捷键（通过“窗口关键词”区分）
   - 例如：
     Chrome-工作:  hotkey=ctrl+alt+1, path=chrome.exe, window_keyword=工作
     Chrome-娱乐:  hotkey=ctrl+alt+2, path=chrome.exe, window_keyword=娱乐
   - 配置写入 config.json，重启后仍生效
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

RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_RUN_NAME = "QuickLauncher"

# ---------------------------
# 配置路径
# ---------------------------
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")


# ---------------------------
# 配置 / 自启
# ---------------------------
def get_launch_command():
    """获取开机自启命令"""
    if getattr(sys, 'frozen', False):
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

            # 兼容旧配置字段
            for p in programs:
                p.setdefault("name", "")
                p.setdefault("path", "")
                p.setdefault("args", "")
                p.setdefault("hotkey", "")
                p.setdefault("window_keyword", "")
            return {"programs": programs, "autostart": autostart}
        except Exception as e:
            print("load_config error:", e)
    return default_data


def save_config(programs, autostart):
    data = {"programs": programs, "autostart": autostart}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


# ---------------------------
# 热键工具
# ---------------------------
def normalize_hotkey(hotkey: str) -> str:
    if not hotkey:
        return ""
    return hotkey.strip().lower().replace(" ", "")


def hotkey_to_mod_vk(hotkey: str):
    """
    解析快捷键到 (modifiers, vk, normalized_hotkey)
    - 单键仅支持 F1-F12
    - 组合键支持 ctrl/alt/shift/win + (a-z,0-9,f1-f12)
    """
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
                raise ValueError("不支持的主键")
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
    """
    把键盘事件转为标准热键字符串。
    返回 "" 表示当前按键不支持作为主键。
    """
    key = event.GetKeyCode()
    main_key = None

    # F1-F12
    if wx.WXK_F1 <= key <= wx.WXK_F12:
        main_key = f"f{key - wx.WXK_F1 + 1}"
    # 0-9
    elif ord('0') <= key <= ord('9'):
        main_key = chr(key).lower()
    # A-Z / a-z
    elif ord('A') <= key <= ord('Z'):
        main_key = chr(key).lower()
    elif ord('a') <= key <= ord('z'):
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
    # Win键
    if event.MetaDown():
        mods.append("win")

    if not mods and not main_key.startswith("f"):
        return ""  # 单键仅允许 F1-F12

    return "+".join(mods + [main_key]) if mods else main_key


# ---------------------------
# 窗口 / 进程工具
# ---------------------------
def get_window_title(hwnd):
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value or ""


def enum_windows_for_exe(exe_name):
    """返回 [(hwnd, title)]，匹配 exe_name"""
    exe_name = exe_name.lower().replace(".exe", "")
    result = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def callback(hwnd, _):
        try:
            if not user32.IsWindowVisible(hwnd) or not user32.IsWindowEnabled(hwnd):
                return True

            pid = ctypes.c_ulong()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            try:
                proc = psutil.Process(pid.value)
                proc_name = proc.name().lower().replace(".exe", "")
                if proc_name == exe_name:
                    title = get_window_title(hwnd)
                    result.append((hwnd, title))
            except Exception:
                pass
        except Exception:
            pass
        return True

    user32.EnumWindows(callback, 0)
    return result


def find_window_for_program(program):
    """
    支持窗口关键词匹配：
    - path -> exe_name
    - window_keyword 非空时，仅匹配标题包含关键词的窗口
    """
    path = program.get("path", "")
    if not path:
        return None

    exe_name = os.path.basename(path).lower().replace(".exe", "")
    keyword = (program.get("window_keyword", "") or "").strip().lower()

    candidates = enum_windows_for_exe(exe_name)
    if keyword:
        candidates = [(h, t) for (h, t) in candidates if keyword in (t or "").lower()]

    if not candidates:
        return None

    fg = user32.GetForegroundWindow()
    for hwnd, _ in candidates:
        if hwnd == fg:
            return hwnd

    # 优先非最小化
    for hwnd, _ in candidates:
        if not user32.IsIconic(hwnd):
            return hwnd

    return candidates[0][0]


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
            # 简易参数拆分（保持常用场景稳定）
            cmd.extend(args.strip().split())
        subprocess.Popen(cmd)


def get_process_title(pid):
    windows = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def callback(hwnd, _):
        try:
            if hwnd and user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
                window_pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                if window_pid.value == pid:
                    title = get_window_title(hwnd)
                    if title:
                        windows.append(title)
        except Exception:
            pass
        return True

    user32.EnumWindows(callback, 0)
    return windows[0] if windows else ""


# ---------------------------
# 快捷键捕获对话框（自动录入）
# ---------------------------
class HotkeyCaptureDialog(wx.Dialog):
    def __init__(self, parent, current_hotkey=""):
        super().__init__(parent, title="设置快捷键（按键自动录入）", size=(500, 250))
        self.captured_hotkey = normalize_hotkey(current_hotkey)

        panel = wx.Panel(self)
        s = wx.BoxSizer(wx.VERTICAL)

        tips = wx.StaticText(
            panel,
            label=(
                "请直接按下目标快捷键（例如 Ctrl+1、Alt+Q、Ctrl+Shift+F2）\n"
                "单键仅支持 F1-F12\n"
                "Esc 取消，Backspace 清空"
            ),
        )
        tips.SetForegroundColour(wx.Colour(90, 90, 90))
        s.Add(tips, 0, wx.ALL | wx.EXPAND, 10)

        self.hk_show = wx.TextCtrl(panel, value=self.captured_hotkey, style=wx.TE_READONLY)
        s.Add(self.hk_show, 0, wx.ALL | wx.EXPAND, 10)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        ok_btn = wx.Button(panel, wx.ID_OK, "确定")
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        clear_btn = wx.Button(panel, wx.ID_ANY, "清空")
        btn_row.Add(ok_btn, 0, wx.ALL, 5)
        btn_row.Add(cancel_btn, 0, wx.ALL, 5)
        btn_row.Add(clear_btn, 0, wx.ALL, 5)
        s.Add(btn_row, 0, wx.ALIGN_CENTER | wx.ALL, 5)

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

        # 再走统一校验和标准化
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
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(840, 530))

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
        self.update_status("运行中... | 全局热键已就绪")

        self.Bind(wx.EVT_CLOSE, self.on_close)

    def init_ui(self):
        panel = wx.Panel(self)
        root = wx.BoxSizer(wx.VERTICAL)

        title = wx.StaticText(panel, label="程序列表")
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        root.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        desc = wx.StaticText(
            panel,
            label=(
                "支持：F1-F12 单键；或 Ctrl/Alt/Shift/Win 组合键\n"
                "同一浏览器多窗口：可为每条记录设置“窗口关键词”来区分不同窗口"
            )
        )
        desc.SetForegroundColour(wx.Colour(100, 100, 100))
        root.Add(desc, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "程序名称", width=130)
        self.list_ctrl.InsertColumn(1, "快捷键", width=120)
        self.list_ctrl.InsertColumn(2, "窗口关键词", width=170)
        self.list_ctrl.InsertColumn(3, "程序路径", width=320)
        self.list_ctrl.InsertColumn(4, "启动参数", width=180)
        root.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        btn_row1 = wx.BoxSizer(wx.HORIZONTAL)

        add_btn = wx.Button(panel, label="手动添加")
        add_btn.Bind(wx.EVT_BUTTON, self.on_manual_add)
        btn_row1.Add(add_btn, 0, wx.ALL, 5)

        add_running_btn = wx.Button(panel, label="从运行程序添加")
        add_running_btn.Bind(wx.EVT_BUTTON, self.on_add_from_running)
        btn_row1.Add(add_running_btn, 0, wx.ALL, 5)

        del_btn = wx.Button(panel, label="删除")
        del_btn.Bind(wx.EVT_BUTTON, self.on_delete)
        btn_row1.Add(del_btn, 0, wx.ALL, 5)

        set_hotkey_btn = wx.Button(panel, label="设置快捷键（按键录入）")
        set_hotkey_btn.Bind(wx.EVT_BUTTON, self.on_set_hotkey)
        btn_row1.Add(set_hotkey_btn, 0, wx.ALL, 5)

        set_kw_btn = wx.Button(panel, label="设置窗口关键词")
        set_kw_btn.Bind(wx.EVT_BUTTON, self.on_set_window_keyword)
        btn_row1.Add(set_kw_btn, 0, wx.ALL, 5)

        hide_btn = wx.Button(panel, label="最小化到托盘")
        hide_btn.Bind(wx.EVT_BUTTON, lambda e: self.hide_to_tray())
        btn_row1.Add(hide_btn, 0, wx.ALL, 5)

        root.Add(btn_row1, 0, wx.ALIGN_CENTER)

        setting_row = wx.BoxSizer(wx.HORIZONTAL)
        self.autostart_cb = wx.CheckBox(panel, label="开机自启")
        self.autostart_cb.SetValue(self.autostart)
        setting_row.Add(self.autostart_cb, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 8)

        apply_setting_btn = wx.Button(panel, label="保存设置")
        apply_setting_btn.Bind(wx.EVT_BUTTON, self.on_apply_settings)
        setting_row.Add(apply_setting_btn, 0, wx.ALL, 5)

        root.Add(setting_row, 0, wx.ALIGN_CENTER)

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
                p.get("window_keyword", ""),
                p.get("path", ""),
                p.get("args", "")
            ])

    # ---------------------------
    # 托盘 / 关闭
    # ---------------------------
    def hide_to_tray(self):
        self.Hide()
        self.update_status("已最小化到托盘（双击托盘图标可恢复）")

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

    # ---------------------------
    # 配置 / 设置
    # ---------------------------
    def persist(self):
        save_config(self.programs, self.autostart_cb.GetValue())

    def on_apply_settings(self, _):
        target = self.autostart_cb.GetValue()
        ok = set_autostart(target)
        if not ok:
            wx.MessageBox("开机自启设置失败（可能权限不足）", "错误", wx.OK | wx.ICON_ERROR)
            self.autostart_cb.SetValue(is_autostart_enabled())
            return

        self.autostart = target
        self.persist()
        wx.MessageBox("设置已保存", "成功", wx.OK | wx.ICON_INFORMATION)

    # ---------------------------
    # 热键注册
    # ---------------------------
    def unregister_all_hotkeys(self):
        for hotkey_id in self.registered_hotkey_ids:
            try:
                self.UnregisterHotKey(hotkey_id)
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
                fail_msgs.append(f"{p.get('name','')}：{hk}（{e}）")
                continue

            if hk in used:
                fail_msgs.append(f"{p.get('name','')}：{hk}（与列表内重复）")
                continue
            used.add(hk)

            hotkey_id = base_id + idx
            ok = self.RegisterHotKey(hotkey_id, mods, vk)
            if ok:
                self.registered_hotkey_ids.append(hotkey_id)
                self.hotkey_id_to_index[hotkey_id] = idx
                self.Bind(wx.EVT_HOTKEY, self.on_hotkey, id=hotkey_id)
            else:
                fail_msgs.append(f"{p.get('name','')}：{hk}（注册失败，可能被系统/其他软件占用）")

        self.persist()
        self.refresh_list()

        if fail_msgs:
            wx.MessageBox(
                "以下快捷键注册失败：\n\n" + "\n".join(fail_msgs),
                "热键提示",
                wx.OK | wx.ICON_WARNING
            )

    def on_hotkey(self, event):
        hotkey_id = event.GetId()
        idx = self.hotkey_id_to_index.get(hotkey_id)
        if idx is None or idx < 0 or idx >= len(self.programs):
            return

        program = self.programs[idx]
        try:
            toggle_program(program)
            self.update_status(f"已切换：{program.get('name', '')}")
        except Exception as e:
            self.update_status("切换失败")
            wx.MessageBox(f"切换失败：{e}", "错误", wx.OK | wx.ICON_ERROR)

    # ---------------------------
    # 列表操作
    # ---------------------------
    def on_manual_add(self, _):
        dialog = wx.Dialog(self, title="添加程序", size=(620, 320))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # 名称
        row1 = wx.BoxSizer(wx.HORIZONTAL)
        row1.Add(wx.StaticText(panel, label="程序名称:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        name_ctrl = wx.TextCtrl(panel, size=(420, -1))
        row1.Add(name_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(row1, 0, wx.ALL | wx.EXPAND, 5)

        # 路径
        row2 = wx.BoxSizer(wx.HORIZONTAL)
        row2.Add(wx.StaticText(panel, label="程序路径:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        path_ctrl = wx.TextCtrl(panel, size=(350, -1))
        row2.Add(path_ctrl, 1, wx.ALL | wx.EXPAND, 5)

        def on_browse(_evt):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                p = dlg.GetPath()
                path_ctrl.SetValue(p)
                if not name_ctrl.GetValue().strip():
                    name_ctrl.SetValue(os.path.basename(p).replace(".exe", ""))
            dlg.Destroy()

        browse_btn = wx.Button(panel, label="浏览")
        browse_btn.Bind(wx.EVT_BUTTON, on_browse)
        row2.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(row2, 0, wx.ALL | wx.EXPAND, 5)

        # 参数
        row3 = wx.BoxSizer(wx.HORIZONTAL)
        row3.Add(wx.StaticText(panel, label="启动参数:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        args_ctrl = wx.TextCtrl(panel, size=(420, -1))
        row3.Add(args_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(row3, 0, wx.ALL | wx.EXPAND, 5)

        # 窗口关键词
        row4 = wx.BoxSizer(wx.HORIZONTAL)
        row4.Add(wx.StaticText(panel, label="窗口关键词:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        kw_ctrl = wx.TextCtrl(panel, size=(420, -1))
        kw_ctrl.SetHint("可选：用于区分同一程序的不同窗口（如 Chrome 多开）")
        row4.Add(kw_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(row4, 0, wx.ALL | wx.EXPAND, 5)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)

        def on_ok(_evt):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            args = args_ctrl.GetValue().strip()
            kw = kw_ctrl.GetValue().strip()

            if not name or not path:
                wx.MessageBox("请填写程序名称和路径", "提示")
                return

            self.programs.append({
                "name": name,
                "path": path,
                "args": args,
                "hotkey": "",
                "window_keyword": kw
            })
            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()
            dialog.Destroy()

        ok_btn = wx.Button(panel, wx.ID_OK, "添加")
        ok_btn.Bind(wx.EVT_BUTTON, on_ok)
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "取消")
        cancel_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())
        btn_row.Add(ok_btn, 0, wx.ALL, 5)
        btn_row.Add(cancel_btn, 0, wx.ALL, 5)
        sizer.Add(btn_row, 0, wx.ALIGN_CENTER | wx.ALL, 8)

        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

    def on_add_from_running(self, _):
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

        dialog = wx.Dialog(self, title="选择运行中的程序", size=(680, 450))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer.Add(wx.StaticText(panel, label="双击添加程序（窗口关键词默认取当前窗口标题，可后续修改）："), 0, wx.ALL, 8)

        listbox = wx.ListBox(panel, size=(-1, 340))
        for p in running:
            text = f"{p['title']}  ——  {p['name']}" if p["title"] else p["name"]
            listbox.Append(text)
        sizer.Add(listbox, 1, wx.EXPAND | wx.ALL, 8)

        def on_double_click(_evt):
            i = listbox.GetSelection()
            if i == wx.NOT_FOUND:
                return
            item = running[i]
            self.programs.append({
                "name": item["name"],
                "path": item["path"],
                "args": "",
                "hotkey": "",
                "window_keyword": item["title"][:80] if item["title"] else ""
            })
            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()
            dialog.Destroy()

        listbox.Bind(wx.EVT_LISTBOX_DCLICK, on_double_click)

        close_btn = wx.Button(panel, wx.ID_CANCEL, "关闭")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dialog.Destroy())
        sizer.Add(close_btn, 0, wx.ALL | wx.ALIGN_CENTER, 6)

        panel.SetSizer(sizer)
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

        current = self.programs[idx].get("hotkey", "")
        dlg = HotkeyCaptureDialog(self, current_hotkey=current)

        if dlg.ShowModal() == wx.ID_OK:
            hk = normalize_hotkey(dlg.captured_hotkey)

            # 清空快捷键
            if not hk:
                self.programs[idx]["hotkey"] = ""
                self.persist()
                self.refresh_list()
                self.register_all_hotkeys()
                wx.MessageBox("已清空快捷键", "提示")
                dlg.Destroy()
                return

            # 格式校验
            try:
                _, _, normalized = hotkey_to_mod_vk(hk)
            except ValueError as e:
                wx.MessageBox(f"快捷键无效：{e}", "错误", wx.OK | wx.ICON_ERROR)
                dlg.Destroy()
                return

            # 列表内重复校验
            for i, p in enumerate(self.programs):
                if i != idx and normalize_hotkey(p.get("hotkey", "")) == normalized:
                    wx.MessageBox("该快捷键已被列表中的其他程序使用", "错误", wx.OK | wx.ICON_ERROR)
                    dlg.Destroy()
                    return

            self.programs[idx]["hotkey"] = normalized
            self.persist()
            self.refresh_list()
            self.register_all_hotkeys()
            wx.MessageBox(f"已设置快捷键：{normalized}", "成功", wx.OK | wx.ICON_INFORMATION)

        dlg.Destroy()

    def on_set_window_keyword(self, _):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        current = self.programs[idx].get("window_keyword", "")
        dlg = wx.TextEntryDialog(
            self,
            "输入窗口关键词（用于区分同一程序的不同窗口）\n"
            "示例：工作、娱乐、Profile 1\n"
            "留空表示不按标题区分",
            "设置窗口关键词",
            current
        )

        if dlg.ShowModal() == wx.ID_OK:
            kw = dlg.GetValue().strip()
            self.programs[idx]["window_keyword"] = kw
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
