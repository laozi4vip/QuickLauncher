# -*- coding: utf-8 -*-
"""
QuickLauncher - Windows任务栏快捷启动器（稳定版）
功能：
- 绑定程序到全局快捷键
- 按快捷键启动或切换窗口
- 窗口在前台时最小化，再按恢复
- 支持单独 F1-F12
- 支持开机自启
- 支持最小化到托盘
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


def get_launch_command():
    """获取开机自启命令"""
    if getattr(sys, 'frozen', False):
        return f'"{sys.executable}"'
    # 非打包模式：python.exe "script.py"
    return f'"{sys.executable}" "{os.path.abspath(__file__)}"'


def is_autostart_enabled():
    """检查开机自启是否开启"""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_READ) as key:
            val, _ = winreg.QueryValueEx(key, APP_RUN_NAME)
            return bool(val)
    except FileNotFoundError:
        return False
    except Exception:
        return False


def set_autostart(enable: bool):
    """设置开机自启"""
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
    """加载配置"""
    default_data = {"programs": [], "autostart": False}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                programs = data.get("programs", [])
                autostart = data.get("autostart", is_autostart_enabled())
                return {"programs": programs, "autostart": autostart}
        except Exception as e:
            print("load_config error:", e)
    # 默认读取当前系统状态
    default_data["autostart"] = is_autostart_enabled()
    return default_data


def save_config(programs, autostart):
    """保存配置"""
    data = {"programs": programs, "autostart": autostart}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def normalize_hotkey(hotkey: str) -> str:
    """标准化快捷键字符串"""
    if not hotkey:
        return ""
    return hotkey.strip().lower().replace(" ", "")


def hotkey_to_mod_vk(hotkey: str):
    """
    解析快捷键到 (modifiers, vk, normalized_hotkey)
    规则：
    - 单独键仅允许 f1-f12
    - 组合键支持 ctrl/alt/shift + (a-z,0-9,f1-f12)
    """
    hotkey = normalize_hotkey(hotkey)
    if not hotkey:
        raise ValueError("快捷键不能为空")

    parts = hotkey.split("+")
    if any(p == "" for p in parts):
        raise ValueError("快捷键格式错误")

    # 键位映射
    key_map = {str(i): 0x30 + i for i in range(10)}
    for i in range(26):
        key_map[chr(ord("a") + i)] = 0x41 + i
    for i in range(1, 13):
        key_map[f"f{i}"] = 0x70 + (i - 1)

    mods = 0
    mod_names = []
    main_key = None

    for p in parts:
        if p == "ctrl":
            if "ctrl" not in mod_names:
                mod_names.append("ctrl")
                mods |= MOD_CONTROL
        elif p == "alt":
            if "alt" not in mod_names:
                mod_names.append("alt")
                mods |= MOD_ALT
        elif p == "shift":
            if "shift" not in mod_names:
                mod_names.append("shift")
                mods |= MOD_SHIFT
        elif p == "win":
            if "win" not in mod_names:
                mod_names.append("win")
                mods |= MOD_WIN
        else:
            if main_key is not None:
                raise ValueError("只能有一个主键")
            if p not in key_map:
                raise ValueError("不支持的主键")
            main_key = p

    if main_key is None:
        raise ValueError("缺少主键")

    # 无修饰键时，只允许 F1-F12
    if mods == 0 and not main_key.startswith("f"):
        raise ValueError("无修饰键仅支持 F1-F12")

    # 统一输出顺序
    ordered_mods = []
    if mods & MOD_CONTROL:
        ordered_mods.append("ctrl")
    if mods & MOD_SHIFT:
        ordered_mods.append("shift")
    if mods & MOD_ALT:
        ordered_mods.append("alt")
    if mods & MOD_WIN:
        ordered_mods.append("win")

    normalized = "+".join(ordered_mods + [main_key]) if ordered_mods else main_key
    return mods, key_map[main_key], normalized


def find_window(exe_name):
    """根据 exe 名查找窗口"""
    exe_name = exe_name.lower().replace(".exe", "")
    candidate_hwnds = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
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

    user32.EnumWindows(callback, 0)

    fg_hwnd = user32.GetForegroundWindow()
    if fg_hwnd in candidate_hwnds:
        return fg_hwnd
    if candidate_hwnds:
        return candidate_hwnds[0]
    return None


def toggle_program(program):
    """切换程序窗口状态"""
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
        else:
            wx.MessageBox(f"程序不存在：\n{path}", "错误", wx.OK | wx.ICON_ERROR)


def get_process_title(pid):
    """获取进程窗口标题"""
    windows = []

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
                        if buf.value:
                            windows.append(buf.value)
        except Exception:
            pass
        return True

    user32.EnumWindows(callback, 0)
    return windows[0] if windows else ""


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

    def on_show(self, event):
        self.frame.show_from_tray()

    def on_hide(self, event):
        self.frame.hide_to_tray()

    def on_exit(self, event):
        self.frame.exit_app()


class QuickLauncherFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="QuickLauncher - 快捷启动器", size=(700, 500))

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
            label="支持：F1-F12 单键；或 Ctrl/Alt/Shift 组合键\n窗口在前台时按快捷键最小化，再按恢复"
        )
        desc.SetForegroundColour(wx.Colour(100, 100, 100))
        root.Add(desc, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "程序名称", width=170)
        self.list_ctrl.InsertColumn(1, "快捷键", width=120)
        self.list_ctrl.InsertColumn(2, "程序路径", width=360)
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

        set_btn = wx.Button(panel, label="设置快捷键")
        set_btn.Bind(wx.EVT_BUTTON, self.on_set_hotkey)
        btn_row1.Add(set_btn, 0, wx.ALL, 5)

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
            self.list_ctrl.Append([p.get("name", ""), p.get("hotkey", ""), p.get("path", "")])

    # ---------------------------
    # 托盘 / 关闭行为
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
    def on_apply_settings(self, event):
        target = self.autostart_cb.GetValue()
        ok = set_autostart(target)
        if not ok:
            wx.MessageBox("开机自启设置失败（可能权限不足）", "错误", wx.OK | wx.ICON_ERROR)
            self.autostart_cb.SetValue(is_autostart_enabled())
            return

        self.autostart = target
        save_config(self.programs, self.autostart)
        wx.MessageBox("设置已保存", "成功", wx.OK | wx.ICON_INFORMATION)

    def persist(self):
        save_config(self.programs, self.autostart_cb.GetValue())

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
                # 写回标准格式
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
                fail_msgs.append(f"{p.get('name','')}：{hk}（注册失败，可能被系统/其它软件占用）")

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
    def on_manual_add(self, event):
        dialog = wx.Dialog(self, title="添加程序", size=(520, 200))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        name_row = wx.BoxSizer(wx.HORIZONTAL)
        name_row.Add(wx.StaticText(panel, label="程序名称:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        name_ctrl = wx.TextCtrl(panel, size=(330, -1))
        name_row.Add(name_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(name_row, 0, wx.ALL | wx.EXPAND, 5)

        path_row = wx.BoxSizer(wx.HORIZONTAL)
        path_row.Add(wx.StaticText(panel, label="程序路径:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        path_ctrl = wx.TextCtrl(panel, size=(280, -1))
        path_row.Add(path_ctrl, 1, wx.ALL | wx.EXPAND, 5)

        def on_browse(_):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                path = dlg.GetPath()
                path_ctrl.SetValue(path)
                if not name_ctrl.GetValue().strip():
                    name_ctrl.SetValue(os.path.basename(path).replace(".exe", ""))
            dlg.Destroy()

        browse_btn = wx.Button(panel, label="浏览")
        browse_btn.Bind(wx.EVT_BUTTON, on_browse)
        path_row.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(path_row, 0, wx.ALL | wx.EXPAND, 5)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)

        def on_ok(_):
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            if not name or not path:
                wx.MessageBox("请填写程序名称和路径", "提示")
                return
            self.programs.append({"name": name, "path": path, "hotkey": ""})
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
        sizer.Add(btn_row, 0, wx.ALIGN_CENTER | wx.ALL, 5)

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

        dialog = wx.Dialog(self, title="选择运行中的程序", size=(560, 420))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer.Add(wx.StaticText(panel, label="双击添加程序："), 0, wx.ALL, 8)

        listbox = wx.ListBox(panel, size=(-1, 320))
        for p in running:
            text = f"{p['title']}  ——  {p['name']}" if p["title"] else p["name"]
            listbox.Append(text)
        sizer.Add(listbox, 1, wx.EXPAND | wx.ALL, 8)

        def on_double_click(_):
            i = listbox.GetSelection()
            if i == wx.NOT_FOUND:
                return
            item = running[i]
            self.programs.append({"name": item["name"], "path": item["path"], "hotkey": ""})
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

    def on_delete(self, event):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择一项", "提示")
            return
        self.programs.pop(idx)
        self.persist()
        self.refresh_list()
        self.register_all_hotkeys()

    def on_set_hotkey(self, event):
        idx = self.list_ctrl.GetFirstSelected()
        if idx < 0:
            wx.MessageBox("请先选择程序", "提示")
            return

        program = self.programs[idx]
        current = program.get("hotkey", "")

        msg = (
            f"当前快捷键：{current}\n\n"
            "输入新快捷键，例如：\n"
            "1) 单键：f1 ~ f12\n"
            "2) 组合：ctrl+1 / alt+q / ctrl+shift+f2\n"
            "（不区分大小写）"
        )
        dlg = wx.TextEntryDialog(self, msg, "设置快捷键", current)

        if dlg.ShowModal() == wx.ID_OK:
            hk = normalize_hotkey(dlg.GetValue())
            if not hk:
                # 清空快捷键
                self.programs[idx]["hotkey"] = ""
                self.persist()
                self.refresh_list()
                self.register_all_hotkeys()
                wx.MessageBox("已清空快捷键", "提示")
                dlg.Destroy()
                return

            # 先做格式校验
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


class QuickLauncherApp(wx.App):
    def OnInit(self):
        self.frame = QuickLauncherFrame()
        self.frame.Show()
        return True


if __name__ == "__main__":
    app = QuickLauncherApp()
    app.MainLoop()
