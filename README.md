# QuickLauncher

Windows 任务栏快捷启动器

## 作者

- **作者**：xinghui
- **GitHub**：https://github.com/laozi4vip/QuickLauncher

## 功能

- 绑定任务栏程序到快捷键
- 按快捷键启动或切换窗口
- 窗口在前台时最小化，再按恢复

## 支持的快捷键

- 数字键：0-9
- 字母键：a-z
- 功能键：F1-F12
- 修饰键：Alt、Ctrl、Shift

## 使用方法

1. 运行 `QuickLauncher.exe`
2. 点击"手动添加"或"从运行程序添加"添加程序
3. 选择程序后，点击"设置快捷键"绑定快捷键
4. 按下设置的快捷键即可快速启动或切换程序
5. 浏览器多用户窗口需要绑定Profile；快捷方式添加目标:"C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Profile xx" 
## 构建

```bash
pip install pyinstaller pywin32 psutil wxpython
pyinstaller --onefile --windowed --collect-all wx --collect-all psutil --name QuickLauncher main/main.py
```

## 版本历史

- v1.0.36 - 修复快捷键问题
- v1.0.35 - 修复语法错误
- v1.0.34 - 使用前台窗口检测改进切换逻辑
- v1.0.33 - 简化窗口切换逻辑
- v1.0.32 - 改进窗口查找和恢复逻辑
- v1.0.31 - 修复查找窗口时同时查找最小化的窗口

## 许可证

MIT
