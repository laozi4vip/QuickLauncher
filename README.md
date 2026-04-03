# QuickLauncher

Windows 任务栏程序快捷启动器


## 作者

- **作者**：laozi4vip
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
5. 想完美实现浏览器多用户窗口需要绑定Profile；
  方式1：永久生效，快捷方式添加目标: "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Profile XX" --user-data-dir="C:\Users\你的用户名\AppData\Local\Google\Chrome\XXXX"   #按实际路径修改
  方式2：临时，每次重启电脑后需要重新添加别名，以谷歌浏览器为例，为每个用户窗口添加别名，然后用title模式绑定。<img width="324" height="312" alt="图片" src="https://github.com/user-attachments/assets/b87525d8-f07a-44fe-b34f-ba2cca9025d6" />
## 构建

```bash
pip install pyinstaller pywin32 psutil wxpython
pyinstaller --onefile --windowed --collect-all wx --collect-all psutil --name QuickLauncher main/main.py
```

## 版本历史

- v3.0 新增窗口及任务栏图标隐藏功能
- v2.0 新增控制同浏览器多用户窗口功能，需要单独设置浏览器，以谷歌浏览器为例，多用户窗口需要绑定Profile；快捷方式添加目标: "C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Profile XX" --user-data-dir="C:\Users\你的用户名\AppData\Local\Google\Chrome\XXXX"   #按实际路径修改
- v1.0 略

## 许可证

MIT
