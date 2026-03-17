from cx_Freeze import setup, Executable
import sys

build_exe_options = {
    "packages": [],
    "excludes": []
}

setup(
    name="QuickLauncher",
    version="1.0",
    description="Windows快捷启动器",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "main/main.py",
            target_name="QuickLauncher.exe",
            base="Win32GUI"
        )
    ]
)
