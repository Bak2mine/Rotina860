import PyInstaller.__main__
import os
import sys

# Get the directory where this script is located
script_dir = os.path.dirname(os.path.abspath(__file__))
script_path = os.path.join(script_dir, "860 auto.py")

# Build the executable
PyInstaller.__main__.run([
    script_path,
    "--onefile",
    "--console",
    "--name=Rotina860",
    f"--distpath={os.path.join(script_dir, 'dist')}",
    f"--workpath={os.path.join(script_dir, 'build')}",
    f"--specpath={os.path.join(script_dir, 'build')}",
    f"--add-data={script_dir}/*.png:.",
    "--hidden-import=pyautogui",
    "--hidden-import=pyperclip",
    "--hidden-import=pandas",
    "--hidden-import=openpyxl",
    "--hidden-import=pygetwindow",
    "--hidden-import=win32gui",
    "--hidden-import=win32con",
    "--hidden-import=keyring",
])

print("\n✅ Build complete! Executable created in 'dist' folder")
