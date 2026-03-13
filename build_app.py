"""
Build Script — Package the Fact Entry Companion App as a standalone executable.

Uses PyInstaller to create a single-file executable for Windows and macOS.

Usage:
    python build_app.py          # Build for current platform
    python build_app.py --onedir # Build as a directory (faster startup)
"""

import argparse
import os
import platform
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description="Build standalone companion app")
    parser.add_argument("--onedir", action="store_true",
                        help="Build as directory instead of single file")
    args = parser.parse_args()

    print("=" * 60)
    print("Building Fact Entry Companion App")
    print(f"Platform: {platform.system()} {platform.machine()}")
    print("=" * 60)

    # Check PyInstaller is available
    try:
        import PyInstaller
        print(f"PyInstaller version: {PyInstaller.__version__}")
    except ImportError:
        print("PyInstaller not installed. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    app_name = "FactEntryApp"
    main_script = "fact_entry_app.py"

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", app_name,
        "--windowed",  # No console window
        "--noconfirm",
        "--clean",
        # Include all needed modules
        "--hidden-import", "msal",
        "--hidden-import", "openpyxl",
        "--hidden-import", "pyperclip",
        "--hidden-import", "requests",
        # Include config and other modules
        "--add-data", f"config{os.pathsep}config",
    ]

    if args.onedir:
        cmd.append("--onedir")
    else:
        cmd.append("--onefile")

    cmd.append(main_script)

    print(f"\nRunning: {' '.join(cmd)}\n")
    result = subprocess.run(cmd, cwd=os.path.dirname(os.path.abspath(__file__)))

    if result.returncode == 0:
        output_dir = "dist"
        print(f"\nBuild successful!")
        print(f"Output: {os.path.join(output_dir, app_name)}")
        if platform.system() == "Darwin":
            print(f"  macOS app: {os.path.join(output_dir, app_name + '.app')}")
        elif platform.system() == "Windows":
            print(f"  Windows exe: {os.path.join(output_dir, app_name + '.exe')}")
        print("\nDistribute this file to your team members.")
    else:
        print(f"\nBuild FAILED with return code {result.returncode}")
        sys.exit(1)


if __name__ == "__main__":
    main()
