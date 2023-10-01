import os
import json
import shutil
import zipfile
import time
import ctypes
import inspect
import subprocess
from tqdm import tqdm
import requests
import logging
import winreg
import psutil
from rich.console import Console
from rich.progress import Progress, BarColumn, TextColumn, DownloadColumn
from rich.live import Live
from enum import Enum

logger_name = 'USE.log'
log_file = os.path.join(os.getenv("TEMP"), logger_name)

# Logging setup
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(levelname)s %(asctime)s %(module)s %(message)s',
)

logger = logging.getLogger(logger_name)


class RegistryTypes(Enum):
    DWORD = winreg.REG_DWORD
    STRING = winreg.REG_SZ
    MULTI_STRING = winreg.REG_MULTI_SZ


def create_folder(*directories):
    """Create folders if they don't exist."""
    for directory in directories:
        os.makedirs(directory, exist_ok=True)


def set_console_title(title):
    """Set console title."""
    ctypes.windll.kernel32.SetConsoleTitleW(title)


def find_exe(directory, exe_name):
    """Find an executable in a directory."""
    for root, dirs, files in os.walk(directory):
        if exe_name.lower() in (file.lower() for file in files):
            return os.path.join(root, exe_name)
    return None


def get_func_name(caller: bool):
    """Get the function name."""
    return inspect.currentframe().f_back.f_back.f_code.co_names if caller else inspect.currentframe().f_back.f_code.co_names


def set_reg_val(key, key_path, val, val_type, new_val):
    """Set registry value."""
    caller_name = inspect.stack()[1][3]
    with winreg.OpenKey(key, key_path, 0, winreg.KEY_SET_VALUE) as sel_key:
        try:
            winreg.SetValueEx(sel_key, val, 0, val_type.value, new_val)
        except Exception as e:
            log_error(caller_name, e)


base_temp_directory = "C:\\UseTemp"
base_install_directory = os.path.join(os.getenv("LOCALAPPDATA"), "Programs")
create_folder(base_temp_directory, base_install_directory)


def stop_process_by_name(process_name):
    """Stop a process by name."""
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == process_name:
            try:
                # Terminate the process
                psutil.Process(process.info['pid']).terminate()
                print(f"Process {process_name} terminated successfully.")
            except psutil.NoSuchProcess:
                print(f"Process {process_name} not found.")
            except psutil.AccessDenied:
                print(f"Access denied to terminate process {process_name}.")


def change_wallpaper(image_path):
    """Change desktop wallpaper."""
    try:
        if image_path.startswith(("http://", "https://")):
            # Download image if it's a web link
            image_data = requests.get(image_path).content
            image_path = os.path.join(os.getenv("USERPROFILE"), "Pictures", "wallpaper.jpg")
            with open(image_path, 'wb') as img_file:
                img_file.write(image_data)

        # Set the wallpaper
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)
        # Notify Windows of the change
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 2)

        logger.info(f"Desktop wallpaper set to '{image_path}' successfully.")
    except Exception as e:
        logger.error(f'Unexpected error occurred at {get_func_name(caller=False)} while being invoked by {get_func_name(caller=True)}: {str(e)}')


# Structure of default JSON
USE_json = {
    "default_binaries": [
        {
            "name": "Mozilla Firefox",
            "enabled": True,
            "url": "https://download.mozilla.org/?product=firefox-latest-ssl&os=win64",
            "path": os.path.join(base_temp_directory, 'FirefoxSetup.exe'),
            "install_command": f'{os.path.join(base_temp_directory, "FirefoxSetup.exe")} /InstallDirectoryPath={os.path.join(base_install_directory, "Firefox")}',
            "shortcut": False,
            "exe_name": "firefox.exe"
        },
        {
            "name": "Process Explorer",
            "enabled": True,
            "url": 'https://download.sysinternals.com/files/ProcessExplorer.zip',
            "path": os.path.join(base_temp_directory, 'ProcessExplorer.zip'),
            "install_command": None,
            "shortcut": True,
            "exe_name": "procexp64.exe"
        },
        {
            "name": "7-Zip",
            "enabled": True,
            "url": 'https://7-zip.org/a/7z2301-x64.exe',
            "path": os.path.join(base_temp_directory, '7z-x64.exe'),
            "install_command": f'{os.path.join(base_temp_directory, "7z-x64.exe")} /S /D={os.path.join(base_install_directory, "7-Zip")}',
            "shortcut": True,
            "exe_name": "7zFM.exe"
        },
        {
            "name": "WinRAR",
            "enabled": False,
            "url": 'https://www.rarlab.com/rar/winrar-x64-624b1.exe',
            "path": os.path.join(base_temp_directory, 'winrar-x64.exe'),
            "install_command": f'{os.path.join(base_temp_directory, "winrar-x64.exe")} /S /D={os.path.join(base_install_directory, "WinRAR")}',
            "shortcut": True,
            "exe_name": "winrar.exe"
        },
        {
            "name": "Notepad++",
            "enabled": True,
            "url": 'https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.5.7/npp.8.5.7.portable.x64.zip',
            "path": os.path.join(base_temp_directory, 'npp.portable.x64.zip'),
            "install_command": None,
            "shortcut": True,
            "exe_name": "notepad++.exe"
        },
        {
            "name": "Cmder",
            "enabled": False,
            "url": 'https://github.com/cmderdev/cmder/releases/download/v1.3.24/cmder_mini.zip',
            "path": os.path.join(base_temp_directory, 'cmder.zip'),
            "install_command": None,
            "shortcut": True,
            "exe_name": "cmder.exe"
        },
        {
            "name": "VLC Media Player",
            "enabled": True,
            "url": 'https://get.videolan.org/vlc/3.0.18/win64/vlc-3.0.18-win64.zip',
            "path": os.path.join(base_temp_directory, 'vlc-win64.zip'),
            "install_command": None,
            "shortcut": True,
            "exe_name": "vlc.exe"
        },
        {
            "name": "RetroArch",
            "enabled": False,
            "url": 'https://buildbot.libretro.com/stable/1.16.0/windows/x86_64/RetroArch-Win64-setup.exe',
            "path": os.path.join(base_temp_directory, 'retroarch-setup.exe'),
            "install_command": f'{os.path.join(base_temp_directory, "retroarch-setup.exe")} /S /D={os.path.join(base_install_directory, "RetroArch")}',
            "shortcut": True,
            "exe_name": "retroarch.exe"
        },
        {
            "name": "Visual Studio Code",
            "enabled": False,
            "url": 'https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user',
            "path": os.path.join(base_temp_directory, 'vscode-setup.exe'),
            "install_command": f'{os.path.join(base_temp_directory, "vscode-setup.exe")} /silent /install',
            "shortcut": True,
            "exe_name": "code.exe"
        },
        {
            "name": "Google Chrome",
            "enabled": False,
            "url": 'https://dl.google.com/chrome/install/latest/chrome_installer.exe',
            "path": os.path.join(base_temp_directory, 'chrome-installer.exe'),
            "install_command": f'{os.path.join(base_temp_directory, "chrome-installer.exe")} /silent /install',
            "shortcut": True,
            "exe_name": "chrome.exe"
        },
        {
            "name": "Arcade",
            "enabled": False,
            "url": 'https://picteon.dev/files/Arcade.exe',
            "path": os.path.join(base_temp_directory, 'Arcade.exe'),
            "install_command": None,
            "shortcut": False,
            "exe_name": "Arcade.exe"
        },
        {
            "name": "DOSBox",
            "enabled": False,
            "url": 'https://downloads.sourceforge.net/project/dosbox/dosbox/0.74-3/DOSBox0.74-3-win32-installer.exe',
            "path": os.path.join(base_temp_directory, 'dosbox.exe'),
            "install_command": None,
            "shortcut": False,
            "exe_name": "dosbox.exe"
        },
        {
            "name": "WinSCP",
            "enabled": False,
            "url": 'https://winscp.net/download/WinSCP-6.1.2-Setup.exe',
            "path": os.path.join(base_temp_directory, 'winscp.exe'),
            "install_command": None,
            "shortcut": False,
            "exe_name": "winscp.exe"
        },
        {
            "name": "WinMerge",
            "enabled": False,
            "url": 'https://downloads.sourceforge.net/project/winmerge/winmerge/2.16.16/WinMerge-2.16.16-Setup.exe',
            "path": os.path.join(base_temp_directory, 'winmerge.exe'),
            "install_command": None,
            "shortcut": False,
            "exe_name": "winmerge.exe"
        },
        {
            "name": "WinXShell",
            "enabled": True,
            "url": 'https://files.rycoh.net/WinXShell.zip',
            "path": os.path.join(base_temp_directory, 'WinXShell.zip'),
            "install_command": None,
            "shortcut": False,
            "exe_name": "explorer.exe"
        },
    ],
    "custom_binaries": [
        {
            "name": "Example Software 1 - do not use please",
            "enabled": False,
            "url": "https://example.com/custom-software1.zip",
            "path": os.path.join(base_temp_directory, 'custom-software1.zip'),
            "install_command": None,
            "shortcut": False,
            "exe_name": "example.exe"
        },
    ],
    "customization": [
        {
            "dark_mode": True,
            "wallpaper_path": "C:\\Windows\\Web\\Wallpaper\\Windows\\img0.jpg"
        }
    ]
}

# Checks for JSON config and generate default if missing
if not os.path.isfile("use_conf.json"):
    with open('use_conf.json', 'w') as json_file:
        json.dump(USE_json, json_file, indent=4)

config = json.load(open('use_conf.json', 'r'))


def create_shortcut(target, shortcut_path):
    """Create a shortcut."""
    try:
        shell = ctypes.windll.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = target
        shortcut.IconLocation = target
        shortcut.save()
        logger.info(f"Shortcut created successfully at {shortcut_path}.")
    except Exception as e:
        log_error(inspect.currentframe().f_code.co_name, e)


def install_winxshell(overall_progress_bar):
<<<<<<< HEAD
    """Install WinXShell."""
    try:
        overall_progress_bar.set_postfix_str("Installing WinXShell...")

        # Source directory where WinXShell is already present
        winxshell_dir = os.path.join(base_install_directory, 'WinXShell')

        # Delete main.bat in base_install_directory/WinXShell
        main_bat_path = os.path.join(winxshell_dir, 'main.bat')
        os.remove(main_bat_path)

        # Delete files with names containing "zh-CN" and "x86" in base_install_directory
        for root, dirs, files in os.walk(winxshell_dir):
            for file in files:
                if "zh-CN" in file or "x86" in file:
                    file_path = os.path.join(root, file)
                    os.remove(file_path)

        # Copy all contents of X_PF/WinXShell to the extracted root folder in base_install_directory
        for root, dirs, files in os.walk(os.path.join(winxshell_dir, 'X_PF', 'WinXShell')):
            for file in files:
                src_path = os.path.join(root, file)
                dest_path = os.path.join(base_install_directory, file)
                shutil.copy2(src_path, dest_path)

        logger.info("WinXShell installed successfully.")
    except Exception as e:
        log_error(inspect.currentframe().f_code.co_name, e)
=======
    overall_progress_bar.set_postfix_str("Installing WinXShell...")

    # Source directory where WinXShell is already extracted
    winxshell_dir = os.path.join(base_install_directory, 'WinXShell')

    # Copy all contents of X_PF/WinXShell to the extracted root folder in base_install_directory
    for root, dirs, files in os.walk(os.path.join(winxshell_dir, 'X_PF', 'WinXShell')):
        for file in files:
            src_path = os.path.join(root, file)
            dst_path = os.path.join(base_install_directory, file)
            shutil.copy(src_path, dst_path)

    # Delete main.bat in base_install_directory
    main_bat_path = os.path.join(base_install_directory, 'main.bat')
    os.remove(main_bat_path)

    # Delete files with names containing "zh-CN" and "x86" in base_install_directory
    for root, dirs, files in os.walk(base_install_directory):
        for file in files:
            if "zh-CN" in file or "x86" in file:
                file_path = os.path.join(root, file)
                os.remove(file_path)

    # Delete X_PF/WinXShell in base_install_directory
    shutil.rmtree(os.path.join(base_install_directory, 'X_PF', 'WinXShell'))

    # Make 2 copies of WinXShell_x64.exe
    winxshell_exe_path = os.path.join(base_install_directory, 'WinXShell_x64.exe')
    explorer_exe_path = os.path.join(base_install_directory, 'explorer.exe')
    gfndesktop_exe_path = os.path.join(base_install_directory, 'gfndesktop.exe')

    shutil.copy(winxshell_exe_path, explorer_exe_path)
    shutil.copy(winxshell_exe_path, gfndesktop_exe_path)

    # Rename WinXShell_x64.exe to WinXShell.exe
    winxshell_dest_path = os.path.join(base_install_directory, 'WinXShell.exe')
    os.rename(winxshell_exe_path, winxshell_dest_path)

    # Make a shortcut to copied explorer.exe on the Desktop
    create_shortcut(explorer_exe_path, os.path.join(os.path.expanduser("~"), 'Desktop', 'WinXShell.lnk'))

    overall_progress_bar.set_postfix_str("WinXShell installed successfully.")
    overall_progress_bar.update(1)
>>>>>>> parent of 1b8fc23 (- WinXShell function fix)


def download_file(url, dest_path, chunk_size=128):
    """Download a file."""
    try:
        response = requests.get(url, stream=True)
        total_size = int(response.headers.get('content-length', 0))
        block_size = chunk_size
        wrote = 0
        with open(dest_path, 'wb') as f, tqdm(total=total_size, unit='B', unit_scale=True, unit_divisor=1024, desc=os.path.basename(dest_path), ncols=100) as bar:
            for data in response.iter_content(chunk_size=block_size):
                wrote += len(data)
                f.write(data)
                bar.update(len(data))
        return dest_path
    except Exception as e:
        log_error(inspect.currentframe().f_code.co_name, e)


def extract_zip(zip_path, extract_path):
    """Extract a ZIP file."""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)
        return extract_path
    except Exception as e:
        log_error(inspect.currentframe().f_code.co_name, e)


def install_binary(binary_info, overall_progress_bar):
    """Install a binary."""
    try:
        overall_progress_bar.set_postfix_str(f"Installing {binary_info['name']}...")

        if binary_info["url"]:
            download_path = download_file(binary_info["url"], binary_info["path"])
            logger.info(f"Downloaded {binary_info['name']} to {download_path}")

        if binary_info["path"].endswith('.zip'):
            extract_path = extract_zip(binary_info["path"], os.path.join(base_temp_directory, binary_info["name"]))
            logger.info(f"Extracted {binary_info['name']} to {extract_path}")
        else:
            extract_path = None

        if binary_info["install_command"]:
            subprocess.run(binary_info["install_command"], shell=True, check=True, cwd=extract_path)

        if binary_info["shortcut"]:
            create_shortcut(find_exe(extract_path or base_install_directory, binary_info["exe_name"]), os.path.join(os.path.join(os.path.join(base_install_directory, binary_info["name"])), f"{binary_info['name']}.lnk"))

        logger.info(f"Installed {binary_info['name']} successfully.")
    except Exception as e:
        log_error(inspect.currentframe().f_code.co_name, e)


def log_error(caller_name, error):
    """Log errors."""
    logger.error(f"Error occurred in {caller_name}: {str(error)}")


def main():
    """Main function."""
    try:
        set_console_title("Unauthorized Software Enabler by SoftwareRat")

        progress_bar_columns = [
            BarColumn(),
            TextColumn("[progress.description]{task.description}"),
            DownloadColumn(),
        ]

        overall_progress_bar = Progress(
            "{task.completed}/{task.total} [{task.percentage:>3.0f}%]",
            auto_refresh=False,
            transient=True,
            console=Console,
            redirect_stdout=True,
            redirect_stderr=True,
            disable=False,
            refresh_per_second=1,
            refresh_callback=None,
            bar_template='{task.completed}/{task.total} [{task.percentage:>3.0f}%] {task.fields[download_bar]}',
            console_width=80,
            console_hide_cursor=True,
            disable_move=False,
            disable_resize=False,
            disable_close=False,
            disable_enter=False,
            disable_clear=False,
            disable_backspace=False,
            disable_close_eof=False,
            disable_styled_output=False,
            refresh_fill_char=' ',
            refresh_empty_char=' ',
            refresh_iterations=None,
            transient_per_second=None,
            transient_iterations=None,
            transient_percent=None,
            update_period=None,
            task_refresh_per_second=1,
            task_refresh_margin=None,
            min_width=None,
            min_delta=None,
            download_token=None,
            refresh_spinner=True,
            progress_bar_width=None,
            transient_time=None,
            task_description=None,
            task_id=None,
            task_set=None,
            tasks=None,
            stats=None,
            cleanup_on_interrupt=True,
            unknown=None,
        )

        # Calculate total tasks
        total_tasks = len(config["default_binaries"]) + len(config["custom_binaries"]) + len(config["customization"])
        overall_progress_bar.total = total_tasks

        # Start installing default binaries
        for binary_info in config["default_binaries"]:
            if binary_info["enabled"]:
                install_binary(binary_info, overall_progress_bar)
                overall_progress_bar.update(advance=1)

        # Start installing custom binaries
        for binary_info in config["custom_binaries"]:
            if binary_info["enabled"]:
                install_binary(binary_info, overall_progress_bar)
                overall_progress_bar.update(advance=1)

        # Start customizations
        for customization_info in config["customization"]:
            if customization_info["dark_mode"]:
                set_reg_val(winreg.HKEY_CURRENT_USER, 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize', 'AppsUseLightTheme', RegistryTypes.DWORD, 0)
                set_reg_val(winreg.HKEY_CURRENT_USER, 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize', 'SystemUsesLightTheme', RegistryTypes.DWORD, 0)
                logger.info("Dark mode enabled.")

            if customization_info["wallpaper_path"]:
                change_wallpaper(customization_info["wallpaper_path"])

            overall_progress_bar.update(advance=1)

        # Special handling for WinXShell
        if find_exe(base_install_directory, "WinXShell.exe"):
            install_winxshell(overall_progress_bar)

        overall_progress_bar.refresh()
        overall_progress_bar.stop()

        # Close the console after 5 seconds
        time.sleep(5)
        subprocess.run("taskkill /F /IM cmd.exe", shell=True, check=False)

    except Exception as e:
        log_error(inspect.currentframe().f_code.co_name, e)


if __name__ == "__main__":
    main()