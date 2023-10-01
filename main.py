import os
import requests
import subprocess
import zipfile
import json
from tqdm import tqdm
import ctypes
import winreg
import logging
import inspect
import shutil
from enum import Enum

logger_name = 'USE.log'
log_file = os.path.join(os.getenv("TEMP"), logger_name)

# Logging setup
LOGGING_CONFIG = {
    'version': 1,
    'disable_existing_loggers': False,
    'handlers': {
        'file': {
            'level': 'DEBUG',
            'class': 'logging.FileHandler',
            'filename': log_file,
            'formatter': 'verbose',
        },
    },
    'formatters': {
        'verbose': {
            'format': '{levelname} {asctime} {module} {message}',
            'style': '{',
        },
    },
    'loggers': {
        logger_name: {
            'handlers': ['file'],
            'level': 'DEBUG',
        },
    },
}

logging.config.dictConfig(LOGGING_CONFIG)
logger = logging.getLogger(logger_name)


class RegistryTypes(Enum):
    DWORD = winreg.REG_DWORD
    STRING = winreg.REG_SZ
    MULTI_STRING = winreg.REG_MULTI_SZ


def create_folder(*directories):
    for dir in directories:
        os.makedirs(dir, exist_ok=True)

def set_console_title(title):
    ctypes.windll.kernel32.SetConsoleTitleW(title)

def find_exe(directory, exe_name):
    for root, dirs, files in os.walk(directory):
        if exe_name.lower() in (file.lower() for file in files):
            return os.path.join(root, exe_name)
    return None

def get_func_name(caller: bool):
    return inspect.currentframe().f_back.f_back.f_code.co_names if caller else inspect.currentframe().f_back.f_code.co_names

def set_reg_val(key, key_path, val, val_type, new_val):
    caller_name = inspect.stack()[1][3]
    with winreg.OpenKey(key, key_path, 0, winreg.KEY_SET_VALUE) as sel_key:
        try:
            winreg.SetValueEx(sel_key, val, 0, val_type, new_val)
        except Exception as e:
            log_error(caller_name, e)

base_temp_directory = "C:\\UseTemp"
base_install_directory = "C:\\Users\\kiosk\\AppData\\Local\\Programs"
create_folder(base_temp_directory, base_install_directory)

def change_wallpaper(image_path):
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
    overall_progress_bar.set_postfix_str("Installing WinXShell...")

    # Source directory where WinXShell is already extracted
    winxshell_dir = os.path.join(base_install_directory, 'WinXShell')

    # Destination directory within the temporary directory
    temp_winxshell_dir = os.path.join(base_temp_directory, 'WinXShell')

    # Ensure the destination directory exists
    create_folder(temp_winxshell_dir)

    # Copy all contents of X_PF/WinXShell to the extracted root folder
    for root, dirs, files in os.walk(os.path.join(winxshell_dir, 'X_PF', 'WinXShell')):
        for file in files:
            src_path = os.path.join(root, file)
            dst_path = os.path.join(base_temp_directory, file)
            shutil.copy(src_path, dst_path)

    # Delete main.bat
    main_bat_path = os.path.join(base_temp_directory, 'main.bat')
    os.remove(main_bat_path)

    # Delete files with names containing "zh-CN" and "x86"
    for root, dirs, files in os.walk(base_temp_directory):
        for file in files:
            if "zh-CN" in file or "x86" in file:
                file_path = os.path.join(root, file)
                os.remove(file_path)

    # Delete X_PF/WinXShell
    shutil.rmtree(os.path.join(base_temp_directory, 'X_PF', 'WinXShell'))

    # Make 2 copies of WinXShell_x64.exe
    winxshell_exe_path = os.path.join(base_temp_directory, 'WinXShell_x64.exe')
    explorer_exe_path = os.path.join(base_temp_directory, 'explorer.exe')
    gfndesktop_exe_path = os.path.join(base_temp_directory, 'gfndesktop.exe')

    shutil.copy(winxshell_exe_path, explorer_exe_path)
    shutil.copy(winxshell_exe_path, gfndesktop_exe_path)

    # Rename WinXShell_x64.exe to WinXShell.exe
    winxshell_dest_path = os.path.join(base_temp_directory, 'WinXShell.exe')
    os.rename(winxshell_exe_path, winxshell_dest_path)

    # Make a shortcut to copied explorer.exe
    create_shortcut(explorer_exe_path, os.path.join(os.path.expanduser("~"), 'Desktop', 'WinXShell.lnk'))

    overall_progress_bar.set_postfix_str("WinXShell installed successfully.")
    overall_progress_bar.update(1)


def log_error(caller_name, error):
    logger.error(f'Unexpected error occurred at {caller_name}: {str(error)}')

def download_file_with_progress(url, save_path, overall_progress_bar):
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()
        total_size = int(response.headers.get('content-length', 0))

        with open(save_path, 'wb') as file, tqdm(
                total=total_size,
                unit='B',
                unit_scale=True,
                unit_divisor=1024,
                desc=os.path.basename(save_path)
        ) as progress_bar:
            chunk_size = 1024
            for data in response.iter_content(chunk_size=chunk_size):
                file.write(data)
                progress_bar.update(len(data))
                overall_progress_bar.update(len(data))
    except requests.RequestException as e:
        log_error(inspect.currentframe().f_code.co_name, e)


def extract_zip(zip_path, extract_path, overall_progress_bar):
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)
            overall_progress_bar.update(os.path.getsize(zip_path))
    except zipfile.BadZipFile as e:
        log_error(inspect.currentframe().f_code.co_name, e)


def install(software, overall_progress_bar):
    if not software["enabled"]:
        return

    overall_progress_bar.set_postfix_str(f"Installing {software['name']}...")

    try:
        if software['url']:
            download_file_with_progress(software['url'], software['path'], overall_progress_bar)

        # Extract ZIP archives to the installation directory
        if software['path'].endswith('.zip'):
            extract_dir = os.path.join(base_install_directory, os.path.basename(software['path']).replace('.zip', ''))
            extract_zip(software['path'], extract_dir, overall_progress_bar)

        # Run installation command
        if software['install_command']:
            subprocess.run(software['install_command'], shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        # Create shortcut for supported applications
        if software.get('shortcut', False) and software["enabled"]:
            exe_path = find_exe(base_install_directory, software['exe_name'])
            if exe_path:
                shortcut_name = f"{software['name']}.lnk"
                shortcut_path = os.path.join(os.path.expanduser("~"), 'Desktop', shortcut_name)
                create_shortcut(exe_path, shortcut_path)
                logger.info(f"Shortcut for {software['name']} created successfully.")
            else:
                logger.error(f"Executable for {software['name']} not found.")

        logger.info(f"{software['name']} installed successfully.")
    except Exception as e:
        log_error(inspect.currentframe().f_code.co_name, e)


def install_software(overall_progress_bar):
    try:
        total_software = len(config['default_binaries']) + len(config['custom_binaries']) + 1  # Adding 1 for WinXShell
        overall_progress_bar.total = total_software

        # Install WinXShell separately
        install_winxshell(overall_progress_bar)

        for software in config['default_binaries'] + config['custom_binaries']:
            logger.info(f"Installing {software['name']}...")
            install(software, overall_progress_bar)
            overall_progress_bar.update(1)
    except Exception as e:
        log_error(inspect.currentframe().f_code.co_name, e)


def apply_customization(overall_progress_bar):
    try:
        logger.info(f"Applying customization settings...")

        # Simulating a progress bar for customization
        overall_progress_bar.set_postfix_str(f"Applying customization settings...")
        overall_progress_bar.update(1)

        # Your customization logic here
        customization_options = config.get("customization", [])
        if customization_options:
            for option, value in customization_options[0].items():
                if option == "dark_mode" and value:
                    set_reg_val(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", RegistryTypes.DWORD, 0)
                    logger.info("Dark mode applied successfully.")
                elif option == "dark_mode" and not value:
                    set_reg_val(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", RegistryTypes.DWORD, 1)
                    logger.info("Light mode applied successfully.")
                elif option == "wallpaper_path" and value:
                    # Set the wallpaper
                    change_wallpaper(value)
                    logger.info(f"Wallpaper set to '{value}' successfully.")
                # Add more conditions for other customization options

        # Simulating the completion of customization
        overall_progress_bar.set_postfix_str(f"Customization settings applied.")
        overall_progress_bar.update(1)

        logger.info(f"Customization settings applied successfully.")
    except Exception as e:
        log_error(inspect.currentframe().f_code.co_name, e)


def main():
    overall_progress_bar = tqdm(total=0, unit='B', unit_scale=True, unit_divisor=1024, desc="Overall Progress")
    set_console_title("Unauthorized Software Enabler by SoftwareRat")
    try:
        install_software(overall_progress_bar)
        apply_customization(overall_progress_bar)
    except Exception as e:
        log_error(inspect.currentframe().f_code.co_name, e)
    finally:
        overall_progress_bar.close()


if __name__ == "__main__":
    main()