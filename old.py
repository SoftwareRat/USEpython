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
from enum import Enum
import comtypes
import comtypes.shelllink
import comtypes.client
import comtypes.persist

# Setup logger
logger_name = 'USE.log'
log_file = os.path.join(os.getenv("TEMP"), logger_name)
logger = logging.getLogger(logger_name)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler = logging.FileHandler(log_file)
file_handler.setFormatter(formatter)
file_handler.setLevel(logging.DEBUG)
stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.addHandler(stream_handler)

def get_caller_name(caller: bool):
    return inspect.currentframe().f_back.f_back.f_code.co_names if caller else inspect.currentframe().f_back.f_code.co_names

def create_folders(*directories):
    for dir in directories:
        os.makedirs(dir, exist_ok=True)  # Use exist_ok to avoid checking existence before creating

def set_console_title(title):
    ctypes.windll.kernel32.SetConsoleTitleW(title)

def download_image(image_path):
    try:
        image_data = requests.get(image_path).content
        image_path = os.path.join(os.getenv("USERPROFILE"), "Pictures", "wallpaper.jpg")
        with open(image_path, 'wb') as img_file:
            img_file.write(image_data)
    except Exception as e:
        logger.error(f"Failed to download image from {image_path}: {str(e)}")
        return None

    return image_path

def find_exe(directory, exe_name) -> str:
    for root, dirs, files in os.walk(directory):
        if exe_name.lower() in (file.lower() for file in files):
            return os.path.join(root, exe_name)
    return None

class reg_Types(Enum):
    dword = winreg.REG_DWORD
    string = winreg.REG_SZ
    multi_string = winreg.REG_MULTI_SZ

def set_reg_val(key: winreg.HKEYType, key_path: str, val: str, val_type: int, new_val):
    caller_name = inspect.stack()[1][3]
    try:
        with winreg.OpenKey(key, key_path, 0, winreg.KEY_SET_VALUE) as sel_key:
            winreg.SetValueEx(sel_key, val, 0, val_type, new_val)
    except Exception as e:
        logger.error(f'Unexpected error occurred at {get_caller_name(caller=False)} while being invoked by {get_caller_name(caller=True)}: {str(e)}')

def create_shortcut(target, shortcut_path):
    try:
        # Create ShellLink object
        shortcut = comtypes.client.CreateObject(comtypes.shelllink.ShellLink)
        shortcut_w = shortcut.QueryInterface(comtypes.shelllink.IShellLinkW)

        # Set the properties
        shortcut_w.SetPath(target)

        # Save the shortcut
        shortcut_file = shortcut.QueryInterface(comtypes.persist.IPersistFile)
        shortcut_file.Save(shortcut_path, True)

        logger.info(f"Shortcut created successfully at {shortcut_path}.")
    except Exception as e:
        logger.error(f'Unexpected error occurred at {get_caller_name(caller=False)} while being invoked by {get_caller_name(caller=True)}: {str(e)}')

base_temp_directory = "C:\\UseTemp"
base_install_directory = os.path.join(os.getenv("LOCALAPPDATA"), "Programs")
create_folders(base_temp_directory, base_install_directory)

def create_software_shortcut(software, config):
    if software.get('shortcut', False) is True and software["enabled"]:
        exe_path = find_exe(base_install_directory, software['exe_name'])
        if exe_path:
            shortcut_name = f"{software['name']}.lnk"
            shortcut_path = os.path.join(os.path.expanduser("~"), 'Desktop', shortcut_name)
            create_shortcut(exe_path, shortcut_path)
            logger.info(f"Shortcut for {software['name']} created successfully.")
        else:
            logger.error(f"Executable for {software['name']} not found.")

def change_wallpaper(image_path):
    try:
        if image_path.startswith(("http://", "https://")):
            # Download image if it's a web link
            image_path = download_image(image_path)

        # Set the wallpaper
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)
        # Notify Windows of the change
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 2)

        logger.info(f"Desktop wallpaper set to '{image_path}' successfully.")
    except Exception as e:
        logger.error(f'Unexpected error occurred at {get_caller_name(caller=False)} while being invoked by {get_caller_name(caller=True)}: {str(e)}')

# Structure of default JSON
DEFAULT_CONFIG = {
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
            "name": "Notepad++",
            "enabled": True,
            "url": 'https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.5.7/npp.8.5.7.portable.x64.zip',
            "path": os.path.join(base_temp_directory, 'npp.portable.x64.zip'),
            "install_command": None,
            "shortcut": True,
            "exe_name": "notepad++.exe"
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
            "name": "Google Chrome",
            "enabled": False,
            "url": 'https://dl.google.com/chrome/install/latest/chrome_installer.exe',
            "path": os.path.join(base_temp_directory, 'chrome-installer.exe'),
            "install_command": f'{os.path.join(base_temp_directory, "chrome-installer.exe")} /silent /install',
            "shortcut": True,
            "exe_name": "chrome.exe"
        },
    ],
    "custom_binaries": [
        {
            "name": "Example Software 1",
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

def read_config():
    # Checks for JSON config and generate default if missing
    if not os.path.isfile("use_conf.json"):
        with open('use_conf.json', 'w') as json_file:
            json.dump(DEFAULT_CONFIG, json_file, indent=4)

    return json.load(open('use_conf.json', 'r'))

def download_file_with_progress(url, save_path, overall_progress_bar):
    response = requests.get(url, stream=True)
    if response.status_code != 200:
        logger.error(f"Failed to download file from {url}. Status code: {response.status_code}")
        return
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

def extract_zip(zip_path, extract_path, overall_progress_bar):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)
        overall_progress_bar.update(os.path.getsize(zip_path))

def install(software, config, overall_progress_bar):
    if not software["enabled"]:
        return

    overall_progress_bar.set_postfix_str(f"Installing {software['name']}...")

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
    create_software_shortcut(software, config)

    logger.info(f"{software['name']} installed successfully.")

def install_software(config, overall_progress_bar):
    total_software = len(config['default_binaries']) + len(config['custom_binaries'])
    overall_progress_bar.total = total_software
    for software in config['default_binaries'] + config['custom_binaries']:
        logger.info(f"Installing {software['name']}...")
        install(software, config, overall_progress_bar)
        overall_progress_bar.update(1)

def apply_customization(config, overall_progress_bar):
    logger.info(f"Applying customization settings...")

    # Simulating a progress bar for customization
    overall_progress_bar.set_postfix_str(f"Applying customization settings...")
    overall_progress_bar.update(1)

    # Your customization logic here
    customization_options = config.get("customization", [])
    if customization_options:
        for option, value in customization_options[0].items():
            if option == "dark_mode" and value:
                set_reg_val(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", reg_Types.dword.value, 0)
                set_reg_val(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme", reg_Types.dword.value, 0)
                logger.info("Dark mode applied successfully.")
            elif option == "dark_mode" and not value:
                set_reg_val(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", reg_Types.dword.value, 1)
                set_reg_val(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme", reg_Types.dword.value, 1)
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

def main():
    overall_progress_bar = tqdm(total=0, unit='B', unit_scale=True, unit_divisor=1024, desc="Overall Progress")
    set_console_title("Unauthorized Software Enabler by SoftwareRat")
    try:
        config = read_config()
        install_software(config, overall_progress_bar)
        apply_customization(config, overall_progress_bar)
    finally:
        overall_progress_bar.close()

if __name__ == "__main__":
    main()