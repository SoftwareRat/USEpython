import os
import requests
import subprocess
import zipfile
import json
from tqdm import tqdm
import ctypes

# Function to create a folder if it doesn't exist
def create_folder(*directories):
    for dir in directories:
        if not os.path.exists(dir):
            os.makedirs(dir)


def set_console_title(title):
    ctypes.windll.kernel32.SetConsoleTitleW(title)


# Function to create a shortcut
def create_shortcut(target, shortcut_path):
    try:
        shell = ctypes.windll.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = target
        shortcut.IconLocation = target
        shortcut.save()
        # Only required for testing purposes
        # print(f"Shortcut created at: {shortcut_path}")
    except Exception as e:
        print(f"Error creating shortcut: {str(e)}")


# Define the directories for temporary downloads and software installation
base_temp_directory = "C:\\UseTemp"
base_install_directory = "C:\\Users\\kiosk\\AppData\\Local\\Programs"
create_folder(base_temp_directory, base_install_directory)


def change_wallpaper(image_path):
    try:
        # Set the wallpaper
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)

        # Notify Windows of the change
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 2)

        print(f"Desktop wallpaper set to '{image_path}' successfully.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

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
    ]
}

# Checks for JSON config and generate default if missing
if not os.path.isfile("use_conf.json"):
    with open('use_conf.json', 'w') as json_file:
        json.dump(USE_json, json_file, indent=4)

config = json.load(open('use_conf.json', 'r'))


def download_file_with_progress(url, save_path):
    response = requests.get(url, stream=True)
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


def extract_zip(zip_path, extract_path):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)


def install(software):
    if not software["enabled"]:
        return

    print(f"Installing {software['name']}...")
    if software['url']:
        print(f"Downloading {software['name']}...")
        download_file_with_progress(software['url'], software['path'])
        print(f"{software['name']} downloaded successfully.")

    # Extract ZIP archives to the installation directory
    if software['path'].endswith('.zip'):
        extract_dir = os.path.join(base_install_directory, os.path.basename(software['path']).replace('.zip', ''))
        extract_zip(software['path'], extract_dir)

    # Run installation command
    if software['install_command']:
        subprocess.run(software['install_command'], shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # Create shortcut for supported applications
    if software.get('shortcut', False) is True and software["enabled"]:
        exe_name = software['exe_name']
        exe_path = find_exe(base_install_directory, exe_name)
        if exe_path:
            shortcut_name = f"{software['name']}.lnk"
            shortcut_path = os.path.join(os.path.expanduser("~"), 'Desktop', shortcut_name)
            create_shortcut(exe_path, shortcut_path)
            # Only required for testing purposes
            #print(f"Shortcut for {software['name']} created successfully.")
        else:
            print(f"Executable for {software['name']} not found.")

    print(f"{software['name']} installed successfully.")


def find_exe(directory, exe_name):
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower() == exe_name.lower():
                return os.path.join(root, file)
    return None


def install_software():
    for software in config['default_binaries'] + config['custom_binaries']:
        install(software)


if __name__ == "__main__":
    set_console_title("Unauthorized Software Enabler by SoftwareRat")
    install_software()
    # TEMP: Set Windows 11 default as wallpaper
    change_wallpaper("C:\\Windows\\Web\\Wallpaper\\Windows\\img0.jpg")
