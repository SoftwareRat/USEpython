import os
import requests
import subprocess
import zipfile
import json
from tqdm import tqdm
import winreg as reg
import ctypes

# Function to create a folder if it doesn't exist
def create_folder(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def set_console_title(title):
    ctypes.windll.kernel32.SetConsoleTitleW(title)

# Define the directories for temporary downloads and software installation
base_temp_directory = "C:\\UseTemp"
base_install_directory = "C:\\Users\\kiosk\\AppData\\Local\\Programs"
create_folder(base_temp_directory, base_install_directory)

# Structure of default JSON
USE_json = {
    "default_binaries": [
        {
            "name": "Mozilla Firefox",
            "enabled": True,
            "url": "https://download.mozilla.org/?product=firefox-latest-ssl&os=win64",
            "path": os.path.join(base_temp_directory, 'FirefoxSetup.exe'),
            "install_command": f'{os.path.join(base_temp_directory, "FirefoxSetup.exe")} /InstallDirectoryPath={os.path.join(base_install_directory, "Firefox")}',
            "progress_message": 'Installing Firefox'
        },
        {
            "name": "Process Explorer",
            "enabled": True,
            "url": 'https://download.sysinternals.com/files/ProcessExplorer.zip',
            "path": os.path.join(base_temp_directory, 'ProcessExplorer.zip'),
            "install_command": None,
            "progress_message": 'Installing Process Explorer'
        },
        {
            "name": "7-Zip",
            "enabled": True,
            "url": 'https://7-zip.org/a/7z2301-x64.exe',
            "path": os.path.join(base_temp_directory, '7z-x64.exe'),
            "install_command": f'{os.path.join(base_temp_directory, "7z-x64.exe")} /S /D={os.path.join(base_install_directory, "7-Zip")}',
            "progress_message": 'Installing 7-Zip'
        },
        {
            "name": "Notepad++",
            "enabled": False,
            "url": '',
            "path": os.path.join(base_temp_directory, 'npp.7.9.5.Installer.x64.exe'),
            "install_command": f'{os.path.join(base_temp_directory, "npp.7.9.5.Installer.x64.exe")} /S',
            "progress_message": 'Installing Notepad++'
        },

    ],
    "custom_binaries": [
        {
            "name": "Example Software 1",
            "enabled": False,
            "url": "https://example.com/custom-software1.zip",
            "path": os.path.join(base_temp_directory, 'custom-software1.zip'),
            "install_command": None,
            "progress_message": 'Installing Custom Software 1'
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

    print(f"Downloading {software['name']}...")
    download_file_with_progress(software['url'], software['path'])
    print(f"{software['name']} downloaded successfully.")

    # Install software
    print(f"{software['progress_message']}...")

    # Extract ZIP archives to the installation directory
    if software['path'].endswith('.zip'):
        extract_dir = os.path.join(base_install_directory, os.path.basename(software['path']).replace('.zip', ''))
        extract_zip(software['path'], extract_dir)

    # Run installation command
    if software['install_command'] is not None:
        subprocess.run(software['install_command'], shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    print(f"{software['name']} installed successfully.")

def install_software():
    for software in config['default_binaries'] + config['custom_binaries']:
        install(software)

if __name__ == "__main__":
    set_console_title("Unauthorized Software Enabler by SoftwareRat")
    install_software()