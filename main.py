import os
import requests
import subprocess
import zipfile
import json 
from tqdm import tqdm
import winreg as reg

# Function to create a folder if it doesn't exist
def create_folder(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

# Define the base directory for temporary downloads
base_temp_directory = "C:\\UseTemp"
create_folder(base_temp_directory)

# Structure of default JSON
USE_json = {
    "default_binaries" : [
        {
            "Mozilla Firefox" : True,
            "Process Explorer" : True,
            "Explorer++": True,
            "7-Zip": True,
            "Notepad++": True,
            "Regcool": True
        }
    ],
    "custom_binaries" : [
        {

        }
    ]
}

# Checks for JSON config and generate default if missing
if not os.path.isfile("use_conf.json"):
    with open('use_conf.json', 'w') as json_file:
        json.dump(USE_json, json_file, indent=4)


config = json.load(open('use_conf.json', 'r'))['default_binaries']


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

def install_software():


    # Define the software installation commands and progress messages
    software_list = [
        ("Mozilla Firefox", 'https://download.mozilla.org/?product=firefox-latest-ssl&os=win64', os.path.join(base_temp_directory, 'FirefoxSetup.exe'), 'Installing Firefox'),
        ("Process Explorer", 'https://download.sysinternals.com/files/ProcessExplorer.zip', os.path.join(base_temp_directory, 'ProcessExplorer.zip'), 'Installing Process Explorer'),
        ("Explorer++", 'https://github.com/derceg/explorerplusplus/releases/download/version-1.4.0-beta-2/explorerpp_x64.zip', os.path.join(base_temp_directory, 'explorerpp_x64.zip'), 'Installing Explorer++'),
        # TODO: Auto-updating link
        ("7-Zip", 'https://7-zip.org/a/7z2301-x64.exe', os.path.join(base_temp_directory, 'install7Zip.exe'), 'Installing 7-Zip'),
        ("Notepad++", 'https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.5.7/npp.8.5.7.portable.x64.zip', os.path.join(base_temp_directory, 'NotepadPlusPlus.zip'), 'Installing Notepad++'),
        ("Regcool", 'https://kurtzimmermann.com/files/RegCoolX64.zip', os.path.join(base_temp_directory, 'RegCoolX64.zip'), 'Installing Regcool')
    ]

    for software_name, url, download_path, progress_message in software_list:
        for name, download_approval in config[0].items(): # This communicate with the config to decide wether or not user wants it
            if name == software_name and download_approval == False:
                break
            print(f"Downloading {software_name}...")
            download_file_with_progress(url, download_path)
            print(f"{software_name} downloaded successfully.")

            # Install software
            print(f"{progress_message}...")

            # Extract ZIP archives natively in Python
            if download_path.endswith('.zip'):
                extract_dir = os.path.dirname(download_path)
                extract_zip(download_path, extract_dir)
            
            # Run installation command
            if not download_path.endswith('.zip'):
                subprocess.run(download_path, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

            print(f"{software_name} installed successfully.")
            input()
            print(f"Downloading {software_name}...")
            download_file_with_progress(url, download_path)
            print(f"{software_name} downloaded successfully.")

            # Install software
            print(f"{progress_message}...")

            # Extract ZIP archives natively in Python
            if download_path.endswith('.zip'):
                extract_dir = os.path.dirname(download_path)
                extract_zip(download_path, extract_dir)
            
            # Run installation command
            if not download_path.endswith('.zip'):
                subprocess.run(download_path, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

            print(f"{software_name} installed successfully.")

    # Enabling Windows dark mode
    try:
        reg_key_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
        with reg.OpenKey(reg.HKEY_CURRENT_USER, reg_key_path, 0, reg.KEY_WRITE) as key:
            reg.SetValueEx(key, 'ListviewShadow', 0, reg.REG_DWORD, 1)
        print("Registry key 'ListviewShadow' set successfully.")
    except Exception as e:
        print(f"Error setting the registry key: {str(e)}")

if __name__ == "__main__":
    install_software()