import os
import requests
import subprocess
import zipfile
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
        ("Regcool", 'https://kurtzimmermann.com/files/RegCoolX64.zip', os.path.join(base_temp_directory, 'RegCoolX64.zip'), 'Installing Regcool'),
        ("VLC media player", 'https://get.videolan.org/vlc/3.0.18/win64/vlc-3.0.18-win64.zip', os.path.join(base_temp_directory, 'vlc-win64.zip'), 'Installing VLC media player')        
    ]

    for software_name, url, download_path, progress_message in software_list:
        print(f"Downloading {software_name}...")
        download_file_with_progress(url, download_path)
        print(f"{software_name} downloaded successfully.")

        # Install software
        print(f"{progress_message}...")

        # Extract ZIP archives to the installation directory
        if download_path.endswith('.zip'):
            extract_dir = os.path.join(base_install_directory, os.path.basename(download_path).replace('.zip', ''))
            extract_zip(download_path, extract_dir)
        
        # Run installation command
        if not download_path.endswith('.zip'):
            if software_name == "Mozilla Firefox":
                install_command = f'{download_path} /InstallDirectoryPath={base_install_directory}\\Firefox'
            elif software_name == "7-Zip":
                install_command = f'{download_path} /S /D={base_install_directory}\\7-Zip'
            else:
                install_command = download_path
            subprocess.run(install_command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        print(f"{software_name} installed successfully.")

if __name__ == "__main__":
    set_console_title("Unauthorized Software Enabler by SoftwareRat")
    install_software()