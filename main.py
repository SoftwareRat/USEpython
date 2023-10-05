import requests
import os
import subprocess
import zipfile
import ctypes
import json
import winreg
import logging
import comtypes
import comtypes.shelllink
import comtypes.client
import comtypes.persist

# Set up logging
logging.basicConfig(filename='install_log.txt', level=logging.DEBUG)

def download_metadata(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an HTTPError for bad responses
        metadata = response.json()
        return metadata
    except requests.RequestException as e:
        logging.error(f"Error downloading metadata: {e}")
        return None

def download_file(url, destination):
    try:
        response = requests.get(url)
        response.raise_for_status()
        with open(destination, 'wb') as file:
            file.write(response.content)
        return True
    except requests.RequestException as e:
        logging.error(f"Error downloading file: {e}")
        return False

def install_exe(file_path, arguments):
    try:
        subprocess.run([file_path] + arguments, check=True)
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"Error installing software: {e}")
        return False

def extract_zip(file_path, destination):
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(destination)
        return True
    except zipfile.BadZipFile as e:
        logging.error(f"Error extracting zip file: {e}")
        return False

def create_shortcut(target, shortcut_path):
    try:
        shortcut = comtypes.client.CreateObject(comtypes.shelllink.ShellLink)
        shortcut_w = shortcut.QueryInterface(comtypes.shelllink.IShellLinkW)
        shortcut_w.SetPath(target)
        shortcut_file = shortcut.QueryInterface(comtypes.persist.IPersistFile)
        shortcut_file.Save(shortcut_path, True)
        return True
    except Exception as e:
        logging.error(f"Error creating shortcut: {e}")
        return False

def download_image(image_path):
    try:
        response = requests.get(image_path)
        response.raise_for_status()
        image_data = response.content
        image_path = os.path.join(os.getenv("USERPROFILE"), "Pictures", "wallpaper.jpg")
        with open(image_path, 'wb') as img_file:
            img_file.write(image_data)
        return image_path
    except requests.RequestException as e:
        logging.error(f"Error downloading image: {e}")
        return None

def find_executable(directory, exe_name) -> str:
    for root, dirs, files in os.walk(directory):
        if exe_name.lower() in (file.lower() for file in files if os.path.isfile(os.path.join(root, file))):
            return os.path.join(root, exe_name)
    return None

def set_wallpaper(image_path):
    try:
        if image_path.startswith(("http://", "https://")):
            image_path = download_image(image_path)
            if image_path is None:
                return False
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 2)
        return True
    except Exception as e:
        logging.error(f"Error setting wallpaper: {e}")
        return False

def set_reg_val(key: winreg.HKEYType, key_path: str, val: str, val_type: int, new_val):
    try:
        with winreg.OpenKey(key, key_path, 0, winreg.KEY_SET_VALUE) as sel_key:
            winreg.SetValueEx(sel_key, val, 0, val_type, new_val)
    except Exception as e:
        logging.error(f"Error setting the registry value: {e}")

def enable_dark_mode():
    try:
        set_reg_val(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", winreg.REG_DWORD.value, 0)
        set_reg_val(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme", winreg.REG_DWORD.value, 0)
        return True
    except Exception as e:
        logging.error(f"Error enabling dark mode: {e}")
        return False

def handle_user_settings(settings):
    if "WallpaperPath" in settings:
        wallpaper_path = settings["WallpaperPath"]
        if set_wallpaper(wallpaper_path):
            print("Wallpaper set successfully.")
        else:
            print("Error setting wallpaper.")

    if "DarkMode" in settings:
        dark_mode_enabled = settings["DarkMode"]
        if dark_mode_enabled:
            if enable_dark_mode():
                print("Dark mode enabled successfully.")
            else:
                print("Error enabling dark mode.")

    # Add more code here to handle other user settings

def load_user_settings(metadata):
    config_file_path = os.path.join(os.getcwd(), "USE_config.json")
    if os.path.exists(config_file_path):
        try:
            with open(config_file_path, "r") as config_file:
                user_settings = json.load(config_file)
                for software in metadata:
                    software_name = software["Name"]
                    if software_name in user_settings:
                        software["Enabled"] = user_settings[software_name].get("Enabled", True)
                        software["CreateShortcut"] = user_settings[software_name].get("CreateShortcut", False)
        except json.JSONDecodeError:
            print("Error loading user settings from config file.")
    else:
        print("Config file not found. Using default settings.")
        # Set default wallpaper path if not provided in USE_config.json
        for software in metadata:
            if "WallpaperPath" not in software:
                software["WallpaperPath"] = "C:\\Windows\\Web\\Wallpaper\\Windows\\img0.jpg"



def main():
    metadata_url = "http://gfnhack.me/use_software_metadata.json"
    metadata = download_metadata(metadata_url)
    if metadata is None:
        print("Error downloading software metadata. Please check your internet connection.")
        return

    load_user_settings(metadata)

    for software in metadata:
        if software.get("Enabled", True):
            file_url = software["FileURL"]
            file_name = os.path.basename(file_url)
            temp_path = os.path.join(os.environ["TEMP"], file_name)
            install_path = os.path.join(os.environ["LOCALAPPDATA"], "Programs", software["Name"])

            if download_file(file_url, temp_path):
                if file_name.endswith(".exe"):
                    if install_exe(temp_path, software.get("Arguments", [])):
                        print(f"Software {software['Name']} installed successfully.")
                    else:
                        print(f"Error installing software {software['Name']}.")
                elif file_name.endswith(".zip"):
                    if extract_zip(temp_path, install_path):
                        print(f"Software {software['Name']} installed successfully.")
                    else:
                        print(f"Error installing software {software['Name']}.")

                if software.get("CreateShortcut", False):
                    executable_path = find_executable(install_path, software.get("Executable", ""))
                    if executable_path:
                        shortcut_name = f"{software['Name']}.lnk"
                        if create_shortcut(executable_path, os.path.join(os.path.expanduser('~'), 'Desktop', shortcut_name)):
                            print(f"Shortcut created for software {software['Name']}.")
                        else:
                            print(f"Error creating shortcut for software {software['Name']}.")
                    else:
                        print(f"Executable not found for software {software['Name']}.")

if __name__ == "__main__":
    main()