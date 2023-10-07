import requests
import sys
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
from urllib.parse import urlparse
from colorama import Fore, Style, init
import shutil

# Set up logging
logging.basicConfig(filename='install_log.txt', level=logging.DEBUG)

# Colorama initialization
init(autoreset=True)
if sys.platform.lower() == 'win32':
    os.system('color')

def print_color(text, color=Fore.WHITE, style=Style.NORMAL, emoji='', end='\n'):
    emoji_str = f"{emoji} " if emoji else ''
    print(f"{style}{color}{emoji_str}{text}{Style.RESET_ALL}", end=end)

def print_ascii_art():
    ascii_art = """
██╗   ██╗███████╗███████╗
██║   ██║██╔════╝██╔════╝
██║   ██║███████╗█████╗  
██║   ██║╚════██║██╔══╝  
╚██████╔╝███████║███████╗
 ╚═════╝ ╚══════╝╚══════╝
"""
    print_color(ascii_art, Fore.RED, Style.BRIGHT)

def download_metadata(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        metadata = response.json()
        return metadata
    except requests.RequestException as e:
        logging.error(f"Error downloading metadata: {e}")
        input()
        return None

def set_console_title(title):
    if sys.platform.lower() == 'win32':
        ctypes.windll.kernel32.SetConsoleTitleW(title)

def download_file(url, destination):
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()
        total_size = int(response.headers.get('content-length', 0))
        block_size = 1024  # 1 Kibibyte
        downloaded_size = 0
        with open(destination, 'wb') as file:
            for data in response.iter_content(block_size):
                file.write(data)
                downloaded_size += len(data)
                progress = min(50, int(50 * downloaded_size / total_size))
                print_color(f"[{'=' * progress}{' ' * (50 - progress)}] {downloaded_size}/{total_size} bytes", Fore.GREEN, Style.BRIGHT, end='\r')
        print()  # Move to the next line after the progress bar
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
        set_reg_val(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", winreg.REG_DWORD, 0)
        set_reg_val(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme", winreg.REG_DWORD, 0)
        return True
    except Exception as e:
        logging.error(f"Error enabling dark mode: {e}")
        return False

def post_winxshell(install_path):
    try:
        # Remove all files including "zh-CN"
        for root, dirs, files in os.walk(install_path):
            for file in files:
                if 'zh-CN' in file:
                    os.remove(os.path.join(root, file))
        # Remove wallpaper.jpg
        wallpaper_path = os.path.join(install_path, 'wallpaper.jpg')
        if os.path.exists(wallpaper_path):
            os.remove(wallpaper_path)
        # Remove WinXShell_x86.exe
        x86_path = os.path.join(install_path, 'WinXShell_x86.exe')
        if os.path.exists(x86_path):
            os.remove(x86_path)
        # Rename WinXShell_x64.exe to WinXShell.exe
        x64_path = os.path.join(install_path, 'WinXShell_x64.exe')
        new_path = os.path.join(install_path, 'WinXShell.exe')
        os.rename(x64_path, new_path)
        # Make a copy of WinXShell.exe to explorer.exe
        explorer_path = os.path.join(install_path, 'explorer.exe')
        shutil.copy(new_path, explorer_path)
        # Terminate any existing Explorer processes currently running
        subprocess.run(['taskkill', '/F', '/IM', 'explorer.exe'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        # Start explorer.exe located on %LOCALAPPDATA%\Programs\WinXShell
        subprocess.Popen([explorer_path])
    except Exception as e:
        logging.error(f"Error in post_winxshell: {e}")

def handle_user_settings(settings):
    if "WallpaperPath" in settings:
        wallpaper_path = settings["WallpaperPath"]
        if set_wallpaper(wallpaper_path):
            print_color("Wallpaper set successfully.", Fore.GREEN, Style.BRIGHT)
        else:
            print_color("ERROR: Error setting wallpaper.", Fore.RED, Style.BRIGHT)

    if "DarkMode" in settings:
        dark_mode_enabled = settings["DarkMode"]
        if dark_mode_enabled:
            if enable_dark_mode():
                print_color("Dark mode enabled successfully.", Fore.GREEN, Style.BRIGHT)
            else:
                print_color("ERROR: Error enabling dark mode.", Fore.RED, Style.BRIGHT)

def load_user_settings(metadata):
    config_file_path = os.path.join(os.getcwd(), "USE_config.json")
    if os.path.exists(config_file_path):
        try:
            with open(config_file_path, "r") as config_file:
                user_settings = json.load(config_file)

                # Update default settings with user-specific settings
                for default_setting in user_settings.get("DefaultSoftwareSettings", []):
                    software_name = default_setting["Name"]
                    matching_software = next((s for s in metadata if s["Name"] == software_name), None)
                    if matching_software:
                        matching_software["Enabled"] = default_setting.get("Enabled", True)
                        matching_software["CreateShortcut"] = default_setting.get("CreateShortcut", False)

                # Add custom software metadata
                metadata.extend(user_settings.get("CustomSoftwareMetadata", []))

                # Handle user-specific settings
                handle_user_settings(user_settings)

        except json.JSONDecodeError:
            print_color("ERROR: Error parsing config file. Using default settings.", Fore.RED, Style.BRIGHT)
            # Set default wallpaper path if not provided in USE_config.json
            for software in metadata:
                if "WallpaperPath" not in software:
                    software["WallpaperPath"] = "C:\\Windows\\Web\\Wallpaper\\Windows\\img0.jpg"
    else:
        print_color("WARNING: Config file not found. Using default settings.", Fore.YELLOW, Style.BRIGHT)
        # Set default wallpaper path if not provided in USE_config.json
        for software in metadata:
            if "WallpaperPath" not in software:
                software["WallpaperPath"] = "C:\\Windows\\Web\\Wallpaper\\Windows\\img0.jpg"

def replace_placeholders(arguments, localappdata):
    return [arg.replace('{{LOCALAPPDATA}}', localappdata) for arg in arguments]

def main():
    # Set console title
    set_console_title("Unauthorized Software Enabler")

    # Display ASCII art
    print_ascii_art()
    metadata_url = "https://gfnhack.me/use_software_metadata.json"
    metadata = download_metadata(metadata_url)
    if metadata is None:
        print_color("ERROR: Error downloading metadata. Press any key to exit.", Fore.RED, Style.BRIGHT)
        input()
        return

    load_user_settings(metadata)

    for software in metadata:
        if software.get("Enabled", True):
            file_url = software["FileURL"]
            file_name = os.path.basename(urlparse(file_url).path)
            temp_path = os.path.join(os.environ["TEMP"], file_name)
            install_path = os.path.join(os.environ["LOCALAPPDATA"], "Programs", software["Name"])

            # Displaying "Installing" with the software name
            print_color(f"Installing {software['Name']}", Fore.CYAN, Style.BRIGHT)

            if download_file(file_url, temp_path):
                if file_name.endswith(".exe"):
                    custom_arguments = replace_placeholders(software.get("Arguments", []), os.environ["LOCALAPPDATA"])
                    if install_exe(temp_path, custom_arguments):
                        print_color(f"{software['Name']} installed successfully.", Fore.GREEN, Style.BRIGHT, '✅')
                    else:
                        print_color(f"ERROR: Error installing {software['Name']}.", Fore.RED, Style.BRIGHT, '❌')
                elif file_name.endswith(".zip"):
                    if extract_zip(temp_path, install_path):
                        print_color(f"{software['Name']} installed successfully.", Fore.GREEN, Style.BRIGHT)
                    else:
                        print_color(f"ERROR: Error installing {software['Name']}.", Fore.RED, Style.BRIGHT)

                if software.get("CreateShortcut", False):
                    executable_path = find_executable(install_path, software.get("Executable", ""))
                    if executable_path:
                        shortcut_name = f"{software['Name']}.lnk"
                        if create_shortcut(executable_path, os.path.join(os.path.expanduser('~'), 'Desktop', shortcut_name)):
                            print_color(f"Shortcut created successfully for {software['Name']}.", Fore.GREEN, Style.BRIGHT)
                        else:
                            print_color(f"ERROR: Error creating shortcut for {software['Name']}.", Fore.RED, Style.BRIGHT)

    # Post-installation steps for WinXShell
    post_winxshell(os.path.join(os.environ["LOCALAPPDATA"], "Programs", "WinXShell"))

if __name__ == "__main__":
    main()