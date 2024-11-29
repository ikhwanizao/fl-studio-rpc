import os
import sys
import time
import psutil
import win32gui
import win32process
import pypresence
import pystray
from PIL import Image
import threading
import winreg
import json
import ctypes

class FLStudioRPC:
    def __init__(self):
        self.CLIENT_ID = self.get_client_id()
        self.rpc = None
        self.start_time = None
        self.last_window_title = None
        self.current_view = "composing"
        self.running = True
        self.icon = None
        self.settings = self.load_settings()
        self.fl_studio_running = False
        
        if self.settings.get('start_with_windows', True):
            self.add_to_startup()
        
    def get_client_id(self):
        try:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            with open(os.path.join(base_path, 'CLIENT_ID.txt'), 'r') as f:
                return f.read().strip()
        except Exception as e:
            return os.environ.get("DISCORD_CLIENT_ID")
    
    def get_settings_path(self):
        """Get the path for settings file"""
        app_data_dir = os.path.join(os.getenv('APPDATA'), 'FLStudioDiscordRPC')
        if not os.path.exists(app_data_dir):
            os.makedirs(app_data_dir)
        return os.path.join(app_data_dir, 'settings.json')
    
    def load_settings(self):
        """Load settings from file or create with defaults"""
        try:
            with open(self.get_settings_path(), 'r') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            default_settings = {
                'start_with_windows': True
            }
            self.save_settings(default_settings)
            return default_settings
    
    def save_settings(self, settings):
        """Save settings to file"""
        try:
            with open(self.get_settings_path(), 'w') as f:
                json.dump(settings, f)
        except Exception as e:
            print(f"Failed to save settings: {e}")
    
    def add_to_startup(self):
        """Add program to Windows startup"""
        try:
            if getattr(sys, 'frozen', False):
                # For compiled executable
                exe_path = f'"{os.path.abspath(sys.executable)}"'
            else:
                # For Python script
                exe_path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'

            # Remove any double quotes that might cause issues
            exe_path = exe_path.replace('""', '"')

            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE
            )
            
            winreg.SetValueEx(key, "FL Studio Discord RPC", 0, winreg.REG_SZ, exe_path)
            winreg.CloseKey(key)
            print(f"Added to startup with path: {exe_path}")
        except Exception as e:
            print(f"Failed to add to startup: {e}")
        
    def remove_from_startup(self):
        """Remove program from Windows startup"""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE
            )
            
            winreg.DeleteValue(key, "FL Studio Discord RPC")
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Failed to remove from startup: {e}")
    
    def toggle_startup(self, icon, item):
        """Toggle start with Windows setting"""
        current_setting = self.settings.get('start_with_windows', True)
        new_setting = not current_setting
        
        if new_setting:
            self.add_to_startup()
        else:
            self.remove_from_startup()
            
        self.settings['start_with_windows'] = new_setting
        self.save_settings(self.settings)
        item.checked = new_setting

    def create_icon(self):
        try:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
                
            icon_path = os.path.join(base_path, 'icon.ico')
            return Image.open(icon_path)
        except Exception:
            return Image.new('RGB', (64, 64), color='black')

    def show_confirmation(self, message, title):
        """Show Windows message box with Yes/No"""
        return ctypes.windll.user32.MessageBoxW(None, message, title, 0x4) == 6 

    def show_message(self, message, title):
        """Show Windows message box"""
        ctypes.windll.user32.MessageBoxW(None, message, title, 0x40)

    def uninstall(self, icon=None, item=None):
        """Completely remove application data and exit"""
        if not self.show_confirmation(
            "Are you sure you want to uninstall FL Studio Discord RPC?\nThis will remove all settings and data.",
            "Confirm Uninstall"
        ):
            return

        success = True
        try:
            # Remove from startup
            try:
                self.remove_from_startup()
            except Exception as e:
                print(f"Failed to remove from startup: {e}")
                success = False

            # Delete settings file
            settings_path = self.get_settings_path()
            try:
                if os.path.exists(settings_path):
                    os.remove(settings_path)
            except Exception as e:
                print(f"Failed to remove settings file: {e}")
                success = False

            # Remove app directory
            try:
                app_dir = os.path.dirname(settings_path)
                if os.path.exists(app_dir):
                    os.rmdir(app_dir)
            except Exception as e:
                print(f"Failed to remove app directory: {e}")
                success = False

            if success:
                self.show_message(
                    "FL Studio Discord RPC has been uninstalled successfully.",
                    "Uninstall Complete"
                )
            else:
                self.show_message(
                    "Some errors occurred during uninstallation.\nCheck the console for details.",
                    "Uninstall Warning"
                )
            self.stop()
        except Exception as e:
            print(f"Critical error during uninstallation: {e}")
            self.show_message(
                "Failed to uninstall. Please check the console for details.",
                "Uninstall Error"
            )
            self.stop()

    def setup_tray(self):
        self.icon = pystray.Icon(
            "FL Studio RPC",
            self.create_icon(),
            "FL Studio Discord RPC",
            menu=pystray.Menu(
                pystray.MenuItem(
                    "Start with Windows",
                    self.toggle_startup,
                    checked=lambda item: self.settings.get('start_with_windows', True)
                ),
                pystray.MenuItem("Uninstall", self.uninstall),
                pystray.MenuItem("Exit", self.stop)
            )
        )
        self.icon.run()

    def stop(self, icon=None, item=None):
        """Ensure clean exit"""
        try:
            self.running = False
            if self.rpc:
                try:
                    self.rpc.close()
                except:
                    pass
            if self.icon:
                try:
                    self.icon.stop()
                except:
                    pass
            # Force exit to ensure cleanup
            os._exit(0)
        except Exception as e:
            print(f"Error during shutdown: {e}")
            os._exit(1)

    def connect(self):
        try:
            self.rpc = pypresence.Presence(self.CLIENT_ID)
            self.rpc.connect()
            print("Connected to Discord RPC")
        except Exception as e:
            print(f"Failed to connect to Discord RPC: {e}")
            return False
        return True

    def get_current_view(self, child_windows):
        """Determine current view based on child windows"""
        for _, title, class_name in child_windows:
            # Piano Roll detection
            if "Piano roll" in str(title):
                return "piano_roll"
            # Mixer detection
            elif "Mixer" in str(title):
                return "mixer"
            # Channel Rack detection
            elif "Channel rack" in str(title):
                return "pattern"
            # Playlist detection (but we'll treat it as composing)
            elif "Playlist" in str(title):
                return "composing"
        return "composing"

    def enum_child_windows(self, hwnd):
        """Enumerate child windows"""
        def callback(child_hwnd, child_windows):
            title = win32gui.GetWindowText(child_hwnd)
            class_name = win32gui.GetClassName(child_hwnd)
            child_windows.append((child_hwnd, title, class_name))
            return True
        
        child_windows = []
        try:
            win32gui.EnumChildWindows(hwnd, callback, child_windows)
        except Exception:
            pass
        return child_windows

    def get_fl_studio_window(self):
        """Find the FL Studio window and its child windows"""
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    process = psutil.Process(pid)
                    if process.name().lower() in ["fl64.exe", "fl.exe"]:
                        title = win32gui.GetWindowText(hwnd)
                        class_name = win32gui.GetClassName(hwnd)
                        child_windows = self.enum_child_windows(hwnd)
                        windows.append((hwnd, title, class_name, child_windows))
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            return True

        windows = []
        win32gui.EnumWindows(callback, windows)
        
        main_window = next((w for w in windows if "FL Studio" in w[1]), None)
        if main_window:
            self.current_view = self.get_current_view(main_window[3])
            return main_window[1]
        return None

    def parse_window_title(self, title):
        if not title or "FL Studio" not in title:
            return None

        project_name = "Untitled"
        if " - " in title:
            parts = title.split(" - ")
            project_name = parts[0]

        state_messages = {
            "piano_roll": "Writing melodies",
            "mixer": "Mixing tracks",
            "pattern": "Creating patterns",
            "composing": "Making music"
        }

        state = {
            "details": f"Project: {project_name}",
            "state": state_messages.get(self.current_view, "Making music"),
            "large_image": "fl_studio_logo",
            "large_text": f"FL Studio - {project_name}"
        }
        return state

    def update_presence(self):
        window_title = self.get_fl_studio_window()
        
        if not window_title:
            if self.rpc:
                self.rpc.clear()
            self.last_window_title = None
            if self.fl_studio_running:
                self.fl_studio_running = False
                self.start_time = None
            return
        
        if not self.fl_studio_running:
            self.fl_studio_running = True
            self.start_time = int(time.time())
        
        if window_title == self.last_window_title and self.current_view == getattr(self, '_last_view', None):
            return

        state = self.parse_window_title(window_title)
        if state:
            try:
                self.rpc.update(
                    start=self.start_time,
                    **state
                )
                self.last_window_title = window_title
                self._last_view = self.current_view
                print(f"Updated presence: {state['details']} - {state['state']}")
            except Exception as e:
                print(f"Failed to update presence: {e}")

    def update_presence_loop(self):
        while self.running:
            self.update_presence()
            time.sleep(15)

    def run(self):
        if not self.connect():
            return

        print("FL Studio Discord RPC running. Right-click tray icon to exit.")
        
        presence_thread = threading.Thread(target=self.update_presence_loop)
        presence_thread.daemon = True
        presence_thread.start()
        
        self.setup_tray()

if __name__ == "__main__":
    fl_rpc = FLStudioRPC()
    fl_rpc.run()