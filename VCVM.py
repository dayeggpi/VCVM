import time
import threading
import sys
import os
import subprocess
import io
import winreg
import configparser
import logging
from datetime import datetime
from pystray import Icon, MenuItem as item
from PIL import Image, ImageDraw
import ctypes
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL, wintypes

class LoggerMaster:
    def __init__(self):
        self.logger = None
        self.logging_enabled = False        
        self.log_file = "VCVM.log" 
        sys.excepthook = self.handle_exception 
        
    def setup_logging(self):
        """Setup logging configuration"""
        if self.logger:  
            for handler in self.logger.handlers[:]:
                self.logger.removeHandler(handler)
                handler.close() 

        if self.logging_enabled:
            self.logger = logging.getLogger('VolumeControl for Voicemeeter') 
            self.logger.setLevel(logging.DEBUG)
            
            try:
                file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
                file_handler.setLevel(logging.DEBUG)
                formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)
            except IOError as e:
                print(f"ERROR: Failed to open log file {self.log_file}: {e}")
            
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO) 
            formatter_console = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter_console)
            self.logger.addHandler(console_handler)

            self.logger.propagate = False
        else:
            self.logger = None 
      
    def log(self, message, level='info', exc_info=None):
        """Log a message"""
        if self.logger:
            if level.lower() == 'debug':
                self.logger.debug(message, exc_info=exc_info)
            elif level.lower() == 'warning':
                self.logger.warning(message, exc_info=exc_info)
            elif level.lower() == 'error':
                self.logger.error(message, exc_info=exc_info)
            else:
                self.logger.info(message, exc_info=exc_info)
        else:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {level.upper()} - {message}")
            if exc_info:
                import traceback
                traceback.print_exception(exc_info[0], exc_info[1], exc_info[2], file=sys.stderr)
            
    def handle_exception(self, exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        self.log("Uncaught exception", level='error', exc_info=(exc_type, exc_value, exc_traceback))
   
        
class VoicemeeterVolumeSync:
    def __init__(self):
        self.running = False
        self.connected = False
        self.vol_interface = None
        self.last_windows_vol = 0
        self.last_vm_gain = 0
        self.last_change_time = 0
        self.sync_thread = None
        self.monitor_thread = None
        self.icon = None
        self.voicemeeter = None
        self.base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        self.config_file = self.get_data_path("config.ini")
        self.log_file = "VCVM.log"
        self.config = configparser.ConfigParser()        
        self.load_config()
        
        self.autostart_enabled = self.is_autostart_enabled()
        self.load_tray_icon()
        self.last_change_source = None  # Can be 'windows' or 'voicemeeter'        
        # Check if we're starting up with the system
        self.is_startup_launch = self.detect_startup_launch()
        
        logclass.log("Application initialized")

    def detect_startup_launch(self):
        """Detect if we're being launched at system startup"""
        # Check if we're within the first few minutes after boot
        try:
            import psutil
            boot_time = psutil.boot_time()
            current_time = time.time()
            time_since_boot = current_time - boot_time
            
            # If less than 5 minutes since boot, likely a startup launch
            if time_since_boot < 300:  # 5 minutes
                logclass.log(f"Detected startup launch - {time_since_boot:.1f}s since boot")
                return True
        except ImportError:
            # Fallback: check if certain processes are still starting
            pass
        
        # Alternative method: check system uptime via Windows API
        try:
            uptime_ms = ctypes.windll.kernel32.GetTickCount64()
            uptime_seconds = uptime_ms / 1000
            if uptime_seconds < 300:  # 5 minutes
                logclass.log(f"Detected startup launch - {uptime_seconds:.1f}s uptime")
                return True
        except:
            pass
        
        return False

    def initialize_com(self):
        """Initialize COM for audio operations"""
        try:
            import comtypes
            comtypes.CoInitialize()
            logclass.log("COM initialized successfully")
            return True
        except Exception as e:
            logclass.log(f"Failed to initialize COM: {e}", 'error')
            return False
        
    def wait_for_system_ready(self):
        """Wait for system to be ready before connecting"""
        if not self.is_startup_launch:
            return
        
        startup_delay = self.config.getint('Startup', 'delay_seconds', fallback=5)
        logclass.log(f"Startup launch detected - waiting {startup_delay}s for system to be ready...")
        time.sleep(startup_delay)
        
        # Initialize COM first
        if not self.initialize_com():
            logclass.log("Failed to initialize COM - audio operations may fail", 'warning')
        
        # Additional checks to ensure audio system is ready
        max_wait = 120  # Maximum 2 minutes additional wait
        wait_interval = 5
        waited = 0
        
        while waited < max_wait:
            try:
                # Try to initialize audio interface as a readiness test
                devices = AudioUtilities.GetSpeakers()
                if devices:
                    self.update_tray_icon("icon_status_on.ico")
                    logclass.log("Audio system appears ready")
                    break
            except Exception as e:
                logclass.log(f"Audio system not ready yet: {e}")
            
            logclass.log(f"Waiting for audio system... ({waited}s/{max_wait}s)")
            time.sleep(wait_interval)
            waited += wait_interval
        
        if waited >= max_wait:
            logclass.log("Maximum wait time reached - proceeding anyway", 'warning')

    def get_resource_path(self, relative_path):
        """Get absolute path to resource, works for dev and for PyInstaller"""
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def get_data_path(self, filename):
        """Get path for data files that should be in the same directory as the exe"""
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
        else:
            exe_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(exe_dir, filename)
    
    def load_config(self):
        """Load configuration from config.ini"""
        default_config = {
            'Voicemeeter': {
                'dll_path': r'c:\Program Files (x86)\VB\Voicemeeter\VoicemeeterRemote64.dll'
            },
            'Logging': {
                'enabled': 'true',
                'verbose': 'false',
                'log_file': 'VCVM.log'
            },
            'Settings': {
                'curve_power': '0.55',
                'sync_interval': '0.3',
                'change_timeout': '4',
                'gain_threshold': '3.0',
                'volume_threshold': '1',
                'bus': '0'
            },
            'Startup': {
                'delay_seconds': '5',
                'max_retry_attempts': '5',
                'retry_interval': '2'
            }
        }

        if not os.path.exists(self.config_file):
            self.config.read_dict(default_config)
            try:
                with open(self.config_file, 'w') as f:
                    self.config.write(f)
                print(f"Created default config file: {self.config_file}")
                logclass.log(f"Created config file at: {self.config_file}")
            except Exception as e:
                print(f"Error creating config file: {e}")
                logclass.log(f"Error creating config file: {e}", 'error')
        else:
            try:
                self.config.read(self.config_file)
                logclass.log(f"Loaded config from: {self.config_file}")
            except Exception as e:
                logclass.log(f"Error reading config file: {e}", 'error')
                self.config.read_dict(default_config)  
            
            config_changed = False
            for section, options in default_config.items():
                if not self.config.has_section(section):
                    self.config.add_section(section)
                    config_changed = True
                for key, value in options.items():
                    if not self.config.has_option(section, key):
                        self.config.set(section, key, value)
                        config_changed = True
            
            if config_changed:
                try:
                    with open(self.config_file, 'w') as f:
                        self.config.write(f)
                    logclass.log("Updated config file with missing entries")
                except Exception as e:
                    logclass.log(f"Error updating config file: {e}", 'error')

        logclass.logging_enabled = self.config.getboolean('Logging', 'enabled', fallback=True)
        log_filename = self.config.get('Logging', 'log_file', fallback="VCVM.log")
        logclass.log_file = self.get_data_path(log_filename)
        logclass.setup_logging()

        self.logging_verbose = self.config.getboolean('Logging', 'verbose', fallback=False)

    def save_config(self):
        """Save current configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                self.config.write(f)
        except Exception as e:
            print(f"Error saving config: {e}")

    def load_voicemeeter_dll(self):
        """Load the Voicemeeter DLL with retry"""
        dll_path = self.config.get('Voicemeeter', 'dll_path')
        max_attempts = 5
        for attempt in range(max_attempts):
            try:
                if os.path.exists(dll_path):
                    self.voicemeeter = ctypes.WinDLL(dll_path)
                    logclass.log(f"Loaded Voicemeeter DLL from: {dll_path} (attempt {attempt+1})")
                    return
                else:
                    logclass.log(f"Voicemeeter DLL not found at: {dll_path}", 'error')
                    break
            except Exception as e:
                logclass.log(f"Failed to load Voicemeeter DLL (attempt {attempt+1}): {e}", 'error')
                time.sleep(2)

        self.voicemeeter = None
        logclass.log("Giving up on loading Voicemeeter DLL after retries", 'error')

    def load_tray_icon(self):
        """Load the tray icon image"""
        try:
            icon_path = self.get_resource_path("icon.ico")
            if os.path.exists(icon_path):
                self.tray_icon_image = Image.open(icon_path)
                logclass.log(f"Loaded tray icon from bundled resource: {icon_path}")
            else:
                icon_path = self.get_data_path("icon.ico")
                if os.path.exists(icon_path):
                    self.tray_icon_image = Image.open(icon_path)
                    logclass.log(f"Loaded tray icon from data path: {icon_path}")
                else:
                    raise FileNotFoundError("Icon not found in either location")
        except Exception as e:
            logclass.log(f"Could not load icon file ({e}), generating fallback", 'warning')
            self.tray_icon_image = Image.new('RGBA', (64, 64), color=(0, 0, 0, 0))
            d = ImageDraw.Draw(self.tray_icon_image)
            d.ellipse((0, 0, 63, 63), fill=(10, 50, 120, 255)) 
            d.rectangle((18, 24, 28, 40), fill=(255, 255, 255, 255)) 
            d.polygon([(28, 24), (38, 20), (38, 44), (28, 40)], fill=(255, 255, 255, 255))
            d.arc((40, 22, 54, 42), start=300, end=60, fill=(255, 255, 255, 180), width=2)
            d.arc((44, 18, 60, 46), start=300, end=60, fill=(255, 255, 255, 100), width=2)
            logclass.log("Generated fallback tray icon (speaker symbol)")

    def connect_voicemeeter(self):
        """Connect to Voicemeeter API with enhanced retry logic"""
        self.load_voicemeeter_dll()
        if not self.voicemeeter:
            logclass.log("Voicemeeter DLL not loaded", 'error')
            return False

        max_attempts = self.config.getint('Startup', 'max_retry_attempts', fallback=5)
        retry_interval = self.config.getint('Startup', 'retry_interval', fallback=2)
        
        for attempt in range(max_attempts):
            try:
                res = self.voicemeeter.VBVMR_Login()
                if res == 0:
                    logclass.log(f"Connected to Voicemeeter on attempt {attempt+1}")
                    return True
                else:
                    error_msg = self.get_voicemeeter_error_message(res)
                    logclass.log(f"VBVMR_Login failed (code: {res} - {error_msg}), attempt {attempt+1}/{max_attempts}")
                    
                    # For startup launches, wait longer between attempts
                    wait_time = retry_interval if self.is_startup_launch else 2
                    if attempt < max_attempts - 1:  # Don't wait after the last attempt
                        logclass.log(f"Waiting {wait_time}s before retry...")
                        time.sleep(wait_time)
                        
            except Exception as e:
                logclass.log(f"Exception in VBVMR_Login: {e}", 'error')
                if attempt < max_attempts - 1:
                    wait_time = retry_interval if self.is_startup_launch else 2
                    time.sleep(wait_time)

        logclass.log(f"Failed to connect to Voicemeeter after {max_attempts} attempts", 'error')
        self.update_tray_icon("icon_status_off.ico")        
        return False

    def get_voicemeeter_error_message(self, error_code):
        """Get human-readable error message for Voicemeeter error codes"""
        error_messages = {
            -1: "Voicemeeter not running or not installed",
            -2: "Voicemeeter DLL not found or incompatible version",
            -3: "Parameter error",
            -4: "Structure mismatch",
            -5: "Connection lost",
            -6: "System error",
            -7: "Unknown error",
            1: "Voicemeeter not running"
        }
        return error_messages.get(error_code, f"Unknown error code: {error_code}")

    def disconnect_voicemeeter(self):
        """Disconnect from Voicemeeter API"""
        if not self.voicemeeter:
            return
        try:
            self.voicemeeter.VBVMR_Logout()
            logclass.log("Disconnected from Voicemeeter")
        except Exception as e:
            logclass.log(f"Error disconnecting from Voicemeeter: {e}", 'error')

    def set_bus_gain(self, bus_index, gain_db):
        """Set gain for a specific bus"""
        if not self.voicemeeter:
            return
        try:
            param_name = f"Bus[{bus_index}].Gain".encode("utf-8")
            self.voicemeeter.VBVMR_SetParameterFloat(ctypes.c_char_p(param_name), ctypes.c_float(gain_db))
        except Exception as e:
            logclass.log(f"Error setting bus {bus_index} gain: {e}", 'error')

    def get_bus_gain(self, bus_index):
        """Get gain for a specific bus"""
        if not self.voicemeeter:
            return None
        try:
            param_name = f"Bus[{bus_index}].Gain".encode("utf-8")
            gain = ctypes.c_float()
            result = self.voicemeeter.VBVMR_GetParameterFloat(ctypes.c_char_p(param_name), ctypes.byref(gain))
            return gain.value if result == 0 else None
        except Exception as e:
            logclass.log(f"Error getting bus {bus_index} gain: {e}", 'error')
            return None

    def map_volume_to_gain(self, volume):
        """Convert Windows volume percentage to Voicemeeter gain in dB"""
        if volume <= 0:
            return -60
        elif volume >= 100:
            return 12
        
        curve_power = self.config.getfloat('Settings', 'curve_power')
        gain = (volume / 100) ** curve_power * 72 - 60
        return round(gain, 2)

    def map_gain_to_volume(self, gain):
        """Convert Voicemeeter gain to Windows volume percentage"""
        if gain <= -60:
            return 0
        elif gain >= 12:
            return 100
        
        curve_power = self.config.getfloat('Settings', 'curve_power')
        volume = ((gain + 60) / 72) ** (1 / curve_power) * 100
        return int(volume)

    def init_windows_volume_interface(self):
        """Initialize Windows volume control interface with retry"""
        # Ensure COM is initialized
        try:
            import comtypes
            comtypes.CoInitialize()
        except:
            pass  # May already be initialized
        
        max_attempts = 5
        for attempt in range(max_attempts):
            try:
                devices = AudioUtilities.GetSpeakers()
                interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                logclass.log(f"Initialized Windows volume interface (attempt {attempt+1})")
                return ctypes.cast(interface, ctypes.POINTER(IAudioEndpointVolume))
            except Exception as e:
                logclass.log(f"Failed to initialize Windows volume interface (attempt {attempt+1}): {e}", 'error')
                if attempt < max_attempts - 1:
                    time.sleep(2)
        
        logclass.log("Failed to initialize Windows volume interface after retries", 'error')
        return None

    def get_windows_volume(self):
        """Get current Windows volume"""
        try:
            if self.vol_interface:
                return int(self.vol_interface.GetMasterVolumeLevelScalar() * 100)
        except Exception as e:
            logclass.log(f"Error getting Windows volume: {e}", 'error')
        return 0

    def set_windows_volume(self, vol_percent):
        """Set Windows volume"""
        try:
            if self.vol_interface:
                scalar = vol_percent / 100
                self.vol_interface.SetMasterVolumeLevelScalar(scalar, None)
        except Exception as e:
            logclass.log(f"Error setting Windows volume: {e}", 'error')

    def is_voicemeeter_ok(self):
        """Check if Voicemeeter is responding"""
        try:
            return self.get_bus_gain(0) is not None
        except:
            return False

    def sync_volumes(self):
        """Main volume synchronization loop with startup handling"""
        logclass.log("Starting volume sync...")
        
        self.wait_for_system_ready()
        
        if not self.connect_voicemeeter():
            logclass.log("Failed to connect to Voicemeeter - sync will not start", 'error')
            return
            
        self.vol_interface = self.init_windows_volume_interface()
        if not self.vol_interface:
            logclass.log("Failed to initialize Windows volume interface - sync will not start", 'error')
            self.disconnect_voicemeeter()
            return

        self.last_windows_vol = self.get_windows_volume()
        self.last_vm_gain = self.get_bus_gain(0) or 0
        self.last_change_time = time.time()

        logclass.log("Volume sync active. Monitoring...")

        sync_interval = self.config.getfloat('Settings', 'sync_interval')
        change_timeout = self.config.getfloat('Settings', 'change_timeout')
        gain_threshold = self.config.getfloat('Settings', 'gain_threshold')
        volume_threshold = self.config.getint('Settings', 'volume_threshold')

        while self.running:
            try:
                current_windows_vol = self.get_windows_volume()
                current_vm_gain = self.get_bus_gain(0) or 0
                time_now = time.time()

                if abs(current_windows_vol - self.last_windows_vol) >= volume_threshold:
                    gain = self.map_volume_to_gain(current_windows_vol)
                    bus_list_str = self.config.get('Settings', 'bus', fallback='0')
                    try:
                        bus_list = [int(x.strip()) for x in bus_list_str.split(',')]
                    except ValueError as e:
                        logclass.log(f"Invalid bus configuration: '{bus_list_str}' - {e}", 'error')
                        bus_list = [0]  # fallback

                    for bus in bus_list:
                        try:
                            self.set_bus_gain(bus, gain)
                            self.last_vm_gain = gain
                        except Exception as e:
                            logclass.log(f"Failed to set gain for bus {bus}: {e}", 'error')
                    
                    if self.logging_verbose:
                        logclass.log(f"Windows volume changed: {current_windows_vol}% → {gain}dB", 'debug')
                    self.last_windows_vol = current_windows_vol
                    self.last_vm_gain = gain
                    self.last_change_time = time_now
                    self.last_change_source = 'windows'                    

                elif time_now - self.last_change_time > change_timeout:
                    if abs(current_vm_gain - self.last_vm_gain) >= gain_threshold:        
                        if self.last_change_source == 'windows':
                            continue  # Prevent bounce back        
                            
                        target_volume = self.map_gain_to_volume(current_vm_gain)
                        gain_diff = current_vm_gain - self.last_vm_gain
                        if self.logging_verbose:
                            logclass.log(f"Voicemeeter gain changed: {self.last_vm_gain}dB → {current_vm_gain}dB (Δ{gain_diff:+.1f}dB) | Target Windows vol: {target_volume}%", 'debug')
                        
                        if abs(target_volume - current_windows_vol) > 10:
                            step = int((target_volume - current_windows_vol) * 0.3)
                            new_volume = current_windows_vol + step
                            self.set_windows_volume(new_volume)
                            self.last_windows_vol = new_volume
                            if self.logging_verbose:
                                logclass.log(f"Applied smooth Windows volume adjustment: {current_windows_vol}% → {new_volume}% (step: {step})", 'debug')
                        else:
                            self.set_windows_volume(target_volume)
                            self.last_windows_vol = target_volume
                            if self.logging_verbose:
                                logclass.log(f"Applied direct Windows volume adjustment: {current_windows_vol}% → {target_volume}%", 'debug')
                        
                        self.last_vm_gain = current_vm_gain
                        self.last_change_time = time_now
                        self.last_change_source = 'voicemeeter'
                    # Don't sync if we were the one who just set the gain
                    if abs(current_vm_gain - self.last_vm_gain) < gain_threshold:
                        continue
                time.sleep(sync_interval)

            except Exception as e:
                logclass.log(f"Error in sync loop: {e}", 'error')
                time.sleep(1)

        logclass.log("Volume sync stopped")
        self.disconnect_voicemeeter()

    def monitor_voicemeeter_status(self):
        """Monitor Voicemeeter connection status"""
        last_status = None
        while self.running:
            try:
                self.connected = self.is_voicemeeter_ok()
                if self.connected != last_status:
                    status_text = 'Connected' if self.connected else 'Disconnected'
                    if self.logging_verbose:
                        logclass.log(f"Voicemeeter status changed: {status_text}")
                    last_status = self.connected

                    if self.connected:
                        self.update_tray_icon("icon_status_on.ico")
                    else:
                        self.update_tray_icon("icon_status_off.ico")

            except Exception as e:
                logclass.log(f"Error monitoring Voicemeeter status: {e}", 'error')
                self.connected = False
            time.sleep(1)
        
    def start_sync(self):
        """Start the volume synchronization"""
        if self.running:
            return
            
        self.running = True
        self.sync_thread = threading.Thread(target=self.sync_volumes, daemon=True)
        self.monitor_thread = threading.Thread(target=self.monitor_voicemeeter_status, daemon=True)
        
        self.sync_thread.start()
        self.monitor_thread.start()
        logclass.log("Volume sync started")

    def stop_sync(self):
        """Stop the volume synchronization"""
        if not self.running:
            return
            
        logclass.log("Stopping volume sync...")
        self.running = False
        
        if self.sync_thread and self.sync_thread.is_alive():
            self.sync_thread.join(timeout=2)
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=2)
            
        logclass.log("Volume sync stopped")

    def delete_startup_task(task_name="VolumeControl for Voicemeeter"):
        try:
            result = subprocess.run(
                ["schtasks", "/Delete", "/TN", task_name, "/F"],
                capture_output=True, text=True
            )
            if result.returncode == 0:
                print(f"Scheduled task '{task_name}' deleted.")
                return True
            else:
                print(f"Error deleting task: {result.stderr.strip()}")
                return False
        except Exception as e:
            print(f"Exception deleting task: {e}")
            return False

    def toggle_autostart(self, enable):
        """Toggle autostart using Windows Task Scheduler with delay"""
        task_name = "VolumeControl for Voicemeeter"
        if enable:
            is_py = not getattr(sys, 'frozen', False)
            python_exe = sys.executable
            script_path = os.path.abspath(__file__)
            cmd = f'"{python_exe}" "{script_path}"' if is_py else f'"{sys.executable}"' #keep as such to AVOID having quotes around python path, otherwise will fail

            schtasks_cmd = [
                "schtasks",
                "/Create",
                "/TN", task_name,
                "/TR", cmd,
                "/SC", "ONLOGON",
                "/RL", "HIGHEST",
                "/DELAY", "0000:05",
                "/F"
            ]
            try:
                result = subprocess.run(schtasks_cmd, capture_output=True, text=True)
                if result.returncode == 0:
                    logclass.log("Autostart task created successfully with 5 seconds delay")
                else:
                    logclass.log(f"Failed to create autostart task: {result.stderr.strip()}", 'error')
            except Exception as e:
                logclass.log(f"Error creating autostart task: {e}", 'error')
        else:
            try:
                result = subprocess.run(
                    ["schtasks", "/Delete", "/TN", task_name, "/F"],
                    capture_output=True, text=True
                )
                if result.returncode == 0:
                    logclass.log("Autostart task removed successfully")
                else:
                    logclass.log(f"Failed to remove autostart task: {result.stderr.strip()}", 'error')
            except Exception as e:
                logclass.log(f"Error removing autostart task: {e}", 'error')

    def is_autostart_enabled(self):
        result = subprocess.run(
            ["schtasks", "/Query", "/TN", "VolumeControl for Voicemeeter"],
            capture_output=True, text=True
        )
        return result.returncode == 0

    def on_quit(self, icon, _):
        """Handle quit from tray menu"""
        logclass.log("Quit requested from tray")
        self.stop_sync()
        if self.icon:
            self.icon.stop()

    def on_reload(self, icon, _):
        """Handle reload from tray menu"""
        logclass.log("Reload requested from tray")
        self.stop_sync()
        time.sleep(1)
        self.load_config()
        self.load_voicemeeter_dll()
        if self.voicemeeter:
            self.start_sync()
        else:
            logclass.log("Voicemeeter DLL not loaded — skipping sync", 'error')

    def on_toggle_autostart(self, icon, _):
        """Handle autostart toggle from tray menu"""
        self.autostart_enabled = not self.autostart_enabled
        self.toggle_autostart(self.autostart_enabled)

    def on_toggle_logging(self, icon, _):
        """Handle logging toggle from tray menu"""
        logclass.logging_enabled = not logclass.logging_enabled 
        self.config.set('Logging', 'enabled', str(logclass.logging_enabled).lower())
        self.save_config()
        logclass.setup_logging()
        logclass.log(f"Logging {'enabled' if logclass.logging_enabled else 'disabled'}")
    
    def on_toggle_logging_verbose(self, icon, _):
        self.logging_verbose = not self.logging_verbose
        self.config.set('Logging', 'verbose', str(self.logging_verbose).lower())
        self.save_config()
        logclass.log(f"Verbose logging {'enabled' if self.logging_verbose else 'disabled'}")
    
    
    def get_autostart_text(self, icon):
        """Get autostart menu text"""
        state = self.is_autostart_enabled()
        return f"Autostart: {'On' if state else 'Off'}"
    
    def get_logging_text(self, icon):
        """Get logging menu text"""
        return f"Logging: {'On' if logclass.logging_enabled else 'Off'}" 
    
    def get_verbose_text(self, icon):
        """Get verbose menu text"""
        return f"Verbose: {'On' if self.logging_verbose else 'Off'}"
    
    def creditsinfo(self, icon=None, item=None):
        """Show About dialog using native Windows MessageBox"""
        try:
            logclass.log("Attempting to show About dialog using Windows MessageBox.", level='debug')
            MB_OK = 0x0
            MB_ICONINFORMATION = 0x40
            
            title = "About VolumeControl for Voicemeeter"
            message = ("VolumeControl for Voicemeeter\n"
                      "Version 1.0.0 of may 2025\n\n"
                      "https://github.com/hycday/VCVM"
                      "Synchronizes Windows volume with Voicemeeter\n"
                      "support them : https://vb-audio.com/\n\n"
                      "by hycday")
            
            result = ctypes.windll.user32.MessageBoxW(
                0,  # hWnd (0 = no parent window)
                message,
                title,
                MB_OK | MB_ICONINFORMATION
            )
            
            logclass.log("About dialog shown successfully using Windows MessageBox.", level='debug')
            
        except Exception as e:
            error_msg = f"Error showing About dialog: {e}"
            logclass.log(error_msg, level='error', exc_info=True)
            print(f"Error: {error_msg}")
            print("\n=== About VolumeControl for Voicemeeter ===")
            print("Version 1.0.0")  
            print("https://github.com/hycday/VCVM")  
            print("Synchronizes Windows volume with Voicemeeter")
            print("by hycday")
            print("=====================================\n")

    def start_tray(self):
        """Start the system tray icon"""
        self.icon = Icon(
            "Voicemeeter", 
            self.tray_icon_image, 
            "VolumeControl for Voicemeeter",
            menu=(
                item("About", self.creditsinfo),
                item(self.get_autostart_text, self.on_toggle_autostart),
                item(self.get_logging_text, self.on_toggle_logging),
                item(self.get_verbose_text, self.on_toggle_logging_verbose),
                item("Reload", self.on_reload),
                item("Quit", self.on_quit),
            )
        )
        logclass.log("Starting tray icon...")
        self.icon.run() 


    def resource_path(self, filename):
        if getattr(sys, 'frozen', False):  # running as .exe (PyInstaller)
            return os.path.join(sys._MEIPASS, filename)
        else:  # running as .py
            base_dir = os.path.dirname(os.path.abspath(__file__))
            return os.path.join(base_dir, filename)

    def update_tray_icon(self, icon_name):
        """Change the tray icon dynamically"""
        try:
            icon_path = self.resource_path(icon_name)
            if os.path.exists(icon_path):
                image = Image.open(icon_path)
                self.icon.icon = image
                logclass.log(f"Tray icon updated to: {icon_name}")
            else:
                logclass.log(f"Icon not found: {icon_path}", 'error')
        except Exception as e:
            logclass.log(f"Failed to update tray icon to {icon_name}: {e}", 'error')


    def run(self):
        """Main application entry point"""
        logclass.log("Starting VolumeControl for Voicemeeter application")
        self.start_sync()
        try:
            self.start_tray()
        except KeyboardInterrupt:
            logclass.log("KeyboardInterrupt received")
        except Exception as e:
            logclass.log(f"Tray error: {e}", 'error')
        finally:
            self.stop_sync()
            logclass.log("Application stopped")

if __name__ == "__main__":
    logclass = LoggerMaster()
    app = VoicemeeterVolumeSync()
    app.run()