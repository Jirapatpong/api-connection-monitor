import subprocess
import os
import sys
import socket
import threading
import time
from datetime import datetime
import schedule
import win32event
import win32api
from winerror import ERROR_ALREADY_EXISTS
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
from pystray import MenuItem as item, Icon
from PIL import Image
import base64
from io import BytesIO

# --- Base64 Encoded Icon (Self-Contained) ---
ICON_B64 = b'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABoSURBVHja7cEBAQAAAIIg/69uSEABAAAAAAAAAAAAAAB8GgIECBAgQIAAAQIECBAgQIAAAQIECBAgQIAAAQIECBAgQIAAAQIECBAgQIAAAQIECBAgQIAAAQIECBAgQIAAAQIECHA3AQSRAAH8jcxgAAAAAElFTSuQmCC'

# --- Single Instance Lock using a Mutex ---
class SingleInstance:
    """Ensures only one instance of the application can run."""
    def __init__(self, name):
        self.mutex_name = name
        self.mutex = win32event.CreateMutex(None, 1, self.mutex_name)
        self.last_error = win32api.GetLastError()

    def is_running(self):
        return self.last_error == ERROR_ALREADY_EXISTS

    def __del__(self):
        if self.mutex:
            win32api.CloseHandle(self.mutex)

# --- Main Application Class ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("API Connection Monitor")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        
        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')

        self.scheduler_thread = None
        self.stop_scheduler = threading.Event()
        self.icon = None

        # --- UI Elements ---
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Configuration Section
        config_frame = ttk.LabelFrame(self.main_frame, text="Configuration", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="API Endpoint:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.host_entry = ttk.Entry(config_frame)
        self.host_entry.insert(0, "tmgposapi.themall.co.th")
        self.host_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)

        ttk.Label(config_frame, text="Schedule Times (HH:MM):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        
        time_frame = ttk.Frame(config_frame)
        time_frame.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.time1_entry = ttk.Entry(time_frame, width=10)
        self.time1_entry.insert(0, "12:00")
        self.time1_entry.pack(side=tk.LEFT, padx=(0, 5))
        self.time2_entry = ttk.Entry(time_frame, width=10)
        self.time2_entry.insert(0, "17:00")
        self.time2_entry.pack(side=tk.LEFT, padx=5)
        self.time3_entry = ttk.Entry(time_frame, width=10)
        self.time3_entry.insert(0, "19:00")
        self.time3_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(config_frame, text="Log File Path:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.log_path_entry = ttk.Entry(config_frame)
        self.log_path_entry.insert(0, "C:\\Latency\\latency test")
        self.log_path_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        self.browse_button = ttk.Button(config_frame, text="Browse...", command=self.select_log_folder)
        self.browse_button.grid(row=2, column=2, padx=5, pady=5)

        # Control Section
        control_frame = ttk.Frame(self.main_frame, padding="10")
        control_frame.pack(fill=tk.X, pady=5)

        self.start_button = ttk.Button(control_frame, text="Start Monitoring", command=self.start_monitoring)
        self.start_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.stop_button = ttk.Button(control_frame, text="Stop Monitoring", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        # Log Section
        log_frame = ttk.LabelFrame(self.main_frame, text="Status Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10, relief="flat")
        self.log_area.pack(fill=tk.BOTH, expand=True)
        
        self.log("Welcome to the API Connection Monitor.")
        self.log("Configure your settings and click 'Start Monitoring'.")

        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        
        # New: Setup and run the tray icon from the start
        self.setup_tray_icon_thread()

    def select_log_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.log_path_entry.delete(0, tk.END)
            self.log_path_entry.insert(0, folder_selected)

    def log(self, message):
        self.root.after(0, self._log_message, message)

    def _log_message(self, message):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{now}] {message}\n")
        self.log_area.see(tk.END)

    def start_monitoring(self):
        self.host = self.host_entry.get()
        self.log_folder = self.log_path_entry.get()
        
        schedule_times_str = []
        time_entries = [self.time1_entry.get(), self.time2_entry.get(), self.time3_entry.get()]

        for t in time_entries:
            if t.strip():
                try:
                    time.strptime(t.strip(), '%H:%M')
                    schedule_times_str.append(t.strip())
                except ValueError:
                    self.log(f"Error: Invalid time format '{t}'. Please use HH:MM (24-hour format).")
                    return

        if not schedule_times_str:
            self.log("Error: Please enter at least one valid schedule time.")
            return

        for widget in [self.start_button, self.host_entry, self.time1_entry, self.time2_entry, self.time3_entry, self.log_path_entry, self.browse_button]:
            widget.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        self.log(f"Monitoring started for {self.host}. Scheduled times: {', '.join(schedule_times_str)}")
        
        schedule.clear()
        for scheduled_time in schedule_times_str:
            schedule.every().day.at(scheduled_time).do(self.run_diagnostics_thread)
        
        self.stop_scheduler.clear()
        self.scheduler_thread = threading.Thread(target=self.run_scheduler, daemon=True)
        self.scheduler_thread.start()

        threading.Timer(2.0, self.run_diagnostics_thread).start()

    def stop_monitoring(self):
        self.stop_scheduler.set()
        schedule.clear()
        
        for widget in [self.start_button, self.host_entry, self.time1_entry, self.time2_entry, self.time3_entry, self.log_path_entry, self.browse_button]:
            widget.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        
        self.log("Monitoring stopped by user.")

    def run_scheduler(self):
        while not self.stop_scheduler.is_set():
            schedule.run_pending()
            time.sleep(1)

    def run_diagnostics_thread(self):
        threading.Thread(target=self.run_diagnostics, daemon=True).start()

    def run_diagnostics(self):
        self.log(f"Running diagnostics for {self.host}...")
        computer_name = socket.gethostname()
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = f"{computer_name}_{timestamp}.txt"
        
        os.makedirs(self.log_folder, exist_ok=True)
        full_path = os.path.join(self.log_folder, file_name)

        bat_content = f"""@echo off
(
    ECHO COMPREHENSIVE NETWORK DIAGNOSTIC REPORT
    ECHO =================================================
    ECHO Report generated on: %date% at %time%
    ECHO Target Host: {self.host}
    ECHO.
    ECHO.
    ECHO ===== 1. TRACEROUTE TO VIEW NETWORK PATH =====
    ECHO.
    tracert {self.host}
    ECHO.
    ECHO.
    ECHO ===== 2. DNS LATENCY & RESOLUTION TEST =====
    ECHO.
    powershell -ExecutionPolicy Bypass -Command "Measure-Command {{Resolve-DnsName {self.host} -Type A -ErrorAction SilentlyContinue}}"
    ECHO.
    ECHO.
    ECHO ===== 3. CURL API CONNECTION TIMING =====
    ECHO.
    curl -o nul -s -w "DNS Lookup:      %%{{time_namelookup}}s\\nTCP Connection:  %%{{time_connect}}s\\nSSL Handshake:   %%{{time_appconnect}}s\\nTTFB:              %%{{time_starttransfer}}s\\nTotal Time:      %%{{time_total}}s\\n" https://{self.host}
) > "{full_path}" 2>&1
"""
        temp_bat_path = os.path.join(os.environ["TEMP"], "diag_script.bat")
        try:
            with open(temp_bat_path, "w", encoding='utf-8') as f:
                f.write(bat_content)
            
            subprocess.run([temp_bat_path], shell=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.log(f"Diagnostics complete. Log saved to {full_path}")
        except Exception as e:
            self.log(f"A critical error occurred during diagnostics: {e}")
        finally:
            if os.path.exists(temp_bat_path):
                os.remove(temp_bat_path)

    # --- System Tray Logic (Rewritten for stability) ---
    def setup_tray_icon_thread(self):
        """Creates and runs the system tray icon in a separate thread from the start."""
        icon_data = base64.b64decode(ICON_B64)
        image = Image.open(BytesIO(icon_data))
        menu = (item('Show', self.show_from_tray, default=True), item('Exit', self.exit_app))
        self.icon = Icon("API_Monitor", image, "API Connection Monitor", menu)
        self.icon.visible = False # Start hidden
        
        # Run the icon in a separate thread
        tray_thread = threading.Thread(target=self.icon.run, daemon=True)
        tray_thread.start()

    def hide_to_tray(self):
        """Hides the main window and makes the tray icon visible."""
        self.log("Minimizing to system tray...")
        self.root.withdraw()
        self.icon.visible = True

    def show_from_tray(self):
        """Hides the tray icon and shows the main window."""
        self.icon.visible = False
        self.root.after(0, self.root.deiconify)

    def exit_app(self):
        """Cleans up and exits the application."""
        self.icon.stop()
        self.stop_monitoring()
        self.root.quit() # Use quit instead of destroy for cleaner exit

# --- Main Application Execution ---
if __name__ == '__main__':
    instance_name = "Global\\API_Monitor_UI_Mutex_v5"
    instance = SingleInstance(instance_name)
    if instance.is_running():
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Application Already Running", 
                               "An instance of the API Connection Monitor is already running.")
        root.destroy()
        sys.exit(1)
        
    root = tk.Tk()
    app = App(root)
    root.mainloop()

