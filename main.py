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
from tkinter import ttk, scrolledtext

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
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        self.scheduler_thread = None
        self.stop_scheduler = threading.Event()

        # --- UI Elements ---
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Configuration Section
        config_frame = ttk.LabelFrame(self.main_frame, text="Configuration", padding="10")
        config_frame.pack(fill=tk.X, pady=5)

        ttk.Label(config_frame, text="API Endpoint:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.host_entry = ttk.Entry(config_frame, width=40)
        self.host_entry.insert(0, "tmgposapi.themall.co.th")
        self.host_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(config_frame, text="Frequency (minutes):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.interval_entry = ttk.Entry(config_frame, width=10)
        self.interval_entry.insert(0, "15")
        self.interval_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)

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

        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.log_area.pack(fill=tk.BOTH, expand=True)
        
        self.log("Welcome to the API Connection Monitor.")
        self.log("Enter your settings and click 'Start Monitoring'.")

    def log(self, message):
        """Adds a message to the log area on the UI thread."""
        self.root.after(0, self._log_message, message)

    def _log_message(self, message):
        """Internal method to update the text widget."""
        now = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{now}] {message}\n")
        self.log_area.see(tk.END)

    def start_monitoring(self):
        """Starts the background scheduling and diagnostic process."""
        self.host = self.host_entry.get()
        try:
            self.interval = int(self.interval_entry.get())
            if self.interval <= 0:
                raise ValueError
        except ValueError:
            self.log("Error: Please enter a valid positive number for the frequency.")
            return

        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.host_entry.config(state=tk.DISABLED)
        self.interval_entry.config(state=tk.DISABLED)

        self.log(f"Monitoring started for {self.host} every {self.interval} minutes.")
        
        # Setup scheduler
        schedule.every(self.interval).minutes.do(self.run_diagnostics_thread)
        
        self.stop_scheduler.clear()
        self.scheduler_thread = threading.Thread(target=self.run_scheduler, daemon=True)
        self.scheduler_thread.start()

        # Run once immediately after a short delay
        threading.Timer(2.0, self.run_diagnostics_thread).start()

    def stop_monitoring(self):
        """Stops the scheduler."""
        self.stop_scheduler.set()
        schedule.clear()
        
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.host_entry.config(state=tk.NORMAL)
        self.interval_entry.config(state=tk.NORMAL)
        
        self.log("Monitoring stopped by user.")

    def run_scheduler(self):
        """The loop that runs the scheduler."""
        while not self.stop_scheduler.is_set():
            schedule.run_pending()
            time.sleep(1)

    def run_diagnostics_thread(self):
        """Wrapper to run diagnostics in a separate thread to avoid freezing the UI."""
        threading.Thread(target=self.run_diagnostics, daemon=True).start()

    def run_diagnostics(self):
        """The core diagnostic function that runs CMD commands."""
        self.log(f"Running diagnostics for {self.host}...")
        log_folder = "C:\\Latency\\latency test"
        computer_name = socket.gethostname()
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = f"{computer_name}_{timestamp}.txt"
        
        os.makedirs(log_folder, exist_ok=True)
        full_path = os.path.join(log_folder, file_name)

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
            with open(temp_bat_path, "w") as f:
                f.write(bat_content)
            
            subprocess.run([temp_bat_path], shell=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.log(f"Diagnostics complete. Log saved to {full_path}")
        except Exception as e:
            self.log(f"A critical error occurred during diagnostics: {e}")
        finally:
            if os.path.exists(temp_bat_path):
                os.remove(temp_bat_path)

# --- Main Application Execution ---
if __name__ == '__main__':
    instance_name = "Global\\API_Monitor_UI_Mutex"
    instance = SingleInstance(instance_name)
    if instance.is_running():
        # You might want to use a proper message box in a real app
        print("Another instance is already running. Exiting.")
        sys.exit(1)
        
    root = tk.Tk()
    app = App(root)
    root.mainloop()

