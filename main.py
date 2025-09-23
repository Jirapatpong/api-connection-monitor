
import subprocess
import os
import sys
import socket
import threading
import time
from datetime import datetime

# Third-party deps expected in your environment:
#   pip install pystray pillow schedule pywin32
import schedule
import win32event
import win32api
from winerror import ERROR_ALREADY_EXISTS

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox

from pystray import MenuItem as item, Icon, Menu
from PIL import Image, ImageDraw
import base64
from io import BytesIO

# =====================================================
# System Tray-Ready Mini App (Close button -> tray)
# =====================================================
# Notes:
# - Clicking the window "X" now hides to tray instead of quitting.
# - Tray menu: Show (restore window), Exit (quit app).
# - Single instance enforced via a Windows mutex.
# - The rest of your app flow is preserved/minimal changes.
# - If Base64 icon fails, we auto-fix padding and/or fall back to a generated icon.
# =====================================================

# --- Base64 Encoded Icon (Self-Contained) ---
# (Tiny 32x32 PNG; padding is auto-corrected to avoid 'Incorrect padding' errors)
ICON_B64 = b'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAACXBIWXMAAAsSAAALEgHS3X78AAABWklEQVRYhe2XzUrDUBSGv7c0rR1w4mQ3WkqS2YgC3UqN5c5Zxj0P1YJkKkqCqR5gH1V4mYgqSxM2Wc4q5rH0g1t2q1fV2cC8c6f8o8x8qg7v3Vg0Z9S7Q4QYb2G7mQy5wWJb1cE5cK5sY6f8n6c0kEoYI7kZl0C2kZgCwq4i6rQ8M2wQyVvC9h4F4vC5vZ1mH3x6w1k7KX+H2G0x8QmJwJX1c2o2qk6q7qj0g3N8Q8xwKpGkJQ5LMm1mC5qkq8rQnJw0qfG4dZqQ7g3pQb7m0lqY8kG6G2Yp5vG0sQ6mFf0Q8o4g6bZxw2nG9S8mJ8j9i6rCq7m4rC0+o7pR5r1f0uD3eYc/3rJqYp5zWw4oZ6rQm2l8tY0w3F2+q2g2m0i9sE6r5y7b1i9mJmKc8E1hR+Y0nHcE1w8z8R6CqE9iV1q6B7+8Q1mGgRkH9cJ3iGqgYbq9h4YH0pJrK1cZ8S5cQKXvJm2zZ7f4y6+Z3W0JQn8E0m4b0bG2cYv2xkB3zY5y2dD3vQy7P6qB4Uqj3lbnD7o+2f6o8fG1xv8T8mL1+Y3mCwS4FzQAAAABJRU5ErkJggg=='

def decode_icon_b64(data: bytes) -> Image.Image:
    """Decode PNG base64 into a PIL Image, auto-fixing padding; if it fails, return a generated fallback icon."""
    try:
        # Auto-fix missing padding (len must be multiple of 4)
        missing = len(data) % 4
        if missing:
            data += b'=' * (4 - missing)
        raw = base64.b64decode(data)
        return Image.open(BytesIO(raw))
    except Exception:
        # Fallback: generate a simple 32x32 accent icon
        img = Image.new('RGBA', (32, 32), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle((2, 2, 30, 30), radius=6, fill=(220, 0, 0, 255))
        draw.line((10, 10, 22, 22), width=3, fill=(255, 255, 255, 255))
        draw.line((22, 10, 10, 22), width=3, fill=(255, 255, 255, 255))
        return img

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
        if getattr(self, "mutex", None):
            win32api.CloseHandle(self.mutex)

# --- Main Application Class ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("API Connection Monitor")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        
        self.style = ttk.Style(self.root)
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        self.scheduler_thread = None
        self.stop_scheduler = threading.Event()

        # --- Tray / Icon state ---
        self.icon = None
        self.icon_visible = False
        self.icon_created = False
        self._icon_lock = threading.Lock()

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
        
        ttk.Label(config_frame, text="Log Folder:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
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

        # Intercept window close to go to tray
        self.root.protocol("WM_DELETE_WINDOW", self.on_close_clicked)

        # Prepare tray icon once so we don't recreate it
        self._prepare_tray_icon()

    # ---------- Tray Helpers ----------
    def _prepare_tray_icon(self):
        """Create the Icon object and menu once."""
        with self._icon_lock:
            if self.icon_created:
                return

            image = decode_icon_b64(ICON_B64)

            menu = Menu(
                item('Show', self._tray_show),
                item('Exit', self._tray_exit)
            )
            # Create but do not show yet
            self.icon = Icon("API_Monitor", image, "API Connection Monitor", menu)
            self.icon_created = True

    def _ensure_tray_visible(self):
        """Show the tray icon if it's not visible yet."""
        with self._icon_lock:
            if not self.icon_created:
                self._prepare_tray_icon()
            if not self.icon_visible:
                # run_detached ensures we don't block Tk's mainloop
                self.icon.run_detached(self._tray_setup)
                self.icon_visible = True

    def _tray_setup(self, icon):
        # Ensure the icon becomes visible immediately
        icon.visible = True

    def _tray_show(self, icon, item):
        # pystray thread -> safely schedule Tk calls
        self.root.after(0, self._show_window_from_tray)

    def _show_window_from_tray(self):
        with self._icon_lock:
            if self.icon and self.icon_visible:
                try:
                    self.icon.visible = False
                    self.icon.stop()
                except Exception:
                    pass
                self.icon_visible = False
        self.root.deiconify()
        self.root.after(0, self.root.lift)

    def _tray_exit(self, icon, item):
        self.root.after(0, self.exit_app)

    def on_close_clicked(self):
        """User clicked the window 'X' button → hide to tray instead of quitting."""
        self.log("Minimizing to system tray...")
        self.root.withdraw()
        self._ensure_tray_visible()

    # ---------- App UI + Logic ----------
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
        self.host = self.host_entry.get().strip()
        self.log_folder = self.log_path_entry.get().strip()
        
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

        if not self.host:
            self.log("Error: Please enter a valid API endpoint/host.")
            return

        if not schedule_times_str:
            self.log("Error: Please enter at least one valid schedule time.")
            return

        # Disable UI components
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

        # optional: run once shortly after start
        threading.Timer(2.0, self.run_diagnostics_thread).start()

    def stop_monitoring(self):
        self.stop_scheduler.set()
        schedule.clear()
        
        # Enable UI components
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
        
        try:
            os.makedirs(self.log_folder, exist_ok=True)
        except Exception as e:
            self.log(f"Error creating log folder '{self.log_folder}': {e}")
            return

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
    curl -o nul -s -w "DNS Lookup:      %%{{time_namelookup}}s\\nTCP Connection:  %%{{time_connect}}s\\nSSL Handshake:   %%{{time_appconnect}}s\\nTTFB:            %%{{time_starttransfer}}s\\nTotal Time:      %%{{time_total}}s\\n" https://{self.host}
) > "{full_path}" 2>&1
"""
        temp_bat_path = os.path.join(os.environ.get("TEMP", "."), "diag_script.bat")
        try:
            with open(temp_bat_path, "w", encoding='utf-8') as f:
                f.write(bat_content)
            
            # CREATE_NO_WINDOW may not exist on some builds; guard it.
            creationflag = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.run([temp_bat_path], shell=True, check=True, creationflags=creationflag)
            self.log(f"Diagnostics complete. Log saved to {full_path}")
        except Exception as e:
            self.log(f"A critical error occurred during diagnostics: {e}")
        finally:
            try:
                if os.path.exists(temp_bat_path):
                    os.remove(temp_bat_path)
            except Exception:
                pass

    def exit_app(self):
        """Cleans up and exits the application."""
        with self._icon_lock:
            if self.icon and self.icon_visible:
                try:
                    self.icon.visible = False
                    self.icon.stop()
                except Exception:
                    pass
            self.icon_visible = False
        self.stop_monitoring()
        self.root.destroy()

# --- Main Application Execution ---
if __name__ == '__main__':
    instance_name = "Global\\API_Monitor_UI_Mutex_v5"
    instance = SingleInstance(instance_name)
    if instance.is_running():
        # Minimal Tk just to show the message box cleanly
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Application Already Running", 
                               "An instance of the API Connection Monitor is already running.")
        root.destroy()
        sys.exit(1)
        
    root = tk.Tk()
    app = App(root)
    root.mainloop()
