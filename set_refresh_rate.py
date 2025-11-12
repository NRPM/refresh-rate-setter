#!/usr/bin/env python3
"""
set_refresh_gui.py

Simple Tkinter GUI that:
- Shows charging/discharging status (via GetSystemPowerStatus)
- If charging -> sets refresh rate to 240 Hz (unless overridden)
- If not charging -> sets refresh rate to 60 Hz (unless overridden)
- Provides a dropdown to override the target refresh rate manually
- Minimize button minimizes to system tray (requires pystray + pillow)
- Exit button exits entirely (removes tray icon if present)

Usage:
  python set_refresh_gui.py

Dependencies (install if missing):
  pip install pystray pillow

Notes:
- Uses Windows API (EnumDisplaySettingsW + ChangeDisplaySettingsExW) to change refresh rate.
- Targets the primary display. For multi-monitor setups you can edit the device parameter
  in EnumDisplaySettingsW / ChangeDisplaySettingsExW calls to target a specific \\.\DISPLAYn device.
"""

import ctypes
from ctypes import wintypes
import threading
import time
import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox

# Try to import pystray & PIL for system tray functionality.
# If not available, the minimize-to-tray feature will be disabled and the user will be prompted.
try:
    import pystray
    from PIL import Image, ImageDraw
    PYSTRAY_AVAILABLE = True
except Exception:
    PYSTRAY_AVAILABLE = False

# -------------------- Win32 Display functions -------------------- #
user32 = ctypes.WinDLL('user32', use_last_error=True)

ENUM_CURRENT_SETTINGS = -1
DM_DISPLAYFREQUENCY = 0x00400000  # DM_DISPLAYFREQUENCY bit
CDS_TEST = 0x00000002
DISP_CHANGE_SUCCESSFUL = 0

class DEVMODEW(ctypes.Structure):
    _fields_ = [
        ("dmDeviceName", wintypes.WCHAR * 32),
        ("dmSpecVersion", wintypes.WORD),
        ("dmDriverVersion", wintypes.WORD),
        ("dmSize", wintypes.WORD),
        ("dmDriverExtra", wintypes.WORD),
        ("dmFields", wintypes.DWORD),
        ("dmOrientation", wintypes.SHORT),
        ("dmPaperSize", wintypes.SHORT),
        ("dmPaperLength", wintypes.SHORT),
        ("dmPaperWidth", wintypes.SHORT),
        ("dmScale", wintypes.SHORT),
        ("dmCopies", wintypes.SHORT),
        ("dmDefaultSource", wintypes.SHORT),
        ("dmPrintQuality", wintypes.SHORT),
        ("dmColor", wintypes.SHORT),
        ("dmDuplex", wintypes.SHORT),
        ("dmYResolution", wintypes.SHORT),
        ("dmTTOption", wintypes.SHORT),
        ("dmCollate", wintypes.SHORT),
        ("dmFormName", wintypes.WCHAR * 32),
        ("dmLogPixels", wintypes.WORD),
        ("dmBitsPerPel", wintypes.DWORD),
        ("dmPelsWidth", wintypes.DWORD),
        ("dmPelsHeight", wintypes.DWORD),
        ("dmDisplayFlags", wintypes.DWORD),
        ("dmDisplayFrequency", wintypes.DWORD),
        ("dmICMMethod", wintypes.DWORD),
        ("dmICMIntent", wintypes.DWORD),
        ("dmMediaType", wintypes.DWORD),
        ("dmDitherType", wintypes.DWORD),
        ("dmReserved1", wintypes.DWORD),
        ("dmReserved2", wintypes.DWORD),
        ("dmPanningWidth", wintypes.DWORD),
        ("dmPanningHeight", wintypes.DWORD),
    ]

EnumDisplaySettingsW = user32.EnumDisplaySettingsW
EnumDisplaySettingsW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, ctypes.POINTER(DEVMODEW)]
EnumDisplaySettingsW.restype  = wintypes.BOOL

ChangeDisplaySettingsExW = user32.ChangeDisplaySettingsExW
ChangeDisplaySettingsExW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(DEVMODEW), wintypes.HWND, wintypes.DWORD, ctypes.c_void_p]
ChangeDisplaySettingsExW.restype  = wintypes.LONG

def list_modes(device_name=None):
    modes = []
    i = 0
    dm = DEVMODEW()
    dm.dmSize = ctypes.sizeof(DEVMODEW)
    while True:
        ok = EnumDisplaySettingsW(device_name, i, ctypes.byref(dm))
        if not ok:
            break
        modes.append((dm.dmPelsWidth, dm.dmPelsHeight, dm.dmBitsPerPel, dm.dmDisplayFrequency))
        i += 1
    # remove duplicates and sort
    modes = sorted(set(modes), key=lambda x: (x[0], x[1], x[3], x[2]))
    return modes

def get_current_mode(device_name=None):
    dm = DEVMODEW()
    dm.dmSize = ctypes.sizeof(DEVMODEW)
    if not EnumDisplaySettingsW(device_name, ENUM_CURRENT_SETTINGS, ctypes.byref(dm)):
        return None
    return dm

def get_available_refresh_rates(device_name=None):
    """
    Get all available refresh rates for the current resolution.
    Returns a sorted tuple of integers (Hz values).
    """
    try:
        # Get current display mode to determine current resolution
        dm = get_current_mode(device_name)
        if dm is None:
            print("Warning: Could not get current display mode. Using default rates.")
            return (60, 120, 144, 165, 240)  # Fallback to common rates
        
        # Get current resolution
        current_resolution = (dm.dmPelsWidth, dm.dmPelsHeight)
        print(f"Current resolution: {current_resolution[0]}x{current_resolution[1]}")
        
        # Get all display modes
        all_modes = list_modes(device_name)
        
        # Filter to only get refresh rates at current resolution
        available_rates = []
        for (width, height, bpp, refresh_rate) in all_modes:
            if (width, height) == current_resolution:
                available_rates.append(refresh_rate)
        
        # Remove duplicates and sort
        available_rates = sorted(set(available_rates))
        
        if not available_rates:
            print("Warning: No refresh rates found. Using default rates.")
            return (60, 120, 144, 165, 240)  # Fallback
        
        print(f"Available refresh rates: {available_rates}")
        
        # Return as tuple
        return tuple(available_rates)
        
    except Exception as e:
        print(f"Error getting available refresh rates: {e}")
        # Fallback to common refresh rates
        return (60, 120, 144, 165, 240)

def set_refresh_rate(target_hz, device_name=None, test_first=True):
    dm = get_current_mode(device_name)
    if dm is None:
        raise RuntimeError("Unable to get current display settings")

    cur_res = (dm.dmPelsWidth, dm.dmPelsHeight)
    modes = list_modes(device_name)
    valid_rates = [r for (w,h,bpp,r) in modes if (w,h)==cur_res]

    if target_hz not in valid_rates:
        raise RuntimeError(f"Requested {target_hz} Hz not supported at current resolution {cur_res}. Available: {sorted(set(valid_rates))}")

    new_dm = DEVMODEW()
    new_dm.dmSize = ctypes.sizeof(DEVMODEW)
    if not EnumDisplaySettingsW(device_name, ENUM_CURRENT_SETTINGS, ctypes.byref(new_dm)):
        raise RuntimeError("Unable to read current settings into new DEVMODE")

    new_dm.dmDisplayFrequency = int(target_hz)
    new_dm.dmFields = new_dm.dmFields | DM_DISPLAYFREQUENCY

    if test_first:
        res = ChangeDisplaySettingsExW(device_name, ctypes.byref(new_dm), None, CDS_TEST, None)
        if res != DISP_CHANGE_SUCCESSFUL:
            raise RuntimeError(f"Test-change failed (code {res}). Driver may not allow this mode.")

    res = ChangeDisplaySettingsExW(device_name, ctypes.byref(new_dm), None, 0, None)
    if res != DISP_CHANGE_SUCCESSFUL:
        raise RuntimeError(f"Change failed (code {res}).")
    return True

# -------------------- Power status (charging) -------------------- #
class SYSTEM_POWER_STATUS(ctypes.Structure):
    _fields_ = [
        ('ACLineStatus', wintypes.BYTE),
        ('BatteryFlag', wintypes.BYTE),
        ('BatteryLifePercent', wintypes.BYTE),
        ('Reserved1', wintypes.BYTE),
        ('BatteryLifeTime', wintypes.DWORD),
        ('BatteryFullLifeTime', wintypes.DWORD),
    ]

GetSystemPowerStatus = ctypes.windll.kernel32.GetSystemPowerStatus
GetSystemPowerStatus.argtypes = [ctypes.POINTER(SYSTEM_POWER_STATUS)]
GetSystemPowerStatus.restype = wintypes.BOOL

def is_plugged_in():
    status = SYSTEM_POWER_STATUS()
    if not GetSystemPowerStatus(ctypes.byref(status)):
        raise RuntimeError("GetSystemPowerStatus failed")
    # ACLineStatus: 0 = offline, 1 = online, 255 = unknown
    return status.ACLineStatus == 1

# -------------------- GUI & Tray -------------------- #
class RefreshGUI:
    POLL_INTERVAL_SECONDS = 1

    def __init__(self, root):
        self.root = root
        self.root.title("RefreshRate Manager")
        self.root.geometry("420x160")
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)  # catch close (exit)

        # State
        self.override_var = tk.BooleanVar(value=False)
        
        # Get available refresh rates dynamically
        self.available_rates = get_available_refresh_rates()
        
        # Set default selected rate to the highest available, or 240 if available
        if 240 in self.available_rates:
            default_rate = 240
        elif self.available_rates:
            default_rate = max(self.available_rates)
        else:
            default_rate = 60
            
        self.selected_rate = tk.IntVar(value=default_rate)
        self.current_status_var = tk.StringVar(value="Unknown")
        self.current_rate_var = tk.StringVar(value="Unknown")
        self.tray_icon = None
        self.tray_thread = None
        self.running = True

        # Build UI
        frm = ttk.Frame(root, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        status_row = ttk.Frame(frm)
        status_row.pack(fill=tk.X, pady=4)
        ttk.Label(status_row, text="Power status:").pack(side=tk.LEFT)
        ttk.Label(status_row, textvariable=self.current_status_var, width=20).pack(side=tk.LEFT, padx=6)
        ttk.Label(status_row, text="Current refresh:").pack(side=tk.LEFT, padx=(12,0))
        ttk.Label(status_row, textvariable=self.current_rate_var, width=10).pack(side=tk.LEFT, padx=6)

        override_row = ttk.Frame(frm)
        override_row.pack(fill=tk.X, pady=6)
        ttk.Checkbutton(override_row, text="Override automatic behavior", variable=self.override_var).pack(side=tk.LEFT)
        ttk.Label(override_row, text="Select refresh rate:").pack(side=tk.LEFT, padx=(8,6))
        self.rate_combo = ttk.Combobox(override_row, textvariable=self.selected_rate, width=8, state="readonly")
        
        # Populate dropdown with dynamically detected refresh rates
        self.rate_combo['values'] = self.available_rates
        self.rate_combo.pack(side=tk.LEFT)

        btn_row = ttk.Frame(frm)
        btn_row.pack(fill=tk.X, pady=(8,0))
        self.min_btn = ttk.Button(btn_row, text="Minimize to tray" if PYSTRAY_AVAILABLE else "Minimize (no pystray)", command=self.on_minimize)
        self.min_btn.pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_row, text="Apply now", command=self.on_apply_clicked).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_row, text="Exit", command=self.on_exit).pack(side=tk.RIGHT, padx=6)

        # Start polling thread
        self.poll_thread = threading.Thread(target=self.poll_loop, daemon=True)
        self.poll_thread.start()

        # Refresh displayed current refresh rate
        self.update_current_refresh_label()

        # If pystray available, create an icon image for tray
        if PYSTRAY_AVAILABLE:
            self.icon_image = self._create_image()
        else:
            self.icon_image = None

    def _create_image(self, width=64, height=64):
        # create a simple PIL image for tray icon
        img = Image.new('RGBA', (width, height), (0,0,0,0))
        d = ImageDraw.Draw(img)
        d.ellipse((8,8,width-8,height-8), fill=(30,144,255,255))
        d.text((18,18), "RR", fill=(255,255,255,255))
        return img

    def on_minimize(self):
        if not PYSTRAY_AVAILABLE:
            messagebox.showwarning("pystray missing", "pystray or pillow not installed. Install with:\n\npip install pystray pillow\n\nMinimize-to-tray won't work until installed.")
            return
        # Hide window and start tray icon
        self.root.withdraw()
        self.start_tray()

    def on_apply_clicked(self):
        # Apply based on current override or charging state immediately
        try:
            if self.override_var.get():
                target = int(self.selected_rate.get())
            else:
                # For automatic mode, use highest available rate when charging, lowest when not
                if is_plugged_in():
                    # Use 240 Hz if available, otherwise use highest available
                    target = 240 if 240 in self.available_rates else max(self.available_rates)
                else:
                    # Use 60 Hz if available, otherwise use lowest available
                    target = 60 if 60 in self.available_rates else min(self.available_rates)
            set_refresh_rate(target)
            self.current_rate_var.set(f"{target} Hz")
        except Exception as e:
            messagebox.showerror("Error applying refresh rate", str(e))

    def start_tray(self):
        if self.tray_icon:
            return  # already running
        menu = pystray.Menu(
            pystray.MenuItem("Open Window", lambda icon, item: self._tray_restore()),
            # pystray.MenuItem("Apply now", lambda icon, item: self._tray_apply()),
            pystray.MenuItem("Exit", lambda icon, item: self._tray_exit())
        )
        self.tray_icon = pystray.Icon("RefreshRateMgr", self.icon_image, "RefreshRate Manager", menu)
        self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
        self.tray_thread.start()

    def _tray_restore(self):
        try:
            self.root.after(0, self._do_restore)
        except Exception:
            pass

    def _do_restore(self):
        if self.tray_icon:
            try:
                self.tray_icon.stop()
            except Exception:
                pass
            self.tray_icon = None
        self.root.deiconify()
        self.root.lift()

    def _tray_apply(self):
        try:
            self.on_apply_clicked()
        except Exception:
            pass

    def _tray_exit(self):
        self.running = False
        try:
            if self.tray_icon:
                self.tray_icon.stop()
        except Exception:
            pass
        self.root.after(0, self.root.destroy)

    def on_exit(self):
        # exit button: remove tray if present and close
        self.running = False
        try:
            if self.tray_icon:
                self.tray_icon.stop()
        except Exception:
            pass
        self.root.destroy()

    def update_current_refresh_label(self):
        try:
            dm = get_current_mode(None)
            if dm is not None:
                self.current_rate_var.set(f"{dm.dmDisplayFrequency} Hz")
            else:
                self.current_rate_var.set("Unknown")
        except Exception:
            self.current_rate_var.set("Unknown")
        # schedule next label update in 5 seconds
        self.root.after(5000, self.update_current_refresh_label)

    def poll_loop(self):
        # Poll power status and apply appropriate refresh rate when changed
        last_plugged = None
        while self.running:
            try:
                plugged = is_plugged_in()
            except Exception:
                plugged = None
            # update UI text (on main thread)
            status_text = "Plugged In (Charging)" if plugged else "On Battery (Discharging)" if plugged is not None else "Unknown"
            try:
                self.root.after(0, self.current_status_var.set, status_text)
            except Exception:
                pass

            if plugged is not None and (last_plugged is None or plugged != last_plugged):
                # state changed -> decide and apply (respect override)
                if not self.override_var.get():
                    # Use dynamic rate selection based on available rates
                    if plugged:
                        # Use 240 Hz if available, otherwise highest
                        target = 240 if 240 in self.available_rates else max(self.available_rates)
                    else:
                        # Use 60 Hz if available, otherwise lowest
                        target = 60 if 60 in self.available_rates else min(self.available_rates)
                    
                    try:
                        set_refresh_rate(target)
                        self.root.after(0, self.current_rate_var.set, f"{target} Hz")
                    except Exception as e:
                        # show error once in UI thread
                        try:
                            self.root.after(0, messagebox.showwarning, "Could not change refresh rate", str(e))
                        except Exception:
                            pass
            last_plugged = plugged
            # poll interval
            for _ in range(int(self.POLL_INTERVAL_SECONDS*2)):
                if not self.running:
                    break
                time.sleep(0.5)

def main():
    if os.name != 'nt':
        print("This script only runs on Windows.")
        return

    root = tk.Tk()
    app = RefreshGUI(root)
    # Show the GUI on startup instead of starting in tray
    # If you want to start in tray, uncomment these lines:
    # root.withdraw()
    # app.start_tray()
    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass

if __name__ == '__main__':
    main()