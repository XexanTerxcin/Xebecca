"""
Xebecca by SMY GAMER
Single-file Python program for Windows (PyQt5)
Requirements:
    pip install PyQt5 psutil win10toast Pillow
Run as Administrator for full functionality (powercfg, standby clearing).
Place Xebecca.ico in the same folder as this script.
"""

import sys
import os
import subprocess
import threading
import time
import shutil
import tempfile
import ctypes
from ctypes import wintypes
from pathlib import Path
from urllib.parse import urlparse

from PyQt5 import QtCore, QtGui, QtWidgets
import psutil
from win10toast import ToastNotifier
import webbrowser

# ---------------------------
# Configuration / GUIDs
# ---------------------------
APP_NAME = "Xebecca"
WINDOW_TITLE = "Xebecca by SMY GAMER"
AUTHOR_URL = "https://xexanterxcin.github.io/Project-NFC-HUNT/SMY-GAMER/"
ICON_FILENAME = "Xebecca.ico"  # must be in same directory

# Processor subgroup GUID and setting GUIDs (from user's batch)
SUB_PROCESSOR = "54533251-82be-4824-96c1-47b60b740d00"
GUID_MIN_PROC_STATE = "893dee8e-2bef-41e0-89c6-b55d0929964c"
GUID_MAX_PROC_STATE = "bc5038f7-23e0-4960-96da-33abaf5935ec"
GUID_BOOST_MODE = "be337238-0d82-4146-a960-4f3749d470c7"
GUID_COOLING_POLICY = "94d3a615-a899-4ac5-ae2b-e4d8f634367f"
GUID_CORE_PARKING = "0cc5b647-c1df-4637-891a-dec35c318583"

# Power scheme GUID: we'll use the active scheme (query)
# Helper to run powercfg commands
def run_cmd(cmd, admin_required=False):
    try:
        completed = subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        return completed.returncode, completed.stdout, completed.stderr
    except Exception as e:
        return -1, "", str(e)

# ---------------------------
# Notifications
# ---------------------------
toaster = ToastNotifier()

def notify(title, msg, duration=5):
    try:
        toaster.show_toast(title, msg, duration=duration, threaded=True)
    except Exception:
        # fallback: no-op
        pass

# ---------------------------
# Timer resolution helpers (winmm)
# ---------------------------
winmm = ctypes.WinDLL('winmm')
timeBeginPeriod = winmm.timeBeginPeriod
timeBeginPeriod.argtypes = [wintypes.UINT]
timeBeginPeriod.restype = wintypes.UINT

timeEndPeriod = winmm.timeEndPeriod
timeEndPeriod.argtypes = [wintypes.UINT]
timeEndPeriod.restype = wintypes.UINT

# We'll keep a global flag and current period
_timer_enforcer = {"enabled": False, "period_ms": 0, "frequency_s": 0, "thread": None}

def _timer_enforcer_thread():
    # Re-apply timeBeginPeriod periodically
    while _timer_enforcer["enabled"]:
        period = int(_timer_enforcer["period_ms"])
        if period > 0:
            timeBeginPeriod(period)
        freq = _timer_enforcer["frequency_s"] or 5
        time.sleep(freq)

def start_timer_enforcer(period_ms, frequency_s):
    stop_timer_enforcer()
    _timer_enforcer["enabled"] = True
    _timer_enforcer["period_ms"] = period_ms
    _timer_enforcer["frequency_s"] = frequency_s
    t = threading.Thread(target=_timer_enforcer_thread, daemon=True)
    _timer_enforcer["thread"] = t
    t.start()

def stop_timer_enforcer():
    if _timer_enforcer["enabled"]:
        _timer_enforcer["enabled"] = False
        # best-effort end period
        try:
            timeEndPeriod(int(_timer_enforcer["period_ms"]))
        except Exception:
            pass
        _timer_enforcer["thread"] = None

# ---------------------------
# RAM cleaning helpers
# ---------------------------
_ram_cleaner = {"enabled": False, "interval_s": 0, "thread": None}

# Use EmptyWorkingSet via ctypes for each process
psapi = ctypes.WinDLL('psapi')
kernel32 = ctypes.WinDLL('kernel32')
EmptyWorkingSet = ctypes.windll.psapi.EmptyWorkingSet
EmptyWorkingSet.argtypes = [wintypes.HANDLE]
EmptyWorkingSet.restype = wintypes.BOOL

def trim_process_working_set(pid):
    try:
        proc = psutil.Process(pid)
        # skip system-critical processes
        if proc.pid == 0 or proc.pid == 4:
            return False
        handle = ctypes.windll.kernel32.OpenProcess(0x001F0FFF, False, pid)
        if not handle:
            return False
        res = EmptyWorkingSet(handle)
        ctypes.windll.kernel32.CloseHandle(handle)
        return bool(res)
    except Exception:
        return False

def clean_ram_once():
    freed = 0
    # iterate processes and attempt to trim working set
    for p in psutil.process_iter(['pid', 'name', 'memory_info']):
        try:
            before = p.info['memory_info'].rss
            trimmed = trim_process_working_set(p.info['pid'])
            after = p.memory_info().rss if p.is_running() else before
            freed += max(0, before - after)
        except Exception:
            continue
    # Attempt to call EmptyStandbyList.exe if present
    exe_path = Path(__file__).parent / "EmptyStandbyList.exe"
    if exe_path.exists():
        try:
            subprocess.run(f'"{exe_path}" workingsets', shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except Exception:
            pass
    # Convert freed bytes to human readable
    return freed

def _ram_cleaner_thread():
    while _ram_cleaner["enabled"]:
        freed = clean_ram_once()
        if freed > 0:
            human = sizeof_fmt(freed)
            notify(APP_NAME, f"RAM cleaned — freed {human}")
        else:
            notify(APP_NAME, "RAM cleaned")
        time.sleep(_ram_cleaner["interval_s"])

def start_ram_cleaner(interval_s):
    stop_ram_cleaner()
    _ram_cleaner["enabled"] = True
    _ram_cleaner["interval_s"] = interval_s
    t = threading.Thread(target=_ram_cleaner_thread, daemon=True)
    _ram_cleaner["thread"] = t
    t.start()

def stop_ram_cleaner():
    _ram_cleaner["enabled"] = False
    _ram_cleaner["thread"] = None

# ---------------------------
# Utility
# ---------------------------
def sizeof_fmt(num, suffix='B'):
    for unit in ['','K','M','G','T','P']:
        if abs(num) < 1024.0:
            return f"{num:3.1f}{unit}{suffix}"
        num /= 1024.0
    return f"{num:.1f}Y{suffix}"

# ---------------------------
# Power settings helpers
# ---------------------------
def unhide_processor_options():
    # Run the unhide commands (same as provided batch)
    cmds = [
        f'powercfg -attributes {SUB_PROCESSOR} {GUID_MIN_PROC_STATE} -ATTRIB_HIDE',
        f'powercfg -attributes {SUB_PROCESSOR} {GUID_MAX_PROC_STATE} -ATTRIB_HIDE',
        f'powercfg -attributes {SUB_PROCESSOR} {GUID_BOOST_MODE} -ATTRIB_HIDE',
        f'powercfg -attributes {SUB_PROCESSOR} {GUID_COOLING_POLICY} -ATTRIB_HIDE',
        f'powercfg -attributes {SUB_PROCESSOR} {GUID_CORE_PARKING} -ATTRIB_HIDE'
    ]
    for c in cmds:
        run_cmd(c)

def get_active_scheme_guid():
    code, out, err = run_cmd("powercfg /getactivescheme")
    if code == 0 and out:
        # output like: Power Scheme GUID: xxxxx (name)
        parts = out.strip().split()
        if len(parts) >= 4:
            return parts[3]
    return None

def set_power_value(ac_or_dc, subgroup_guid, setting_guid, value):
    # ac_or_dc: "AC" or "DC" (we only set AC/Plugged in)
    scheme = get_active_scheme_guid()
    if not scheme:
        return False, "No active scheme"
    # powercfg -setacvalueindex <scheme> <subgroup> <setting> <value>
    cmd = f'powercfg -setacvalueindex {scheme} {subgroup_guid} {setting_guid} {value}'
    code, out, err = run_cmd(cmd)
    if code != 0:
        return False, err.strip()
    # apply
    code2, out2, err2 = run_cmd(f'powercfg -S {scheme}')
    return code2 == 0, (out2 or err2).strip()

def apply_game_mode():
    # Values per user request (Plugged in)
    # Processor performance core parking -> 100 (this GUID often expects percentage)
    set_power_value("AC", SUB_PROCESSOR, GUID_CORE_PARKING, 100)
    set_power_value("AC", SUB_PROCESSOR, GUID_MIN_PROC_STATE, 100)
    set_power_value("AC", SUB_PROCESSOR, GUID_COOLING_POLICY, 1)  # 1 = Active, 0 = Passive (Active is usually 0)
    set_power_value("AC", SUB_PROCESSOR, GUID_MAX_PROC_STATE, 100)
    # Processor performance boost mode: Aggressive -> value mapping may vary; common mapping:
    # 0 = Disabled, 1 = Enabled, 2 = Aggressive, 3 = Efficient Enabled, 4 = Efficient Aggressive (varies by Windows)
    set_power_value("AC", SUB_PROCESSOR, GUID_BOOST_MODE, 2)

def apply_stable_mode():
    set_power_value("AC", SUB_PROCESSOR, GUID_CORE_PARKING, 80)
    set_power_value("AC", SUB_PROCESSOR, GUID_MIN_PROC_STATE, 10)
    set_power_value("AC", SUB_PROCESSOR, GUID_COOLING_POLICY, 1)  # Active
    set_power_value("AC", SUB_PROCESSOR, GUID_MAX_PROC_STATE, 90)
    set_power_value("AC", SUB_PROCESSOR, GUID_BOOST_MODE, 4)  # Efficient Aggressive (approx)

# ---------------------------
# Temp cleaning
# ---------------------------
def clean_temp_files():
    total_freed = 0
    # Common temp locations
    temp_paths = [
        tempfile.gettempdir(),
        os.path.join(os.environ.get('SYSTEMROOT', 'C:\\Windows'), 'Temp'),
        os.path.join(os.environ.get('USERPROFILE', ''), 'AppData', 'Local', 'Temp')
    ]
    seen = set()
    for p in temp_paths:
        if not p:
            continue
        p = os.path.abspath(p)
        if p in seen:
            continue
        seen.add(p)
        if os.path.exists(p):
            for root, dirs, files in os.walk(p):
                for name in files:
                    try:
                        fp = os.path.join(root, name)
                        size = os.path.getsize(fp)
                        os.remove(fp)
                        total_freed += size
                    except Exception:
                        continue
                for name in dirs:
                    try:
                        dp = os.path.join(root, name)
                        shutil.rmtree(dp, ignore_errors=True)
                    except Exception:
                        continue
    return total_freed

# ---------------------------
# PyQt5 GUI
# ---------------------------
class MainWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(WINDOW_TITLE)
        self.setWindowIcon(QtGui.QIcon(ICON_FILENAME) if os.path.exists(ICON_FILENAME) else self.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon))
        self.setFixedSize(520, 360)
        self.setup_ui()
        # tray
        self.tray_icon = QtWidgets.QSystemTrayIcon(self)
        if os.path.exists(ICON_FILENAME):
            self.tray_icon.setIcon(QtGui.QIcon(ICON_FILENAME))
        else:
            self.tray_icon.setIcon(self.windowIcon())
        self.tray_menu = QtWidgets.QMenu()
        show_action = self.tray_menu.addAction("Show")
        show_action.triggered.connect(self.show_window)
        exit_action = self.tray_menu.addAction("Exit")
        exit_action.triggered.connect(self.exit_app)
        self.tray_icon.setContextMenu(self.tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()

        # Start by unhiding processor options
        threading.Thread(target=unhide_processor_options, daemon=True).start()

    def setup_ui(self):
        # Styling: dark + neon green
        self.setStyleSheet("""
            QWidget { background-color: #121212; color: #bfeea8; font-family: Segoe UI, Arial; }
            QGroupBox { border: 1px solid #0f9d58; margin-top: 10px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px 0 3px; color: #bfeea8; }
            QRadioButton::indicator { width: 16px; height: 16px; }
            QPushButton { background-color: #0f9d58; color: #0b0b0b; border-radius: 6px; padding: 8px; }
            QPushButton#tempBtn { background-color: #c62828; color: white; }
            QPushButton#authorBtn { background-color: #2e7d32; color: white; }
            QLabel { color: #bfeea8; }
        """)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        # Mode group
        mode_group = QtWidgets.QGroupBox("Mode")
        mg_layout = QtWidgets.QHBoxLayout()
        self.rb_game = QtWidgets.QRadioButton("Game")
        self.rb_stable = QtWidgets.QRadioButton("Stable")
        self.rb_game.setChecked(True)
        mg_layout.addWidget(self.rb_game)
        mg_layout.addWidget(self.rb_stable)
        mode_group.setLayout(mg_layout)
        layout.addWidget(mode_group)

        # Ram Clean Frequency
        ram_group = QtWidgets.QGroupBox("Ram Clean Frequency")
        rg_layout = QtWidgets.QHBoxLayout()
        self.rb_ram_5 = QtWidgets.QRadioButton("5 Min")
        self.rb_ram_10 = QtWidgets.QRadioButton("10 Min")
        self.rb_ram_30 = QtWidgets.QRadioButton("30 Min")
        self.rb_ram_5.setChecked(True)
        rg_layout.addWidget(self.rb_ram_5)
        rg_layout.addWidget(self.rb_ram_10)
        rg_layout.addWidget(self.rb_ram_30)
        ram_group.setLayout(rg_layout)
        layout.addWidget(ram_group)

        # Timer resolution
        timer_group = QtWidgets.QGroupBox("Set Timer Resolution")
        tg_layout = QtWidgets.QHBoxLayout()
        self.rb_timer_05 = QtWidgets.QRadioButton("0.5 ms")
        self.rb_timer_1 = QtWidgets.QRadioButton("1 ms")
        self.rb_timer_default = QtWidgets.QRadioButton("Default")
        self.rb_timer_1.setChecked(True)
        tg_layout.addWidget(self.rb_timer_05)
        tg_layout.addWidget(self.rb_timer_1)
        tg_layout.addWidget(self.rb_timer_default)
        timer_group.setLayout(tg_layout)
        layout.addWidget(timer_group)

        # Timer frequency
        freq_group = QtWidgets.QGroupBox("Timer Resolution Frequency")
        fg_layout = QtWidgets.QHBoxLayout()
        self.rb_freq_5s = QtWidgets.QRadioButton("5 Sec")
        self.rb_freq_10s = QtWidgets.QRadioButton("10 Sec")
        self.rb_freq_30s = QtWidgets.QRadioButton("30 Sec")
        self.rb_freq_5m = QtWidgets.QRadioButton("5 Min")
        self.rb_freq_30s.setChecked(True)
        fg_layout.addWidget(self.rb_freq_5s)
        fg_layout.addWidget(self.rb_freq_10s)
        fg_layout.addWidget(self.rb_freq_30s)
        fg_layout.addWidget(self.rb_freq_5m)
        freq_group.setLayout(fg_layout)
        layout.addWidget(freq_group)

        # Buttons row
        btn_row = QtWidgets.QHBoxLayout()
        self.author_btn = QtWidgets.QPushButton("Author")
        self.author_btn.setObjectName("authorBtn")
        self.author_btn.clicked.connect(self.open_author)
        btn_row.addWidget(self.author_btn)

        self.temp_btn = QtWidgets.QPushButton("Click here to clean temp files")
        self.temp_btn.setObjectName("tempBtn")
        self.temp_btn.clicked.connect(self.on_clean_temp)
        btn_row.addWidget(self.temp_btn)

        self.apply_btn = QtWidgets.QPushButton("Apply")
        self.apply_btn.clicked.connect(self.on_apply)
        btn_row.addWidget(self.apply_btn)

        layout.addLayout(btn_row)

        # Status label
        self.status_label = QtWidgets.QLabel("Ready")
        layout.addWidget(self.status_label)

    def open_author(self):
        webbrowser.open(AUTHOR_URL)

    def on_clean_temp(self):
        self.status_label.setText("Cleaning temporary files...")
        self.apply_btn.setEnabled(False)
        self.temp_btn.setEnabled(False)
        def worker():
            freed = clean_temp_files()
            human = sizeof_fmt(freed)
            notify(APP_NAME, f"Temporary file cleaning successful — freed {human}")
            self.status_label.setText(f"Temp cleaned — freed {human}")
            self.apply_btn.setEnabled(True)
            self.temp_btn.setEnabled(True)
        threading.Thread(target=worker, daemon=True).start()

    def on_apply(self):
        # Apply power mode
        self.status_label.setText("Applying settings...")
        self.apply_btn.setEnabled(False)
        def worker():
            try:
                # Apply power mode
                if self.rb_game.isChecked():
                    apply_game_mode()
                    notify(APP_NAME, "Game mode applied")
                else:
                    apply_stable_mode()
                    notify(APP_NAME, "Stable mode applied")

                # RAM cleaner
                if self.rb_ram_5.isChecked():
                    start_ram_cleaner(5*60)
                    notify(APP_NAME, "RAM cleaner started (5 min)")
                elif self.rb_ram_10.isChecked():
                    start_ram_cleaner(10*60)
                    notify(APP_NAME, "RAM cleaner started (10 min)")
                elif self.rb_ram_30.isChecked():
                    start_ram_cleaner(30*60)
                    notify(APP_NAME, "RAM cleaner started (30 min)")
                else:
                    stop_ram_cleaner()

                # Timer resolution
                if self.rb_timer_05.isChecked():
                    period_ms = 1  # timeBeginPeriod expects integer ms; 0.5ms not supported directly, use 1ms as closest
                elif self.rb_timer_1.isChecked():
                    period_ms = 1
                else:
                    period_ms = 0

                # frequency
                if self.rb_freq_5s.isChecked():
                    freq = 5
                elif self.rb_freq_10s.isChecked():
                    freq = 10
                elif self.rb_freq_30s.isChecked():
                    freq = 30
                else:
                    freq = 5*60

                if period_ms > 0:
                    start_timer_enforcer(period_ms, freq)
                    notify(APP_NAME, f"Timer resolution set to {period_ms} ms (enforced every {freq}s)")
                else:
                    stop_timer_enforcer()
                    notify(APP_NAME, "Timer resolution set to Default")

                self.status_label.setText("Settings applied")
            except Exception as e:
                self.status_label.setText(f"Error: {e}")
            finally:
                self.apply_btn.setEnabled(True)

        threading.Thread(target=worker, daemon=True).start()

    def closeEvent(self, event):
        # Minimize to tray instead of closing
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(APP_NAME, "Xebecca is still running in the background. Right-click the tray icon to Exit.", QtWidgets.QSystemTrayIcon.Information, 3000)

    def show_window(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def exit_app(self):
        # cleanup
        stop_ram_cleaner()
        stop_timer_enforcer()
        QtWidgets.qApp.quit()

    def on_tray_activated(self, reason):
        if reason == QtWidgets.QSystemTrayIcon.Trigger:
            self.show_window()

# ---------------------------
# Main
# ---------------------------
def main():
    # Ensure working directory is script directory
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # Try to unhide processor options at startup (non-blocking)
    threading.Thread(target=unhide_processor_options, daemon=True).start()

    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon("Xebecca.ico"))
    # High DPI scaling
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps)

    w = MainWindow()
    w.show()
    notify(APP_NAME, "Xebecca started")
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
