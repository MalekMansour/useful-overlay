import ctypes
ctypes.windll.user32.SetProcessDPIAware()

import sys
import asyncio
import pythoncom
import psutil
import time
import datetime
import numpy as np
import sounddevice as sd
import winsdk.windows.media.control as wmc
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QHBoxLayout
from pynput import keyboard
import threading
import win32gui, win32process
import psutil as ps

# TRY importing GPUtil, but DO NOT CRASH if not installed
try:
    import GPUtil
    gpu_available = True
except:
    gpu_available = False

import os

# ─────────────────────────────────────────────
# COLORS
# ─────────────────────────────────────────────
COLOR_CYCLE = [
    "white",
    "#9ef39e",
    "#8de6ee",
    "#e73535",
    "#fd75d0",
    "#8943cf"
]
color_index = 0
current_color = COLOR_CYCLE[color_index]

# ─────────────────────────────────────────────
# SHARED STATS
# ─────────────────────────────────────────────
stats_lock = threading.Lock()
stats = {
    "battery": "Battery: --%",
    "ram": "RAM: --%",
    "gpu": "GPU: --%",
    "cpu": "CPU: --%",
    "app": "App: —",
    "date": "Date: --/--/----",
    "time": "Time: --:--",
    "timer_seconds": 0,
    "mic_bars": 0,
    "mic_percent": 0,
    "spotify": "Spotify: —"
}
_worker_stop = False

# ─────────────────────────────────────────────
# SPOTIFY
# ─────────────────────────────────────────────
async def spotify_now_playing():
    try:
        pythoncom.CoInitialize()
        sessions = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        current = sessions.get_current_session()
        if current is None:
            return "Spotify: —"

        info = await current.try_get_media_properties_async()
        source = str(current.source_app_user_model_id).lower()

        if "spotify" not in source:
            return "Spotify: —"

        artist = info.artist or ""
        title = info.title or ""

        if not artist and not title:
            return "Spotify: —"

        return f"{artist} – {title}"
    except Exception as e:
        print("Spotify Error:", e)
        return "Spotify: —"


def fetch_spotify_sync():
    try:
        return asyncio.run(spotify_now_playing())
    except Exception as e:
        print("Spotify Sync Error:", e)
        return "Spotify: —"

# ─────────────────────────────────────────────
# MIC LEVEL
# ─────────────────────────────────────────────
def get_mic_level_blocking():
    try:
        audio = sd.rec(int(0.05 * 44100), samplerate=44100, channels=1,
                       dtype='float32', blocking=True)
        rms = float(np.sqrt(np.mean(np.square(audio))))
        scaled = min(max(rms * 180.0, 0.0), 1.0)
        percent = int(scaled * 100)
        bars = int((percent / 100) * 10)
        return bars, percent
    except Exception as e:
        print("MIC Error:", e)
        return 0, 0

# ─────────────────────────────────────────────
# WORKER THREAD
# ─────────────────────────────────────────────
def stats_worker_loop():
    pythoncom.CoInitialize()
    global _worker_stop

    last_spotify = 0

    while not _worker_stop:
        try:
            # Battery
            battery = psutil.sensors_battery()
            batt = f"Battery: {battery.percent}%" if battery else "Battery: --%"

            ram = f"RAM: {psutil.virtual_memory().percent}%"
            cpu = f"CPU: {psutil.cpu_percent(interval=0.1)}%"

            # GPU SAFE
            if gpu_available:
                try:
                    g = GPUtil.getGPUs()[0]
                    gpu = f"GPU: {g.load*100:.0f}%"
                except:
                    gpu = "GPU: N/A"
            else:
                gpu = "GPU: N/A"

            # Foreground app
            try:
                hwnd = win32gui.GetForegroundWindow()
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                app_name = ps.Process(pid).name()
                app = f"App: {app_name}"
            except:
                app = "App: —"

            # Date / Time
            now = datetime.datetime.now()
            date = f"Date: {now.day:02}/{now.month:02}/{now.year}"
            tim = now.strftime("Time: %I:%M %p")

            # MIC
            mic_bars, mic_percent = get_mic_level_blocking()

            # Spotify (every 1 second)
            if time.time() - last_spotify >= 1:
                sp = fetch_spotify_sync()
                last_spotify = time.time()
            else:
                sp = stats.get("spotify", "Spotify: —")

            # WRITE ALL STATS
            with stats_lock:
                stats.update({
                    "battery": batt,
                    "ram": ram,
                    "gpu": gpu,
                    "cpu": cpu,
                    "app": app,
                    "date": date,
                    "time": tim,
                    "mic_bars": mic_bars,
                    "mic_percent": mic_percent,
                    "spotify": sp
                })

        except Exception as e:
            print("Worker error:", e)

        time.sleep(0.05)

# ─────────────────────────────────────────────
# OVERLAY UI
# ─────────────────────────────────────────────
class Overlay(QWidget):
    def __init__(self):
        super().__init__()
        global current_color

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setStyleSheet("background-color: black;")

        scr_w = QApplication.primaryScreen().size().width()
        self.setGeometry(0, 0, scr_w, 24)

        layout = QHBoxLayout()
        layout.setContentsMargins(8, 2, 8, 2)
        layout.setSpacing(40)

        self.battery_label = QLabel()
        self.ram_label = QLabel()
        self.gpu_label = QLabel()
        self.cpu_label = QLabel()
        self.app_label = QLabel()
        self.date_label = QLabel()
        self.time_label = QLabel()
        self.mic_label = QLabel()
        self.spotify_label = QLabel()

        self.labels = [
            self.battery_label, self.ram_label, self.gpu_label, self.cpu_label,
            self.app_label, self.date_label, self.time_label,
            self.mic_label, self.spotify_label
        ]

        for lbl in self.labels:
            lbl.setStyleSheet(f"color: {current_color}; font-size: 11px;")
            layout.addWidget(lbl)

        layout.addStretch(1)
        self.setLayout(layout)

        # Update timer
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_overlay)
        self.timer.start(200)

        # Start worker
        self.worker = threading.Thread(target=stats_worker_loop, daemon=True)
        self.worker.start()

    def update_overlay(self):
        with stats_lock:
            self.battery_label.setText(stats["battery"])
            self.ram_label.setText(stats["ram"])
            self.gpu_label.setText(stats["gpu"])
            self.cpu_label.setText(stats["cpu"])
            self.app_label.setText(stats["app"])
            self.date_label.setText(stats["date"])
            self.time_label.setText(stats["time"])

            bars = max(0, min(10, stats["mic_bars"]))
            mic_bar = "█" * bars + "░" * (10 - bars)
            self.mic_label.setText(f"Mic: {mic_bar} {stats['mic_percent']}%")

            self.spotify_label.setText(stats["spotify"])


# ─────────────────────────────────────────────
# RUN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    o = Overlay()
    o.show()
    try:
        sys.exit(app.exec_())
    finally:
        _worker_stop = True
