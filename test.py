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
import GPUtil
import win32gui
import win32process
import psutil as ps
import os
import threading
import traceback

# ─────────────────────────────────────────────
# COLOR CYCLING
# ─────────────────────────────────────────────
COLOR_CYCLE = [
    "white",
    "#9ef39e",  # green
    "#8de6ee",  # cyan
    "#e73535",  # red
    "#fd75d0",  # pink
    "#8943cf"   # purple
]
color_index = 0
current_color = COLOR_CYCLE[color_index]

# ─────────────────────────────────────────────
# SHARED STATS (populated by background thread)
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
# SPOTIFY (called from worker) - kept defensive
# ─────────────────────────────────────────────
async def spotify_now_playing():
    pythoncom.CoInitialize()
    try:
        sessions = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        current = sessions.get_current_session()
        if current is None:
            return "Spotify: —"
        info = await current.try_get_media_properties_async()
        source = (current.source_app_user_model_id or "").lower()
        if "spotify" not in source:
            return "Spotify: —"
        title = info.title or ""
        artist = ", ".join((info.artist or "").split(";"))
        return f"{artist} – {title}" if (artist or title) else "Spotify: —"
    except Exception:
        return "Spotify: —"

def fetch_spotify_sync_worker():
    try:
        return asyncio.run(spotify_now_playing())
    except Exception:
        return "Spotify: —"

# ─────────────────────────────────────────────
# MIC LEVEL (blocking - ok in worker)
# ─────────────────────────────────────────────
def get_mic_level_blocking():
    try:
        duration = 0.05
        sample_rate = 44100
        audio = sd.rec(int(duration * sample_rate), samplerate=sample_rate,
                       channels=1, dtype='float32', blocking=True)
        rms = float(np.sqrt(np.mean(np.square(audio))))
        scaled = min(max(rms * 180.0, 0.0), 1.0)
        percent = int(scaled * 100)
        bars = int((percent / 100) * 10)
        return bars, percent
    except Exception:
        return 0, 0

# ─────────────────────────────────────────────
# BACKGROUND WORKER THREAD
# Collects stats off the GUI thread and updates `stats` dict.
# spotify_interval set to 5.0 seconds (your request)
# ─────────────────────────────────────────────
def stats_worker_loop(spotify_interval=5.0):
    last_spotify = 0.0
    global _worker_stop
    # warm up cpu_percent baseline
    try:
        psutil.cpu_percent(interval=None)
    except Exception:
        pass

    # do one immediate pass so UI doesn't wait
    first_pass = True
    while not _worker_stop:
        try:
            # Battery
            try:
                battery = psutil.sensors_battery()
                batt_text = f"Battery: {battery.percent}%" if battery else "Battery: --%"
            except Exception:
                batt_text = "Battery: --%"

            # RAM
            try:
                ram_percent = psutil.virtual_memory().percent
                ram_text = f"RAM: {ram_percent}%"
            except Exception:
                ram_text = "RAM: --%"

            # CPU (use a short blocking interval here on worker thread)
            try:
                cpu_percent = psutil.cpu_percent(interval=0.05)
                cpu_text = f"CPU: {cpu_percent}%"
            except Exception:
                cpu_text = "CPU: --%"

            # GPU (GPUtil can be slow; it's OK on worker thread)
            try:
                gpus = GPUtil.getGPUs()
                if gpus:
                    g = gpus[0]
                    gpu_temp = getattr(g, "temperature", "N/A")
                    gpu_text = f"GPU: {g.load*100:.0f}% {gpu_temp}°C"
                else:
                    gpu_text = "GPU: N/A"
            except Exception:
                gpu_text = "GPU: N/A"

            # Foreground app (fast)
            try:
                hwnd = win32gui.GetForegroundWindow()
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                proc = ps.Process(pid)
                app_text = f"App: {proc.name()}"
            except Exception:
                app_text = "App: —"

            # Date/Time
            now = datetime.datetime.now()
            date_text = f"Date: {now.day:02}/{now.month:02}/{now.year}"
            time_text = f"Time: {now.strftime('%I:%M %p')}"

            # Mic (blocking short recording)
            mic_bars, mic_percent = get_mic_level_blocking()

            # Spotify (throttled)
            now_t = time.time()
            if now_t - last_spotify >= spotify_interval or first_pass:
                try:
                    sp = fetch_spotify_sync_worker()
                except Exception:
                    sp = "Spotify: —"
                last_spotify = now_t
            else:
                with stats_lock:
                    sp = stats.get("spotify", "Spotify: —")

            # Save into shared stats dict (always update, even on partial failures)
            with stats_lock:
                stats["battery"] = batt_text
                stats["ram"] = ram_text
                stats["gpu"] = gpu_text
                stats["cpu"] = cpu_text
                stats["app"] = app_text
                stats["date"] = date_text
                stats["time"] = time_text
                stats["mic_bars"] = mic_bars
                stats["mic_percent"] = mic_percent
                stats["spotify"] = sp

            first_pass = False

        except Exception:
            # log unexpected error to console (so you can see what's failing)
            print("stats_worker_loop exception:\n", traceback.format_exc())
        # small sleep to avoid tight loop; mic and cpu_percent already add blocking
        time.sleep(0.05)

# ─────────────────────────────────────────────
# OVERLAY CLASS (UI reads cached `stats`)
# ─────────────────────────────────────────────
class Overlay(QWidget):
    def __init__(self):
        super().__init__()

        global current_color

        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Tool
        )
        self.setStyleSheet("background-color: black;")

        screen_width = QApplication.primaryScreen().size().width()
        self.setGeometry(0, 0, screen_width, 26)

        layout = QHBoxLayout()
        layout.setContentsMargins(12, 2, 12, 2)
        layout.setSpacing(90)

        self.battery_label = QLabel("Battery: --%")
        self.ram_label = QLabel("RAM: --%")
        self.gpu_label = QLabel("GPU: --%")
        self.cpu_label = QLabel("CPU: --%")
        self.app_label = QLabel("App: —")
        self.date_label = QLabel("Date: --/--/----")
        self.time_label = QLabel("Time: --:--")
        self.timer_label = QLabel("Timer: 00:00")
        self.mic_label = QLabel("Mic: ░░░░░░░░░░ 0%")
        self.spotify_label = QLabel("Spotify: —")

        self.labels = [
            self.battery_label, self.ram_label, self.gpu_label, self.cpu_label,
            self.app_label, self.date_label, self.time_label, self.timer_label,
            self.mic_label, self.spotify_label
        ]

        for lbl in self.labels:
            lbl.setStyleSheet(f"color: {current_color}; font-size: 12px;")
            layout.addWidget(lbl)

        layout.addStretch(1)
        self.setLayout(layout)

        # timer state (UI side)
        self.timer_running = False
        self.start_time = None
        self.seconds = 0

        self._last_spotify_time = 0
        # change the UI's idea of spotify interval to match worker (not required)
        self._spotify_interval = 5.0

        self.keys_down = set()

        # start background worker thread (5 second spotify interval)
        self.worker_thread = threading.Thread(target=stats_worker_loop, args=(self._spotify_interval,), daemon=True)
        self.worker_thread.start()
        print("stats worker started")

        # UI update timer reads cached stats (cheap)
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_overlay)
        self.update_timer.start(300)  # UI refresh rate (ms)

        # keyboard listener
        self.listener = keyboard.Listener(
            on_press=self.key_press,
            on_release=self.key_release
        )
        self.listener.start()

    # HOTKEYS
    def key_press(self, key):
        global current_color, color_index, _worker_stop

        if hasattr(key, "vk"):
            vk = key.vk
        else:
            return

        if vk in self.keys_down:
            return
        self.keys_down.add(vk)

        # NumPad 2 → color cycle
        if vk == 98:
            color_index = (color_index + 1) % len(COLOR_CYCLE)
            current_color = COLOR_CYCLE[color_index]
            self.apply_colors()

        # NumPad 3 → restart program (graceful)
        if vk == 99:
            _worker_stop = True
            time.sleep(0.05)
            os.execv(sys.executable, [sys.executable] + sys.argv)

        # NumPad 1 → timer toggle
        if vk == 97:
            if not self.timer_running:
                self.start_time = time.time()
                self.timer_running = True
            else:
                self.timer_running = False
                self.start_time = None
                self.seconds = 0

    def key_release(self, key):
        if hasattr(key, "vk"):
            self.keys_down.discard(key.vk)

    def apply_colors(self):
        for lbl in self.labels:
            lbl.setStyleSheet(f"color: {current_color}; font-size: 12px;")

    # UI update: cheap reads from stats and update labels
    def update_overlay(self):
        # copy cached stats under lock
        with stats_lock:
            batt = stats.get("battery", "Battery: --%")
            ram = stats.get("ram", "RAM: --%")
            gpu = stats.get("gpu", "GPU: --%")
            cpu = stats.get("cpu", "CPU: --%")
            app = stats.get("app", "App: —")
            date = stats.get("date", "Date: --/--/----")
            tim = stats.get("time", "Time: --:--")
            mic_bars = stats.get("mic_bars", 0)
            mic_percent = stats.get("mic_percent", 0)
            spotify_text = stats.get("spotify", "Spotify: —")

        # set UI
        self.battery_label.setText(batt)
        self.ram_label.setText(ram)
        self.gpu_label.setText(gpu)
        self.cpu_label.setText(cpu)
        self.app_label.setText(app)
        self.date_label.setText(date)
        self.time_label.setText(tim)

        # Timer (UI-managed)
        if self.timer_running and self.start_time is not None:
            self.seconds = int(time.time() - self.start_time)
        mins = self.seconds // 60
        secs = self.seconds % 60
        self.timer_label.setText(f"Timer: {mins:02}:{secs:02}")

        # Mic
        bars = max(0, min(10, mic_bars))
        mic_bar = "█" * bars + "░" * (10 - bars)
        self.mic_label.setText(f"Mic: {mic_bar} {mic_percent}%")

        # Spotify
        self.spotify_label.setText(spotify_text)

# ─────────────────────────────────────────────
# RUN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    overlay = Overlay()
    overlay.show()
    try:
        sys.exit(app.exec_())
    finally:
        # request worker stop on exit
        _worker_stop = True
