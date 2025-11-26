# overlay_fixed.py
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
    # COM must be initialized on the thread that uses it
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
        # ensure COM cleaned up if something went wrong
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return "Spotify: —"

def fetch_spotify_sync_worker():
    try:
        return asyncio.run(spotify_now_playing())
    except Exception:
        return "Spotify: —"

# ─────────────────────────────────────────────
# MIC LEVEL
# ─────────────────────────────────────────────
def get_mic_level_blocking():
    """
    Record a very short snippet and return (bars, percent).
    This function is defensive: if sounddevice fails (no device, permission),
    it returns (0,0) rather than crashing the worker.
    """
    try:
        # small buffer, lower sample rate to reduce CPU
        duration = 0.05
        sample_rate = 22050
        # ensure a default input device exists
        sd.check_input_settings(samplerate=sample_rate, channels=1)
        audio = sd.rec(int(duration * sample_rate), samplerate=sample_rate,
                       channels=1, dtype='float32', blocking=True)
        if audio is None or audio.size == 0:
            return 0, 0
        rms = float(np.sqrt(np.mean(np.square(audio))))
        # Map rms to a reasonable 0-100 percent.
        # quiet mic -> small percent; loud -> near 100
        # tune multiplier empirically
        scaled = min(max(rms * 200.0, 0.0), 1.0)
        percent = int(scaled * 100)
        bars = int(round((percent / 100) * 10))
        return bars, percent
    except Exception:
        # silent failure is better than crash; print for debug
        # (don't spam repeatedly)
        return 0, 0

# ─────────────────────────────────────────────
# BACKGROUND WORKER THREAD
# ─────────────────────────────────────────────
def stats_worker_loop(spotify_interval=1.0):
    # Initialize COM for this worker thread
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    last_spotify = 0.0
    global _worker_stop

    # warm up cpu percent (psutil requires a call first)
    try:
        psutil.cpu_percent(interval=None)
    except Exception:
        pass

    # small sleep to let UI start
    time.sleep(0.05)

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

            # CPU
            try:
                cpu_percent = psutil.cpu_percent(interval=0.05)
                cpu_text = f"CPU: {cpu_percent}%"
            except Exception:
                cpu_text = "CPU: --%"

            # GPU
            try:
                gpus = GPUtil.getGPUs()
                if gpus:
                    g = gpus[0]
                    # safe getattr for temperature
                    temp = getattr(g, "temperature", None)
                    temp_text = f" {temp}°C" if temp is not None else ""
                    gpu_text = f"GPU: {g.load*100:.0f}%{temp_text}"
                else:
                    gpu_text = "GPU: N/A"
            except Exception:
                gpu_text = "GPU: N/A"

            # Foreground app
            try:
                hwnd = win32gui.GetForegroundWindow()
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                proc = ps.Process(pid)
                app_text = f"App: {proc.name()}"
            except Exception:
                app_text = "App: —"

            # Date / Time
            now = datetime.datetime.now()
            date_text = f"Date: {now.day:02}/{now.month:02}/{now.year}"
            time_text = f"Time: {now.strftime('%I:%M %p')}"

            # Mic
            mic_bars, mic_percent = get_mic_level_blocking()

            # Spotify (refresh every spotify_interval seconds)
            now_t = time.time()
            if now_t - last_spotify >= spotify_interval:
                try:
                    sp = fetch_spotify_sync_worker()
                except Exception:
                    sp = "Spotify: —"
                last_spotify = now_t
            else:
                with stats_lock:
                    sp = stats.get("spotify", "Spotify: —")

            # Write stats
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

        except Exception:
            # print traceback once per loop iteration so you can see what's wrong
            traceback.print_exc()

        # short sleep to avoid busy loop but keep UI snappy
        time.sleep(0.05)

    # Cleanup COM for this thread
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

# ─────────────────────────────────────────────
# OVERLAY CLASS
# ─────────────────────────────────────────────
class Overlay(QWidget):
    def __init__(self):
        super().__init__()

        global current_color

        # make window frameless, always on top, not focused
        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Tool
        )
        # nicer look
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowOpacity(0.92)

        # get screen width in a safe way (QApplication instance must exist)
        screen_width = 800
        try:
            app_instance = QApplication.instance()
            if app_instance is not None and app_instance.primaryScreen() is not None:
                screen_width = app_instance.primaryScreen().size().width()
        except Exception:
            pass

        # small height for a compact bar
        self.setGeometry(0, 0, screen_width, 28)
        self.setFixedHeight(28)

        layout = QHBoxLayout()
        layout.setContentsMargins(8, 2, 8, 2)
        layout.setSpacing(12)  # reduced spacing to avoid labels pushed off

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
            lbl.setFixedHeight(20)
            layout.addWidget(lbl)

        layout.addStretch(1)
        self.setLayout(layout)

        self.timer_running = False
        self.start_time = None
        self.seconds = 0

        self._spotify_interval = 1

        self.keys_down = set()

        # Start worker thread
        self.worker_thread = threading.Thread(target=stats_worker_loop, args=(self._spotify_interval,), daemon=True)
        self.worker_thread.start()

        # UI update timer
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_overlay)
        self.update_timer.start(300)

        # Hotkeys
        self.listener = keyboard.Listener(
            on_press=self.key_press,
            on_release=self.key_release
        )
        self.listener.start()

        # Attempt a quick spotify fetch to avoid initial "—" if possible
        try:
            spinitial = fetch_spotify_sync_worker()
            with stats_lock:
                stats["spotify"] = spinitial
        except Exception:
            pass

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

        # NumPad 3 → restart program
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

    def update_overlay(self):
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

        self.battery_label.setText(batt)
        self.ram_label.setText(ram)
        self.gpu_label.setText(gpu)
        self.cpu_label.setText(cpu)
        self.app_label.setText(app)
        self.date_label.setText(date)
        self.time_label.setText(tim)

        if self.timer_running and self.start_time:
            self.seconds = int(time.time() - self.start_time)
        mins = self.seconds // 60
        secs = self.seconds % 60
        self.timer_label.setText(f"Timer: {mins:02}:{secs:02}")

        bars = max(0, min(10, mic_bars))
        mic_bar = "█" * bars + "░" * (10 - bars)
        self.mic_label.setText(f"Mic: {mic_bar} {mic_percent}%")

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
        _worker_stop = True
