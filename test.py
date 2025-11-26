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

# ─────────────────────────────────────────────
# COLOR PRESETS (NumPad)
# ─────────────────────────────────────────────
COLORS = {
    7: "white",
    8: "#fd75d0",   # Pink
    9: "#e73535",   # Red
    4: "#9ef39e",   # Light Green
    5: "#8de6ee",   # Cyan
    6: "#8943cf"    # Purple
}
current_color = "white"

# ─────────────────────────────────────────────
# SPOTIFY NOW PLAYING
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
    except:
        return "Spotify: —"

def fetch_spotify_sync():
    try:
        return asyncio.run(spotify_now_playing())
    except:
        return "Spotify: —"

# ─────────────────────────────────────────────
# MIC FUNCTION USING SOUNDEVICE
# ─────────────────────────────────────────────
def get_mic_level():
    """Returns a tuple (bars, percent) for mic bar with 10 bars."""
    try:
        duration = 0.05  
        sample_rate = 44100
        audio = sd.rec(int(duration * sample_rate), samplerate=sample_rate, channels=1, dtype='float32', blocking=True)
        rms = float(np.sqrt(np.mean(np.square(audio))))
        scaled = min(max(rms * 180.0, 0.0), 1.0)  
        percent = int(scaled * 100)
        bars = int((percent / 100) * 10)  
        return bars, percent
    except Exception:
        return 0, 0

# ─────────────────────────────────────────────
# OVERLAY CLASS
# ─────────────────────────────────────────────
class Overlay(QWidget):
    def __init__(self):
        super().__init__()

        # Window config
        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Tool
        )
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        self.setStyleSheet("background-color: black;")

        # Top bar dimensions
        screen_width = QApplication.primaryScreen().size().width()
        self.setGeometry(0, 0, screen_width, 26)

        # Layout
        layout = QHBoxLayout()
        layout.setContentsMargins(12, 2, 12, 2)
        layout.setSpacing(90)
        # Labels
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

        # Apply color
        for lbl in [self.battery_label, self.ram_label, self.gpu_label, self.cpu_label,
                    self.app_label, self.date_label, self.time_label, self.timer_label,
                    self.mic_label, self.spotify_label]:
            lbl.setStyleSheet(f"color: {current_color}; font-size: 12px;")
            layout.addWidget(lbl)

        # Stretch to keep Spotify at right
        layout.addStretch(1)

        self.setLayout(layout)

        # Timer logic
        self.timer_running = False
        self.start_time = None
        self.seconds = 0

        # Spotify throttle
        self._last_spotify_time = 0.0
        self._spotify_interval = 1.0  # seconds

        # Key cooldown for color changes
        self.keys_down = set()

        # Update loop
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_overlay)
        self.update_timer.start(500)

        # Keyboard listener
        self.listener = keyboard.Listener(
            on_press=self.key_press,
            on_release=self.key_release
        )
        self.listener.start()

    # ─────────────────────────────────────────────
    # HOTKEYS
    # ─────────────────────────────────────────────
    def key_press(self, key):
        global current_color
        if hasattr(key, "vk"):
            vk = key.vk
        else:
            return

        if vk in self.keys_down:
            return
        self.keys_down.add(vk)

        if vk == 97:  # Numpad 1 → start/stop timer
            if not self.timer_running:
                self.start_time = time.time()
                self.timer_running = True
            else:
                self.timer_running = False
                self.start_time = None
                self.seconds = 0

        # Color hotkeys 7-9 / 4-6
        num = None
        if vk in (103, 104, 105):
            num = {103:7, 104:8, 105:9}[vk]
        elif vk in (100, 101, 102):
            num = {100:4, 101:5, 102:6}[vk]

        if num and num in COLORS:
            current_color = COLORS[num]
            self.apply_colors()

    def key_release(self, key):
        if hasattr(key, "vk"):
            self.keys_down.discard(key.vk)

    def apply_colors(self):
        style = f"color: {current_color}; font-size: 12px;"
        for lbl in [self.battery_label, self.ram_label, self.gpu_label, self.cpu_label,
                    self.app_label, self.date_label, self.time_label, self.timer_label,
                    self.mic_label, self.spotify_label]:
            lbl.setStyleSheet(style)

    # ─────────────────────────────────────────────
    # UPDATE LOOP
    # ─────────────────────────────────────────────
    def update_overlay(self):
        # Battery
        try:
            battery = psutil.sensors_battery()
            self.battery_label.setText(f"Battery: {battery.percent}%" if battery else "Battery: --%")
        except:
            self.battery_label.setText("Battery: --%")

        # RAM
        ram_percent = psutil.virtual_memory().percent
        self.ram_label.setText(f"RAM: {ram_percent}%")

        # GPU
        try:
            gpus = GPUtil.getGPUs()
            if gpus:
                gpu = gpus[0]
                self.gpu_label.setText(f"GPU: {gpu.load*100:.0f}% {gpu.temperature}°C")
            else:
                self.gpu_label.setText("GPU: N/A")
        except:
            self.gpu_label.setText("GPU: N/A")

        # CPU
        cpu_percent = psutil.cpu_percent(interval=None)
        self.cpu_label.setText(f"CPU: {cpu_percent}%")

        # Current App
        try:
            hwnd = win32gui.GetForegroundWindow()
            tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = ps.Process(pid)
            app_name = proc.name()
            self.app_label.setText(f"App: {app_name}")
        except:
            self.app_label.setText("App: —")

        # Date & Time
        now = datetime.datetime.now()
        self.date_label.setText(f"Date: {now.day:02}/{now.month:02}/{now.year}")
        self.time_label.setText(f"Time: {now.strftime('%I:%M %p')}")

        # Timer
        if self.timer_running and self.start_time is not None:
            self.seconds = int(time.time() - self.start_time)
        mins = self.seconds // 60
        secs = self.seconds % 60
        self.timer_label.setText(f"Timer: {mins:02}:{secs:02}")

        # Mic
        bars, percent = get_mic_level()
        mic_bar = "█" * bars + "░" * (10 - bars)
        self.mic_label.setText(f"Mic: {mic_bar} {percent}%")

        # Spotify (throttled)
        now_time = time.time()
        if now_time - self._last_spotify_time >= self._spotify_interval:
            self._last_spotify_time = now_time
            song = fetch_spotify_sync()
            self.spotify_label.setText(song or "Spotify: —")

# ─────────────────────────────────────────────
# RUN APP
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    overlay = Overlay()
    overlay.show()
    sys.exit(app.exec_())
