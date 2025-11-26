import sys
import asyncio
import pythoncom
import psutil
import time
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QHBoxLayout
import winsdk.windows.media.control as wmc
from pycaw.pycaw import AudioUtilities, IAudioMeterInformation
from comtypes import CLSCTX_ALL
from ctypes import POINTER, cast
from pynput import keyboard

# ─────────────────────────────────────────────
# COLOR PRESETS (NumPad)
# ─────────────────────────────────────────────
COLORS = {
    7: "white",
    8: "#ff66cc",   # Pink
    9: "red",
    4: "#66ff66",   # Light Green
    5: "#00eaff",   # Cyan
    6: "#b266ff"    # Purple
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
            return ""
        info = await current.try_get_media_properties_async()
        source = (current.source_app_user_model_id or "").lower()
        if "spotify" not in source:
            return ""
        title = info.title or ""
        artist = ", ".join((info.artist or "").split(";"))
        return f"{artist} – {title}" if (artist or title) else ""
    except:
        return ""

def fetch_spotify_sync():
    try:
        return asyncio.run(spotify_now_playing())
    except:
        return ""

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
        layout.setSpacing(30)

        # Battery
        self.battery_label = QLabel("Batt: --%")
        self.battery_label.setStyleSheet(f"color: {current_color}; font-size: 13px;")
        layout.addWidget(self.battery_label)

        # Mic level
        self.mic_label = QLabel("Mic: ░░░░░ 0% (No Mic)")
        self.mic_label.setStyleSheet(f"color: {current_color}; font-size: 13px;")
        layout.addWidget(self.mic_label)

        # Timer
        self.timer_label = QLabel("Timer: 00:00")
        self.timer_label.setStyleSheet(f"color: {current_color}; font-size: 13px;")
        layout.addWidget(self.timer_label)

        # Spacer then Spotify right aligned
        layout.addStretch(1)

        # Spotify
        self.spotify_label = QLabel("")
        self.spotify_label.setStyleSheet(f"color: {current_color}; font-size: 13px;")
        layout.addWidget(self.spotify_label)

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

        # Initialize devices
        self.mic_device, self.mic_name = self.get_mic_device()

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

        # Prevent repeated triggers
        if vk in self.keys_down:
            return
        self.keys_down.add(vk)

        # Numpad 1 → start/stop & reset timer
        if vk == 97:
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
        style = f"color: {current_color}; font-size: 13px;"
        self.battery_label.setStyleSheet(style)
        self.mic_label.setStyleSheet(style)
        self.timer_label.setStyleSheet(style)
        self.spotify_label.setStyleSheet(style)

    # ─────────────────────────────────────────────
    # MIC FUNCTIONS
    # ─────────────────────────────────────────────
    def get_mic_device(self):
        try:
            devices = AudioUtilities.GetMicrophone()
            if not devices:
                return None, "No Mic"
            mic = devices[0]  # first mic (default Discord mic)
            return mic, mic.FriendlyName
        except:
            return None, "No Mic"

    def get_mic_level(self):
        try:
            if not self.mic_device:
                return 0
            interface = self.mic_device.Activate(
                IAudioMeterInformation._iid_, CLSCTX_ALL, None
            )
            meter = cast(interface, POINTER(IAudioMeterInformation))
            level = meter.GetPeakValue() * 100
            return int(level)
        except:
            return 0

    # ─────────────────────────────────────────────
    # UPDATE LOOP
    # ─────────────────────────────────────────────
    def update_overlay(self):
        # Battery
        try:
            battery = psutil.sensors_battery()
            batt_text = f"Batt: {battery.percent}%" if battery else "Batt: --%"
        except:
            batt_text = "Batt: --%"
        self.battery_label.setText(batt_text)

        # Mic
        mic_val = self.get_mic_level()
        bars = int(mic_val / 20)
        mic_bar = "█" * bars + "░" * (5 - bars)
        self.mic_label.setText(f"Mic: {mic_bar} {mic_val}% ({self.mic_name})")

        # Timer
        if self.timer_running and self.start_time is not None:
            self.seconds = int(time.time() - self.start_time)
        mins = self.seconds // 60
        secs = self.seconds % 60
        self.timer_label.setText(f"Timer: {mins:02}:{secs:02}")

        # Spotify (throttled)
        now = time.time()
        if now - self._last_spotify_time >= self._spotify_interval:
            self._last_spotify_time = now
            song = fetch_spotify_sync()
            self.spotify_label.setText(song or "")

# ─────────────────────────────────────────────
# RUN APP
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    overlay = Overlay()
    overlay.show()
    sys.exit(app.exec_())
