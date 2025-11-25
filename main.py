import sys
import asyncio
import pythoncom
import psutil
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QHBoxLayout
import winsdk.windows.media.control as wmc
import win32api
from pynput import keyboard
import time

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
    # CoInitialize for COM (winsdk)
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

# synchronous wrapper (safe to call from PyQt thread)
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

        # Smaller bar height
        screen_width = QApplication.primaryScreen().size().width()
        self.setGeometry(0, 0, screen_width, 26)

        # Layout
        layout = QHBoxLayout()
        layout.setContentsMargins(12, 2, 12, 2)
        layout.setSpacing(35)

        # Battery
        self.battery_label = QLabel("Batt: --%")
        self.battery_label.setStyleSheet(f"color: {current_color}; font-size: 13px;")
        layout.addWidget(self.battery_label)

        # Mic level (kept exactly as you had it)
        self.mic_label = QLabel("Mic: ▄▄▄▄▄ 0%")
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
        self.seconds = 0

        # Spotify throttle
        self._last_spotify_time = 0.0
        self._spotify_interval = 1.0  # seconds

        # Update loop
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_overlay)
        self.update_timer.start(500)

        # Keyboard listener (pynput)
        self.listener = keyboard.Listener(on_press=self.key_press)
        self.listener.start()

    # ─────────────────────────────────────────────
    # HOTKEY LISTENER (NUMPAD)
    # ─────────────────────────────────────────────
    def key_press(self, key):
        global current_color

        # Convert key to number (pynput KeyCode has .vk on Windows)
        if hasattr(key, "vk"):
            vk = key.vk
        else:
            return

        # NumPad 1 → toggle timer
        # Common virtual-key for NumPad1 is 97
        if vk == 97:
            self.timer_running = not self.timer_running
            # do not reset seconds when toggling; preserve elapsed (same behaviour as before)

        # Color Hotkeys: NumPad 7-8-9 and 4-5-6
        num = None
        if vk in (103, 104, 105):  # numpad 7,8,9
            num = {103: 7, 104: 8, 105: 9}[vk]
        elif vk in (100, 101, 102):  # numpad 4,5,6
            num = {100: 4, 101: 5, 102: 6}[vk]

        if num and num in COLORS:
            current_color = COLORS[num]
            self.apply_colors()

    # Apply chosen color to all labels
    def apply_colors(self):
        style = f"color: {current_color}; font-size: 13px;"
        self.battery_label.setStyleSheet(style)
        self.mic_label.setStyleSheet(style)
        self.timer_label.setStyleSheet(style)
        self.spotify_label.setStyleSheet(style)

    # ─────────────────────────────────────────────
    # UPDATE LOOP
    # ─────────────────────────────────────────────
    def update_overlay(self):
        # Battery
        try:
            battery = psutil.sensors_battery()
            batt_text = f"Batt: {battery.percent}%" if battery is not None else "Batt: --%"
        except:
            batt_text = "Batt: --%"
        self.battery_label.setText(batt_text)

        # Mic (kept exactly the same simulation you had: GetAsyncKeyState)
        try:
            mic_level = win32api.GetAsyncKeyState(0x41)  # placeholder method you had
            mic_val = abs(mic_level) % 100
            bars = int(mic_val / 20)
            mic_bar = "█" * bars + "░" * (5 - bars)
            self.mic_label.setText(f"Mic: {mic_bar} {mic_val}%")
        except:
            # fallback if win32api fails
            self.mic_label.setText("Mic: ░░░░░ 0%")

        # Timer
        if self.timer_running:
            self.seconds += 1
        mins = self.seconds // 60
        secs = self.seconds % 60
        self.timer_label.setText(f"Timer: {mins:02}:{secs:02}")

        # Spotify (throttled to once per _spotify_interval seconds)
        now = time.time()
        if now - self._last_spotify_time >= self._spotify_interval:
            self._last_spotify_time = now
            song = fetch_spotify_sync()
            # Only set if non-empty; otherwise keep previous or blank
            self.spotify_label.setText(song or "")


# ─────────────────────────────────────────────
# RUN APP
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    overlay = Overlay()
    overlay.show()
    sys.exit(app.exec_())
