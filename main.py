import sys
import psutil
import time
import pythoncom
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QHBoxLayout
from PyQt5.QtCore import Qt, QTimer
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CoInitialize

# --------------------------
# REALTEK VOLUME (WORKS 100%)
# --------------------------
def get_volume_realtek():
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        return int(volume.GetMasterVolumeLevelScalar() * 100)
    except:
        return -1


# --------------------------
# MICROPHONE INPUT BAR
# --------------------------
import sounddevice as sd
import numpy as np

def get_mic_level():
    try:
        duration = 0.05
        sample_rate = 44100
        audio = sd.rec(int(duration * sample_rate), samplerate=sample_rate, channels=1, dtype='float32')
        sd.wait()
        level = np.abs(audio).mean() * 10
        level = min(level, 1.0)
        bars = int(level * 5)
        return bars, int(level * 100)
    except:
        return 0, 0


# --------------------------
# SPOTIFY (WINDOWS GLOBAL MEDIA)
# --------------------------
import winsdk.windows.media.control as wmc
import asyncio

async def get_spotify():
    try:
        sessions = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        current = sessions.get_current_session()
        if not current:
            return "No Music"
        info = await current.try_get_media_properties_async()
        artist = info.artist
        title = info.title
        return f"{artist} – {title}" if artist and title else "No Music"
    except:
        return "No Music"

def get_spotify_sync():
    return asyncio.run(get_spotify())


# --------------------------
# TIMER
# --------------------------
timer_running = False
timer_start = 0

def toggle_timer():
    global timer_running, timer_start
    if timer_running:
        timer_running = False
    else:
        timer_start = time.time()
        timer_running = True

def get_timer():
    if timer_running:
        elapsed = int(time.time() - timer_start)
        return f"{elapsed}s"
    return "Stopped"


# --------------------------
# OVERLAY UI
# --------------------------
class Overlay(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Tool
        )

        self.setAttribute(Qt.WA_TranslucentBackground, False)

        # --------------------------
        # SMALL BAR (20px height)
        # --------------------------
        self.setStyleSheet("background-color: black;")
        self.setGeometry(0, 0, 1920, 24)

        layout = QHBoxLayout()
        layout.setContentsMargins(6, 0, 6, 0)
        layout.setSpacing(20)

        font = "font-size: 11px; color: white;"

        self.volume_label = QLabel("Vol: --%")
        self.volume_label.setStyleSheet(font)

        self.mic_label = QLabel("Mic: ░░░░░ (0%)")
        self.mic_label.setStyleSheet(font)

        self.battery_label = QLabel("Battery: --%")
        self.battery_label.setStyleSheet(font)

        self.timer_label = QLabel("Timer: Stopped")
        self.timer_label.setStyleSheet(font)

        self.song_label = QLabel("Song: No Music")
        self.song_label.setStyleSheet(font)

        layout.addWidget(self.volume_label)
        layout.addWidget(self.mic_label)
        layout.addWidget(self.battery_label)
        layout.addWidget(self.timer_label)
        layout.addWidget(self.song_label)

        self.setLayout(layout)

        timer = QTimer(self)
        timer.timeout.connect(self.update_overlay)
        timer.start(200)

    def update_overlay(self):
        pythoncom.CoInitialize()

        # volume
        vol = get_volume_realtek()
        self.volume_label.setText(f"Vol: {vol}%")

        # mic
        bars, percent = get_mic_level()
        bar_str = "█" * bars + "░" * (5 - bars)
        self.mic_label.setText(f"Mic: {bar_str} ({percent}%)")

        # battery
        battery = psutil.sensors_battery()
        self.battery_label.setText(f"Battery: {battery.percent}%")

        # timer
        self.timer_label.setText(f"Timer: {get_timer()}")

        # song
        song = get_spotify_sync()
        self.song_label.setText(f"Song: {song}")


# --------------------------
# HOTKEY: NUMPAD 1 → TOGGLE TIMER
# --------------------------
from pynput import keyboard

def on_press(key):
    try:
        if key == keyboard.Key.num_pad1:
            toggle_timer()
    except:
        pass

listener = keyboard.Listener(on_press=on_press)
listener.start()


# --------------------------
# RUN APP
# --------------------------
app = QApplication(sys.argv)
overlay = Overlay()
overlay.show()
sys.exit(app.exec_())
