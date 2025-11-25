import sys
import time
import numpy as np
import sounddevice as sd
import keyboard
import psutil

from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL

from pycaw.pycaw import IAudioEndpointVolume, IMMDeviceEnumerator
from PySide6.QtWidgets import QApplication, QWidget, QHBoxLayout, QLabel
from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QFont

# ---------- TIMER (NumPad 1) ----------
timer_running = False
timer_start = 0

def toggle_timer():
    global timer_running, timer_start
    if not timer_running:
        timer_running = True
        timer_start = time.time()
    else:
        timer_running = False

# register hotkey; this runs in background thread inside keyboard lib
keyboard.add_hotkey("num 1", toggle_timer)

# ---------- VOLUME (robust) ----------
def get_volume_percent():
    try:
        enumerator = IMMDeviceEnumerator()
        device = enumerator.GetDefaultAudioEndpoint(0, 1)  # eRender, eMultimedia
        interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        vol = volume.GetMasterVolumeLevelScalar()  # 0.0 - 1.0
        return int(max(0, min(100, round(vol * 100))))
    except Exception:
        return 0

# ---------- BATTERY ----------
def get_battery_percent():
    try:
        batt = psutil.sensors_battery()
        return int(batt.percent) if batt is not None else 0
    except Exception:
        return 0

# ---------- MICROPHONE LEVEL (RMS -> percent -> 5-bars) ----------
def get_mic_rms_percent():
    try:
        # Read a short buffer (non-blocking style with blocking=True for simplicity)
        duration = 0.04  # seconds
        frames = int(44100 * duration)
        audio = sd.rec(frames, samplerate=44100, channels=1, dtype='float32', blocking=True)
        rms = float(np.sqrt(np.mean(np.square(audio))))
        # map RMS to 0-1 using an empirically chosen scale factor
        # tweak 70.0 if values are too low/high on your mic
        scaled = rms * 200.0
        scaled = max(0.0, min(1.0, scaled))
        percent = int(round(scaled * 100))
        return percent
    except Exception:
        return 0

def mic_bar_and_percent(percent):
    bars = int((percent / 100) * 5)
    bars = max(0, min(5, bars))
    bar_str = "█" * bars + "-" * (5 - bars)
    return bar_str, percent

# ---------- OVERLAY (solid black top bar) ----------
class OverlayBar(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Tool
        )
        # solid black background (non-transparent)
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        self.setStyleSheet("background-color: black;")

        # full width, fixed height
        screen_width = QApplication.primaryScreen().size().width()
        bar_height = 36
        self.setGeometry(0, 0, screen_width, bar_height)

        # layout
        layout = QHBoxLayout()
        layout.setContentsMargins(12, 4, 12, 4)
        layout.setSpacing(28)

        font = QFont("Segoe UI", 11)

        # labels
        self.volume_label = QLabel("Vol: --%")
        self.volume_label.setFont(font)
        self.volume_label.setStyleSheet("color: white;")

        self.battery_label = QLabel("Battery: --%")
        self.battery_label.setFont(font)
        self.battery_label.setStyleSheet("color: white;")

        self.mic_label = QLabel("Mic: [-----] 0%")
        self.mic_label.setFont(font)
        self.mic_label.setStyleSheet("color: white;")

        self.timer_label = QLabel("Timer: 0s")
        self.timer_label.setFont(font)
        self.timer_label.setStyleSheet("color: white;")

        layout.addWidget(self.battery_label)
        layout.addWidget(self.volume_label)
        layout.addWidget(self.mic_label)
        layout.addWidget(self.timer_label)

        self.setLayout(layout)

        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_info)
        self.update_timer.start(250)  # update every 0.25s

    def update_info(self):
        # volume
        vol = get_volume_percent()
        self.volume_label.setText(f"Vol: {vol}%")

        # battery
        batt = get_battery_percent()
        self.battery_label.setText(f"Battery: {batt}%")

        # mic
        mic_pct = get_mic_rms_percent()
        bar, pct = mic_bar_and_percent(mic_pct)
        self.mic_label.setText(f"Mic: [{bar}] {pct}%")

        # timer
        if timer_running:
            sec = int(time.time() - timer_start)
        else:
            sec = 0
        self.timer_label.setText(f"Timer: {sec}s")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    overlay = OverlayBar()
    overlay.show()
    sys.exit(app.exec())
