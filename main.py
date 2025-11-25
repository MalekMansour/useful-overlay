# overlay.py
import sys
import time
import asyncio
import numpy as np
import sounddevice as sd
import keyboard
import psutil

from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL

# Pycaw / audio
try:
    from pycaw.pycaw import IAudioEndpointVolume, IMMDeviceEnumerator, AudioUtilities
    PYCAW_OK = True
except Exception:
    PYCAW_OK = False

# winsdk for Windows Media Sessions (Spotify)
try:
    import winsdk.windows.media.control as wmc
    WINSKD_OK = True
except Exception:
    WINSKD_OK = False

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

# register hotkey
keyboard.add_hotkey("num 1", toggle_timer)

# ---------- VOLUME (robust, tries several ways) ----------
# If you want to force a device name match, change this or remove.
PREFERRED_OUTPUT_NAME = "Headphones (Realtek(R) Audio)".lower()

def get_volume_percent():
    """Try multiple methods to obtain the real output device master volume as 0-100."""
    # 1) Try to find a device by friendly name using pycaw enumerator (best)
    if PYCAW_OK:
        try:
            enumerator = IMMDeviceEnumerator()
            # EnumAudioEndpoints(eRender=0, DEVICE_STATE_ACTIVE=1) -> returns collection
            collection = enumerator.EnumAudioEndpoints(0, 1)  # 0=eRender, 1=active
            count = collection.GetCount()
            for i in range(count):
                dev = collection.Item(i)
                name = dev.GetId()  # fallback id
                # Try FriendlyName from property store
                try:
                    props = dev.OpenPropertyStore(0)
                    # PKEY_Device_FriendlyName guid: use 14 (property key 14) typical pattern, but easiest is try AudioUtilities wrapper:
                    # Use AudioUtilities.GetAllDevices fallback below if friendly name missing.
                    friendly = None
                    try:
                        # pycaw's AudioUtilities.GetDeviceById may not exist in older versions - wrap in try
                        friendly = AudioUtilities.GetDeviceById(dev.GetId()).FriendlyName
                    except Exception:
                        friendly = None
                except Exception:
                    friendly = None

                if friendly and PREFERRED_OUTPUT_NAME in friendly.lower():
                    # found Realtek device; get its volume
                    try:
                        interface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                        volume = cast(interface, POINTER(IAudioEndpointVolume))
                        val = volume.GetMasterVolumeLevelScalar()
                        return int(round(max(0.0, min(1.0, val)) * 100))
                    except Exception:
                        pass
            # 2) fallback: use default audio endpoint
            try:
                device = enumerator.GetDefaultAudioEndpoint(0, 1)
                interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                volume = cast(interface, POINTER(IAudioEndpointVolume))
                val = volume.GetMasterVolumeLevelScalar()
                return int(round(max(0.0, min(1.0, val)) * 100))
            except Exception:
                pass
        except Exception:
            pass

    # 3) Try AudioUtilities.GetSpeakers() fallback
    if PYCAW_OK:
        try:
            dev = AudioUtilities.GetSpeakers()
            interface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            val = volume.GetMasterVolumeLevelScalar()
            return int(round(max(0.0, min(1.0, val)) * 100))
        except Exception:
            pass

    # 4) As a last resort, return 0
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
        duration = 0.04  # seconds (short)
        frames = int(44100 * duration)
        audio = sd.rec(frames, samplerate=44100, channels=1, dtype='float32', blocking=True)
        rms = float(np.sqrt(np.mean(np.square(audio))))
        # Map RMS to 0..1 with a scale factor calibrated for typical USB mics.
        scaled = rms * 180.0   # tweak this number if your mic is always too low/high
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

# ---------- SPOTIFY (Windows Media Session) ----------
async def _get_spotify_async():
    try:
        sessions = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        current = sessions.get_current_session()
        if current is None:
            return "No Spotify"

        # Ensure source is Spotify
        source = current.source_app_user_model_id or ""
        if "spotify" not in source.lower():
            return "No Spotify"

        props = await current.try_get_media_properties_async()
        title = getattr(props, "title", "") or ""
        artist = getattr(props, "artist", "") or ""
        if title == "" and artist == "":
            return "No Spotify"
        display = f"{artist} - {title}" if artist else title
        return display
    except Exception:
        return "No Spotify"

def get_spotify_now():
    """Sync wrapper for the async winsdk call; returns a string."""
    if not WINSKD_OK:
        return "No Spotify"
    try:
        return asyncio.run(_get_spotify_async())
    except Exception:
        return "No Spotify"

# ---------- UI: Top solid black bar ----------
class OverlayBar(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Tool
        )
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        self.setStyleSheet("background-color: black;")

        screen_width = QApplication.primaryScreen().size().width()
        bar_height = 36
        self.setGeometry(0, 0, screen_width, bar_height)

        layout = QHBoxLayout()
        layout.setContentsMargins(12, 4, 12, 4)
        layout.setSpacing(28)

        font = QFont("Segoe UI", 11)

        self.battery_label = QLabel("Battery: --%")
        self.battery_label.setFont(font)
        self.battery_label.setStyleSheet("color: white;")

        self.volume_label = QLabel("Vol: --%")
        self.volume_label.setFont(font)
        self.volume_label.setStyleSheet("color: white;")

        self.mic_label = QLabel("Mic: [-----] 0%")
        self.mic_label.setFont(font)
        self.mic_label.setStyleSheet("color: white;")

        self.timer_label = QLabel("Timer: 0s")
        self.timer_label.setFont(font)
        self.timer_label.setStyleSheet("color: white;")

        self.spotify_label = QLabel("Spotify: —")
        self.spotify_label.setFont(font)
        self.spotify_label.setStyleSheet("color: white;")

        # Add in order you want to see them
        layout.addWidget(self.battery_label)
        layout.addWidget(self.volume_label)
        layout.addWidget(self.mic_label)
        layout.addWidget(self.timer_label)
        layout.addWidget(self.spotify_label)

        self.setLayout(layout)

        # Update loop
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_info)
        self.update_timer.start(300)  # refresh rate (ms)

    def update_info(self):
        # Battery
        self.battery_label.setText(f"Battery: {get_battery_percent()}%")

        # Volume (robust)
        vol = get_volume_percent()
        self.volume_label.setText(f"Vol: {vol}%")

        # Mic
        mic_pct = get_mic_rms_percent()
        bar, pct = mic_bar_and_percent(mic_pct)
        self.mic_label.setText(f"Mic: [{bar}] {pct}%")

        # Timer
        if timer_running:
            sec = int(time.time() - timer_start)
        else:
            sec = 0
        self.timer_label.setText(f"Timer: {sec}s")

        # Spotify (non-blocking-ish)
        # To avoid calling winsdk too often we call only every ~1 second
        now = int(time.time() * 10)  # decisecond buckets
        if getattr(self, "_last_spotify_bucket", None) != now // 10:
            # update label
            sp = get_spotify_now()
            self.spotify_label.setText(f"Spotify: {sp}")
            self._last_spotify_bucket = now // 10


if __name__ == "__main__":
    app = QApplication(sys.argv)
    overlay = OverlayBar()
    overlay.show()
    sys.exit(app.exec())
