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
# SPOTIFY (robust & permissive)
# ─────────────────────────────────────────────
# Cache last good string, and timestamp it
_spotify_cache = {"text": "Spotify: —", "ts": 0.0}
_SPOTIFY_CACHE_TTL = 3.0  # seconds to keep last good value on transient failure

async def _spotify_get_async():
    """
    Async helper that talks to winsdk and returns a string or None on no useful data.
    This runs inside asyncio.run(...) where called.
    """
    # Important: do *not* rely on source_app_user_model_id being 'spotify' exactly.
    sessions = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
    current = sessions.get_current_session()
    if current is None:
        return None

    try:
        props = await current.try_get_media_properties_async()
    except Exception:
        return None

    # Try artist/title first (most common)
    artist = getattr(props, "artist", "") or ""
    title = getattr(props, "title", "") or ""

    # Sometimes other fields exist; try album or albumartist fallback
    if not artist and not title:
        # Some implementations use 'album' or other properties; attempt safe attribute reads
        album = getattr(props, "album", "") or ""
        albumartist = getattr(props, "albumArtist", "") or getattr(props, "albumArtist", None) or ""
        # If album has something, use it as 'title' fallback
        if album and not title:
            title = album
        if albumartist and not artist:
            artist = albumartist

    # If still nothing useful, return None
    if not artist and not title:
        return None

    # Return a nice string
    a = artist.strip()
    t = title.strip()
    if a and t:
        return f"{a} – {t}"
    elif t:
        return f"{t}"
    elif a:
        return f"{a}"
    return None

def spotify_now_playing():
    try:
        # Ensure COM in this thread (safe to call multiple times)
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        # asyncio.run the async helper; safe here because it's called in worker thread
        try:
            result = asyncio.run(_spotify_get_async())
        except Exception as e:
            # Sometimes asyncio.run will fail if an event loop is already running (rare in worker thread)
            # Fallback: create new loop manually
            try:
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                result = loop.run_until_complete(_spotify_get_async())
            except Exception:
                result = None
        return result
    finally:
        # CoUninitialize to be tidy (ignore failures)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

# ─────────────────────────────────────────────
# MIC LEVEL
# ─────────────────────────────────────────────
def get_mic_level_blocking():
    try:
        duration = 0.05
        sample_rate = 44100
        # guard: if sounddevice throws due to no input device, return zeros
        audio = sd.rec(int(duration * sample_rate), samplerate=sample_rate,
                       channels=1, dtype='float32', blocking=True)
        if audio is None or audio.size == 0:
            return 0, 0
        rms = float(np.sqrt(np.mean(np.square(audio))))
        scaled = min(max(rms * 180.0, 0.0), 1.0)
        percent = int(scaled * 100)
        bars = int((percent / 100) * 10)
        return bars, percent
    except Exception as e:
        # Print for debugging but don't crash
        # (If you want less spam, remove this print later)
        print("MIC Error:", e)
        return 0, 0

# ─────────────────────────────────────────────
# WORKER THREAD
# ─────────────────────────────────────────────
def stats_worker_loop(spotify_interval=1.0):
    # Initialize COM on this worker thread
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    last_spotify = 0.0
    global _worker_stop, _spotify_cache

    psutil.cpu_percent(interval=None)  # warmup

    while not _worker_stop:
        try:
            # Battery
            try:
                battery = psutil.sensors_battery()
                batt = f"Battery: {battery.percent}%" if battery else "Battery: --%"
            except Exception:
                batt = "Battery: --%"

            # RAM
            try:
                ram = f"RAM: {psutil.virtual_memory().percent}%"
            except Exception:
                ram = "RAM: --%"

            # CPU
            try:
                cpu = f"CPU: {psutil.cpu_percent(interval=0.05)}%"
            except Exception:
                cpu = "CPU: --%"

            # GPU
            try:
                if gpu_available:
                    gpus = GPUtil.getGPUs()
                    if gpus:
                        g = gpus[0]
                        gpu = f"GPU: {g.load*100:.0f}%"
                    else:
                        gpu = "GPU: N/A"
                else:
                    gpu = "GPU: N/A"
            except Exception:
                gpu = "GPU: N/A"

            # Foreground app
            try:
                hwnd = win32gui.GetForegroundWindow()
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                proc = ps.Process(pid)
                app = f"App: {proc.name()}"
            except Exception:
                app = "App: —"

            # Date / Time
            now = datetime.datetime.now()
            date = f"Date: {now.day:02}/{now.month:02}/{now.year}"
            tim = f"Time: {now.strftime('%I:%M %p')}"

            # Mic
            mic_bars, mic_percent = get_mic_level_blocking()

            # Spotify (refresh every spotify_interval seconds)
            now_t = time.time()
            if now_t - last_spotify >= spotify_interval:
                try:
                    sp = spotify_now_playing()  # returns string or None
                except Exception as e:
                    print("Spotify worker error:", e)
                    sp = None
                last_spotify = now_t
            else:
                with stats_lock:
                    sp = stats.get("spotify", None)

            # Use cache logic:
            if sp and sp.strip():
                # got new valid text; store it
                _spotify_cache["text"] = sp
                _spotify_cache["ts"] = time.time()
                spotify_to_write = sp
            else:
                # no new text; if cache recent, use it; else show default
                if time.time() - _spotify_cache["ts"] <= _SPOTIFY_CACHE_TTL and _spotify_cache["text"]:
                    spotify_to_write = _spotify_cache["text"]
                else:
                    spotify_to_write = "Spotify: —"

            # Write stats
            with stats_lock:
                stats["battery"] = batt
                stats["ram"] = ram
                stats["gpu"] = gpu
                stats["cpu"] = cpu
                stats["app"] = app
                stats["date"] = date
                stats["time"] = tim
                stats["mic_bars"] = mic_bars
                stats["mic_percent"] = mic_percent
                stats["spotify"] = spotify_to_write

        except Exception:
            # Print stack so we can see if something keeps failing
            import traceback
            traceback.print_exc()

        time.sleep(0.05)

    # Uninitialize COM for thread cleanup (best-effort)
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

# ─────────────────────────────────────────────
# OVERLAY UI
# ─────────────────────────────────────────────
class Overlay(QWidget):
    def __init__(self):
        super().__init__()
        global current_color

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setStyleSheet("background-color: black;")

        # safe screen width access
        scr_w = 800
        try:
            scr_w = QApplication.primaryScreen().size().width()
        except Exception:
            pass

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
            lbl.setStyleSheet(f"color: {current_color}; font-size: 12px;")
            layout.addWidget(lbl)

        layout.addStretch(1)
        self.setLayout(layout)

        # Update timer
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_overlay)
        self.timer.start(200)

        # Start worker thread
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
        print("Overlay started — watch the console for Spotify debug.")
        sys.exit(app.exec_())
    finally:
        _worker_stop = True
