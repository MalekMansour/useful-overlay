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

# optional GPUtil
try:
    import GPUtil
    gpu_available = True
except:
    gpu_available = False

import os
import traceback

# ─────────────────────────────────────────────
# COLORS (order requested)
# ─────────────────────────────────────────────
COLOR_CYCLE = [
    "white",
    "#e73535",  # red
    "#8de6ee",  # cyan
    "#9ef39e",  # green
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
    "mic_bars": 0,
    "mic_percent": 0,
    "spotify": "Spotify: —"
}
_worker_stop = False

# ─────────────────────────────────────────────
# Timer state (controlled by numpad 7)
# ─────────────────────────────────────────────
timer_lock = threading.Lock()
_timer_running = False   # True when actively counting
_timer_offset = 0.0      # seconds accumulated while previously running
_timer_start_time = None # wall-clock when current run began

def timer_start():
    global _timer_running, _timer_start_time
    with timer_lock:
        if not _timer_running:
            _timer_start_time = time.time()
            _timer_running = True

def timer_pause():
    global _timer_running, _timer_offset, _timer_start_time
    with timer_lock:
        if _timer_running:
            _timer_offset += time.time() - (_timer_start_time or time.time())
            _timer_running = False
            _timer_start_time = None

def timer_reset_and_start():
    global _timer_running, _timer_offset, _timer_start_time
    with timer_lock:
        _timer_offset = 0.0
        _timer_start_time = time.time()
        _timer_running = True

def timer_get_seconds_int():
    with timer_lock:
        total = _timer_offset
        if _timer_running and _timer_start_time is not None:
            total += time.time() - _timer_start_time
        return int(total)

# ─────────────────────────────────────────────
# Spotify helper (permissive)
# ─────────────────────────────────────────────
async def _spotify_get_async():
    sessions = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
    current = sessions.get_current_session()
    if current is None:
        return None
    try:
        props = await current.try_get_media_properties_async()
    except Exception:
        return None
    artist = getattr(props, "artist", "") or ""
    title = getattr(props, "title", "") or ""
    if not artist and not title:
        album = getattr(props, "album", "") or ""
        albumartist = getattr(props, "albumArtist", "") or ""
        if album and not title:
            title = album
        if albumartist and not artist:
            artist = albumartist
    if not artist and not title:
        return None
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
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            return asyncio.run(_spotify_get_async())
        except Exception:
            # fallback loop
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                return loop.run_until_complete(_spotify_get_async())
            finally:
                try:
                    loop.close()
                except:
                    pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return None

# Mic level
MIC_NOISE_FLOOR = 0.005      
MIC_ATTACK = 0.9
MIC_RELEASE = 0.4
_smoothed_level = 0.0

def get_mic_level_blocking():
    global _smoothed_level

    try:
        duration = 0.04
        sr = 16000
        audio = sd.rec(int(duration * sr), samplerate=sr, channels=1,
                       dtype='float32', blocking=True)

        if audio is None or audio.size == 0:
            return 0, 0

        rms = float(np.sqrt(np.mean(audio**2)))

        # Silence
        if rms < MIC_NOISE_FLOOR:
            target = 0.0
        else:
            # EXTREME BOOST
            boosted = rms * 100.0   

            # Soft compression so it doesn't instantly hit 100%
            compressed = boosted / (1 + boosted)

            target = max(0.0, min(compressed, 1.0))

        # Smooth attack/release
        if target > _smoothed_level:
            _smoothed_level = (
                _smoothed_level * (1 - MIC_ATTACK)
                + target * MIC_ATTACK
            )
        else:
            _smoothed_level = (
                _smoothed_level * (1 - MIC_RELEASE)
                + target * MIC_RELEASE
            )

        percent = int(_smoothed_level * 100)
        bars = int(_smoothed_level * 10)

        return bars, percent

    except Exception:
        return 0, 0

# Worker thread: collects sensors and spotify
def stats_worker_loop(spotify_interval=1.0):
    global _worker_stop
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    psutil.cpu_percent(interval=None) 
    last_spotify = 0.0
    spotify_cache = {"text": "Spotify: —", "ts": 0.0}
    SPOTIFY_TTL = 3.0

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
                ram_text = f"RAM: {psutil.virtual_memory().percent}%"
            except Exception:
                ram_text = "RAM: --%"

            # CPU
            try:
                cpu_text = f"CPU: {psutil.cpu_percent(interval=0.05)}%"
            except Exception:
                cpu_text = "CPU: --%"

            # GPU
            try:
                if gpu_available:
                    gpus = GPUtil.getGPUs()
                    if gpus:
                        g = gpus[0]
                        gpu_text = f"GPU: {g.load*100:.0f}%"
                    else:
                        gpu_text = "GPU: N/A"
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

            # Spotify 
            now_t = time.time()
            sp = None
            if now_t - last_spotify >= spotify_interval:
                try:
                    sp = spotify_now_playing()
                except Exception as e:
                    sp = None
                last_spotify = now_t
            else:
                sp = None

            # cache logic
            spotify_to_write = spotify_cache["text"]
            if sp and sp.strip():
                spotify_cache["text"] = sp
                spotify_cache["ts"] = time.time()
                spotify_to_write = sp
            else:
                if time.time() - spotify_cache["ts"] > SPOTIFY_TTL:
                    spotify_to_write = "Spotify: —"

            # write all stats
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
                stats["spotify"] = spotify_to_write

        except Exception:
            traceback.print_exc()

        time.sleep(0.05)

    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

# Overlay UI
class Overlay(QWidget):
    def __init__(self):
        super().__init__()
        global current_color

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setStyleSheet("background-color: black;")

        try:
            scr_w = QApplication.primaryScreen().size().width()
        except Exception:
            scr_w = 800
        self.setGeometry(0, 0, scr_w, 26)

        layout = QHBoxLayout()
        layout.setContentsMargins(8, 2, 8, 2)
        layout.setSpacing(90)

        self.battery_label = QLabel()
        self.ram_label = QLabel()
        self.gpu_label = QLabel()
        self.cpu_label = QLabel()
        self.app_label = QLabel()
        self.date_label = QLabel()
        self.time_label = QLabel()
        self.timer_label = QLabel()
        self.mic_label = QLabel()
        self.spotify_label = QLabel()

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

        # UI update timer
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_overlay)
        self.update_timer.start(200)

        # Start worker thread
        self.worker_thread = threading.Thread(target=stats_worker_loop, daemon=True)
        self.worker_thread.start()

        # Hotkeys listener
        self.keys_down = set()
        self.listener = keyboard.Listener(on_press=self.key_press, on_release=self.key_release)
        self.listener.start()

    # hotkey handling
    def key_press(self, key):
        global color_index, current_color, _worker_stop

        if not hasattr(key, "vk"):
            return
        vk = key.vk

        if vk in self.keys_down:
            return
        self.keys_down.add(vk)

        # Numpad 7 -> timer cycle
        if vk == 103:
            secs = timer_get_seconds_int()
            with timer_lock:
                running = _timer_running
            if not running and secs == 0:
                timer_start()
            elif running:
                timer_pause()
            else:
                timer_reset_and_start()

        # Numpad 8 -> color cycle
        if vk == 104:
            color_index = (color_index + 1) % len(COLOR_CYCLE)
            current_color = COLOR_CYCLE[color_index]
            self.apply_colors()

        # Numpad 9 -> full restart
        if vk == 105:
            try:
                _worker_stop = True
                time.sleep(0.05)
            except Exception:
                pass
            os.execv(sys.executable, [sys.executable] + sys.argv)

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

        # timer display 
        secs = timer_get_seconds_int()
        self.timer_label.setText(f"Timer: {secs:03d}")

        bars = max(0, min(10, mic_bars))
        mic_bar = "█" * bars + "░" * (10 - bars)
        self.mic_label.setText(f"Mic: {mic_bar} {mic_percent}%")

        self.spotify_label.setText(spotify_text)

# RUN
if __name__ == "__main__":
    app = QApplication(sys.argv)
    overlay = Overlay()
    overlay.show()
    try:
        print("Overlay started. Use Numpad 7 to control timer, 8 to change color, 9 to restart.")
        sys.exit(app.exec_())
    finally:
        _worker_stop = True
