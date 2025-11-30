# UsefulOverlay

A lightweight real-time overlay that shows useful system information, media playback, microphone activity, and more, all in a clean bar always visible on screen.

## Features

### Live Microphone Level
- Shows your current microphone volume in real time  
- Helps instantly see if you're muted or if your mic disconnected  
- Smooth attack/release visualizer  

### Media Now-Playing Display
- Shows the current song or video you're listening to  
- Works with **Spotify, YouTube, YouTube Music, Twitch, SoundCloud**, and anything using Windows Media Sessions  
- Automatically switches between apps depending on what’s playing

### 🖥 System Stats
- CPU usage  
- RAM usage  
- GPU usage (if available — laptops may show N/A)  
- Temperature reading supported on desktop PCs

### ⏰ Time & Date
Clean digital time + formatted date.

### ⏱ Built-in Timer
Simple, always-visible countdown timer for games or tasks.

### 🪟 Focused Window Detection
Displays the name of the currently focused application instantly.

### 🔢 Numpad Controls
All controls use numpad keys so you don’t interfere with typing:
- **Numpad 8** → Cycle through color themes  
- **Numpad 2** → Switch microphones  
- More custom keybinds can be added

### 🎨 Custom Themes
- Multiple color themes  
- Smooth transitions  
- No flickering or glitches even when switching rapidly

---

## 🧩 Technologies Used
- Python  
- Tkinter (UI)  
- psutil  
- sounddevice  
- numpy  
- asyncio  
- windows-media-controller  
- pyinstaller (for building .exe)
