# FetchPro

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue?style=for-the-badge&logo=python)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/moshepinhasi/fetchpro?style=for-the-badge)](https://github.com/moshepinhasi/fetchpro)

> 🚀 **Enterprise-grade download manager** with automation superpowers, built with Python & CustomTkinter.

Tired of slow downloads and unreliable managers? **FetchPro** is the ultimate download solution—lightning-fast multipart acceleration, smart automation, REST API control, and a modern dark/light interface.

**Perfect for:** Developers, content creators, system admins, and power users who demand speed and reliability.

---

## 🎯 Why FetchPro?

| Feature | FetchPro | IDM | DownThemAll | Browser |
|---------|----------|-----|-------------|---------|
| **Multipart Acceleration** | ✅ | ✅ | ❌ | ❌ |
| **YouTube Downloads** | ✅ | ❌ | ✅ | ❌ |
| **BitTorrent Support** | ✅ | ❌ | ❌ | ❌ |
| **REST API** | ✅ | ❌ | ❌ | ❌ |
| **Chrome Extension** | ✅ | ✅ | ✅ | ✅ |
| **Open Source & Free** | ✅ | ❌ | ✅ | ✅ |
| **No Ads** | ✅ | ❌ | ✅ | ✅ |
| **VirusTotal Scanning** | ✅ | ❌ | ❌ | ❌ |
| **Scheduled Downloads** | ✅ | ✅ | ❌ | ❌ |

---

## ⚡ Key Features

### 📥 Core Downloads
- **HTTP/HTTPS** — multipart acceleration (up to 16 parallel segments), resume support, auto-retry
- **FTP/FTPS** — full FTP download support
- **YouTube & Media** — via `yt-dlp`: video quality selection, audio/MP3 extraction, playlist downloads
- **BitTorrent** — via `libtorrent`: full torrent support with seeding

### 📂 File Management
- Hash verification (MD5, SHA256)
- Automatic archive extraction (ZIP, RAR, 7Z, TAR, etc.)
- Auto-categorize files by type
- Auto-open downloaded files
- File tags and notes

### ⚙️ Performance & Control
- Global bandwidth throttling + per-download speed limits
- Priority-based queue (HIGH / NORMAL / LOW)
- Scheduled downloads with time-range restrictions
- Persistent queue — survives app restarts
- Watchdog service to restart stalled downloads

### 🔒 Security
- VirusTotal API integration for malware scanning
- File quarantine for detected threats
- Hash verification against checksums

### 🤖 Automation & Integration
- **REST API** on `http://127.0.0.1:9100` — control via n8n, scripts, or any HTTP client
- **Chrome Extension Bridge** on `http://127.0.0.1:9099` — one-click browser capture
- Clipboard monitoring — auto-detects copied URLs

### 🎨 UI & UX
- Dark / Light theme
- Drag-and-drop support
- Real-time progress with speed graphs
- Tabbed view: All · Active · Done · Failed
- Download history with search & CSV export
- System tray integration
- **Multi-language:** 🇮🇱 Hebrew · 🇬🇧 English · 🇦🇪 Arabic · 🇷🇺 Russian · 🇪🇸 Spanish · 🇫🇷 French

---

## 🖥️ System Requirements

- **OS:** Windows, macOS, or Linux
- **Python:** 3.8 or higher
- **RAM:** 512 MB minimum (1 GB recommended)
- **Disk:** 100 MB for installation
- **FFmpeg:** Required for audio conversion & format conversion (see installation)

---

## 🚀 Installation

### Step 1: Install FFmpeg (REQUIRED for full functionality)

**Windows:**
```bash
# Using Chocolatey
choco install ffmpeg

# Or using Scoop
scoop install ffmpeg
```

**macOS:**
```bash
# Using Homebrew
brew install ffmpeg
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get update
sudo apt-get install ffmpeg
```

**Linux (Fedora/RHEL):**
```bash
sudo dnf install ffmpeg
```

### Step 2: Install Python Dependencies

**Basic Setup:**
```bash
pip install requests
```

**Full Features (Recommended):**
```bash
pip install pystray pillow plyer yt-dlp libtorrent
```

| Package | Feature |
|---------|---------|
| `pystray pillow` | System tray icon |
| `plyer` | Native notifications |
| `yt-dlp` | YouTube & media downloads |
| `libtorrent` | BitTorrent support |
| `ffmpeg` | Audio/format conversion |

### Step 3: Verify Installation

```bash
ffmpeg -version
python fetchpro.py
```

---

## ⚡ Quick Start

```bash
python fetchpro.py
```

1. Paste a URL or drag-and-drop
2. Click Download
3. Watch the magic happen! ✨

---

## 📡 REST API

The API runs on `http://127.0.0.1:9100` (localhost only).

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/status` | List all downloads |
| GET | `/stats` | Usage statistics |
| GET | `/queue` | Current queue |
| POST | `/add` | Add a download (`{"url": "..."}`) |
| POST | `/pause_all` | Pause all downloads |
| POST | `/resume_all` | Resume all downloads |
| POST | `/cancel_all` | Cancel all downloads |

### Example

```bash
curl -X POST -H "Content-Type: application/json" \
  -d '{"url":"https://example.com/file.zip"}' \
  http://127.0.0.1:9100/add
```

---

## ⚙️ Configuration

Data stored in `~/.fetchpro/`:

| File | Purpose |
|------|---------|
| `settings.json` | User preferences |
| `history.db` | SQLite download history |
| `queue.json` | Persistent queue |
| `stats.json` | Usage statistics |
| `resume/` | Resume states |

---

## 🔌 Chrome Extension

### Features
🎥 **Download Button on Every Video**
- The extension adds a download button directly next to every video on YouTube, Twitch, Instagram, and other video platforms
- Click the button and choose your preferred format (MP4, WebM, MP3, M4A, WAV, etc.)
- Supports multiple quality options for video downloads
- One-click audio extraction (MP3, M4A, WAV, OPUS, VORBIS)

### Supported Formats
**Video Formats:**
- MP4 (H.264)
- WebM (VP9)
- FLV
- 3GP

**Audio Formats:**
- MP3 (MPEG Audio)
- M4A (AAC)
- WAV (PCM)
- OPUS
- VORBIS

### Installation
1. Download `FetchPro-Extension-v1.3.zip` from the repository
2. Extract to a folder
3. Open Chrome → `chrome://extensions/`
4. Enable **Developer Mode** (top-right corner)
5. Click **Load unpacked** → select the extracted folder
6. Done! 🎉

### How to Use
1. Browse to any video platform (YouTube, Twitch, Instagram, etc.)
2. Look for the **FetchPro download button** next to the video player
3. Click the button to see available formats
4. Select your preferred:
   - **Video quality** (360p, 720p, 1080p, etc.)
   - **Format** (MP4, WebM, MP3, etc.)
5. Download starts automatically in FetchPro!

### Troubleshooting
- **Button not appearing?** Make sure FetchPro desktop app is running (extension communicates via REST API)
- **Wrong format selected?** Make sure FFmpeg is installed for proper conversion
- **Download not starting?** Verify FetchPro is running and REST API is accessible on `http://127.0.0.1:9099`

---

## ❓ Troubleshooting

**Q: Downloads are slow**  
A: FetchPro uses multipart acceleration by default. Enable 16 parallel segments in settings for maximum speed.

**Q: yt-dlp not working?**  
A: Install it: `pip install yt-dlp` or update: `pip install --upgrade yt-dlp`

**Q: Can't access REST API?**  
A: Ensure FetchPro is running. The API only listens on localhost for security.

**Q: FFmpeg not found?**  
A: Make sure FFmpeg is installed on your system and added to PATH. Run `ffmpeg -version` to verify.

**Q: Chrome Extension button not showing?**  
A: Ensure FetchPro desktop app is running on the same machine. The extension communicates via local REST API.

**Q: Video conversion is slow**  
A: FFmpeg conversion speed depends on your CPU. Higher quality or longer videos take more time to convert.

---

## 🤝 Contributing

We welcome:
- 🐛 Bug reports
- 💡 Feature suggestions
- 🔧 Pull requests
- 📖 Documentation improvements

---

## 📄 License

MIT License — Use freely in personal and commercial projects.

---

## 🌟 Show Your Support

- ⭐ Star this repository
- 📢 Share with friends
- 💬 Join discussions
- 🐛 Report issues

---

Made with ❤️ by Moshe Pinhasi
