# FetchPro

Enterprise-grade download manager built with Python & CustomTkinter.

## Features

### Core Downloads
- **HTTP/HTTPS** — multipart acceleration (up to 16 parallel segments), resume support, auto-retry with exponential backoff
- **FTP/FTPS** — full FTP download support
- **YouTube & Media** — via `yt-dlp`: video quality selection, audio/MP3 extraction, playlist downloads
- **BitTorrent** — via `libtorrent`: full torrent support with seeding

### File Management
- Hash verification (MD5, SHA256)
- Automatic archive extraction (ZIP, RAR, 7Z, TAR, and more)
- Auto-categorize files by type into subfolders
- Auto-open downloaded file on completion
- File tags and notes

### Performance & Control
- Global bandwidth throttling + per-download speed limits
- Priority-based queue (HIGH / NORMAL / LOW)
- Scheduled downloads with time-range restrictions
- Persistent queue — survives app restarts
- Watchdog service to restart stalled downloads

### Security
- VirusTotal API integration for malware scanning
- File quarantine for detected threats
- Hash verification against user-supplied checksums

### Automation & Integration
- **REST API** on `http://127.0.0.1:9100` — control via n8n, scripts, or any HTTP client
- **Chrome Extension Bridge** on `http://127.0.0.1:9099` — one-click capture from the browser
- Clipboard monitoring — auto-detects copied URLs

### UI
- Dark / Light theme
- Drag-and-drop for URLs and torrent files
- Download cards with real-time progress, speed graph, and ETA
- Tabbed view: All · Active · Done · Failed
- Download history with search and CSV export
- System tray integration
- Multi-language: Hebrew, English, Arabic, Russian, Spanish, French

---

## Installation

### Requirements

```bash
pip install requests
```

### Optional (enables extra features)

| Package | Feature |
|---------|---------|
| `pystray pillow` | System tray icon |
| `plyer` | Native OS notifications |
| `yt-dlp` | YouTube & media downloads |
| `libtorrent` | BitTorrent support |

```bash
pip install pystray pillow plyer yt-dlp libtorrent
```

---

## Usage

```bash
python fetchpro.py
```

---

## REST API

The REST API listens on `http://127.0.0.1:9100` (localhost only).

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/status` | List all downloads |
| GET | `/stats` | Usage statistics |
| GET | `/queue` | Current queue |
| POST | `/add` | Add a download (`{"url": "..."}`) |
| POST | `/pause_all` | Pause all downloads |
| POST | `/resume_all` | Resume all downloads |
| POST | `/cancel_all` | Cancel all downloads |

---

## Configuration

Settings and data are stored in `~/.fetchpro/`:

| File | Description |
|------|-------------|
| `settings.json` | User preferences |
| `history.db` | SQLite download history |
| `queue.json` | Persistent download queue |
| `stats.json` | Usage statistics |
| `resume/` | Partial download resume states |

---

## Chrome Extension

Install the included `FetchPro-Extension-v1.3.zip` as an unpacked Chrome extension to capture download links directly from the browser with one click.
