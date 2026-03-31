# JARVIS Voice Assistant

A Windows-focused voice assistant project with:
- Python backend (Flask)
- Browser UI (voiceAssistant.html)
- Wake-word flow (for example: "hey jarvis")
- Voice + text command support

The backend is served from Flask, and the frontend is loaded at `http://127.0.0.1:5000`.

## Features

- Wake-word style interaction (hands-free mode in UI)
- Voice command processing through browser microphone + backend command API
- Text command input from chat UI
- Timers and reminders (set, status, cancel)
- Weather and news queries
- App launch/close support (Word, Excel, PowerPoint, Outlook, Notepad, Calculator, WhatsApp, browser targets)
- Screenshot capture endpoint
- Optional AI response fallback via provider keys in `.env`

## Project Structure

- `app.py` - Flask backend, API routes, command processing bridge
- `VoiceAssistantgen8.py` - Core assistant logic (speech parsing, wake-word parsing, system actions)
- `voiceAssistant.html` - Frontend UI and browser speech integration
- `.env.example` - Environment variable template

## Requirements

- Windows 10/11
- Python 3.9+
- Working microphone
- Internet connection for weather/news/AI provider features

## Installation

1. Open terminal in project folder.
2. (Optional but recommended) create a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\activate
```

3. Install dependencies:

```powershell
pip install -r requirements.txt
```

4. If PyAudio fails to compile during install on your machine, install a compatible wheel for your Python version and retry:

```powershell
pip install pyaudio
```

The project is Windows-first, so PyAudio is included for microphone input support.

## Environment Setup

1. Copy `.env.example` to `.env`.
2. Add your API keys (optional for basic local commands, required for AI fallback and external services if configured through env).

Example:

```env
DEEPSEEK_API_KEY=your_key_here
# or use LLM_KEYS for multiple keys
```

## Run

Start backend:

```powershell
python app.py
```

Then open:

- `http://127.0.0.1:5000`

## Usage Flow

1. Click START SYSTEM.
2. Click ENABLE MIC in the browser if prompted.
3. Use one of these methods:
   - Manual: click LISTEN and speak.
   - Hands-free mode: say wake phrase (for example "hey jarvis").
4. You can also type commands in the chat input.

## Example Commands

- "hey jarvis what time is it"
- "weather in delhi"
- "set a timer for 5 minutes"
- "remind me to drink water in 30 minutes"
- "open youtube"
- "close spotify"
- "take screenshot"

## Main API Routes

- `GET /` - Serves UI
- `GET /api/status` - Assistant status
- `POST /api/toggle-system` - Start/stop assistant
- `POST /api/listen` - Trigger backend listening cycle
- `POST /api/process-command` - Process text command
- `GET /api/responses` - Poll queued responses
- `GET /api/check-mic` - Microphone check
- `POST /api/screenshot` - Capture screenshot
- `POST /api/timer/set` - Set timer
- `GET /api/timer/status` - Timer status
- `POST /api/timer/cancel` - Cancel timer

## Troubleshooting

- Microphone not working in browser:
  - Grant mic permission in browser site settings.
  - Use Chrome or Edge.
- Backend not starting:
  - Verify dependency install.
  - Check Python version and active environment.
- `ModuleNotFoundError`:
  - Install missing package with `pip install <package_name>`.
- Frontend shows old behavior after updates:
  - Restart backend.
  - Hard refresh browser with `Ctrl+F5`.

## Notes

- This is a development setup (Flask debug server).
- Some actions are Windows-specific by design.
- Keep secrets in `.env`; do not commit real API keys.