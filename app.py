"""
JARVIS Backend Server - Integration with VoiceAssistantgen8.py
This Flask app bridges the HTML UI with the Python voice assistant backend
"""

from flask import Flask, render_template, jsonify, request, send_from_directory
from flask_cors import CORS
import os
import sys
import json
import threading
import queue
from datetime import datetime
import time
import requests
import re

# Import from VoiceAssistantgen8
sys.path.insert(0, os.path.dirname(__file__))

# Create Flask app
app = Flask(__name__, 
            static_folder=os.path.dirname(__file__),
            template_folder=os.path.dirname(__file__))
CORS(app)
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
app.config['TEMPLATES_AUTO_RELOAD'] = True


def load_local_env(env_path='.env'):
    """Lightweight .env loader to avoid requiring extra dependencies."""
    file_path = os.path.join(os.path.dirname(__file__), env_path)
    if not os.path.exists(file_path):
        return

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, value = line.split('=', 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = value
    except Exception as env_err:
        print(f".env load warning: {env_err}")


load_local_env('.env')


@app.after_request
def add_no_cache_headers(response):
    """Prevent stale cached frontend/API responses during rapid iteration."""
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0, private'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    response.headers['Surrogate-Control'] = 'no-store'
    response.headers['Vary'] = 'Origin, Accept-Encoding'
    return response

# Import voice assistant functions
try:
    from VoiceAssistantgen8 import (
        speak, tell_time, tell_joke, get_weather, handle_weather_request,
        get_news_headlines, handle_news_request, open_office_app,
        open_music_app, play_music, open_whatsapp_app, close_application, system_shutdown,
        tell_date, list_capabilities, empty_recycle_bin, open_recycle_bin,
        take_command, check_microphone, greet_user, take_screenshot, set_timer, get_timer_status,
        cancel_timer, active_timers, extract_command_after_wake_word
    )
except ImportError as e:
    print(f"Error importing voice assistant: {e}")
    print("Make sure VoiceAssistantgen8.py is in the same directory")

# Global state
assistant_state = {
    'active': False,
    'listening': False,
    'command_count': 0,
    'start_time': None,
    'last_command': '',
    'personality': 'butler'
}

# Map timer_id -> reminder text so reminder alarms can surface the original task.
reminder_notes = {}


def _parse_csv_env(name):
    """Parse comma-separated environment variable values."""
    raw = os.getenv(name, '')
    return [item.strip() for item in raw.split(',') if item.strip()]


def _build_llm_config(api_key, model_override=''):
    """Build one provider config from key + optional model."""
    is_openrouter = api_key.startswith('sk-or-v1')
    if is_openrouter:
        return {
            'api_key': api_key,
            'url': os.getenv('LLM_API_URL', 'https://openrouter.ai/api/v1/chat/completions'),
            'model': model_override or os.getenv('LLM_MODEL', 'google/gemma-3-27b-it:free'),
            'provider': 'openrouter'
        }

    return {
        'api_key': api_key,
        'url': os.getenv('LLM_API_URL', 'https://api.deepseek.com/chat/completions'),
        'model': model_override or os.getenv('LLM_MODEL', 'deepseek-chat'),
        'provider': 'deepseek'
    }


def get_llm_configs():
    """Resolve all LLM provider candidates from environment variables."""
    keys = _parse_csv_env('LLM_KEYS')
    models = _parse_csv_env('LLM_MODELS')

    # Backward compatibility: single-key setup from older env names.
    if not keys:
        single_key = (
            os.getenv('DEEPSEEK_API_KEY', '').strip()
            or os.getenv('OPENROUTER_API_KEY', '').strip()
            or os.getenv('LLM_API_KEY', '').strip()
        )
        if not single_key:
            return []
        return [_build_llm_config(single_key, os.getenv('LLM_MODEL', '').strip())]

    configs = []
    for i, key in enumerate(keys):
        model_override = ''
        if models:
            model_override = models[i] if i < len(models) else models[-1]
        configs.append(_build_llm_config(key, model_override))

    return configs


def get_llm_config():
    """Return the first configured LLM provider (legacy compatibility helper)."""
    configs = get_llm_configs()
    return configs[0] if configs else None


def _should_try_next_provider(error_text, status_code=0):
    """Decide if we should rotate to the next account/model candidate."""
    if status_code in (402, 408, 409, 425, 429, 500, 502, 503, 504):
        return True

    err = (error_text or '').lower()
    retry_signals = [
        'rate limit',
        'rate-limited',
        'temporarily rate-limited',
        'free-models-per-day',
        'daily limit',
        'per day',
        'quota',
        'insufficient credits',
        'temporarily unavailable',
        'overloaded',
        'capacity',
        'busy',
        'timeout',
        'http 429'
    ]
    return any(signal in err for signal in retry_signals)


def _ask_llm_once(prompt_text, cfg):
    """Send one LLM request using a single provider config."""
    headers = {
        'Authorization': f"Bearer {cfg['api_key']}",
        'Content-Type': 'application/json'
    }

    # Optional headers recommended by OpenRouter are harmless for DeepSeek too.
    headers['HTTP-Referer'] = os.getenv('LLM_HTTP_REFERER', 'http://localhost:5000')
    headers['X-Title'] = os.getenv('LLM_APP_TITLE', 'JARVIS Local Assistant')

    payload = {
        'model': cfg['model'],
        'messages': [
            {
                'role': 'system',
                'content': (
                    "You are JARVIS, a concise, friendly voice assistant. "
                    "Answer naturally in plain text, suitable for text-to-speech."
                )
            },
            {'role': 'user', 'content': prompt_text}
        ],
        'temperature': 0.7,
        'max_tokens': 350
    }

    # Some free reasoning-heavy models may consume output budget in reasoning and return content=None.
    # For OpenRouter, force reasoning budget to 0 so we reliably get final answer text.
    if 'openrouter.ai' in cfg.get('url', ''):
        payload['reasoning'] = {'max_tokens': 0}

    try:
        resp = requests.post(cfg['url'], headers=headers, json=payload, timeout=25)
        if resp.status_code >= 400:
            err_msg = ''
            try:
                err_obj = resp.json().get('error', {})
                err_msg = err_obj.get('message', '')
                raw_detail = err_obj.get('metadata', {}).get('raw', '')
                if raw_detail:
                    err_msg = f"{err_msg}. {raw_detail}" if err_msg else raw_detail
            except Exception:
                err_msg = ''
            final_msg = err_msg or f"Provider request failed with HTTP {resp.status_code}."
            return None, final_msg, _should_try_next_provider(final_msg, resp.status_code)

        data = resp.json()
        content = (
            data.get('choices', [{}])[0]
            .get('message', {})
            .get('content', '')
            .strip()
        )
        if content:
            return content, '', False
        return None, 'Provider returned an empty response.', True
    except Exception as llm_err:
        err_text = str(llm_err)
        return None, err_text, _should_try_next_provider(err_text)


def ask_llm(prompt_text):
    """Send a prompt to configured LLM providers with automatic account/model rotation."""
    configs = get_llm_configs()
    if not configs:
        return None

    last_error = 'No LLM response.'
    for idx, cfg in enumerate(configs):
        key_hint = cfg['api_key'][-6:] if len(cfg['api_key']) >= 6 else '******'
        print(f"[AI] Trying provider {idx + 1}/{len(configs)} ({cfg['provider']}, key ...{key_hint}, model {cfg['model']})")

        content, err_msg, should_rotate = _ask_llm_once(prompt_text, cfg)
        if content:
            return content

        if err_msg:
            last_error = err_msg
            print(f"[AI] Candidate failed: {err_msg[:220]}")

        if not should_rotate:
            break

    return f"__AI_ERROR__:{last_error}"


def local_chat_fallback(user_text, provider_error=''):
    """Return a lightweight conversational fallback when external LLM is unavailable."""
    text = (user_text or '').strip()
    lower = text.lower()

    if any(k in lower for k in ['sad', 'upset', 'hurt', 'bad day', 'depressed', 'cry']):
        return "I am here with you. Want to tell me what happened? We can take it one step at a time."

    if any(k in lower for k in ['gf', 'girlfriend', 'relationship', 'boyfriend', 'breakup', 'fight']):
        return "That sounds personal and important. Start with what happened, how you feel, and what outcome you want, then I can help you frame a calm message."

    if any(k in lower for k in ['thank you', 'thanks']):
        return "Always happy to help."

    if any(k in lower for k in ['who are you', 'what are you']):
        return "I am JARVIS, your assistant. I can help with tasks and basic conversation even when AI provider is temporarily unavailable."

    # Keep a concise and helpful generic response.
    if provider_error:
        err_lower = provider_error.lower()

        if (
            'rate-limited' in err_lower
            or 'temporarily rate-limited' in err_lower
            or 'rate limit' in err_lower
            or 'free-models-per-day' in err_lower
            or 'code":429' in err_lower
            or 'http 429' in err_lower
        ):
            if 'free-models-per-day' in err_lower or 'per day' in err_lower:
                return "Your free model daily limit is reached. Please wait for reset or add provider credits to continue AI replies."
            return "The free AI model is temporarily busy. Please retry in a few seconds, and I will answer through AI."

    if provider_error and ('insufficient credits' in provider_error.lower() or 'quota' in provider_error.lower()):
        return "I can still assist with basic replies right now. For full AI answers, your provider credits need to be topped up."

    return "I am listening. Please share a bit more detail and I will help with the best answer I can."

# Commands processing queue
command_queue = queue.Queue()
response_queue = queue.Queue()

# Note: Voice output handled by browser Speech Synthesis API, not pyttsx3

@app.route('/')
def index():
    """Serve the main HTML file"""
    return send_from_directory(os.path.dirname(__file__), 'voiceAssistant.html')

@app.route('/api/status', methods=['GET'])
def get_status():
    """Get current assistant status"""
    uptime = "00:00:00"
    if assistant_state['active'] and assistant_state['start_time']:
        elapsed = time.time() - assistant_state['start_time']
        hours = int(elapsed // 3600)
        minutes = int((elapsed % 3600) // 60)
        seconds = int(elapsed % 60)
        uptime = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    
    return jsonify({
        'active': assistant_state['active'],
        'listening': assistant_state['listening'],
        'command_count': assistant_state['command_count'],
        'uptime': uptime,
        'last_command': assistant_state['last_command'],
        'personality': assistant_state['personality']
    })

@app.route('/api/toggle-system', methods=['POST'])
def toggle_system():
    """Toggle assistant on/off"""
    assistant_state['active'] = not assistant_state['active']
    
    if assistant_state['active']:
        # Start each session clean so no previous timers/reminders auto-appear.
        try:
            cancel_timer('all')
        except Exception:
            pass
        reminder_notes.clear()
        assistant_state['start_time'] = time.time()
        assistant_state['command_count'] = 0
        return jsonify({
            'status': 'active', 
            'message': 'System started',
            'speak': 'System activated. I am JARVIS, ready to serve you. How may I assist?'
        })
    else:
        # Clear all timers/reminders on shutdown so next session starts fresh.
        try:
            cancel_timer('all')
        except Exception:
            pass
        reminder_notes.clear()
        return jsonify({
            'status': 'inactive', 
            'message': 'System stopped',
            'speak': 'System shutting down. Until next time, sir.'
        })

@app.route('/api/listen', methods=['POST'])
def start_listening():
    """Start voice recognition"""
    if not assistant_state['active']:
        return jsonify({'error': 'System not active'}), 400
    
    assistant_state['listening'] = True
    
    def listen_thread():
        try:
            command = take_command(wait_for_wake_word=True, silence_reply=True)
            if command and command != 'error':
                assistant_state['last_command'] = command
                assistant_state['command_count'] += 1
                process_command(command)
                response_queue.put({'type': 'command', 'text': command})
            else:
                response_queue.put({'type': 'error', 'text': 'Could not understand'})
        finally:
            assistant_state['listening'] = False
    
    thread = threading.Thread(target=listen_thread, daemon=True)
    thread.start()
    
    return jsonify({'status': 'listening'})

@app.route('/api/process-command', methods=['POST'])
def process_text_command():
    """Process text command from UI"""
    data = request.get_json(silent=True) or {}
    command = str(data.get('command', '')).strip().lower()

    if not command:
        return jsonify({
            'command': '',
            'response': 'Please say or type a command.',
            'action': 'info',
            'speak': 'Please say or type a command.',
            'command_count': assistant_state['command_count']
        })

    # Auto-recover instead of returning 400 when UI/backend state is out of sync
    if not assistant_state['active']:
        assistant_state['active'] = True
        if not assistant_state['start_time']:
            assistant_state['start_time'] = time.time()
    
    assistant_state['last_command'] = command
    assistant_state['command_count'] += 1
    
    result = process_command(command)
    
    return jsonify({
        'command': command,
        'response': result['response'],
        'action': result.get('action'),
        'speak': result.get('speak', result['response']),
        'command_count': assistant_state['command_count']
    })

@app.route('/api/responses', methods=['GET'])
def get_responses():
    """Get queued responses"""
    responses = []
    while not response_queue.empty():
        try:
            responses.append(response_queue.get_nowait())
        except queue.Empty:
            break
    
    return jsonify({'responses': responses})

@app.route('/api/check-mic', methods=['GET'])
def check_mic():
    """Check microphone availability"""
    available = check_microphone()
    return jsonify({'available': available})

@app.route('/api/weather/current-location', methods=['GET'])
def weather_current_location():
    """Get weather by browser geolocation coordinates for widget sync."""
    lat = request.args.get('lat', type=float)
    lon = request.args.get('lon', type=float)

    if lat is None or lon is None:
        return jsonify({'error': 'lat and lon are required'}), 400

    try:
        api_key = os.getenv('OPENWEATHER_API_KEY', '13828b8798daaef3ccba7c6b8cbb55fe')
        url = 'https://api.openweathermap.org/data/2.5/weather'
        params = {
            'lat': lat,
            'lon': lon,
            'appid': api_key,
            'units': 'metric'
        }

        resp = requests.get(url, params=params, timeout=10)
        data = resp.json()

        if str(data.get('cod')) != '200':
            return jsonify({'error': data.get('message', 'Unable to fetch weather')}), 502

        city = data.get('name', 'Current location')
        temp = data.get('main', {}).get('temp')
        desc = data.get('weather', [{}])[0].get('description', 'Unknown')

        return jsonify({
            'city': city,
            'temperature_c': temp,
            'description': desc,
            'lat': lat,
            'lon': lon
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def extract_weather_city(command_norm):
    """Extract a likely city phrase from weather commands."""
    keys = [
        'weather in ', 'weather at ', 'weather for ', 'weather of ',
        'temperature in ', 'temperature at ', 'temperature for ', 'temperature of '
    ]

    city = ''
    for key in keys:
        if key in command_norm:
            city = command_norm.split(key, 1)[1].strip()
            break

    if not city:
        return None

    # Remove conversational fillers and locality descriptors.
    for stop in [' right now', ' please', ' today', ' now']:
        city = city.replace(stop, '')

    city = re.sub(
        r'\b(near me|nearby|around me|around here|in my area|my area|my village|a village|the village|village|local area|my location)\b',
        ' ',
        city
    )
    city = re.sub(r'\b(which|that)\b.*$', ' ', city)
    city = re.sub(r'\b(a|an|the)\b', ' ', city)
    city = re.sub(r'\s+', ' ', city)
    city = city.strip(' .,!?:;')

    return city or None


def build_weather_city_candidates(city):
    """Build normalized location candidates from informal speech phrases."""
    if not city:
        return []

    candidates = [city]

    simplified = re.sub(r'\b(near me|nearby|around me|around here|in my area|my area|my village|village|local area|my location)\b', ' ', city)
    simplified = re.sub(r'\b(which|that)\b.*$', ' ', simplified)
    simplified = re.sub(r'\b(a|an|the)\b', ' ', simplified)
    simplified = re.sub(r'\s+', ' ', simplified).strip(' .,!?:;')

    if simplified and simplified not in candidates:
        candidates.append(simplified)

    # Fallback to first two words, then first word, for noisy inputs.
    parts = simplified.split()
    if len(parts) >= 2:
        first_two = ' '.join(parts[:2]).strip()
        if first_two and first_two not in candidates:
            candidates.append(first_two)
    if parts:
        first_one = parts[0].strip()
        if first_one and first_one not in candidates:
            candidates.append(first_one)

    return candidates


def get_precise_weather(city):
    """Resolve city with geocoding and fetch weather by coordinates for higher accuracy."""
    api_key = os.getenv('OPENWEATHER_API_KEY', '13828b8798daaef3ccba7c6b8cbb55fe')
    if not api_key:
        return None, 'Weather service key is missing on the server.'

    try:
        geo_url = 'https://api.openweathermap.org/geo/1.0/direct'
        loc = None
        for candidate in build_weather_city_candidates(city):
            geo_params = {
                'q': candidate,
                'limit': 1,
                'appid': api_key
            }
            geo_resp = requests.get(geo_url, params=geo_params, timeout=10)
            geo_data = geo_resp.json()

            if isinstance(geo_data, list) and geo_data:
                loc = geo_data[0]
                break

        if not loc:
            return None, f"I couldn't find a valid city for '{city}'. Please say a city name like 'weather in Delhi'."

        lat = loc.get('lat')
        lon = loc.get('lon')
        city_name = loc.get('name', city.title())
        country = loc.get('country', '')

        weather_url = 'https://api.openweathermap.org/data/2.5/weather'
        weather_params = {
            'lat': lat,
            'lon': lon,
            'appid': api_key,
            'units': 'metric'
        }
        weather_resp = requests.get(weather_url, params=weather_params, timeout=10)
        weather_data = weather_resp.json()

        if str(weather_data.get('cod')) != '200':
            return None, weather_data.get('message', 'Unable to fetch weather right now.')

        temp = weather_data.get('main', {}).get('temp')
        humidity = weather_data.get('main', {}).get('humidity')
        desc = weather_data.get('weather', [{}])[0].get('description', 'unknown')
        wind = weather_data.get('wind', {}).get('speed', 'unknown')

        place = f"{city_name}, {country}" if country else city_name
        report = (
            f"Current weather in {place}: {desc}. "
            f"Temperature {temp} degrees Celsius, humidity {humidity} percent, wind {wind} meters per second."
        )
        return report, None
    except Exception as err:
        return None, f"Weather service error: {err}"


def parse_reminder_command(command_norm):
    """Parse reminder intent like 'remind me to X in 1 hour' or 'set reminder to X after 30 minutes'."""
    patterns = [
        r'^remind me to\s+(?P<task>.+?)\s+(?:in|after)\s+(?P<when>.+?)(?:\s+from\s+now)?$',
        r'^set reminder to\s+(?P<task>.+?)\s+(?:in|after)\s+(?P<when>.+?)(?:\s+from\s+now)?$',
        r'^remind me\s+(?P<task>.+?)\s+(?:in|after)\s+(?P<when>.+?)(?:\s+from\s+now)?$',
        r'^remind me to\s+(?P<task>.+?)\s+(?P<when>\d+\s+(?:second|seconds|minute|minutes|hour|hours|day|days))\s+from\s+now$',
        r'^set reminder to\s+(?P<task>.+?)\s+(?P<when>\d+\s+(?:second|seconds|minute|minutes|hour|hours|day|days))\s+from\s+now$',
        r'^remind me\s+(?P<task>.+?)\s+(?P<when>\d+\s+(?:second|seconds|minute|minutes|hour|hours|day|days))\s+from\s+now$'
    ]

    for pat in patterns:
        m = re.match(pat, command_norm)
        if m:
            task = (m.group('task') or '').strip(' .,!?:;')
            when = (m.group('when') or '').strip(' .,!?:;')
            if task and when:
                return task, when

    return None, None


def is_time_query(command_norm):
    """Match explicit time questions, not casual mentions like 'lunch time'."""
    text = (command_norm or '').strip()
    if not text:
        return False

    patterns = [
        r'^time$',
        r'^what time(?: is it)?$',
        r'^whats the time$',
        r'^what is the time$',
        r'^tell me the time$',
        r'^current time$',
        r'^time now$'
    ]
    return any(re.match(p, text) for p in patterns)


def is_date_query(command_norm):
    """Match explicit date questions only."""
    text = (command_norm or '').strip().lower()
    if not text:
        return False

    # Normalize apostrophes so both "today's" and "todays" match.
    text = re.sub(r"[’']", '', text)

    patterns = [
        r'^date$',
        r'^what date(?: is it)?$',
        r'^what is the date$',
        r'^what is todays date$',
        r'^whats todays date$',
        r'^what is today date$',
        r'^todays date$',
        r'^today(?: date)?$',
        r'^tell me (?:the )?date$',
        r'^tell me todays date$'
    ]
    return any(re.match(p, text) for p in patterns)


def is_weather_query(command_norm):
    """Match explicit weather intents, not casual mentions of weather in free text."""
    text = (command_norm or '').strip().lower()
    if not text:
        return False

    text = re.sub(r"[’']", '', text)

    patterns = [
        r'^weather$',
        r'^weather\s+(?:in|at|for|of)\s+.+$',
        r'^temperature$',
        r'^temperature\s+(?:in|at|for|of)\s+.+$',
        r'^whats the weather(?:\s+(?:in|at|for)\s+.+)?$',
        r'^what is the weather(?:\s+(?:in|at|for)\s+.+)?$',
        r'^hows the weather(?:\s+(?:in|at|for)\s+.+)?$',
        r'^how is the weather(?:\s+(?:in|at|for)\s+.+)?$',
        r'^tell me the weather(?:\s+(?:in|at|for)\s+.+)?$'
    ]
    return any(re.match(p, text) for p in patterns)

@app.route('/api/set-personality', methods=['POST'])
def set_personality():
    """Set AI personality"""
    data = request.json
    personality = data.get('personality', 'butler')
    assistant_state['personality'] = personality
    return jsonify({'personality': personality, 'speak': f'Personality changed to {personality}'})

# ==================== Command Processing ====================

def process_command(command):
    """Process voice/text commands"""
    raw_command = (command or '').lower().strip()
    heard_wake_word, wake_command = extract_command_after_wake_word(raw_command)

    if heard_wake_word and not wake_command:
        return {
            'response': "Yes, I am listening. Please tell me your command.",
            'action': 'wake_word',
            'speak': "Yes, I am listening. Please tell me your command.",
            'timestamp': datetime.now().isoformat()
        }

    command_lower = wake_command if (heard_wake_word and wake_command) else raw_command
    command_norm = command_lower.replace('?', '').replace('.', '').replace(',', '').strip()
    command_norm = re.sub(r"[’']", '', command_norm)
    command_norm = re.sub(r'\s+', ' ', command_norm).strip()
    command_compact = command_norm.replace(' ', '')
    whatsapp_aliases = ('whatsapp', 'whatsup', 'whatsap', 'watsapp')
    has_whatsapp = any(alias in command_compact for alias in whatsapp_aliases)
    response_text = "Command processed"
    action = None
    
    try:
        # Reminder commands (real scheduled reminders using existing timer engine)
        if (
            'remind me' in command_lower
            or 'set reminder' in command_lower
            or 'show reminder' in command_lower
            or 'show reminders' in command_lower
            or 'reminder status' in command_lower
            or 'cancel reminder' in command_lower
        ):
            task, when_text = parse_reminder_command(command_norm)

            if task and when_text:
                result = set_timer(when_text)
                if result.get('success'):
                    timer_id = result.get('timer_id')
                    if timer_id is not None:
                        reminder_notes[timer_id] = task

                    ring_at = datetime.fromtimestamp(result['end_time']).strftime('%I:%M:%S %p').lstrip('0').lower()
                    response_text = f"Reminder set: {task}. I will remind you at {ring_at}."
                    action = 'reminder'
                else:
                    response_text = result.get('error', "Couldn't set reminder. Try 'remind me to take pills in 1 hour'.")
                    action = 'error'
            elif 'show reminder' in command_lower or 'reminder status' in command_lower:
                statuses = get_timer_status()
                reminder_statuses = [s for s in statuses if s.get('id') in reminder_notes]
                if reminder_statuses:
                    lines = []
                    for s in reminder_statuses:
                        if s.get('status') == 'running':
                            lines.append(f"- {reminder_notes.get(s['id'], 'Reminder')}: {s.get('remaining', 0)} seconds left")
                        else:
                            lines.append(f"- {reminder_notes.get(s['id'], 'Reminder')}: completed")
                    response_text = "Active reminders:\n" + "\n".join(lines)
                else:
                    response_text = "No active reminders"
                action = 'reminder'
            elif 'cancel reminder' in command_lower:
                if 'all' in command_lower:
                    result = cancel_timer('all')
                    reminder_notes.clear()
                    response_text = result.get('message', 'Cancelled all reminders')
                else:
                    reminder_ids = list(reminder_notes.keys())
                    if reminder_ids:
                        first_id = reminder_ids[0]
                        result = cancel_timer(first_id)
                        reminder_notes.pop(first_id, None)
                        response_text = result.get('message', 'Cancelled reminder')
                    else:
                        response_text = 'No active reminders to cancel'
                action = 'reminder'
            else:
                response_text = "Please say it like: remind me to take pills in 1 hour"
                action = 'reminder'

        # Timer commands (must be checked before generic time commands)
        elif (
            ('timer' in command_lower and ('set' in command_lower or 'start' in command_lower or 'create' in command_lower))
            or 'timer for' in command_lower
            or 'timer of' in command_lower
        ):
            duration = command_lower
            for token in ['set a timer for', 'set timer for', 'set a timer of', 'set timer of', 'set timer', 'start timer for', 'start timer of', 'start timer', 'create timer for', 'create timer of', 'create timer']:
                duration = duration.replace(token, '')
            duration = duration.replace('please', '').replace('a ', ' ').strip()
            if duration:
                result = set_timer(duration)
                if 'success' in result and result['success']:
                    ring_at = datetime.fromtimestamp(result['end_time']).strftime('%I:%M:%S %p').lstrip('0').lower()
                    response_text = f"{result['timer_name']} set for {result['duration']} seconds. Time's up at {ring_at}."
                    action = 'timer'
                else:
                    response_text = result.get('error', 'Could not set timer')
                    action = 'error'
            else:
                response_text = "Please specify a duration like '5 minutes' or '30 seconds'"
                action = 'timer'

        elif 'show timer' in command_lower or 'get timer' in command_lower or 'timer status' in command_lower:
            statuses = get_timer_status()
            if statuses:
                response_text = f"You have {len(statuses)} active timer(s)"
                action = 'timer'
            else:
                response_text = "No active timers"
                action = 'timer'

        elif 'cancel timer' in command_lower or 'stop timer' in command_lower:
            if 'all' in command_lower:
                result = cancel_timer('all')
                response_text = result.get('message', 'Timers cancelled')
            else:
                result = cancel_timer()
                response_text = result.get('message', 'Timer cancelled')
            action = 'timer'

        # Time commands
        elif is_time_query(command_norm):
            now = datetime.now()
            hour_24 = now.hour
            minute = now.minute
            am_pm = 'am' if hour_24 < 12 else 'pm'
            hour_12 = hour_24 % 12
            if hour_12 == 0:
                hour_12 = 12

            # Example: "Current time is 12 25 am"
            response_text = f"Current time is {hour_12} {minute:02d} {am_pm}"
            action = 'time'
        
        # Date commands
        elif is_date_query(command_norm):
            now = datetime.now()
            response_text = f"Current date is {now.strftime('%A, %B %d, %Y')}"
            action = 'date'
        
        # Weather commands - GUI-friendly natural response
        elif is_weather_query(command_norm):
            try:
                city = extract_weather_city(command_norm)
                if city:
                    report, weather_err = get_precise_weather(city)
                    if report:
                        response_text = report
                    else:
                        response_text = weather_err or "I could not fetch weather for that place."
                else:
                    response_text = "Please tell me a city, for example: weather in Delhi"

            except Exception as weather_err:
                print(f"Weather error: {weather_err}")
                response_text = f"I had trouble fetching the weather: {str(weather_err)}"
            action = 'weather'
        
        # News commands
        elif 'news' in command_lower or 'headline' in command_lower or \
             'what\'s the news' in command_lower or 'what is the news' in command_lower:
            try:
                category_map = {
                    'business': 'business',
                    'entertainment': 'entertainment',
                    'health': 'health',
                    'science': 'science',
                    'sports': 'sports',
                    'technology': 'technology',
                    'general': 'general'
                }

                source = 'both'
                if 'gnews' in command_lower:
                    source = 'gnews'
                elif 'newsapi' in command_lower or 'news api' in command_lower:
                    source = 'newsapi'

                category = 'general'
                for key, value in category_map.items():
                    if key in command_lower:
                        category = value
                        break

                num_headlines = 5
                if 'three' in command_lower or '3' in command_lower:
                    num_headlines = 3
                elif 'ten' in command_lower or '10' in command_lower:
                    num_headlines = 10

                headlines = get_news_headlines(source=source, category=category, num_headlines=num_headlines)
                response_text = f"Here are the latest {category} headlines:\n{headlines}"
            except Exception as news_err:
                print(f"News error: {news_err}")
                response_text = f"I had trouble fetching the news: {str(news_err)}"
            action = 'news'
        
        # Office apps
        elif 'open' in command_lower and 'word' in command_lower:
            open_office_app('word')
            response_text = "Opening Microsoft Word"
            action = 'open_app'
        
        elif 'open' in command_lower and 'excel' in command_lower:
            open_office_app('excel')
            response_text = "Opening Microsoft Excel"
            action = 'open_app'
        
        elif 'open' in command_lower and 'powerpoint' in command_lower:
            open_office_app('powerpoint')
            response_text = "Opening Microsoft PowerPoint"
            action = 'open_app'
        
        elif 'open' in command_lower and 'onenote' in command_lower:
            open_office_app('onenote')
            response_text = "Opening Microsoft OneNote"
            action = 'open_app'
        
        elif 'open' in command_lower and 'outlook' in command_lower:
            open_office_app('outlook')
            response_text = "Opening Microsoft Outlook"
            action = 'open_app'
        
        elif 'open' in command_lower and 'notepad' in command_lower:
            open_office_app('notepad')
            response_text = "Opening Notepad"
            action = 'open_app'
        
        elif 'open' in command_lower and 'calculator' in command_lower:
            open_office_app('calculator')
            response_text = "Opening Calculator"
            action = 'open_app'
        
        elif 'open' in command_lower and 'wordpad' in command_lower:
            open_office_app('wordpad')
            response_text = "Opening WordPad"
            action = 'open_app'

        elif (
            (('open' in command_norm or 'start' in command_norm or 'launch' in command_norm)
             and has_whatsapp)
            or command_norm in ('whatsapp', 'whats app')
        ):
            opened = open_whatsapp_app()
            response_text = "Opening WhatsApp" if opened else "WhatsApp Desktop is not installed on this system."
            action = 'open_app'
        
        # Music
        elif 'play' in command_lower and 'music' in command_lower:
            play_music()
            response_text = "Playing music"
            action = 'music'
        
        elif 'open spotify' in command_norm or command_norm == 'spotify':
            open_music_app('spotify')
            response_text = "Opening Spotify"
            action = 'open_app'
        
        elif 'open youtube music' in command_norm or 'open yt music' in command_norm or \
             ('youtube music' in command_norm and 'open' in command_norm):
            open_music_app('youtube music')
            response_text = "Opening YouTube Music"
            action = 'open_app'

        elif 'open youtube' in command_norm:
            import webbrowser
            webbrowser.open("https://youtube.com")
            response_text = "Opening YouTube"
            action = 'open_app'

        elif 'open google' in command_norm:
            import webbrowser
            webbrowser.open("https://google.com")
            response_text = "Opening Google"
            action = 'open_app'

        elif 'close youtube music' in command_norm or 'close yt music' in command_norm:
            close_application('youtube music')
            response_text = "Closing YouTube Music"
            action = 'close_app'

        elif 'close youtube' in command_norm:
            close_application('youtube')
            response_text = "Closing YouTube"
            action = 'close_app'

        elif 'close google' in command_norm:
            close_application('google')
            response_text = "Closing Google"
            action = 'close_app'

        elif 'close spotify' in command_norm:
            close_application('spotify')
            response_text = "Closing Spotify"
            action = 'close_app'

        elif ('close' in command_norm or 'stop' in command_norm or 'exit' in command_norm) and has_whatsapp:
            close_application('whatsapp')
            response_text = "Closing WhatsApp"
            action = 'close_app'

        elif 'close' in command_lower and ('word' in command_lower or 'ms word' in command_lower or 'microsoft word' in command_lower):
            close_application('word')
            response_text = "Closing Microsoft Word"
            action = 'close_app'

        elif 'close' in command_lower and ('excel' in command_lower or 'microsoft excel' in command_lower):
            close_application('excel')
            response_text = "Closing Microsoft Excel"
            action = 'close_app'

        elif 'close' in command_lower and ('powerpoint' in command_lower or 'ppt' in command_lower or 'microsoft powerpoint' in command_lower):
            close_application('powerpoint')
            response_text = "Closing Microsoft PowerPoint"
            action = 'close_app'

        elif 'close' in command_lower and ('onenote' in command_lower or 'one note' in command_lower):
            close_application('onenote')
            response_text = "Closing Microsoft OneNote"
            action = 'close_app'

        elif 'close' in command_lower and ('outlook' in command_lower or 'microsoft outlook' in command_lower):
            close_application('outlook')
            response_text = "Closing Microsoft Outlook"
            action = 'close_app'

        elif 'close' in command_lower and 'notepad' in command_lower:
            close_application('notepad')
            response_text = "Closing Notepad"
            action = 'close_app'

        elif 'close' in command_lower and ('wordpad' in command_lower or 'word pad' in command_lower):
            close_application('wordpad')
            response_text = "Closing WordPad"
            action = 'close_app'

        elif 'close' in command_lower and ('calculator' in command_lower or 'calc' in command_lower):
            close_application('calculator')
            response_text = "Closing Calculator"
            action = 'close_app'
        
        # Jokes
        elif 'joke' in command_lower or 'funny' in command_lower:
            joke_text = tell_joke()
            response_text = joke_text or "Here is a joke for you."
            action = 'joke'
        
        # Web search
        elif 'search' in command_lower or ('google' in command_lower and 'open google' not in command_norm):
            search_term = command_lower.replace('search', '').replace('google', '').strip()
            if search_term:
                import webbrowser
                webbrowser.open(f"https://www.google.com/search?q={search_term}")
                response_text = f"Searching for {search_term}"
                action = 'search'
            else:
                response_text = "What should I search on Google?"
                action = 'search'
        
        # Recycle bin
        elif 'empty' in command_lower and 'recycle' in command_lower:
            empty_recycle_bin()
            response_text = "Recycle bin emptied"
            action = 'system'
        
        elif 'open' in command_lower and 'recycle' in command_lower:
            open_recycle_bin()
            response_text = "Recycle bin opened"
            action = 'system'

        elif 'screenshot' in command_lower or 'screen shot' in command_lower:
            shot = take_screenshot()
            if shot.get('success'):
                response_text = f"Screenshot saved at: {shot['path']}"
                action = 'screenshot'
            else:
                response_text = f"Could not take screenshot: {shot.get('error', 'Unknown error')}"
                action = 'error'
        
        # System commands
        elif 'shutdown' in command_lower:
            system_shutdown('shutdown')
            response_text = "Shutdown initiated"
            action = 'system'
        
        elif 'restart' in command_lower:
            system_shutdown('restart')
            response_text = "Restart initiated"
            action = 'system'
        
        # Capabilities
        elif 'can you do' in command_lower or 'capabilities' in command_lower or 'what can you do' in command_lower:
            capabilities_lines = [
                "Here is what I can do:",
                "- Tell time and date",
                "- Get weather by city (example: weather in delhi)",
                "- Read latest news and category-wise headlines",
                "- Set, check, and cancel timers",
                "- Set reminders (example: remind me to take pills in 1 hour)",
                "- Open apps and sites: YouTube, YouTube Music, Google, Spotify, WhatsApp, Calculator, Notepad, Word, Excel, PowerPoint, Outlook",
                "- Close apps/sites: YouTube, YouTube Music, Google, Spotify, WhatsApp and others",
                "- Tell jokes",
                "- Recycle Bin actions: open and empty",
                "- Search Google",
                "- Listen to your voice commands with live mic level",
            ]
            response_text = "\n".join(capabilities_lines)
            action = 'info'
        
        # Default response
        else:
            llm_reply = ask_llm(command)
            if llm_reply:
                if llm_reply.startswith('__AI_ERROR__:'):
                    provider_error = llm_reply.replace('__AI_ERROR__:', '', 1).strip()
                    response_text = local_chat_fallback(command, provider_error)
                    action = 'ai_fallback'
                else:
                    response_text = llm_reply
                    action = 'ai'
            else:
                if get_llm_config():
                    response_text = local_chat_fallback(command)
                    action = 'ai_fallback'
                else:
                    response_text = "AI fallback is off. Add DEEPSEEK_API_KEY or LLM_KEYS to .env, then restart the backend."
                    action = 'ai_setup_required'
    
    except Exception as e:
        print(f"Error processing command: {e}")
        response_text = f"Error: {str(e)}"
        action = 'error'
    
    return {
        'response': response_text,
        'action': action,
        'speak': response_text,  # Browser will speak this
        'timestamp': datetime.now().isoformat()
    }

@app.route('/api/speak', methods=['POST'])
def api_speak():
    """API endpoint to speak text (browser handles voice synthesis)"""
    data = request.json
    text = data.get('text', '')
    
    if text:
        return jsonify({'status': 'spoken', 'speak': text})
    
    return jsonify({'error': 'No text provided'}), 400

@app.route('/api/close-app', methods=['POST'])
def close_app():
    """Close an application"""
    data = request.json
    app_name = data.get('app', '')
    
    if app_name:
        close_application(app_name)
        return jsonify({'status': 'closed', 'app': app_name})
    
    return jsonify({'error': 'No app specified'}), 400


@app.route('/api/screenshot', methods=['POST'])
def api_take_screenshot():
    """Capture a real screenshot and return where it was saved."""
    result = take_screenshot()
    if result.get('success'):
        return jsonify(result)
    return jsonify(result), 500

# ==================== Timer APIs ====================

@app.route('/api/timer/set', methods=['POST'])
def api_set_timer():
    """Set a timer via API"""
    data = request.json
    duration = data.get('duration', '')
    
    if not duration:
        return jsonify({'error': 'No duration specified'}), 400
    
    result = set_timer(duration)
    
    if 'success' in result and result['success']:
        return jsonify(result)
    else:
        return jsonify(result), 400

@app.route('/api/timer/status', methods=['GET'])
def api_timer_status():
    """Get all active timers status"""
    statuses = get_timer_status()

    # Attach reminder text to timers and clean stale reminder mappings.
    completed_ids = set()
    for timer in statuses:
        tid = timer.get('id')
        if tid in reminder_notes:
            timer['reminder'] = reminder_notes[tid]
        if timer.get('status') == 'completed':
            completed_ids.add(tid)

    for tid in completed_ids:
        reminder_notes.pop(tid, None)

    active_ids = set(active_timers.keys())
    for tid in list(reminder_notes.keys()):
        if tid not in active_ids and tid not in completed_ids:
            reminder_notes.pop(tid, None)

    return jsonify({
        'timers': statuses,
        'count': len(statuses),
        'active': len(statuses) > 0
    })

@app.route('/api/timer/cancel', methods=['POST'])
def api_cancel_timer():
    """Cancel a specific timer or all timers"""
    data = request.json
    timer_id = data.get('timer_id')  # None for first, 'all' for all
    
    result = cancel_timer(timer_id)
    
    if 'success' in result and result['success']:
        return jsonify(result)
    else:
        return jsonify(result), 400

@app.route('/api/timer/cancel-all', methods=['POST'])
def api_cancel_all_timers():
    """Cancel all timers"""
    result = cancel_timer('all')
    
    if 'success' in result and result['success']:
        return jsonify(result)
    else:
        return jsonify(result), 400

# ==================== Error Handlers ====================

@app.errorhandler(404)
def not_found(error):
    return jsonify({'error': 'Not found'}), 404

@app.errorhandler(500)
def server_error(error):
    return jsonify({'error': 'Server error'}), 500

# ==================== Main ====================

if __name__ == '__main__':
    print("=" * 60)
    print("[JARVIS] Backend Server Starting...")
    print("=" * 60)
    print("Voice Assistant Gen 8 Backend")
    print("UI + Python Integration")
    print("=" * 60)
    print("\n[OK] Server configuration:")
    print(f"   - Host: 127.0.0.1")
    print(f"   - Port: 5000")
    print(f"   - Debug: True")
    print(f"\n[WEB] Access the UI at: http://localhost:5000")
    print(f"\n[MIC] Voice Recognition: Enabled")
    print(f"[API] API Endpoints: Active")
    print(f"[AI] LLM Fallback: {'Enabled' if get_llm_config() else 'Disabled (set DEEPSEEK_API_KEY or LLM_KEYS)'}")
    print(f"\nPress CTRL+C to stop the server\n")
    
    try:
        app.run(
            host='127.0.0.1',
            port=5000,
            debug=True,
            use_reloader=False,
            threaded=True
        )
    except KeyboardInterrupt:
        print("\n\n[STOP] Server shutdown requested")
        print("[EXIT] Goodbye!")
