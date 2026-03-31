# updated on 2024-07-17
# with weather functionality
# And the child of Gen 7
# Can now open Calculator
# Now with latest news headlines from multiple sources
# Removed the feature to open vs code will do give in future updates
# now also can open WhatsApp desktop app ( updated 2026 - 03 - 23)
import speech_recognition as sr
import pyttsx3
import webbrowser
import datetime
import os
import time
import random
import sys
import subprocess
import psutil
from threading import Thread
import pygetwindow as gw
import pyautogui
import winshell
import ctypes
from ctypes import wintypes
import winreg  # For checking installed programs
import requests
import json
import re

WEATHER_API_KEY = "13828b8798daaef3ccba7c6b8cbb55fe"  # Replace with your actual API key
WEATHER_BASE_URL = "https://api.openweathermap.org/data/2.5/weather?"
# News API Configuration (get from https://newsapi.org/)
NEWS_API_KEY = "ca9cfc9c2faf469fa712106c835505cd"  # Replace with your actual key
NEWS_API_URL = "https://newsapi.org/v2/top-headlines"
# GNews API Configuration (get from https://gnews.io/)
GNEWS_API_KEY = "512f0c5b2c63cd86adbe915305ed23d5"
GNEWS_API_URL = "https://gnews.io/api/v4/top-headlines"

def get_news_headlines(source="both", category="general", num_headlines=5):
    """Fetch news headlines using multiple sources"""
    headlines = []
    
    # Get NewsAPI headlines
    if source in ["newsapi", "both"]:
        try:
            params = {
                "apiKey": NEWS_API_KEY,
                "category": category,
                "pageSize": num_headlines,
                "language": "en"
            }
            response = requests.get(NEWS_API_URL, params=params)
            news_data = response.json()
            
            if news_data.get("status") == "ok":
                for article in news_data.get("articles", [])[:num_headlines]:
                    title = article.get("title", "")
                    # Clean up title: remove special characters and source names
                    title = re.sub(r' - [^-]+$', '', title)  # Remove source attribution
                    title = re.sub(r'[^\x00-\x7F]+', ' ', title)  # Replace non-ASCII chars
                    headlines.append(title.strip())
        except Exception as e:
            print(f"NewsAPI error: {e}")
    
    # Get GNews headlines if we don't have enough yet
    if len(headlines) < num_headlines and source in ["gnews", "both"]:
        try:
            params = {
                "token": GNEWS_API_KEY,
                "topic": category,
                "max": num_headlines,
                "lang": "en"
            }
            response = requests.get(GNEWS_API_URL, params=params)
            gnews_data = response.json()
            
            for article in gnews_data.get("articles", [])[:num_headlines - len(headlines)]:
                title = article.get("title", "")
                # Clean up title: remove special characters and source names
                title = re.sub(r' - [^-]+$', '', title)  # Remove source attribution
                title = re.sub(r'[^\x00-\x7F]+', ' ', title)  # Replace non-ASCII chars
                if title not in headlines:  # Avoid duplicates
                    headlines.append(title.strip())
        except Exception as e:
            print(f"GNews error: {e}")
    
    # Format results
    if headlines:
        return "\n".join([f"{idx+1}. {title}" for idx, title in enumerate(headlines)])
    return "Sorry, I couldn't fetch news headlines at the moment."

def handle_news_request(query):
    """Handle news-related queries with natural language processing"""
    # Category mapping for both APIs
    category_map = {
        "business": "business",
        "entertainment": "entertainment",
        "health": "health",
        "science": "science",
        "sports": "sports",
        "technology": "technology",
        "general": "general",
        "world": "general",
        "national": "general",
        "politics": "general"
    }
    
    # Source preference
    source = "both"
    if 'gnews' in query:
        source = "gnews"
    elif 'newsapi' in query or 'news api' in query:
        source = "newsapi"
    
    # Default settings
    category = "general"
    num_headlines = 5
    
    # Detect category in query
    for key, value in category_map.items():
        if key in query:
            category = value
            break
    
    # Detect number of headlines requested
    if 'five' in query:
        num_headlines = 5
    elif 'three' in query:
        num_headlines = 3
    elif 'ten' in query or '10' in query:
        num_headlines = 10
    
    speak(f"Getting the latest {category} news. Please wait...")
    headlines = get_news_headlines(source, category, num_headlines)
    speak(headlines)

def get_weather(city_name):
    """Get weather data from OpenWeatherMap API"""
    try:
        # Build the complete URL
        complete_url = f"{WEATHER_BASE_URL}appid={WEATHER_API_KEY}&q={city_name}&units=metric"
        
        # Make the API request
        response = requests.get(complete_url)
        weather_data = response.json()
        
        # Check if city is found
        if weather_data["cod"] != "404":
            main_data = weather_data["main"]
            current_temperature = main_data["temp"]
            current_humidity = main_data["humidity"]
            weather_desc = weather_data["weather"][0]["description"]
            
            # Get wind speed if available
            wind_speed = weather_data.get("wind", {}).get("speed", "unknown")
            
            # Get visibility if available (converted from meters to km)
            visibility = weather_data.get("visibility")
            if visibility:
                visibility = f"{visibility/1000:.1f} km"
            else:
                visibility = "unknown"
            
            # Prepare the weather report
            report = (
                f"The current weather in {city_name} is {weather_desc}. "
                f"The temperature is {current_temperature:.1f}°C with humidity at {current_humidity}%. "
                f"Wind speed is {wind_speed} m/s and visibility is {visibility}."
            )
            return report
        else:
            return f"Weather information not available for {city_name}"
            
    except Exception as e:
        print(f"Weather API error: {e}")
        return "Sorry, I couldn't fetch the weather information right now."
    
def handle_weather_request(query):
    """Handle weather-related queries"""
    # Check if the query already contains a location
    location_keywords = ['weather in', 'weather at', 'weather for', 'temperature in']
    location = None
    
    for keyword in location_keywords:
        if keyword in query:
            location = query.split(keyword)[-1].strip()
            break
    
    # If no location found, ask for one
    if not location:
        speak("Which city's weather would you like to know?")
        location = take_command()
    
    if location and location.lower() not in ['cancel', 'never mind']:
        speak(f"Getting weather information for {location}. Please wait...")
        weather_report = get_weather(location)
        speak(weather_report)
    else:
        speak("Weather check cancelled.")    
    
# Initialize speech engine (disabled - using browser Web Speech API instead)
# engine = pyttsx3.init()
# engine.setProperty('rate', 160)
# voices = engine.getProperty('voices')
# engine.setProperty('voice', voices[1].id)

# Global state
assistant_active = True
processing = False

WAKE_WORD_ALIASES = (
    'hey jarvis',
    'hi jarvis',
    'okay jarvis',
    'ok jarvis',
    'jarvis',
    'hey jervis',
)


def normalize_spoken_text(text):
    """Normalize recognized speech for easier intent and wake-word parsing."""
    normalized = (text or '').lower().strip()
    normalized = re.sub(r"[’']", '', normalized)
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized


def extract_command_after_wake_word(text):
    """Return (heard_wake_word, command_text_after_wake_word)."""
    cleaned = normalize_spoken_text(text)
    if not cleaned:
        return False, ''

    for alias in WAKE_WORD_ALIASES:
        if cleaned == alias:
            return True, ''

        for separator in (' ', ',', ':', '-', '.'):
            prefix = f"{alias}{separator}"
            if cleaned.startswith(prefix):
                return True, cleaned[len(prefix):].strip()

    return False, ''

# Dictionary of applications with their common names and typical paths
APPS = {
    # Microsoft Office apps (unchanged from your original code)
    'word': {
        'name': 'Microsoft Word',
        'paths': [
            r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
            r'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE',
            r'C:\Program Files\Microsoft Office\Office15\WINWORD.EXE'
        ]
    },
    # ... (other Office apps remain the same)
    
    # Music apps
    'spotify': {
        'name': 'Spotify',
        'paths': [
            r'C:\Users\{username}\AppData\Roaming\Spotify\Spotify.exe',
            r'C:\Program Files\WindowsApps\SpotifyAB.SpotifyMusic_*\Spotify.exe',
            r'C:\Program Files\Spotify\Spotify.exe'
        ],
        'check_registry': True,
        'registry_path': r'Software\Microsoft\Windows\CurrentVersion\Uninstall\Spotify'
    },
    'youtube music': {
        'name': 'YouTube Music',
        'web_url': 'https://music.youtube.com',
        'desktop_path': r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --app=https://music.youtube.com'
    }
}

def is_app_installed(app_name):
    """Check if an application is installed by looking in registry"""
    try:
        if app_name == 'spotify':
            # Check both 32-bit and 64-bit registry
            reg_path = APPS['spotify']['registry_path']
            try:
                winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
                return True
            except WindowsError:
                try:
                    winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
                    return True
                except WindowsError:
                    return False
        return False
    except Exception as e:
        print(f"Registry check error for {app_name}: {e}")
        return False

def open_music_app(app_key):
    """Handle opening music applications with proper checks"""
    if app_key not in APPS:
        speak(f"I don't recognize the music application {app_key}")
        return
    
    app_info = APPS[app_key]
    app_name = app_info['name']
    
    # Special handling for Spotify
    if app_key == 'spotify':
        if not is_app_installed('spotify'):
            speak("Spotify doesn't appear to be installed. Would you like to open it in your browser instead?")
            response = take_command()
            if response and 'yes' in response:
                webbrowser.open("https://open.spotify.com")
            return
        
        # Try all possible paths for Spotify
        username = os.getlogin()
        paths = [p.format(username=username) for p in app_info['paths']]
        
        found = False
        for path in paths:
            # Handle wildcards in path (common for Windows Store apps)
            if '*' in path:
                import glob
                matches = glob.glob(path)
                if matches:
                    path = matches[0]
            
            if os.path.exists(path):
                try:
                    speak(f"Opening {app_name}")
                    os.startfile(path)
                    found = True
                    break
                except Exception as e:
                    print(f"Error opening {app_name}: {e}")
        
        if not found:
            speak(f"Sorry, I couldn't find {app_name}. It might not be properly installed.")
    
    # Handling for YouTube Music
    elif app_key == 'youtube music':
        if 'web_url' in app_info:
            speak(f"Opening {app_name} in your browser")
            webbrowser.open(app_info['web_url'])
        elif 'desktop_path' in app_info:
            try:
                speak(f"Opening {app_name}")
                os.system(app_info['desktop_path'])
            except Exception as e:
                print(f"Error opening {app_name}: {e}")
                speak(f"Sorry, I couldn't open {app_name}")

def play_music():
    """Enhanced music playback with options"""
    speak("Would you like to play music on Spotify or YouTube Music?")
    choice = take_command()
    
    if choice and 'spotify' in choice:
        open_music_app('spotify')
    elif choice and ('youtube' in choice or 'music' in choice):
        open_music_app('youtube music')
    else:
        # Fall back to original music folder behavior
        music_dir = os.path.join(os.path.expanduser('~'), 'Music')
        if os.path.exists(music_dir):
            try:
                speak("Playing music from your Music folder")
                os.startfile(music_dir)
            except Exception as e:
                speak("Sorry, I couldn't play music. Please check your Music folder.")
        else:
            speak("I couldn't find your Music folder. Where do you keep your music files?")


# Dictionary of Microsoft Office applications with their common names and typical paths
OFFICE_APPS = {
    'word': {
        'name': 'Microsoft Word',
        'paths': [
            r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
            r'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE',
            r'C:\Program Files\Microsoft Office\Office15\WINWORD.EXE'
        ]
    },
    'excel': {
        'name': 'Microsoft Excel',
        'paths': [
            r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE',
            r'C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE',
            r'C:\Program Files\Microsoft Office\Office15\EXCEL.EXE'
        ]
    },
    'powerpoint': {
        'name': 'Microsoft PowerPoint',
        'paths': [
            r'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE',
            r'C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE',
            r'C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE'
        ]
    },
    'onenote': {
        'name': 'Microsoft OneNote',
        'paths': [
            r'C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE',
            r'C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE'
        ]
    },
    'outlook': {
        'name': 'Microsoft Outlook',
        'paths': [
            r'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE',
            r'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE'
        ]
    },
    'notepad': {
        'name': 'Notepad',
        'paths': [
            r'C:\Windows\system32\notepad.exe',
            r'C:\Windows\notepad.exe'
        ]
    },
    # Add this entry to your OFFICE_APPS dictionary
    'calculator': {
        'name': 'Calculator',
        'paths': [
            r'C:\Windows\System32\calc.exe',
            r'C:\Windows\SysWOW64\calc.exe'  # For 32-bit systems
        ]
    },
    'wordpad': {
        'name': 'WordPad',
        'paths': [
            r'C:\Program Files\Windows NT\Accessories\wordpad.exe',
            r'C:\Program Files (x86)\Windows NT\Accessories\wordpad.exe'
        ]
    },
    'whatsapp': {
        'name': 'WhatsApp',
        'paths': [
            r'C:\Users\{username}\AppData\Local\WhatsApp\WhatsApp.exe',
            r'C:\Program Files\WhatsApp\WhatsApp.exe',
            r'C:\Program Files (x86)\WhatsApp\WhatsApp.exe'
        ]
    }
}

def speak(text, wait=False):
    """Print to console - Browser handles voice output via Web Speech API"""
    # Clean text for display
    safe_text = re.sub(r'[^\x00-\x7F]+', ' ', text)
    print(f"JARVIS: {safe_text}")
    # Note: Web Speech Synthesis API in browser handles voice output
    if wait:
        time.sleep(0.3)

def play_sound(frequency=500, duration=100):
    """Play a simple beep sound"""
    import winsound
    winsound.Beep(frequency, duration)

def play_timer_completion_alert():
    """Play a clear multi-beep alert when a timer/reminder completes."""
    try:
        for freq, dur in [(880, 180), (1040, 180), (1320, 260)]:
            play_sound(freq, dur)
            time.sleep(0.06)
    except Exception as e:
        print(f"Timer alert sound failed: {e}")

def check_microphone():
    """Verify microphone is available"""
    try:
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Testing microphone...")
            r.adjust_for_ambient_noise(source, duration=0.5)
            print("✅ Microphone OK")
            return True
    except sr.MicrophoneError as e:
        print(f"❌ Microphone not found: {e}")
        return False
    except Exception as e:
        print(f"❌ Microphone error: {e}")
        return False

def take_command(wait_for_wake_word=False, timeout=8, silence_reply=False):
    """Listen for commands, optionally waiting for wake word before returning command text."""
    r = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            print("\nListening...")
            play_sound(600, 150)  # Start listening sound
            r.pause_threshold = 1.5
            r.energy_threshold = 4000
            r.adjust_for_ambient_noise(source, duration=0.8)
            audio = r.listen(source, timeout=timeout)
            
            print("Processing...")
            play_sound(400, 150)  # Processing sound
            query = r.recognize_google(audio, language='en-US')
            print(f"You said: {query}")
            query_lower = normalize_spoken_text(query)

            if not wait_for_wake_word:
                return query_lower

            heard_wake_word, inline_command = extract_command_after_wake_word(query_lower)
            if not heard_wake_word:
                return ''

            if inline_command:
                return inline_command

            speak("Yes, I am listening.")
            return take_command(wait_for_wake_word=False, timeout=10, silence_reply=True)
            
    except sr.WaitTimeoutError:
        if not processing and not silence_reply and not wait_for_wake_word:
            speak("I didn't hear anything. Are you still there?")
        return ""
    except sr.UnknownValueError:
        if not silence_reply and not wait_for_wake_word:
            speak("Sorry, I didn't catch that. Could you repeat?")
        return ""
    except Exception as e:
        print(f"Recognition error: {e}")
        if not silence_reply:
            speak("I'm having trouble with my audio. Please check your microphone.")
        return "error"

def open_office_app(app_key):
    """Open Microsoft Office applications with intelligent path finding"""
    if app_key not in OFFICE_APPS:
        speak(f"I don't recognize the application {app_key}")
        return
    
    app_info = OFFICE_APPS[app_key]
    app_name = app_info['name']
    username = os.getlogin()
    paths = [p.format(username=username) for p in app_info['paths']]
    
    found = False
    for path in paths:
        if os.path.exists(path):
            try:
                speak(f"Opening {app_name}")
                os.startfile(path)
                found = True
                break
            except Exception as e:
                print(f"Error opening {app_name}: {e}")
    
    if not found:
        # Last attempt using shell command
        try:
            speak(f"Attempting to open {app_name}")
            os.system(f'start {app_name.replace(" ", "").lower()}:')
        except Exception as e:
            print(f"Final attempt failed: {e}")
            speak(f"Sorry, I couldn't open {app_name}. It might not be installed or the path is different.")

def open_application(app_name, path):
    """Open applications with feedback"""
    if os.path.exists(path):
        speak(f"Opening {app_name}")
        os.startfile(path)
    else:
        speak(f"I couldn't find {app_name}. Is it installed?")


def open_whatsapp_app():
    """Open WhatsApp Desktop ONLY (no browser fallback)."""
    try:
        # Best method: URI scheme (works if app installed)
        os.startfile('whatsapp:')
        speak("Opening WhatsApp")
        return True
    except Exception:
        pass

    # Try common installation paths
    username = os.getlogin()
    path_candidates = [
        rf'C:\Users\{username}\AppData\Local\WhatsApp\WhatsApp.exe',
        r'C:\Program Files\WhatsApp\WhatsApp.exe',
        r'C:\Program Files (x86)\WhatsApp\WhatsApp.exe'
    ]

    for exe_path in path_candidates:
        if os.path.exists(exe_path):
            try:
                os.startfile(exe_path)
                speak("Opening WhatsApp")
                return True
            except Exception as e:
                print(f"Error opening WhatsApp from {exe_path}: {e}")

    # If not found -> NO browser fallback
    speak("WhatsApp Desktop is not installed on this system.")
    return False

def close_application(app_name):
    """Close applications by name with improved Edge browser tab handling"""
    closed = False
    app_name_lower = app_name.lower()
    
    # Mapping of common names to process names with Edge-specific settings
    app_mapping = {
        'youtube': {
            'process': 'msedge',  # Edge's process name
            'url': 'youtube.com',
            'close_method': 'tab'
        },
        'google': {
            'process': 'msedge',
            'url': 'google.com',
            'close_method': 'tab'
        },
        # Music applications
        'spotify': {
            'process': 'Spotify',
            'close_method': 'process'
        },
        'youtube music': {
            'process': 'msedge',  # Edge's process name
            'url': 'music.youtube.com',
            'close_method': 'tab'
        },
        'yt music': {
            'process': 'msedge',  # Edge's process name
            'url': 'music.youtube.com',
            'close_method': 'tab'
        },
        'edge': {
            'process': 'msedge',
            'close_method': 'process'
        },
        'microsoft edge': {
            'process': 'msedge',
            'close_method': 'process'
        },
        'word': {
            'process': 'winword',
            'close_method': 'process'
        },
        'excel': {
            'process': 'excel',
            'close_method': 'process'
        },
        'powerpoint': {
            'process': 'powerpnt',
            'close_method': 'process'
        },
        'onenote': {
            'process': 'onenote',
            'close_method': 'process'
        },
        'outlook': {
            'process': 'outlook',
            'close_method': 'process'
        },
        'notepad': {
            'process': 'notepad',
            'close_method': 'process'
        },
        'calculator': {
            'process': 'calculator',
            'close_method': 'process'
        },
        'calc': {
            'process': 'calculator',
            'close_method': 'process'
        },
        'calc culator': {
            'process': 'calculator',
            'close_method': 'process'
        },
        'wordpad': {
            'process': 'wordpad',
            'close_method': 'process'
        }
    }
    
    app_info = app_mapping.get(app_name_lower, {'process': app_name_lower, 'close_method': 'process'})
    
    try:
        if app_info.get('close_method') == 'tab' and 'url' in app_info:
            # Try to close Edge tab by URL
            try:                
                # Find Edge window with the URL
                for window in gw.getWindowsWithTitle(app_info['url']):
                    if window.title.lower().find(app_info['url']) != -1:
                        window.activate()
                        time.sleep(0.3)  # Small delay for window activation
                        pyautogui.hotkey('ctrl', 'w')  # Close tab
                        closed = True
                        time.sleep(0.5)  # Wait for tab to close
                        break  # Stop after closing one matching tab
            except Exception as e:
                print(f"Tab closing error: {e}")
                # Fall back to process closing if tab closing fails
        
        # If tab closing didn't work or it's a regular app
        if not closed:
            for proc in psutil.process_iter(['pid', 'name']):
                if app_info['process'].lower() in proc.info['name'].lower():
                    proc.kill()
                    closed = True
        # Process closing (for Spotify and fallback)
        if not closed:
            process_name = app_info['process']
            for proc in psutil.process_iter(['pid', 'name']):
                if process_name.lower() in proc.info['name'].lower():
                    # Special handling for Spotify (it has multiple processes)
                    if app_name_lower == 'spotify':
                        # Close all Spotify related processes
                        for spotify_proc in psutil.process_iter(['pid', 'name']):
                            if 'spotify' in spotify_proc.info['name'].lower():
                                spotify_proc.kill()
                        closed = True
                        break
                    else:
                        proc.kill()
                        closed = True
        if closed:
            speak(f"Closed {app_name}")
        else:
            speak(f"I couldn't find {app_name} running")
    except Exception as e:
        print(f"Error closing application: {e}")
        speak(f"Sorry, I couldn't close {app_name}")

# Recycle Bin Functions (unchanged from your original code)
def empty_recycle_bin():
    """Empty the Recycle Bin"""
    try:
        winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=False)
        speak("Recycle Bin has been emptied successfully")
    except Exception as e:
        print(f"Error emptying Recycle Bin: {e}")
        speak("Sorry, I couldn't empty the Recycle Bin")

def open_recycle_bin():
    """Open the Recycle Bin window"""
    try:
        os.system('explorer.exe shell:RecycleBinFolder')
        speak("Opened Recycle Bin")
    except Exception as e:
        print(f"Error opening Recycle Bin: {e}")
        speak("Sorry, I couldn't open the Recycle Bin")

def select_all_in_recycle_bin():
    """Select all items in Recycle Bin (simulates Ctrl+A)"""
    try:
        # Get Recycle Bin window
        for window in gw.getWindowsWithTitle('Recycle Bin'):
            if 'Recycle Bin' in window.title:
                window.activate()
                time.sleep(0.5)
                pyautogui.hotkey('ctrl', 'a')
                speak("Selected all items in Recycle Bin")
                return
        speak("Couldn't find the Recycle Bin window")
    except Exception as e:
        print(f"Error selecting items: {e}")
        speak("Sorry, I couldn't select items in the Recycle Bin")

def delete_selected_items():
    """Delete selected items (simulates Delete key)"""
    try:
        # Get Recycle Bin window
        for window in gw.getWindowsWithTitle('Recycle Bin'):
            if 'Recycle Bin' in window.title:
                window.activate()
                time.sleep(0.5)
                pyautogui.press('delete')
                speak("Deleted selected items")
                return
        speak("Couldn't find the Recycle Bin window")
    except Exception as e:
        print(f"Error deleting items: {e}")
        speak("Sorry, I couldn't delete the selected items")

def close_recycle_bin():
    """Close the Recycle Bin window more reliably"""
    try:
        # Try multiple methods to ensure the window closes
        method_used = None
        
        # Method 1: Close via window handle
        for window in gw.getWindowsWithTitle('Recycle Bin'):
            if 'Recycle Bin' in window.title:
                window.close()
                method_used = "window close"
                break
        
        # Method 2: Send Alt+F4 if window is still open
        if method_used is None:
            for window in gw.getWindowsWithTitle('Recycle Bin'):
                if 'Recycle Bin' in window.title:
                    window.activate()
                    time.sleep(0.3)
                    pyautogui.hotkey('alt', 'f4')
                    method_used = "alt+f4"
                    time.sleep(0.5)
                    break
        
        # Method 3: Kill explorer process if needed
        if method_used is None:
            for proc in psutil.process_iter(['pid', 'name']):
                if 'explorer' in proc.info['name'].lower():
                    proc.kill()
                    method_used = "explorer restart"
                    # Restart explorer after killing it
                    subprocess.Popen('explorer.exe')
                    time.sleep(1)  # Wait for explorer to restart
                    break
        
        if method_used:
            speak(f"Closed Recycle Bin using {method_used.replace('_', ' ')} method")
        else:
            speak("Couldn't find or close the Recycle Bin window")
    except Exception as e:
        print(f"Error closing Recycle Bin: {e}")
        speak("Sorry, I encountered an error while trying to close the Recycle Bin")
        
def tell_time():
    """Natural time telling with date included"""
    now = datetime.datetime.now()
    hour = now.hour
    minute = now.minute
    
    # Time formatting
    if minute == 0:
        time_str = f"{hour} o'clock"
    elif minute < 10:
        time_str = f"{hour} oh {minute}"
    else:
        time_str = f"{hour} {minute}"
    
    # Period of day
    if 5 <= hour < 12:
        period = "in the morning"
    elif 12 <= hour < 17:
        period = "in the afternoon"
    elif 17 <= hour < 21:
        period = "in the evening"
    else:
        period = "at night"

    # Date formatting
    day = now.day
    month = now.strftime("%B")
    year = now.year
    weekday = now.strftime("%A")
    
    # Ordinal date (1st, 2nd, 3rd, etc.)
    if 4 <= day <= 20 or 24 <= day <= 30:
        day_str = f"{day}th"
    else:
        suffixes = {1: "st", 2: "nd", 3: "rd"}
        day_str = f"{day}{suffixes.get(day % 10, 'th')}"
    
    speak(f"It's {time_str} {period} on {weekday}, {month} {day_str}, {year}")

def tell_date():
    """Dedicated date reporting"""
    now = datetime.datetime.now()
    day = now.day
    month = now.strftime("%B")
    year = now.year
    weekday = now.strftime("%A")
    
    # Ordinal date
    if 4 <= day <= 20 or 24 <= day <= 30:
        day_str = f"{day}th"
    else:
        suffixes = {1: "st", 2: "nd", 3: "rd"}
        day_str = f"{day}{suffixes.get(day % 10, 'th')}"
    
    speak(f"Today is {weekday}, {month} {day_str}, {year}")

# Timer dictionary to store active timers {timer_id: {'duration': seconds, 'start_time': time.time(), 'name': 'Timer 1'}}
active_timers = {}
timer_counter = 0

def set_timer(duration_text):
    """Set a timer with natural language like '5 minutes', '30 seconds'"""
    global timer_counter
    
    try:
        duration_seconds = 0
        duration_lower = duration_text.lower().strip()

        # Normalize common spoken-number words (one, two, a minute, an hour, etc.)
        word_to_num = {
            'zero': '0', 'one': '1', 'two': '2', 'three': '3', 'four': '4',
            'five': '5', 'six': '6', 'seven': '7', 'eight': '8', 'nine': '9',
            'ten': '10', 'eleven': '11', 'twelve': '12', 'thirteen': '13',
            'fourteen': '14', 'fifteen': '15', 'sixteen': '16', 'seventeen': '17',
            'eighteen': '18', 'nineteen': '19', 'twenty': '20'
        }
        for word, num in word_to_num.items():
            duration_lower = re.sub(rf'\b{word}\b', num, duration_lower)

        # Handle phrases like "a minute", "an hour", "a second"
        duration_lower = re.sub(r'\ba\s+(minute|min|m)\b', '1 minute', duration_lower)
        duration_lower = re.sub(r'\ban\s+(hour|hr|h)\b', '1 hour', duration_lower)
        duration_lower = re.sub(r'\ba\s+(second|sec|s)\b', '1 second', duration_lower)
        
        # Parse duration (e.g., "5 minutes", "30 seconds", "1 hour and 20 minutes")
        
        # Extract hours
        hours_match = re.search(r'(\d+)\s*(?:hours?|hrs?|h)\b', duration_lower)
        if hours_match:
            duration_seconds += int(hours_match.group(1)) * 3600
        
        # Extract minutes
        minutes_match = re.search(r'(\d+)\s*(?:minutes?|mins?|m)\b', duration_lower)
        if minutes_match:
            duration_seconds += int(minutes_match.group(1)) * 60
        
        # Extract seconds
        seconds_match = re.search(r'(\d+)\s*(?:seconds?|secs?|s)\b', duration_lower)
        if seconds_match:
            duration_seconds += int(seconds_match.group(1))
        
        if duration_seconds <= 0:
            return {"error": "Could not parse timer duration. Try '5 minutes' or '30 seconds'"}
        
        # Create timer
        timer_counter += 1
        timer_id = timer_counter
        timer_name = f"Timer {timer_id}"
        
        start_time = time.time()
        active_timers[timer_id] = {
            'duration': duration_seconds,
            'start_time': start_time,
            'name': timer_name,
            'end_time': start_time + duration_seconds
        }
        
        # Format duration for speech
        if duration_seconds >= 3600:
            hours = duration_seconds // 3600
            remaining = duration_seconds % 3600
            minutes = remaining // 60
            if minutes > 0:
                speak(f"Timer set for {hours} hour and {minutes} minutes")
            else:
                speak(f"Timer set for {hours} hour")
        elif duration_seconds >= 60:
            minutes = duration_seconds // 60
            seconds = duration_seconds % 60
            if seconds > 0:
                speak(f"Timer set for {minutes} minute and {seconds} seconds")
            else:
                speak(f"Timer set for {minutes} minute")
        else:
            speak(f"Timer set for {duration_seconds} seconds")
        
        return {
            "success": True,
            "timer_id": timer_id,
            "timer_name": timer_name,
            "duration": duration_seconds,
            "end_time": start_time + duration_seconds,
            "speak": f"Timer set for {duration_seconds} seconds"
        }
    
    except Exception as e:
        print(f"Timer error: {e}")
        return {"error": f"Failed to set timer: {str(e)}"}

def get_timer_status():
    """Get status of all active timers"""
    current_time = time.time()
    timer_statuses = []
    
    # Check for completed timers
    completed_timers = []
    for timer_id, timer_info in list(active_timers.items()):
        elapsed = current_time - timer_info['start_time']
        remaining = timer_info['duration'] - elapsed
        
        if remaining <= 0:
            # Timer completed
            play_timer_completion_alert()
            speak(f"{timer_info['name']} is done!")
            timer_statuses.append({
                "id": timer_id,
                "name": timer_info['name'],
                "status": "completed",
                "remaining": 0
            })
            completed_timers.append(timer_id)
        else:
            # Timer still running
            timer_statuses.append({
                "id": timer_id,
                "name": timer_info['name'],
                "status": "running",
                "remaining": int(remaining),
                "total": timer_info['duration']
            })
    
    # Remove completed timers
    for timer_id in completed_timers:
        del active_timers[timer_id]
    
    return timer_statuses

def cancel_timer(timer_id=None):
    """Cancel a specific timer or all timers"""
    try:
        if timer_id is None or timer_id == "all":
            if active_timers:
                count = len(active_timers)
                active_timers.clear()
                speak(f"Cancelled all {count} timer(s)")
                return {"success": True, "message": f"Cancelled {count} timers"}
            else:
                speak("No timers are active")
                return {"error": "No timers to cancel"}
        else:
            if timer_id in active_timers:
                timer_name = active_timers[timer_id]['name']
                del active_timers[timer_id]
                speak(f"Cancelled {timer_name}")
                return {"success": True, "message": f"Cancelled {timer_name}"}
            else:
                return {"error": f"Timer {timer_id} not found"}
    except Exception as e:
        return {"error": str(e)}

def list_capabilities():
    """List all available commands"""
    capabilities = [
        "Here's what I can do for you:",
        "- Set timers: Say 'set timer for 5 minutes' or 'set timer for 30 seconds'",
        "- Get timer status: Say 'show timers' to see active timers",
        "- Cancel timers: Say 'cancel timer' or 'cancel all timers'",
        "- Tell you the current time and date",
        "- Open Microsoft Office apps: Word, Excel, PowerPoint, OneNote, Outlook",
        "- Open Notepad and WordPad",
        "- Open websites: YouTube, Google, Youtube Music, Spotify",
        "- Open WhatsApp",
        "- Close applications and browser tabs",
        "- Manage Recycle Bin: open, empty, select all, delete",
        "- Search the web for anything",
        "- Play music (try 'play music')",
        "- Tell you a joke",
        "- Shut down or restart your computer",
        "- Provide weather updates for any city",
        "- Open Calculator",
        "- Get latest news headlines (try 'what's the news?' or 'sports news')",
        "- Set reminders (try 'remind me to take pills in 1 hour')",
        "- Specify categories like business, technology, sports, etc.",
        "- Request specific number of headlines (3, 5, or 10)",
        "- Use different news sources (say 'use GNews' or 'use NewsAPI')",
        "- And more!",
        "",
        "You can ask me 'what can you do' anytime to see this list again."
    ]
    speak('\n'.join(capabilities), wait=True)

def tell_joke():
    """Tell a random joke"""
    jokes = [
        "Why don't scientists trust atoms? Because they make up everything!",
        "Did you hear about the mathematician who's afraid of negative numbers? He'll stop at nothing to avoid them.",
        "Why don't skeletons fight each other? They don't have the guts.",
        "The cemetery is so popular... people are dying to get in there.",
        "Today I decided to visit my childhood home. I asked the residents if I could come inside because I was feeling nostalgic, but they refused and slammed the door in my face. My parents are the worst.",
        "I was reading a book about helium. I just couldn't put it down.",
        "I'm reading a book about anti-gravity. It's impossible to put down!",
        "Did you hear about the claustrophobic astronaut? He just needed a little space.",
        "Why do programmers prefer dark mode? Because light attracts bugs.",
        "I finally found my purpose in life... just kidding, I'm still scrolling through Netflix.",
        "They say 'live every day like it's your last' - which explains why I haven't started my taxes.",
        "My therapist asked what I'd do differently if I had unlimited time. I said I'd wait until Thursday to kill myself.",
        "I told my wife she was drawing her eyebrows too high. She looked surprised.",
        "The coroner called my dad's autopsy results 'inconclusive'... which is what dad always said about my career choices.",
        "I asked God for a bike, but I know God doesn't work that way. So I stole a bike and asked for forgiveness.",
        "The only thing keeping me from ending it all is knowing I'd miss the next season of my favorite show.",
        "They found my suicide note and laughed. Good thing I didn't include the punchline.",
        "I used to think life was meaningless... then I realized that realization was meaningless too.",
        "My gravestone will say: 'Finally found peace'... because I'll be too dead to correct them.",
        "The doctor said my blood type is 'Never Positive'. Explains why I'm O-negative about everything.",
        "I donated my body to science... but they returned it saying 'no observable signs of intelligence'.",
        "I've started charging my phone in the bathroom - that way if I slip in the shower, at least my corpse will be at 100%.",
        "They say 'suicide is a permanent solution to a temporary problem'... but have you met my landlord?",
        "I finally achieved inner peace... right after the overdose kicked in.",
        "The best part about being dead? No more 'live, laugh, love' signs.",
        "My will states: 'If you're reading this, you've wasted your life watching me die'.",
        "I told my mom I wanted to be a nihilist when I grow up. She said 'whatever'.",
        "The good news? After I die, climate change won't bother me. The bad news? It never did.",
        "I put my existential dread in a box... then realized the box was also meaningless.",
        "They asked what I'd bring to a desert island. I said cyanide - it solves both food and shelter problems.",
        "My autobiography will be titled: 'I Told You I Was Sick'... with empty pages after chapter 1.",
        "I don't fear death... I fear the 3 minutes of awkward silence before it happens.",
        "The only thing worse than dying alone? Realizing you've been alone the whole time.",
        "I asked Death for more time. He laughed and said 'That's what everyone asks for... right before they waste it'.",
        "My wife accused me of being immature. I told her to get out of my fort.",
        "I told my computer I needed a break, and now it won't stop sending me Kit-Kat ads.",
        "I asked my Alexa if it could help me with my homework. It said, 'I'm sorry, I'm not programmed to do that.' Then it ordered a pizza.",
        "I have a stepladder because my real ladder left when I was 5.",
        "I told my therapist I keep thinking about suicide. He told me from now on I have to pay in advance.",
        "I was wondering why the frisbee was getting bigger, then it hit me.",
        "I have an EpiPen. My friend gave it to me when he was dying. It seemed very important to him that I have it.",
        "I used to play piano by ear, but now I use my hands.",
        "Parallel lines have so much in common. It is a shame they all never meet."
    ]
    selected_joke = random.choice(jokes)
    speak(selected_joke)
    return selected_joke


def take_screenshot():
    """Capture the current screen and save it to configured screenshot folder."""
    try:
        default_dir = r'C:\Users\hp\OneDrive\Pictures\Screenshots'
        screenshot_dir = os.getenv('SCREENSHOT_SAVE_DIR', default_dir).strip() or default_dir
        os.makedirs(screenshot_dir, exist_ok=True)

        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"screenshot_{timestamp}.png"
        file_path = os.path.join(screenshot_dir, filename)

        image = pyautogui.screenshot()
        image.save(file_path)

        speak("Screenshot captured and saved")
        return {
            'success': True,
            'filename': filename,
            'path': file_path
        }
    except Exception as e:
        print(f"Screenshot error: {e}")
        return {
            'success': False,
            'error': str(e)
        }

def play_music():
    """Play music from a default directory"""
    music_dir = os.path.join(os.path.expanduser('~'), 'Music')
    if os.path.exists(music_dir):
        try:
            speak("Playing music from your Music folder")
            os.startfile(music_dir)
        except Exception as e:
            speak("Sorry, I couldn't play music. Please check your Music folder.")
    else:
        speak("I couldn't find your Music folder. Where do you keep your music files?")

def system_shutdown(action='shutdown'):
    """Shutdown or restart the computer"""
    try:
        if action == 'shutdown':
            speak("Shutting down the computer in 30 seconds. Say cancel to abort.")
            time.sleep(30)
            os.system("shutdown /s /t 1")
        elif action == 'restart':
            speak("Restarting the computer in 30 seconds. Say cancel to abort.")
            time.sleep(30)
            os.system("shutdown /r /t 1")
    except Exception as e:
        speak("Sorry, I couldn't perform the system operation.")

def greet_user():
    """Personalized greeting"""
    hour = datetime.datetime.now().hour
    username = os.getlogin()
    
    greetings = [
        f"Hello {username}, how can I assist you today?",
        f"Hi {username}, what can I do for you?",
        f"Ready to help, {username}. What do you need?"
    ]
    
    if hour < 5:
        greetings.append(f"Up late, {username}? How can I help?")
    elif hour > 22:
        greetings.append(f"Working late tonight, {username}?")
    
    speak(random.choice(greetings))

def execute_command(query):
    """Handle all commands"""
    global assistant_active, processing
    
    if not query:
        return

    query = normalize_spoken_text(query)
    heard_wake_word, command_after_wake = extract_command_after_wake_word(query)
    if heard_wake_word:
        if command_after_wake:
            query = command_after_wake
        else:
            speak("Yes, tell me your command.")
            return
    
    processing = True
    query_norm = normalize_spoken_text(query)
    query_compact = query_norm.replace(' ', '')
    whatsapp_aliases = ('whatsapp', 'whatsup', 'whatsap', 'watsapp')
    has_whatsapp = any(alias in query_compact for alias in whatsapp_aliases)
    
    try:
        # Microsoft Office applications
        if 'open word' in query or 'open microsoft word' in query:
            open_office_app('word')
            
        # Add this new condition for weather requests
        elif 'weather' in query or 'temperature' in query:
            handle_weather_request(query)
            
        elif 'open excel' in query or 'open microsoft excel' in query:
            open_office_app('excel')
            
        elif 'open powerpoint' in query or 'open microsoft powerpoint' in query or 'open ppt' in query or 'open power point' in query:
            open_office_app('powerpoint')
            
        elif 'open onenote' in query or 'open microsoft onenote' in query or 'open one note' in query:
            open_office_app('onenote')
            
        elif 'open outlook' in query or 'open microsoft outlook' in query or 'open out look' in query:
            open_office_app('outlook')
            
        elif 'open notepad' in query or 'open note pad' in query:
            open_office_app('notepad')
            
        elif 'open wordpad' in query or 'open word pad' in query:
            open_office_app('wordpad')

        elif 'open calculator' in query or 'open calc' in query or 'open calc culator' in query:
            open_office_app('calculator')

        elif ((('open' in query_norm or 'start' in query_norm or 'launch' in query_norm) and has_whatsapp)
              or query_norm in ('whatsapp', 'whats app')):
            open_whatsapp_app()
            
        # Close commands for Office apps
        elif 'close word' in query:
            close_application('word')
            
        elif 'close excel' in query:
            close_application('excel')
            
        elif 'close powerpoint' in query or 'close ppt' in query or 'close power point' in query:
            close_application('powerpoint')
            
        elif 'close onenote' in query or 'close one note' in query:
            close_application('onenote')
            
        elif 'close outlook' in query or 'close out look' in query:
            close_application('outlook')
            
        elif 'close notepad' in query or 'close note pad' in query:
            close_application('notepad')
            
        elif 'close wordpad' in query or 'close word pad' in query:
            close_application('wordpad')
            
        elif 'close calculator' in query or 'close calc' in query or 'close calc culator' in query:
            close_application('calculator')
            
        # New music commands
        elif 'open spotify' in query:
            open_music_app('spotify')
            
        elif 'open youtube music' in query or 'open yt music' in query:
            open_music_app('youtube music')
            
        elif 'play music' in query:
            play_music()
        
        #News requests
        elif any(word in query for word in ['news', 'headlines', 'headline', 'current affairs']):
            handle_news_request(query)
            
        # New close commands for music apps
        elif 'close spotify' in query:
            close_application('spotify')
            
        elif 'close youtube music' in query or 'close yt music' in query:
            close_application('youtube music')
            
        # Existing commands from your original code
        elif 'open youtube' in query:
            speak("Opening YouTube")
            webbrowser.open("https://youtube.com")
            
        elif 'open google' in query:
            speak("Opening Google")
            webbrowser.open("https://google.com")
            
        elif 'close youtube' in query:
            close_application('youtube')
            
        elif 'close google' in query:
            close_application('google')
            
        elif 'close browser' in query:
            close_application('edge')
            
        elif 'close' in query and ('application' in query or 'app' in query):
            speak("Which application would you like to close?")
            app_to_close = take_command()
            if app_to_close and app_to_close != "error":
                close_application(app_to_close)
                
        # Recycle Bin Commands
        elif 'open recycle bin' in query or 'open bin' in query:
            open_recycle_bin()
            
        elif 'empty recycle bin' in query or 'empty bin' in query:
            empty_recycle_bin()
            
        elif 'select all' in query and ('recycle bin' in query or 'bin' in query):
            select_all_in_recycle_bin()
            
        elif 'delete selected' in query or 'delete all' in query:
            delete_selected_items()
            
        elif 'close recycle bin' in query or 'close bin' in query:
            close_recycle_bin()
            
        elif 'time' in query or 'what time is it' in query or 'what is current time' in query:
            tell_time()
        
        elif 'date' in query or 'what date is it' in query or 'what day is it' in query or 'what is todays date' in query:
            tell_date()
            
        elif 'search' in query:
            speak("What would you like me to search for?")
            search_term = take_command()
            if search_term and search_term != "error":
                webbrowser.open(f"https://google.com/search?q={search_term}")
            
        elif 'play music' in query:
            play_music()
            
        elif 'tell me a joke' in query or 'make me laugh' in query:
            tell_joke()
            
        elif 'shutdown' in query or 'turn off' in query:
            system_shutdown('shutdown')
            
        elif 'restart' in query or 'reboot' in query:
            system_shutdown('restart')
            
        elif 'what can you do' in query or 'help' in query or 'what are your features' in query or 'list capabilities' in query or 'list features' in query or 'what can you do for me' in query :
            list_capabilities()
            
        elif any(word in query for word in ['exit', 'quit', 'stop', 'goodbye', 'bye', 'Close yourself']):
            farewells = [
                "Goodbye! Have a great day.",
                "Signing off. Call me if you need anything.",
                "Until next time!",
                "See you later! I'm here if you need me.",
                "Take care! I'll be here when you need me.",
                "It was nice assisting you. Goodbye!",
                "That's all for now ! hope to meet you soon ! till then Goodbye!",
                "Assistant shutting down."
            ]
            speak(random.choice(farewells))
            assistant_active = False
            
        else:
            responses = [
                "I'm not programmed for that yet. Say 'what can you do' to see my capabilities.",
                "I don't understand that command. Ask me 'what can you do' for options.",
                "Could you try something else? Say 'help' to see what I can do.",
                "That's beyond my current skills. I can show you my capabilities if you ask."
            ]
            speak(random.choice(responses))
            
    except Exception as e:
        print(f"Command error: {e}")
        speak("Sorry, I encountered an error processing that request.")
    finally:
        processing = False

def main():
    """Main assistant loop"""
    print("DEBUG: Script starting")  # Add this line
    try:
        print("DEBUG: Checking microphone")  # Add this line
        if not check_microphone():
            speak("Microphone not detected. Please check your audio settings.")
            input("Press Enter to exit...")
            return
        
        print("DEBUG: Greeting user")  # Add this line
        greet_user()
        
        while assistant_active:
            print("DEBUG: Waiting for command")  # Add this line
            command = take_command(wait_for_wake_word=True, silence_reply=True)
            if command == "error":
                break
            if not command:
                continue
            execute_command(command)
            
            # Small delay between commands
            time.sleep(0.5)
    except Exception as e:
        print(f"DEBUG: Main loop error: {e}")  # Add this line
        speak("Assistant shutting down unexpectedly.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        speak("Assistant shutting down unexpectedly.")
    finally:
        print("Assistant session ended")