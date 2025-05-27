import tkinter
import tkinter.filedialog
import customtkinter as ctk
import edge_tts
import asyncio
import threading
import os
import tempfile
import json
from playsound import playsound # Using playsound 1.2.2
import time # For small delay in search
import docx  # For DOCX files
import chardet  # For detecting text file encodings
import tkinter.ttk as ttk
import pygame  # For advanced audio playback
from voice_cache import load_cached_voices, save_voices_to_cache, get_cache_status

# --- Global Variables ---
WINDOW_TITLE = "🎙️ Edge TTS Studio"  # More professional name
WINDOW_SIZE = "1000x700"  # Larger initial size
DEFAULT_APPEARANCE_MODE = "System"
DEFAULT_COLOR_THEME = "blue"
TEMP_AUDIO_FILENAME = "temp_audio_edge_tts1.mp3"  # Keep MP3 as default for temp files
DEFAULT_TEXT = "Hello, this is a test of Microsoft Edge Text-to-Speech with CustomTkinter."
DEFAULT_VOICE = "JennyNeural (en-US)"  # Default voice to select when loading voices
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".edge_tts_gui_config.json")  # Config file in user's home directory

# Color scheme
COLORS = {
    "primary": "#0078D4",      # Microsoft blue
    "secondary": "#2B88D8",    # Lighter blue
    "accent": "#00B294",       # Teal
    "error": "#E74C3C",        # Red
    "success": "#2ECC71",      # Green
    "warning": "#F39C12",      # Orange
    "background_light": "#F5F5F5",
    "background_dark": "#1E1E1E",
    "text_light": "#2D2D2D",
    "text_dark": "#FFFFFF"
}

# Icons (Unicode)
ICONS = {
    "MICROPHONE": "🎙️",
    "SPEAKER": "🔊",
    "SAVE": "💾",
    "STOP": "⏹️",
    "FILE": "📄",
    "THEME": "🌓",
    "STATUS": "ℹ️",
    "CACHE": "📦",
    "VOICE": "🗣️",
    "CLOCK": "🕒",
    "COUNT": "🔢",
    "PLAY": "▶️",
    "PAUSE": "⏸️",
    "RESUME": "⏵️",
    "VOLUME": "🔈",
    "SEEK": "⏩",
    "TEXT": "📝",
    "LOAD": "📂",
}

# Supported input file formats
SUPPORTED_INPUT_FORMATS = [
    ("Text files", "*.txt"),
    ("Word documents", "*.docx"),
    ("Rich Text Format", "*.rtf"),
    ("All files", "*.*")
]

# Supported audio formats and their extensions
SUPPORTED_FORMATS = [
    ("MP3 audio file", "*.mp3"),
    ("WAV audio file", "*.wav"),
    ("OGG audio file", "*.ogg"),
    ("M4A audio file", "*.m4a"),
    ("All files", "*.*")
]

# Format to MIME type mapping
FORMAT_MIME_TYPES = {
    ".mp3": "audio/mp3",
    ".wav": "audio/wav",
    ".ogg": "audio/ogg",
    ".m4a": "audio/m4a"
}

# Add this near the top, after global variables
LOCALE_NAME_MAP = {
    "af-ZA": "Afrikaans (South Africa)",
    "sq-AL": "Albanian (Albania)",
    "am-ET": "Amharic (Ethiopia)",
    "ar-AE": "Arabic (United Arab Emirates)",
    "ar-BH": "Arabic (Bahrain)",
    "ar-DZ": "Arabic (Algeria)",
    "ar-EG": "Arabic (Egypt)",
    "ar-IQ": "Arabic (Iraq)",
    "ar-JO": "Arabic (Jordan)",
    "ar-KW": "Arabic (Kuwait)",
    "ar-LB": "Arabic (Lebanon)",
    "ar-LY": "Arabic (Libya)",
    "ar-MA": "Arabic (Morocco)",
    "ar-OM": "Arabic (Oman)",
    "ar-QA": "Arabic (Qatar)",
    "ar-SA": "Arabic (Saudi Arabia)",
    "ar-SY": "Arabic (Syria)",
    "ar-TN": "Arabic (Tunisia)",
    "ar-YE": "Arabic (Yemen)",
    "az-AZ": "Azerbaijani (Azerbaijan)",
    "bn-BD": "Bengali (Bangladesh)",
    "bn-IN": "Bengali (India)",
    "bs-BA": "Bosnian (Bosnia and Herzegovina)",
    "bg-BG": "Bulgarian (Bulgaria)",
    "ca-ES": "Catalan (Spain)",
    "zh-CN": "Chinese (Mandarin, Simplified)",
    "zh-CN-liaoning": "Chinese (Mandarin, Liaoning)",
    "zh-CN-shaanxi": "Chinese (Mandarin, Shaanxi)",
    "zh-HK": "Chinese (Cantonese, Hong Kong)",
    "zh-TW": "Chinese (Mandarin, Taiwan)",
    "hr-HR": "Croatian (Croatia)",
    "cs-CZ": "Czech (Czech Republic)",
    "da-DK": "Danish (Denmark)",
    "nl-BE": "Dutch (Belgium)",
    "nl-NL": "Dutch (Netherlands)",
    "en-AU": "English (Australia)",
    "en-CA": "English (Canada)",
    "en-GB": "English (United Kingdom)",
    "en-HK": "English (Hong Kong)",
    "en-IE": "English (Ireland)",
    "en-IN": "English (India)",
    "en-KE": "English (Kenya)",
    "en-NG": "English (Nigeria)",
    "en-NZ": "English (New Zealand)",
    "en-PH": "English (Philippines)",
    "en-SG": "English (Singapore)",
    "en-TZ": "English (Tanzania)",
    "en-US": "English (United States)",
    "en-ZA": "English (South Africa)",
    "et-EE": "Estonian (Estonia)",
    "fil-PH": "Filipino (Philippines)",
    "fi-FI": "Finnish (Finland)",
    "fr-BE": "French (Belgium)",
    "fr-CA": "French (Canada)",
    "fr-CH": "French (Switzerland)",
    "fr-FR": "French (France)",
    "gl-ES": "Galician (Spain)",
    "ka-GE": "Georgian (Georgia)",
    "de-AT": "German (Austria)",
    "de-CH": "German (Switzerland)",
    "de-DE": "German (Germany)",
    "el-GR": "Greek (Greece)",
    "gu-IN": "Gujarati (India)",
    "he-IL": "Hebrew (Israel)",
    "hi-IN": "Hindi (India)",
    "hu-HU": "Hungarian (Hungary)",
    "is-IS": "Icelandic (Iceland)",
    "id-ID": "Indonesian (Indonesia)",
    "it-IT": "Italian (Italy)",
    "iu-Cans-CA": "Inuktitut (Canada, Syllabics)",
    "iu-Latn-CA": "Inuktitut (Canada, Latin)",
    "ja-JP": "Japanese (Japan)",
    "jv-ID": "Javanese (Indonesia)",
    "kn-IN": "Kannada (India)",
    "kk-KZ": "Kazakh (Kazakhstan)",
    "km-KH": "Khmer (Cambodia)",
    "ko-KR": "Korean (South Korea)",
    "lo-LA": "Lao (Laos)",
    "lv-LV": "Latvian (Latvia)",
    "lt-LT": "Lithuanian (Lithuania)",
    "mk-MK": "Macedonian (North Macedonia)",
    "ms-MY": "Malay (Malaysia)",
    "ml-IN": "Malayalam (India)",
    "mt-MT": "Maltese (Malta)",
    "mr-IN": "Marathi (India)",
    "mn-MN": "Mongolian (Mongolia)",
    "my-MM": "Burmese (Myanmar)",
    "ne-NP": "Nepali (Nepal)",
    "nb-NO": "Norwegian Bokmål (Norway)",
    "fa-IR": "Persian (Iran)",
    "pl-PL": "Polish (Poland)",
    "pt-BR": "Portuguese (Brazil)",
    "pt-PT": "Portuguese (Portugal)",
    "pa-IN": "Punjabi (India)",
    "ps-AF": "Pashto (Afghanistan)",
    "ro-RO": "Romanian (Romania)",
    "ru-RU": "Russian (Russia)",
    "si-LK": "Sinhala (Sri Lanka)",
    "sk-SK": "Slovak (Slovakia)",
    "sl-SI": "Slovenian (Slovenia)",
    "so-SO": "Somali (Somalia)",
    "es-AR": "Spanish (Argentina)",
    "es-BO": "Spanish (Bolivia)",
    "es-CL": "Spanish (Chile)",
    "es-CO": "Spanish (Colombia)",
    "es-CR": "Spanish (Costa Rica)",
    "es-CU": "Spanish (Cuba)",
    "es-DO": "Spanish (Dominican Republic)",
    "es-EC": "Spanish (Ecuador)",
    "es-ES": "Spanish (Spain)",
    "es-GQ": "Spanish (Equatorial Guinea)",
    "es-GT": "Spanish (Guatemala)",
    "es-HN": "Spanish (Honduras)",
    "es-MX": "Spanish (Mexico)",
    "es-NI": "Spanish (Nicaragua)",
    "es-PA": "Spanish (Panama)",
    "es-PE": "Spanish (Peru)",
    "es-PR": "Spanish (Puerto Rico)",
    "es-PY": "Spanish (Paraguay)",
    "es-SV": "Spanish (El Salvador)",
    "es-US": "Spanish (United States)",
    "es-UY": "Spanish (Uruguay)",
    "es-VE": "Spanish (Venezuela)",
    "su-ID": "Sundanese (Indonesia)",
    "sw-KE": "Swahili (Kenya)",
    "sw-TZ": "Swahili (Tanzania)",
    "sv-SE": "Swedish (Sweden)",
    "ta-IN": "Tamil (India)",
    "ta-LK": "Tamil (Sri Lanka)",
    "ta-MY": "Tamil (Malaysia)",
    "ta-SG": "Tamil (Singapore)",
    "te-IN": "Telugu (India)",
    "th-TH": "Thai (Thailand)",
    "tr-TR": "Turkish (Turkey)",
    "uk-UA": "Ukrainian (Ukraine)",
    "ur-IN": "Urdu (India)",
    "ur-PK": "Urdu (Pakistan)",
    "uz-UZ": "Uzbek (Uzbekistan)",
    "vi-VN": "Vietnamese (Vietnam)",
    "zu-ZA": "Zulu (South Africa)",
}

class EdgeTTSApp(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Initialize pygame mixer
        pygame.mixer.init()
        
        self.title(WINDOW_TITLE)
        self.geometry(WINDOW_SIZE)
        self.minsize(800, 600)  # Set minimum window size
        
        ctk.set_appearance_mode(DEFAULT_APPEARANCE_MODE)
        ctk.set_default_color_theme(DEFAULT_COLOR_THEME)

        # Audio playback state
        self.is_paused = False
        self.current_audio_file = None
        self.audio_length = 0
        self.update_progress_id = None
        
        # Configure grid layout (3x1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Create sidebar frame with widgets
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=10)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew", padx=10, pady=10)
        self.sidebar_frame.grid_rowconfigure(1, weight=1)  # Adjust row weight for text frame
        self.sidebar_frame.grid_columnconfigure(0, weight=1)

        # Create header frame for logo and theme button
        header_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        header_frame.grid_columnconfigure(1, weight=1)  # Make space between logo and button expandable

        # App logo/name
        self.logo_label = ctk.CTkLabel(
            header_frame, 
            text=WINDOW_TITLE,
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=COLORS["primary"]
        )
        self.logo_label.grid(row=0, column=0, sticky="w")

        # Theme toggle button next to logo
        self.theme_button = ctk.CTkButton(
            header_frame,
            text=f"{ICONS['THEME']} Toggle Theme",
            command=self.toggle_theme,
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            width=120
        )
        self.theme_button.grid(row=0, column=1, sticky="e")

        # Text input section in sidebar
        text_frame = ctk.CTkFrame(self.sidebar_frame, corner_radius=10)
        text_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        text_frame.grid_columnconfigure(0, weight=1)
        text_frame.grid_rowconfigure(1, weight=1)  # Make text input row expandable

        # Header frame with modern styling
        header_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        header_frame.grid_columnconfigure(1, weight=1)

        # Text input header with icon and modern font
        header_label = ctk.CTkLabel(
            header_frame,
            text=f"{ICONS['TEXT']} Text Input",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS["primary"]
        )
        header_label.grid(row=0, column=0, sticky="w")

        # Load file button with improved styling
        self.load_file_button = ctk.CTkButton(
            header_frame,
            text=f"{ICONS['LOAD']} Load File",
            width=120,
            height=32,
            command=self.on_load_file,
            font=ctk.CTkFont(size=13),
            fg_color=COLORS["secondary"],
            hover_color=COLORS["primary"]
        )
        self.load_file_button.grid(row=0, column=2, padx=(10, 0))

        # Text input with modern styling
        self.text_input = ctk.CTkTextbox(
            text_frame,
            wrap="word",
            font=ctk.CTkFont(size=13, family="Segoe UI"),
            border_width=2,
            border_color=COLORS["primary"],
            fg_color=("white", "gray10"),
            text_color=("gray10", "white")
        )
        self.text_input.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.text_input.insert("1.0", DEFAULT_TEXT)

        # Stats frame for word and character count
        stats_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        stats_frame.grid(row=2, column=0, sticky="e", padx=10, pady=(0, 5))

        # Word count label
        self.word_count_label = ctk.CTkLabel(
            stats_frame,
            text="Words: 0",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray60")
        )
        self.word_count_label.grid(row=0, column=0, padx=(0, 15))

        # Character count label
        self.char_count_label = ctk.CTkLabel(
            stats_frame,
            text="Characters: 0",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray60")
        )
        self.char_count_label.grid(row=0, column=1)

        # Bind text changes to update both counters
        self.text_input.bind('<KeyRelease>', self.update_text_stats)
        self.update_text_stats(None)  # Initial count

        # Create main content frame
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=1, rowspan=4, sticky="nsew", padx=(0, 10), pady=10)
        self.main_frame.grid_rowconfigure(3, weight=1)  # Give extra space to the last row
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Initialize other components
        self.voices_list_full = []
        self.voice_map = {}
        self.display_voices_full = []
        self.last_selected_voice = self.load_config().get('last_voice', DEFAULT_VOICE)
        self.is_speaking = False
        self.stop_requested = threading.Event()

        # Setup UI components
        self.setup_voice_selection()
        self.setup_controls()
        self.setup_status_section()

        self.load_initial_voices()

    def adjust_window_size(self):
        """Adjust window size to fit components"""
        self.update_idletasks()  # Ensure all widgets are rendered
        width = self.winfo_width()  # Keep current width
        height = self.main_frame.winfo_reqheight() + 40  # Add padding
        self.geometry(f"{width}x{height}")

    def update_detailed_status(self, message, cache_info=None):
        """Update all status components with detailed information"""
        self.loading_status.configure(text=f"{ICONS['SPEAKER']} Status: {message}")
        
        if cache_info:
            self.cache_status.configure(
                text=f"{ICONS['CACHE']} Cache: {cache_info['message']}"
                + (f" (Expires in: {cache_info['expires_in']})" if cache_info['expires_in'] else "")
            )
            
            if cache_info.get('voice_count'):
                self.voice_count.configure(text=f"{ICONS['COUNT']} Voices: {cache_info['voice_count']}")
            
            if cache_info.get('last_updated'):
                self.last_updated.configure(text=f"{ICONS['CLOCK']} Last Updated: {cache_info['last_updated']}")

    def load_initial_voices(self):
        """Initialize voice loading process"""
        self.update_detailed_status("Starting voice load process...")
        self.progress_bar.set(0.1)
        self.speak_button.configure(state="disabled")
        self.save_button.configure(state="disabled")
        threading.Thread(target=self.load_voices_threaded, daemon=True).start()

    def load_voices_threaded(self):
        try:
            # Check cache status
            cache_info = get_cache_status()
            self.after(0, self.update_detailed_status, "Checking cache...", cache_info)
            self.after(0, self.progress_bar.set, 0.2)

            # Try loading from cache
            cached_voices = load_cached_voices()
            if cached_voices:
                self.after(0, self.update_detailed_status, "Loading from cache...", cache_info)
                self.after(0, self.progress_bar.set, 0.6)
                self.voices_list_full = cached_voices
                self.after(0, self.process_loaded_voices)
                return

            # If no cache, load from network
            self.after(0, self.update_detailed_status, "Cache not available, fetching from network...", cache_info)
            self.after(0, self.progress_bar.set, 0.3)

            async def get_voices_async():
                return await edge_tts.VoicesManager.create()

            self.after(0, self.update_detailed_status, "Connecting to Microsoft Edge TTS service...", cache_info)
            self.after(0, self.progress_bar.set, 0.4)
            
            voices_manager = asyncio.run(get_voices_async())
            self.voices_list_full = voices_manager.voices
            
            self.after(0, self.update_detailed_status, "Saving to cache...", cache_info)
            self.after(0, self.progress_bar.set, 0.7)
            
            # Save to cache and get updated status
            save_voices_to_cache(self.voices_list_full)
            cache_info = get_cache_status()
            
            self.after(0, self.process_loaded_voices)
            self.after(0, self.update_detailed_status, "Processing voices...", cache_info)
            self.after(0, self.progress_bar.set, 0.9)

        except Exception as e:
            error_msg = f"Error loading voices: {str(e)}"
            self.after(0, self.update_detailed_status, error_msg, cache_info)
            self.after(0, self.progress_bar.set, 0)
            self.after(0, self.voice_combobox.configure, {"values": ["Error loading voices"]})
            self.voice_combobox.set("Error loading voices")

    def process_loaded_voices(self):
        """Process loaded voices and update UI"""
        try:
            self.progress_bar.set(0.95)

            # Sort voices: by mapped language name, then gender, then first name
            def sort_key(v):
                locale_name = LOCALE_NAME_MAP.get(v['Locale'], v['Locale'])
                friendly_name = v['FriendlyName']
                if friendly_name.startswith("Microsoft "):
                    friendly_name = friendly_name[len("Microsoft ") :]
                first_name = friendly_name.split()[0]
                return (locale_name, v['Gender'], first_name)

            self.voices_list_full.sort(key=sort_key)

            self.voice_map.clear()
            self.display_voices_full.clear()

            for voice in self.voices_list_full:
                locale_name = LOCALE_NAME_MAP.get(voice['Locale'], voice['Locale'])
                friendly_name = voice['FriendlyName']
                if friendly_name.startswith("Microsoft "):
                    friendly_name = friendly_name[len("Microsoft ") :]
                first_name = friendly_name.split()[0]
                display_name = f"{locale_name} - {voice['Gender']} - {first_name}"
                self.voice_map[display_name] = voice['Name']
                self.display_voices_full.append(display_name)

            cache_info = get_cache_status()
            self.update_detailed_status("Ready", cache_info)
            self.progress_bar.set(1.0)
            self.update_voice_combobox_post_load()

        except Exception as e:
            error_msg = f"Error processing voices: {str(e)}"
            self.update_detailed_status(error_msg)
            self.progress_bar.set(0)
            self.voice_combobox.configure(values=["Error processing voices"])
            self.voice_combobox.set("Error processing voices")

    def load_config(self):
        """Load configuration from file"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading config: {e}")
        return {}

    def save_config(self):
        """Save configuration to file"""
        try:
            config = {
                'last_voice': self.voice_combobox.get()
            }
            os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            print(f"Error saving config: {e}")

    def update_voice_combobox_post_load(self):
        """Update the combobox with loaded voices"""
        if self.display_voices_full:
            # Update the combobox values
            self.voice_combobox.configure(values=self.display_voices_full)
            
            # Try to set last selected voice, fallback to default or first voice
            default_voice = next(
                (v for v in self.display_voices_full if self.last_selected_voice in v),
                next(
                    (v for v in self.display_voices_full if DEFAULT_VOICE in v),
                    self.display_voices_full[0]
                )
            )
            self.voice_combobox.set(default_voice)
            
            # Enable the combobox and buttons
            self.voice_combobox.configure(state="normal")
            self.speak_button.configure(state="normal")
            self.save_button.configure(state="normal")
            self.update_detailed_status("Voices loaded. Ready.")
        else:
            # No voices found - set appropriate messages
            no_voices_msg = "No voices found"
            self.voice_combobox.configure(values=[no_voices_msg])
            self.voice_combobox.set(no_voices_msg)
            self.voice_combobox.configure(state="disabled")
            self.speak_button.configure(state="disabled")
            self.save_button.configure(state="disabled")
            self.update_detailed_status("No voices found.")

    def on_voice_selected_from_combobox(self, choice):
        # Save the selected voice to config when changed
        self.save_config()

    def on_speak(self):
        if self.is_speaking: return

        text = self.text_input.get("1.0", "end-1c").strip()
        if not text:
            self.update_detailed_status("Error: Text input is empty.")
            return

        selected_voice_short_name = self.get_selected_voice_short_name()
        if not selected_voice_short_name:
            self.update_detailed_status("Error: No valid voice selected.")
            return

        self._set_speaking_state(True)
        self.update_detailed_status(f"Synthesizing with {selected_voice_short_name}...")

        temp_dir = tempfile.gettempdir()
        temp_audio_path = os.path.join(temp_dir, TEMP_AUDIO_FILENAME)

        def synthesis_and_playback_thread():
            try:
                success = self._synthesize_speech(text, selected_voice_short_name, temp_audio_path)

                if self.stop_requested.is_set():
                    self.after(0, self.update_detailed_status, "Speak operation stopped.")
                    if os.path.exists(temp_audio_path): os.remove(temp_audio_path)
                    return

                if success:
                    self.after(0, self.update_detailed_status, "Playing audio...")
                    if self.stop_requested.is_set():
                        self.after(0, self.update_detailed_status, "Speak operation stopped before playback.")
                        if os.path.exists(temp_audio_path): os.remove(temp_audio_path)
                        return

                    try:
                        self.current_audio_file = temp_audio_path
                        self.play_audio(temp_audio_path)
                        if not self.stop_requested.is_set():
                            self.after(0, self.update_detailed_status, "Playback finished. Ready.")
                        else:
                            self.after(0, self.update_detailed_status, "Playback stopped/skipped. Ready.")
                    except Exception as e:
                        if not self.stop_requested.is_set():
                            self.after(0, self.update_detailed_status, f"Error playing audio: {e}")
                    finally:
                        if os.path.exists(temp_audio_path) and not pygame.mixer.music.get_busy():
                            try: os.remove(temp_audio_path)
                            except Exception as e_del: print(f"Error deleting temp file: {e_del}")
            finally:
                if not self.stop_requested.is_set():
                    self.after(0, lambda: self._set_speaking_state(False))

        threading.Thread(target=synthesis_and_playback_thread, daemon=True).start()

    def play_audio(self, audio_path):
        """Play audio using pygame mixer"""
        try:
            pygame.mixer.music.load(audio_path)
            pygame.mixer.music.set_volume(self.volume_slider.get() / 100)
            pygame.mixer.music.play()
            
            # Get audio length
            audio = pygame.mixer.Sound(audio_path)
            self.audio_length = audio.get_length()
            self.total_time.configure(text=self.format_time(self.audio_length))
            
            # Start progress updates
            self.update_progress()
            
        except Exception as e:
            self.update_detailed_status(f"Error playing audio: {e}")

    def update_progress(self):
        """Update progress bar and time display"""
        if pygame.mixer.music.get_busy() and not self.stop_requested.is_set():
            current_pos = pygame.mixer.music.get_pos() / 1000  # Convert to seconds
            progress = current_pos / self.audio_length if self.audio_length > 0 else 0
            self.progress_bar.set(progress)
            self.current_time.configure(text=self.format_time(current_pos))
            self.update_progress_id = self.after(100, self.update_progress)
        else:
            self.progress_bar.set(0)
            self.current_time.configure(text="0:00")
            if self.update_progress_id:
                self.after_cancel(self.update_progress_id)
                self.update_progress_id = None

    def format_time(self, seconds):
        """Format seconds to MM:SS"""
        minutes = int(seconds // 60)
        seconds = int(seconds % 60)
        return f"{minutes}:{seconds:02d}"

    def on_pause_resume(self):
        """Handle pause/resume button click"""
        if self.is_paused:
            pygame.mixer.music.unpause()
            self.pause_button.configure(text=f"{ICONS['PAUSE']} Pause")
            self.is_paused = False
            self.update_detailed_status("Playback resumed.")
        else:
            pygame.mixer.music.pause()
            self.pause_button.configure(text=f"{ICONS['RESUME']} Resume")
            self.is_paused = True
            self.update_detailed_status("Playback paused.")

    def on_volume_change(self, value):
        """Handle volume slider change"""
        pygame.mixer.music.set_volume(float(value) / 100)

    def on_progress_click(self, event):
        """Handle click on progress bar for seeking"""
        if pygame.mixer.music.get_busy():
            # Calculate relative position
            width = self.progress_bar.winfo_width()
            relative_pos = event.x / width
            # Set position
            pygame.mixer.music.set_pos(relative_pos * self.audio_length)
            # Update display
            self.progress_bar.set(relative_pos)
            self.current_time.configure(text=self.format_time(relative_pos * self.audio_length))

    def on_save_as(self):
        if self.is_speaking: return

        text = self.text_input.get("1.0", "end-1c").strip()
        if not text:
            self.update_detailed_status("Error: Text input is empty.")
            return

        selected_voice_short_name = self.get_selected_voice_short_name()
        if not selected_voice_short_name:
            self.update_detailed_status("Error: No valid voice selected.")
            return

        filepath = tkinter.filedialog.asksaveasfilename(
            defaultextension=".mp3",
            filetypes=SUPPORTED_FORMATS,
            title="Save Speech As Audio"
        )
        if not filepath:
            self.update_detailed_status("Save cancelled. Ready.")
            return

        self._set_speaking_state(True) # Use speaking state to manage buttons
        self.update_detailed_status(f"Synthesizing and saving to {os.path.basename(filepath)}...")

        def synthesis_thread():
            try:
                success = self._synthesize_speech(text, selected_voice_short_name, filepath)
                if self.stop_requested.is_set():
                    self.after(0, self.update_detailed_status, "Save operation stopped.")
                    return

                if success:
                    self.after(0, self.update_detailed_status, f"Audio saved to {os.path.basename(filepath)}. Ready.")
            finally:
                if not self.stop_requested.is_set():
                    self.after(0, lambda: self._set_speaking_state(False))

        threading.Thread(target=synthesis_thread, daemon=True).start()

    def on_stop(self):
        """Handle stop button click"""
        self.update_detailed_status("Stop request received...")
        self.stop_requested.set()
        pygame.mixer.music.stop()
        self._set_speaking_state(False)
        self.progress_bar.set(0)
        self.current_time.configure(text="0:00")
        if self.update_progress_id:
            self.after_cancel(self.update_progress_id)
            self.update_progress_id = None
        self.update_detailed_status("Operation stopped. Ready.")

    def on_load_file(self):
        """Handle loading text from a file."""
        if self.is_speaking:
            return

        filepath = tkinter.filedialog.askopenfilename(
            filetypes=SUPPORTED_INPUT_FORMATS,
            title="Select Text File"
        )
        
        if not filepath:
            return

        try:
            text = self._read_file_content(filepath)
            if text:
                self.text_input.delete("1.0", "end")
                self.text_input.insert("1.0", text)
                self.update_detailed_status(f"Loaded text from {os.path.basename(filepath)}")
            else:
                self.update_detailed_status("Error: Could not read text from file.")
        except Exception as e:
            self.update_detailed_status(f"Error loading file: {str(e)}")

    def _read_file_content(self, filepath):
        """Read content from various file types."""
        file_ext = os.path.splitext(filepath)[1].lower()
        
        try:
            if file_ext == '.docx':
                return self._read_docx(filepath)
            elif file_ext == '.rtf':
                return self._read_rtf(filepath)
            else:  # Default to text file
                return self._read_text_file(filepath)
        except Exception as e:
            raise Exception(f"Error reading file: {str(e)}")

    def _read_text_file(self, filepath):
        """Read content from a text file with encoding detection."""
        with open(filepath, 'rb') as file:
            raw_data = file.read()
            detected = chardet.detect(raw_data)
            encoding = detected['encoding'] or 'utf-8'
            
        with open(filepath, 'r', encoding=encoding) as file:
            return file.read()

    def _read_docx(self, filepath):
        """Read content from a DOCX file."""
        doc = docx.Document(filepath)
        return '\n'.join([paragraph.text for paragraph in doc.paragraphs])

    def _read_rtf(self, filepath):
        """Read content from an RTF file."""
        # For RTF files, we'll use a simple text reading approach
        # as proper RTF parsing would require additional dependencies
        return self._read_text_file(filepath)

    def _synthesize_speech(self, text, voice_short_name, output_filepath):
        try:
            if self.stop_requested.is_set():
                self.after(0, self.update_detailed_status, "Operation stopped before synthesis.")
                return False

            communicate = edge_tts.Communicate(text, voice_short_name)
            asyncio.run(communicate.save(output_filepath))

            if self.stop_requested.is_set(): # Check again after potentially long synthesis
                self.after(0, self.update_detailed_status, "Operation stopped after synthesis, before playback/save completion.")
                if os.path.exists(output_filepath) and output_filepath.endswith(TEMP_AUDIO_FILENAME):
                    try: os.remove(output_filepath) # Clean up temp file if stop requested
                    except Exception: pass
                return False
            return True
        except Exception as e:
            self.after(0, self.update_detailed_status, f"Synthesis error: {e}")
            return False

    def _set_speaking_state(self, speaking: bool):
        self.is_speaking = speaking
        self.stop_requested.clear()
        self.is_paused = False

        if speaking:
            self.speak_button.grid_remove()
            self.save_button.grid_remove()
            self.stop_button.grid(row=0, column=0, sticky="ew", padx=5)
            self.pause_button.grid(row=0, column=1, sticky="ew", padx=5)
            self.speak_button.configure(state="disabled")
            self.save_button.configure(state="disabled")
            self.voice_combobox.configure(state="disabled")
        else:
            self.stop_button.grid_remove()
            self.pause_button.grid_remove()
            self.speak_button.grid(row=0, column=0, sticky="ew", padx=5)
            self.save_button.grid(row=0, column=2, sticky="ew", padx=5)
            self.speak_button.configure(state="normal")
            self.save_button.configure(state="normal")
            self.voice_combobox.configure(state="normal")
            self.update_idletasks()

    def get_selected_voice_short_name(self):
        selected_display_name = self.voice_combobox.get()
        if selected_display_name in ["Loading voices...", "Error loading voices", "No voices found", "No match found"]:
            return None
        return self.voice_map.get(selected_display_name)

    def setup_voice_selection(self):
        """Setup voice selection with modern UI"""
        voice_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        voice_frame.grid(row=0, column=0, sticky="new", padx=10, pady=5)
        voice_frame.grid_columnconfigure(0, weight=1)

        # Voice selection header with icon and modern font
        header_label = ctk.CTkLabel(
            voice_frame,
            text=f"{ICONS['VOICE']} Voice Selection",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS["primary"]
        )
        header_label.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))

        # Voice selection container
        selection_frame = ctk.CTkFrame(voice_frame, fg_color="transparent")
        selection_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        selection_frame.grid_columnconfigure(0, weight=1)

        # Modern combobox styling
        style = ttk.Style()
        style.configure(
            'Custom.TCombobox',
            background=COLORS["background_light"],
            fieldbackground=COLORS["background_light"],
            selectbackground=COLORS["primary"],
            selectforeground=COLORS["text_dark"],
            padding=8,
            arrowsize=15
        )

        # Voice combobox with improved styling
        self.voice_combobox = ttk.Combobox(
            selection_frame,
            values=["Loading voices..."],
            state="readonly",
            style='Custom.TCombobox',
            height=15
        )
        self.voice_combobox.grid(row=0, column=0, sticky="ew", pady=(5, 10))
        self.voice_combobox.set("Loading voices...")
        
        # Bind voice selection
        self.voice_combobox.bind('<<ComboboxSelected>>', 
            lambda e: self.on_voice_selected_from_combobox(self.voice_combobox.get()))

        # Add search functionality
        search_frame = ctk.CTkFrame(voice_frame, fg_color="transparent")
        search_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        search_frame.grid_columnconfigure(1, weight=1)

        # Search icon and label
        search_label = ctk.CTkLabel(
            search_frame,
            text="🔍",
            font=ctk.CTkFont(size=14)
        )
        search_label.grid(row=0, column=0, padx=(0, 5))

        # Search entry
        self.voice_search = ctk.CTkEntry(
            search_frame,
            placeholder_text="Search voices...",
            font=ctk.CTkFont(size=13),
            height=32
        )
        self.voice_search.grid(row=0, column=1, sticky="ew")
        
        # Bind search functionality
        self.voice_search.bind('<KeyRelease>', self.filter_voices)

    def filter_voices(self, event):
        """Filter voices based on search text"""
        search_text = self.voice_search.get().lower()
        if not search_text:
            self.voice_combobox.configure(values=self.display_voices_full)
            return

        filtered_voices = [
            voice for voice in self.display_voices_full
            if search_text in voice.lower()
        ]
        
        if filtered_voices:
            self.voice_combobox.configure(values=filtered_voices)
        else:
            self.voice_combobox.configure(values=["No match found"])
            self.voice_combobox.set("No match found")

    def setup_controls(self):
        """Setup control buttons with modern styling"""
        controls_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        controls_frame.grid(row=1, column=0, sticky="new", padx=10, pady=5)
        controls_frame.grid_columnconfigure((0, 1), weight=1)

        # Controls header
        header_label = ctk.CTkLabel(
            controls_frame,
            text=f"{ICONS['MICROPHONE']} Controls",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS["primary"]
        )
        header_label.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 5))

        # Button container
        button_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        button_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 10))
        button_frame.grid_columnconfigure((0, 1, 2), weight=1)

        # Modern button styling
        button_font = ctk.CTkFont(size=14, weight="bold")
        
        # Speak button with hover animation
        self.speak_button = ctk.CTkButton(
            button_frame,
            text=f"{ICONS['PLAY']} Speak",
            command=self.on_speak,
            font=button_font,
            height=45,
            fg_color=COLORS["primary"],
            hover_color=COLORS["secondary"],
            corner_radius=8
        )
        self.speak_button.grid(row=0, column=0, sticky="ew", padx=5)

        # Stop button (initially hidden)
        self.stop_button = ctk.CTkButton(
            button_frame,
            text=f"{ICONS['STOP']} Stop",
            command=self.on_stop,
            font=button_font,
            height=45,
            fg_color=COLORS["error"],
            hover_color="#C0392B",
            corner_radius=8
        )

        # Pause/Resume button (initially hidden)
        self.pause_button = ctk.CTkButton(
            button_frame,
            text=f"{ICONS['PAUSE']} Pause",
            command=self.on_pause_resume,
            font=button_font,
            height=45,
            fg_color=COLORS["warning"],
            hover_color="#E67E22",
            corner_radius=8
        )

        # Save button with modern styling
        self.save_button = ctk.CTkButton(
            button_frame,
            text=f"{ICONS['SAVE']} Save Audio",
            command=self.on_save_as,
            font=button_font,
            height=45,
            fg_color=COLORS["accent"],
            hover_color="#00A080",
            corner_radius=8
        )
        self.save_button.grid(row=0, column=2, sticky="ew", padx=5)

        # Add playback controls frame
        playback_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        playback_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 10))
        playback_frame.grid_columnconfigure(1, weight=1)  # Give weight to progress bar

        # Progress bar and time labels
        self.current_time = ctk.CTkLabel(playback_frame, text="0:00", font=("Helvetica", 10))
        self.current_time.grid(row=0, column=0, padx=(0, 5))

        self.progress_bar = ctk.CTkProgressBar(
            playback_frame,
            mode="determinate",
            height=10,
            corner_radius=3
        )
        self.progress_bar.grid(row=0, column=1, sticky="ew", padx=5)
        self.progress_bar.set(0)

        self.total_time = ctk.CTkLabel(playback_frame, text="0:00", font=("Helvetica", 10))
        self.total_time.grid(row=0, column=2, padx=(5, 0))

        # Volume control
        volume_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        volume_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 10))
        volume_frame.grid_columnconfigure(1, weight=1)

        volume_label = ctk.CTkLabel(
            volume_frame,
            text=f"{ICONS['VOLUME']} Volume",
            font=("Helvetica", 12)
        )
        volume_label.grid(row=0, column=0, padx=(0, 10))

        self.volume_slider = ctk.CTkSlider(
            volume_frame,
            from_=0,
            to=100,
            number_of_steps=100,
            command=self.on_volume_change
        )
        self.volume_slider.grid(row=0, column=1, sticky="ew")
        self.volume_slider.set(70)  # Default volume

        # Bind progress bar click for seeking
        self.progress_bar.bind("<Button-1>", self.on_progress_click)

    def setup_status_section(self):
        """Setup status section with modern visualization"""
        status_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        status_frame.grid(row=3, column=0, sticky="new", padx=10, pady=5)  # Changed row to 3 since controls now use row 2
        status_frame.grid_columnconfigure(0, weight=1)

        # Status header
        header_label = ctk.CTkLabel(
            status_frame,
            text=f"{ICONS['STATUS']} Status Information",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS["primary"]
        )
        header_label.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))

        # Status container
        info_frame = ctk.CTkFrame(status_frame, fg_color="transparent")
        info_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        info_frame.grid_columnconfigure(0, weight=1)

        # Status labels with modern styling
        label_font = ctk.CTkFont(size=13)
        
        self.loading_status = ctk.CTkLabel(
            info_frame,
            text=f"{ICONS['SPEAKER']} Status: Initializing...",
            anchor="w",
            font=label_font,
            text_color=("gray20", "gray80")
        )
        self.loading_status.grid(row=0, column=0, sticky="w", pady=2)

        self.cache_status = ctk.CTkLabel(
            info_frame,
            text=f"{ICONS['CACHE']} Cache: Checking...",
            anchor="w",
            font=label_font,
            text_color=("gray20", "gray80")
        )
        self.cache_status.grid(row=1, column=0, sticky="w", pady=2)

        self.voice_count = ctk.CTkLabel(
            info_frame,
            text=f"{ICONS['COUNT']} Voices: -",
            anchor="w",
            font=label_font,
            text_color=("gray20", "gray80")
        )
        self.voice_count.grid(row=2, column=0, sticky="w", pady=2)

        self.last_updated = ctk.CTkLabel(
            info_frame,
            text=f"{ICONS['CLOCK']} Last Updated: -",
            anchor="w",
            font=label_font,
            text_color=("gray20", "gray80")
        )
        self.last_updated.grid(row=3, column=0, sticky="w", pady=2)

    def toggle_theme(self):
        """Toggle between light and dark theme"""
        if ctk.get_appearance_mode() == "Dark":
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("Dark")

    def update_text_stats(self, event):
        """Update both word and character count labels"""
        text = self.text_input.get("1.0", "end-1c")
        char_count = len(text)
        # Split by whitespace and filter out empty strings
        word_count = len([word for word in text.split() if word.strip()])
        
        self.char_count_label.configure(text=f"Characters: {char_count}")
        self.word_count_label.configure(text=f"Words: {word_count}")

if __name__ == "__main__":
    app = EdgeTTSApp()
    app.mainloop()