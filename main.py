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
from voice_cache import load_cached_voices, save_voices_to_cache, get_cache_status

# --- Global Variables ---
WINDOW_TITLE = "🎙️ Edge TTS GUI"  # Added microphone icon to window title
WINDOW_SIZE = "800x600"  # Increased initial size
DEFAULT_APPEARANCE_MODE = "System"
DEFAULT_COLOR_THEME = "blue"
TEMP_AUDIO_FILENAME = "temp_audio_edge_tts1.mp3"  # Keep MP3 as default for temp files
DEFAULT_TEXT = "Hello, this is a test of Microsoft Edge Text-to-Speech with CustomTkinter."
DEFAULT_VOICE = "JennyNeural (en-US)"  # Default voice to select when loading voices
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".edge_tts_gui_config.json")  # Config file in user's home directory

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
    "bg-BG": "Bulgarian (Bulgaria)",
    "bn-BD": "Bengali (Bangladesh)",
    "bn-IN": "Bengali (India)",
    "bs-BA": "Bosnian (Bosnia and Herzegovina)",
    "ca-ES": "Catalan (Spain)",
    "cs-CZ": "Czech (Czech Republic)",
    "cy-GB": "Welsh (United Kingdom)",
    "da-DK": "Danish (Denmark)",
    "de-AT": "German (Austria)",
    "de-CH": "German (Switzerland)",
    "de-DE": "German (Germany)",
    "el-GR": "Greek (Greece)",
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
    "et-EE": "Estonian (Estonia)",
    "fa-IR": "Persian (Iran)",
    "fi-FI": "Finnish (Finland)",
    "fil-PH": "Filipino (Philippines)",
    "fr-BE": "French (Belgium)",
    "fr-CA": "French (Canada)",
    "fr-CH": "French (Switzerland)",
    "fr-FR": "French (France)",
    "ga-IE": "Irish (Ireland)",
    "gl-ES": "Galician (Spain)",
    "gu-IN": "Gujarati (India)",
    "he-IL": "Hebrew (Israel)",
    "hi-IN": "Hindi (India)",
    "hr-HR": "Croatian (Croatia)",
    "hu-HU": "Hungarian (Hungary)",
    "id-ID": "Indonesian (Indonesia)",
    "is-IS": "Icelandic (Iceland)",
    "it-IT": "Italian (Italy)",
    "iu-Cans-CA": "Inuktitut (Canada, Syllabics)",
    "iu-Latn-CA": "Inuktitut (Canada, Latin)",
    "ja-JP": "Japanese (Japan)",
    "jv-ID": "Javanese (Indonesia)",
    "ka-GE": "Georgian (Georgia)",
    "kk-KZ": "Kazakh (Kazakhstan)",
    "km-KH": "Khmer (Cambodia)",
    "kn-IN": "Kannada (India)",
    "ko-KR": "Korean (South Korea)",
    "lo-LA": "Lao (Laos)",
    "lt-LT": "Lithuanian (Lithuania)",
    "lv-LV": "Latvian (Latvia)",
    "mk-MK": "Macedonian (North Macedonia)",
    "ml-IN": "Malayalam (India)",
    "mn-MN": "Mongolian (Mongolia)",
    "mr-IN": "Marathi (India)",
    "ms-MY": "Malay (Malaysia)",
    "mt-MT": "Maltese (Malta)",
    "my-MM": "Burmese (Myanmar)",
    "nb-NO": "Norwegian Bokmål (Norway)",
    "ne-NP": "Nepali (Nepal)",
    "nl-BE": "Dutch (Belgium)",
    "nl-NL": "Dutch (Netherlands)",
    "pl-PL": "Polish (Poland)",
    "ps-AF": "Pashto (Afghanistan)",
    "pt-BR": "Portuguese (Brazil)",
    "pt-PT": "Portuguese (Portugal)",
    "ro-RO": "Romanian (Romania)",
    "ru-RU": "Russian (Russia)",
    "si-LK": "Sinhala (Sri Lanka)",
    "sk-SK": "Slovak (Slovakia)",
    "sl-SI": "Slovenian (Slovenia)",
    "so-SO": "Somali (Somalia)",
    "sq-AL": "Albanian (Albania)",
    "sr-RS": "Serbian (Serbia)",
    "su-ID": "Sundanese (Indonesia)",
    "sv-SE": "Swedish (Sweden)",
    "sw-KE": "Swahili (Kenya)",
    "sw-TZ": "Swahili (Tanzania)",
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
    "zh-CN": "Chinese (Mandarin, Simplified)",
    "zh-CN-liaoning": "Chinese (Mandarin, Liaoning)",
    "zh-CN-shaanxi": "Chinese (Mandarin, Shaanxi)",
    "zh-HK": "Chinese (Cantonese, Hong Kong)",
    "zh-TW": "Chinese (Mandarin, Taiwan)",
    "zu-ZA": "Zulu (South Africa)",
}

class EdgeTTSApp(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title(WINDOW_TITLE)
        self.geometry(WINDOW_SIZE)
        self.minsize(700, 500)  # Set minimum window size
        
        ctk.set_appearance_mode(DEFAULT_APPEARANCE_MODE)
        ctk.set_default_color_theme(DEFAULT_COLOR_THEME)

        self.voices_list_full = [] # Full list of voice dicts
        self.voice_map = {}        # Maps display name to short name
        self.display_voices_full = [] # Full list of display names for combobox
        self.last_selected_voice = self.load_config().get('last_voice', DEFAULT_VOICE)  # Load last selected voice

        self.is_speaking = False
        self.stop_requested = threading.Event() # For stopping speak/save operations

        # Create main container with padding
        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True, padx=20, pady=20)

        # Add theme toggle in top right
        self.theme_button = ctk.CTkButton(
            self.container,
            text="🌓",
            width=40,
            command=self.toggle_theme,
            fg_color="transparent",
            border_width=1
        )
        self.theme_button.pack(side="top", anchor="e", padx=10, pady=(0, 10))

        # Create main content area
        self.main_frame = ctk.CTkFrame(self.container)
        self.main_frame.pack(fill="both", expand=True)

        # Text input section with improved styling
        self.setup_text_input_section()
        
        # Voice selection section with modern styling
        self.setup_voice_selection()
        
        # Controls section with improved layout
        self.setup_controls()
        
        # Status section with better visualization
        self.setup_status_section()

        self.load_initial_voices()
        self.after(100, self.adjust_window_size)

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
            
            # Sort voices: Primarily by Locale (e.g., 'en-US'), then by FriendlyName
            self.voices_list_full.sort(key=lambda v: (v['Locale'], v['FriendlyName']))

            self.voice_map.clear()
            self.display_voices_full.clear()

            for voice in self.voices_list_full:
                locale_name = LOCALE_NAME_MAP.get(voice['Locale'], voice['Locale'])
                friendly_name = voice['FriendlyName']
                if friendly_name.startswith("Microsoft "):
                    friendly_name = friendly_name[len("Microsoft "):]
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
                    if self.stop_requested.is_set(): # Check one more time before playsound
                        self.after(0, self.update_detailed_status, "Speak operation stopped before playback.")
                        if os.path.exists(temp_audio_path): os.remove(temp_audio_path)
                        return

                    try:
                        playsound(temp_audio_path) # This blocks this thread
                        if not self.stop_requested.is_set(): # Only update if not stopped
                             self.after(0, self.update_detailed_status, "Playback finished. Ready.")
                        else:
                             self.after(0, self.update_detailed_status, "Playback stopped/skipped. Ready.")
                    except Exception as e:
                        if not self.stop_requested.is_set():
                            self.after(0, self.update_detailed_status, f"Error playing audio: {e}")
                    finally:
                        if os.path.exists(temp_audio_path):
                            try: os.remove(temp_audio_path)
                            except Exception as e_del: print(f"Error deleting temp file: {e_del}")
            finally:
                # Ensure UI is reset regardless of how the thread exits
                # if stop was requested, _set_speaking_state(False) might have been called by on_stop
                if not self.stop_requested.is_set():
                    self.after(0, lambda: self._set_speaking_state(False))
                # If stop was requested, on_stop handles resetting the UI.

        threading.Thread(target=synthesis_and_playback_thread, daemon=True).start()

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
        self.update_detailed_status("Stop request received...")
        self.stop_requested.set()
        self._set_speaking_state(False) # Immediately reset UI buttons
        # Note: This won't instantly kill the playsound or edge-tts network call if already deep into it,
        # but it will prevent new actions and update UI state.
        # The running threads will check the stop_requested event at their earliest convenience.
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
        self.stop_requested.clear() # Clear stop flag when starting a new operation

        if speaking:
            self.speak_button.pack_forget()
            self.save_button.pack_forget()
            self.stop_button.pack(side="left", padx=5, expand=True, fill="x")
            self.speak_button.configure(state="disabled")
            self.save_button.configure(state="disabled")
            self.voice_combobox.configure(state="disabled")
        else:
            self.stop_button.pack_forget()
            self.speak_button.pack(side="left", padx=5, expand=True, fill="x")
            self.save_button.pack(side="left", padx=5, expand=True, fill="x")
            self.speak_button.configure(state="normal")
            self.save_button.configure(state="normal")
            self.voice_combobox.configure(state="normal")
            self.update_idletasks() # Ensure UI updates layout changes

    def get_selected_voice_short_name(self):
        selected_display_name = self.voice_combobox.get()
        if selected_display_name in ["Loading voices...", "Error loading voices", "No voices found", "No match found"]:
            return None
        return self.voice_map.get(selected_display_name)

    def setup_text_input_section(self):
        """Setup the text input section with improved styling"""
        text_frame = ctk.CTkFrame(self.main_frame)
        text_frame.pack(fill="both", expand=True, padx=10, pady=5)

        header_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(5, 0))

        # Text input header with icon
        ctk.CTkLabel(
            header_frame,
            text=f"{ICONS['TEXT']} Text Input",
            font=("Helvetica", 14, "bold")
        ).pack(side="left")

        self.load_file_button = ctk.CTkButton(
            header_frame,
            text=f"{ICONS['LOAD']} Load File",
            width=100,
            command=self.on_load_file,
            font=("Helvetica", 12)
        )
        self.load_file_button.pack(side="right")

        # Text input with modern styling
        self.text_input = ctk.CTkTextbox(
            text_frame,
            height=150,
            wrap="word",
            font=("Helvetica", 12),
            border_width=2
        )
        self.text_input.pack(fill="both", expand=True, pady=10)
        self.text_input.insert("1.0", DEFAULT_TEXT)

    def setup_voice_selection(self):
        """Setup voice selection with improved UI"""
        voice_frame = ctk.CTkFrame(self.main_frame)
        voice_frame.pack(fill="x", padx=10, pady=5)

        # Voice selection header with icon
        ctk.CTkLabel(
            voice_frame,
            text=f"{ICONS['VOICE']} Voice Selection",
            font=("Helvetica", 14, "bold")
        ).pack(anchor="w", pady=5)

        # Modern combobox styling
        style = ttk.Style()
        style.configure(
            'Custom.TCombobox',
            background='white',
            fieldbackground='white',
            selectbackground='#0078D4',
            selectforeground='white',
            padding=5
        )

        self.voice_combobox = ttk.Combobox(
            voice_frame,
            values=["Loading voices..."],
            state="readonly",
            width=60,
            height=15,
            style='Custom.TCombobox'
        )
        self.voice_combobox.pack(fill="x", pady=5)
        self.voice_combobox.set("Loading voices...")
        self.voice_combobox.bind('<<ComboboxSelected>>', 
            lambda e: self.on_voice_selected_from_combobox(self.voice_combobox.get()))

    def setup_controls(self):
        """Setup control buttons with modern styling"""
        controls_frame = ctk.CTkFrame(self.main_frame)
        controls_frame.pack(fill="x", padx=10, pady=10)

        # Modern button styling
        button_font = ("Helvetica", 12, "bold")
        
        self.speak_button = ctk.CTkButton(
            controls_frame,
            text=f"{ICONS['PLAY']} Speak",
            command=self.on_speak,
            font=button_font,
            height=40
        )
        self.speak_button.pack(side="left", padx=5, expand=True, fill="x")

        self.stop_button = ctk.CTkButton(
            controls_frame,
            text=f"{ICONS['STOP']} Stop",
            command=self.on_stop,
            font=button_font,
            fg_color="#E74C3C",
            hover_color="#C0392B",
            height=40
        )

        self.save_button = ctk.CTkButton(
            controls_frame,
            text=f"{ICONS['SAVE']} Save Audio",
            command=self.on_save_as,
            font=button_font,
            height=40
        )
        self.save_button.pack(side="left", padx=5, expand=True, fill="x")

    def setup_status_section(self):
        """Setup status section with improved visualization"""
        status_frame = ctk.CTkFrame(self.main_frame)
        status_frame.pack(fill="x", padx=10, pady=5)

        # Status header with icon
        ctk.CTkLabel(
            status_frame,
            text=f"{ICONS['STATUS']} Status Information",
            font=("Helvetica", 14, "bold")
        ).pack(anchor="w", pady=5)

        # Status labels with improved styling and icons
        self.loading_status = ctk.CTkLabel(
            status_frame,
            text=f"{ICONS['SPEAKER']} Status: Initializing...",
            anchor="w",
            font=("Helvetica", 12)
        )
        self.loading_status.pack(fill="x", pady=2)

        self.cache_status = ctk.CTkLabel(
            status_frame,
            text=f"{ICONS['CACHE']} Cache: Checking...",
            anchor="w",
            font=("Helvetica", 12)
        )
        self.cache_status.pack(fill="x", pady=2)

        self.voice_count = ctk.CTkLabel(
            status_frame,
            text=f"{ICONS['COUNT']} Voices: -",
            anchor="w",
            font=("Helvetica", 12)
        )
        self.voice_count.pack(fill="x", pady=2)

        self.last_updated = ctk.CTkLabel(
            status_frame,
            text=f"{ICONS['CLOCK']} Last Updated: -",
            anchor="w",
            font=("Helvetica", 12)
        )
        self.last_updated.pack(fill="x", pady=2)

        # Modern progress bar
        self.progress_bar = ctk.CTkProgressBar(status_frame)
        self.progress_bar.pack(fill="x", pady=10)
        self.progress_bar.set(0)

    def toggle_theme(self):
        """Toggle between light and dark theme"""
        if ctk.get_appearance_mode() == "Dark":
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("Dark")

if __name__ == "__main__":
    app = EdgeTTSApp()
    app.mainloop()