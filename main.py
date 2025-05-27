import tkinter
import tkinter.filedialog
import customtkinter as ctk
import edge_tts
import asyncio
import threading
import os
import tempfile
from playsound import playsound # Using playsound 1.2.2
import time # For small delay in search
import docx  # For DOCX files
import chardet  # For detecting text file encodings
import tkinter.ttk as ttk
from CTkScrollableDropdown.ctk_scrollable_dropdown_keyboard import CTkScrollableDropdownKeyboard

# --- Global Variables ---
WINDOW_TITLE = "Edge TTS GUI"
WINDOW_SIZE = "700x650"
DEFAULT_APPEARANCE_MODE = "System"
DEFAULT_COLOR_THEME = "blue"
TEMP_AUDIO_FILENAME = "temp_audio_edge_tts1.mp3"  # Keep MP3 as default for temp files
DEFAULT_TEXT = "Hello, this is a test of Microsoft Edge Text-to-Speech with CustomTkinter."
DEFAULT_VOICE = "JennyNeural (en-US)"  # Default voice to select when loading voices

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

        ctk.set_appearance_mode(DEFAULT_APPEARANCE_MODE)
        ctk.set_default_color_theme(DEFAULT_COLOR_THEME)

        self.voices_list_full = [] # Full list of voice dicts
        self.voice_map = {}        # Maps display name to short name
        self.display_voices_full = [] # Full list of display names for combobox

        self.is_speaking = False
        self.stop_requested = threading.Event() # For stopping speak/save operations

        # --- Main Frame ---
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # --- Text Input Frame ---
        self.text_input_frame = ctk.CTkFrame(self.main_frame)
        self.text_input_frame.pack(fill="x", expand=False)

        # --- Text Input Header ---
        self.text_header_frame = ctk.CTkFrame(self.text_input_frame)
        self.text_header_frame.pack(fill="x", pady=(0, 5))

        self.text_label = ctk.CTkLabel(self.text_header_frame, text="Enter Text:")
        self.text_label.pack(side="left", padx=(0, 10))

        self.load_file_button = ctk.CTkButton(self.text_header_frame, text="Load from File", 
                                            command=self.on_load_file, width=100)
        self.load_file_button.pack(side="right")

        self.text_input = ctk.CTkTextbox(self.text_input_frame, height=100, wrap="word")
        self.text_input.insert("1.0", DEFAULT_TEXT)
        self.text_input.pack(fill="x", expand=False)

        # --- Voice Search and Selection ---
        self.voice_selection_frame = ctk.CTkFrame(self.main_frame)
        self.voice_selection_frame.pack(pady=(0,10), fill="x")

        self.voice_label = ctk.CTkLabel(self.voice_selection_frame, text="Select Voice:")
        self.voice_label.pack(side="left", padx=(0, 10))

        self.voice_combobox = ctk.CTkComboBox(
            self.voice_selection_frame,
            values=["Loading voices..."],  # Start with loading message
            state="disabled",  # Start disabled until voices are loaded
            width=400,
            command=self.on_voice_selected_from_combobox
        )
        self.voice_combobox.pack(fill="x", expand=True)
        self.voice_combobox.set("Loading voices...")
        
        # Initialize the scrollable dropdown with autocomplete
        self.voice_dropdown = CTkScrollableDropdownKeyboard(
            self.voice_combobox,
            values=["Loading voices..."],  # Start with loading message
            command=self.on_voice_selected_from_dropdown,
            autocomplete=True,
            justify="left",
            button_height=30,
            height=200
        )

        # --- Controls ---
        self.controls_frame = ctk.CTkFrame(self.main_frame)
        self.controls_frame.pack(pady=10, fill="x")

        self.speak_button = ctk.CTkButton(self.controls_frame, text="Speak", command=self.on_speak)
        self.speak_button.pack(side="left", padx=5, expand=True, fill="x")

        self.stop_button = ctk.CTkButton(self.controls_frame, text="Stop", command=self.on_stop, fg_color="red", hover_color="darkred")
        # Stop button will be packed/unpacked dynamically

        self.save_button = ctk.CTkButton(self.controls_frame, text="Save Audio", command=self.on_save_as)
        self.save_button.pack(side="left", padx=5, expand=True, fill="x")

        # --- Status Bar ---
        self.status_label = ctk.CTkLabel(self.main_frame, text="Ready", anchor="w")
        self.status_label.pack(pady=(10, 0), fill="x")

        self.load_initial_voices()

    def load_initial_voices(self):
        self.update_status("Loading voices...")
        self.speak_button.configure(state="disabled")
        self.save_button.configure(state="disabled")
        threading.Thread(target=self.load_voices_threaded, daemon=True).start()

    def load_voices_threaded(self):
        try:
            async def get_voices_async():
                return await edge_tts.VoicesManager.create()
            voices_manager = asyncio.run(get_voices_async())
            self.voices_list_full = voices_manager.voices

            # Sort voices: Primarily by Locale (e.g., 'en-US'), then by FriendlyName
            self.voices_list_full.sort(key=lambda v: (v['Locale'], v['FriendlyName']))

            self.voice_map.clear()
            self.display_voices_full.clear()

            for voice in self.voices_list_full:
                # Use LOCALE_NAME_MAP for full language/region name, fallback to Locale code
                locale_name = LOCALE_NAME_MAP.get(voice['Locale'], voice['Locale'])
                # Extract first name after 'Microsoft' if present, otherwise use first word
                friendly_name = voice['FriendlyName']
                if friendly_name.startswith("Microsoft "):
                    friendly_name = friendly_name[len("Microsoft "):]
                first_name = friendly_name.split()[0]
                display_name = f"{locale_name} - {voice['Gender']} - {first_name}"
                self.voice_map[display_name] = voice['Name']
                self.display_voices_full.append(display_name)

            self.after(0, self.update_voice_combobox_post_load)

        except Exception as e:
            self.after(0, self.update_status, f"Error loading voices: {e}")
            self.after(0, self.voice_combobox.configure, {"values": ["Error loading voices"]})
            self.after(0, self.voice_combobox.set, "Error loading voices")
            self.after(0, self.voice_search_entry.configure, {"state": "normal"}) # Allow typing even if error


    def update_voice_combobox_post_load(self):
        """Update the combobox with loaded voices"""
        if self.display_voices_full:
            # Update the combobox values
            self.voice_combobox.configure(values=self.display_voices_full)
            
            # Update the dropdown values
            self.voice_dropdown.configure(values=self.display_voices_full)
            
            # Set default voice if available, otherwise use first voice
            default_voice = next((v for v in self.display_voices_full if DEFAULT_VOICE in v), self.display_voices_full[0])
            self.voice_combobox.set(default_voice)
            
            # Enable the combobox and buttons
            self.voice_combobox.configure(state="normal")
            self.speak_button.configure(state="normal")
            self.save_button.configure(state="normal")
            self.update_status("Voices loaded. Ready.")
        else:
            # No voices found - set appropriate messages
            no_voices_msg = "No voices found"
            self.voice_combobox.configure(values=[no_voices_msg])
            self.voice_dropdown.configure(values=[no_voices_msg])
            self.voice_combobox.set(no_voices_msg)
            self.voice_combobox.configure(state="disabled")
            self.speak_button.configure(state="disabled")
            self.save_button.configure(state="disabled")
            self.update_status("No voices found.")

    def on_voice_selected_from_combobox(self, choice):
        # This callback is useful if you need to do something specific
        # when a voice is picked, other than just it being set.
        # For now, we don't need to do much here as the `get()` method
        # will retrieve the current selection.
        # print(f"Voice selected: {choice}")
        pass

    def on_voice_selected_from_dropdown(self, choice):
        """Callback for when a voice is selected from the dropdown"""
        self.voice_combobox.set(choice)  # Update the combobox text
        # The regular combobox callback will handle the rest
        self.on_voice_selected_from_combobox(choice)

    def update_status(self, message):
        self.status_label.configure(text=message)
        self.update_idletasks()

    def get_selected_voice_short_name(self):
        selected_display_name = self.voice_combobox.get()
        if selected_display_name in ["Loading voices...", "Error loading voices", "No voices found", "No match found"]:
            return None
        return self.voice_map.get(selected_display_name)

    def _synthesize_speech(self, text, voice_short_name, output_filepath):
        try:
            if self.stop_requested.is_set():
                self.after(0, self.update_status, "Operation stopped before synthesis.")
                return False

            communicate = edge_tts.Communicate(text, voice_short_name)
            asyncio.run(communicate.save(output_filepath))

            if self.stop_requested.is_set(): # Check again after potentially long synthesis
                self.after(0, self.update_status, "Operation stopped after synthesis, before playback/save completion.")
                if os.path.exists(output_filepath) and output_filepath.endswith(TEMP_AUDIO_FILENAME):
                    try: os.remove(output_filepath) # Clean up temp file if stop requested
                    except Exception: pass
                return False
            return True
        except Exception as e:
            self.after(0, self.update_status, f"Synthesis error: {e}")
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

    def on_speak(self):
        if self.is_speaking: return

        text = self.text_input.get("1.0", "end-1c").strip()
        if not text:
            self.update_status("Error: Text input is empty.")
            return

        selected_voice_short_name = self.get_selected_voice_short_name()
        if not selected_voice_short_name:
            self.update_status("Error: No valid voice selected.")
            return

        self._set_speaking_state(True)
        self.update_status(f"Synthesizing with {selected_voice_short_name}...")

        temp_dir = tempfile.gettempdir()
        temp_audio_path = os.path.join(temp_dir, TEMP_AUDIO_FILENAME)

        def synthesis_and_playback_thread():
            try:
                success = self._synthesize_speech(text, selected_voice_short_name, temp_audio_path)

                if self.stop_requested.is_set():
                    self.after(0, self.update_status, "Speak operation stopped.")
                    if os.path.exists(temp_audio_path): os.remove(temp_audio_path)
                    return

                if success:
                    self.after(0, self.update_status, "Playing audio...")
                    if self.stop_requested.is_set(): # Check one more time before playsound
                        self.after(0, self.update_status, "Speak operation stopped before playback.")
                        if os.path.exists(temp_audio_path): os.remove(temp_audio_path)
                        return

                    try:
                        playsound(temp_audio_path) # This blocks this thread
                        if not self.stop_requested.is_set(): # Only update if not stopped
                             self.after(0, self.update_status, "Playback finished. Ready.")
                        else:
                             self.after(0, self.update_status, "Playback stopped/skipped. Ready.")
                    except Exception as e:
                        if not self.stop_requested.is_set():
                            self.after(0, self.update_status, f"Error playing audio: {e}")
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
            self.update_status("Error: Text input is empty.")
            return

        selected_voice_short_name = self.get_selected_voice_short_name()
        if not selected_voice_short_name:
            self.update_status("Error: No valid voice selected.")
            return

        filepath = tkinter.filedialog.asksaveasfilename(
            defaultextension=".mp3",
            filetypes=SUPPORTED_FORMATS,
            title="Save Speech As Audio"
        )
        if not filepath:
            self.update_status("Save cancelled. Ready.")
            return

        self._set_speaking_state(True) # Use speaking state to manage buttons
        self.update_status(f"Synthesizing and saving to {os.path.basename(filepath)}...")

        def synthesis_thread():
            try:
                success = self._synthesize_speech(text, selected_voice_short_name, filepath)
                if self.stop_requested.is_set():
                    self.after(0, self.update_status, "Save operation stopped.")
                    return

                if success:
                    self.after(0, self.update_status, f"Audio saved to {os.path.basename(filepath)}. Ready.")
            finally:
                if not self.stop_requested.is_set():
                    self.after(0, lambda: self._set_speaking_state(False))

        threading.Thread(target=synthesis_thread, daemon=True).start()

    def on_stop(self):
        self.update_status("Stop request received...")
        self.stop_requested.set()
        self._set_speaking_state(False) # Immediately reset UI buttons
        # Note: This won't instantly kill the playsound or edge-tts network call if already deep into it,
        # but it will prevent new actions and update UI state.
        # The running threads will check the stop_requested event at their earliest convenience.
        self.update_status("Operation stopped. Ready.")

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
                self.update_status(f"Loaded text from {os.path.basename(filepath)}")
            else:
                self.update_status("Error: Could not read text from file.")
        except Exception as e:
            self.update_status(f"Error loading file: {str(e)}")

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

if __name__ == "__main__":
    app = EdgeTTSApp()
    app.mainloop()