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

# --- Global Variables ---
WINDOW_TITLE = "Edge TTS GUI"
WINDOW_SIZE = "700x650"
DEFAULT_APPEARANCE_MODE = "System"
DEFAULT_COLOR_THEME = "blue"
TEMP_AUDIO_FILENAME = "temp_audio_edge_tts1.mp3"  # Keep MP3 as default for temp files
DEFAULT_TEXT = "Hello, this is a test of Microsoft Edge Text-to-Speech with CustomTkinter."
DEFAULT_VOICE = "JennyNeural (en-US)"  # Default voice to select when loading voices

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

        # --- Text Input ---
        self.text_label = ctk.CTkLabel(self.main_frame, text="Enter Text:")
        self.text_label.pack(pady=(0, 5), anchor="w")

        self.text_input = ctk.CTkTextbox(self.main_frame, height=100, wrap="word") # Reduced height
        self.text_input.insert("1.0", DEFAULT_TEXT)
        self.text_input.pack(fill="x", expand=False) # Changed expand to False for y-axis

        # --- Voice Search and Selection ---
        self.voice_search_label = ctk.CTkLabel(self.main_frame, text="Search Voice:")
        self.voice_search_label.pack(pady=(10, 0), anchor="w")

        self.voice_search_entry = ctk.CTkEntry(self.main_frame, placeholder_text="Type to filter voices...")
        self.voice_search_entry.pack(fill="x", pady=(0,5))
        self.voice_search_entry.bind("<KeyRelease>", self.filter_voices)

        self.voice_selection_frame = ctk.CTkFrame(self.main_frame)
        self.voice_selection_frame.pack(pady=(0,10), fill="x")

        self.voice_label = ctk.CTkLabel(self.voice_selection_frame, text="Select Voice:")
        self.voice_label.pack(side="left", padx=(0, 10))

        self.voice_combobox = ctk.CTkComboBox(self.voice_selection_frame, values=["Loading voices..."], state="readonly",
                                              command=self.on_voice_selected_from_combobox)
        self.voice_combobox.pack(side="left", fill="x", expand=True)
        self.voice_combobox.set("Loading voices...")


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
        self.voice_search_entry.configure(state="disabled")
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
                display_name = f"{voice['FriendlyName']} ({voice['Locale']}) - {voice['Gender']}"
                self.voice_map[display_name] = voice['Name']
                self.display_voices_full.append(display_name)

            self.after(0, self.update_voice_combobox_post_load)

        except Exception as e:
            self.after(0, self.update_status, f"Error loading voices: {e}")
            self.after(0, self.voice_combobox.configure, {"values": ["Error loading voices"]})
            self.after(0, self.voice_combobox.set, "Error loading voices")
            self.after(0, self.voice_search_entry.configure, {"state": "normal"}) # Allow typing even if error


    def update_voice_combobox_post_load(self):
        if self.display_voices_full:
            self.voice_combobox.configure(values=self.display_voices_full)
            default_selection = next((v for v in self.display_voices_full if DEFAULT_VOICE in v), self.display_voices_full[0])
            self.voice_combobox.set(default_selection)
            self.update_status("Voices loaded. Ready.")
            self.speak_button.configure(state="normal")
            self.save_button.configure(state="normal")
        else:
            self.voice_combobox.configure(values=["No voices found"])
            self.voice_combobox.set("No voices found")
            self.update_status("No voices found.")
        self.voice_search_entry.configure(state="normal")

    def filter_voices(self, event=None):
        search_term = self.voice_search_entry.get().lower()
        current_selection = self.voice_combobox.get()

        if not search_term:
            filtered_display_voices = self.display_voices_full
        else:
            filtered_display_voices = [
                name for name in self.display_voices_full if search_term in name.lower()
            ]

        if not filtered_display_voices:
            self.voice_combobox.configure(values=["No match found"])
            self.voice_combobox.set("No match found")
        else:
            self.voice_combobox.configure(values=filtered_display_voices)
            if current_selection in filtered_display_voices:
                self.voice_combobox.set(current_selection)
            elif filtered_display_voices: # Select first if current is not in filtered
                self.voice_combobox.set(filtered_display_voices[0])

    def on_voice_selected_from_combobox(self, choice):
        # This callback is useful if you need to do something specific
        # when a voice is picked, other than just it being set.
        # For now, we don't need to do much here as the `get()` method
        # will retrieve the current selection.
        # print(f"Voice selected: {choice}")
        pass

    def update_status(self, message):
        self.status_label.configure(text=message)
        self.update_idletasks()

    def get_selected_voice_short_name(self):
        selected_display_name = self.voice_combobox.get()
        if selected_display_name in ["Loading voices...", "Error loading voices", "No voices found", "No match found"]:
            return None
        return self.voice_map.get(selected_display_name)

    def _synthesize_speech(self, text, voice_short_name, output_filepath, mime_type="audio/mp3"):
        try:
            if self.stop_requested.is_set():
                self.after(0, self.update_status, "Operation stopped before synthesis.")
                return False

            communicate = edge_tts.Communicate(text, voice_short_name)
            asyncio.run(communicate.save(output_filepath, mime_type=mime_type))

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
            self.voice_search_entry.configure(state="disabled")
            self.voice_combobox.configure(state="disabled")
        else:
            self.stop_button.pack_forget()
            self.speak_button.pack(side="left", padx=5, expand=True, fill="x")
            self.save_button.pack(side="left", padx=5, expand=True, fill="x")
            self.speak_button.configure(state="normal")
            self.save_button.configure(state="normal")
            self.voice_search_entry.configure(state="normal")
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

        # Get the file extension and corresponding MIME type
        file_ext = os.path.splitext(filepath)[1].lower()
        mime_type = FORMAT_MIME_TYPES.get(file_ext, "audio/mp3")  # Default to MP3 if unknown

        self._set_speaking_state(True) # Use speaking state to manage buttons
        self.update_status(f"Synthesizing and saving to {os.path.basename(filepath)}...")

        def synthesis_thread():
            try:
                success = self._synthesize_speech(text, selected_voice_short_name, filepath, mime_type)
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


if __name__ == "__main__":
    app = EdgeTTSApp()
    app.mainloop()