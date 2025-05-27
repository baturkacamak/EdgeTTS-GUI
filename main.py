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
from PIL import Image, ImageTk  # For icon support

# --- Global Variables ---
WINDOW_TITLE = "🎙️ Edge TTS Studio"  # More professional name
WINDOW_SIZE = "1000x700"  # Larger initial size
DEFAULT_APPEARANCE_MODE = "System"
DEFAULT_COLOR_THEME = "blue"
TEMP_AUDIO_FILENAME = "temp_audio_edge_tts1.mp3"  # Keep MP3 as default for temp files
DEFAULT_TEXT = "Hello, this is a test of Microsoft Edge Text-to-Speech with CustomTkinter."
DEFAULT_VOICE = "JennyNeural (en-US)"  # Default voice to select when loading voices
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".edge_tts_gui_config.json")  # Config file in user's home directory
ICON_FILES = ["icon.ico", "icon.png"]  # Application icon files

# Color scheme
COLORS = {
    "primary": "#0078D4",      # Microsoft blue
    "secondary": "#2B88D8",    # Lighter blue
    "accent": "#00B294",       # Teal
    "error": "#E74C3C",        # Red
    "error_dark": "#C0392B",   # Darker red for hover
    "success": "#2ECC71",      # Green
    "warning": "#D35400",      # Darker orange for pause button
    "warning_dark": "#A04000", # Even darker orange for hover
    "background_light": "#F5F5F5",
    "background_dark": "#1E1E1E",
    "text_light": "#2D2D2D",
    "text_dark": "#FFFFFF",
    "header_light": "#0078D4",  # Primary color for light mode
    "header_dark": "#4CC2FF"    # Lighter blue for dark mode
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

# Preview texts for different languages
PREVIEW_TEXTS = {
    # Germanic languages
    "en": {"rate": "This is {speed}× speed", "pitch": "Testing pitch adjustment"},  # English
    "de": {"rate": "Dies ist {speed}× Geschwindigkeit", "pitch": "Test der Tonhöhenanpassung"},  # German
    "nl": {"rate": "Dit is {speed}× snelheid", "pitch": "Toonhoogte aanpassing testen"},  # Dutch
    "af": {"rate": "Dit is {speed}× spoed", "pitch": "Toets toonhoogte aanpassing"},  # Afrikaans
    "is": {"rate": "Þetta er {speed}× hraði", "pitch": "Prófun tónhæðarbreytingar"},  # Icelandic
    
    # Romance languages
    "fr": {"rate": "Ceci est la vitesse {speed}×", "pitch": "Test d'ajustement de la hauteur"},  # French
    "es": {"rate": "Esta es velocidad {speed}×", "pitch": "Prueba de ajuste de tono"},  # Spanish
    "pt": {"rate": "Esta é a velocidade {speed}×", "pitch": "Teste de ajuste de tom"},  # Portuguese
    "it": {"rate": "Questa è velocità {speed}×", "pitch": "Test regolazione tono"},  # Italian
    "ro": {"rate": "Aceasta este viteza {speed}×", "pitch": "Test ajustare înălțime"},  # Romanian
    "ca": {"rate": "Aquesta és la velocitat {speed}×", "pitch": "Prova d'ajust de to"},  # Catalan
    "gl": {"rate": "Esta é a velocidade {speed}×", "pitch": "Proba de axuste de ton"},  # Galician
    
    # Slavic languages
    "ru": {"rate": "Это скорость {speed}×", "pitch": "Проверка настройки высоты"},  # Russian
    "pl": {"rate": "To jest prędkość {speed}×", "pitch": "Test regulacji wysokości"},  # Polish
    "cs": {"rate": "Toto je rychlost {speed}×", "pitch": "Test nastavení výšky"},  # Czech
    "uk": {"rate": "Це швидкість {speed}×", "pitch": "Тест налаштування висоти"},  # Ukrainian
    "bg": {"rate": "Това е скорост {speed}×", "pitch": "Тест за настройка на височината"},  # Bulgarian
    "hr": {"rate": "Ovo je brzina {speed}×", "pitch": "Test podešavanja visine tona"},  # Croatian
    "sr": {"rate": "Ово је брзина {speed}×", "pitch": "Тест подешавања висине тона"},  # Serbian
    "sk": {"rate": "Toto je rýchlosť {speed}×", "pitch": "Test nastavenia výšky"},  # Slovak
    "sl": {"rate": "To je hitrost {speed}×", "pitch": "Test nastavitve višine"},  # Slovenian
    "mk": {"rate": "Ова е брзина {speed}×", "pitch": "Тест за прилагодување на висината"},  # Macedonian
    
    # Baltic languages
    "lt": {"rate": "Tai yra {speed}× greitis", "pitch": "Aukščio reguliavimo testas"},  # Lithuanian
    "lv": {"rate": "Šis ir {speed}× ātrums", "pitch": "Augstuma regulēšanas tests"},  # Latvian
    "et": {"rate": "See on {speed}× kiirus", "pitch": "Helikõrguse reguleerimise test"},  # Estonian
    
    # East Asian languages
    "zh": {"rate": "这是{speed}倍速", "pitch": "音调调整测试"},  # Chinese
    "ja": {"rate": "これは{speed}倍速です", "pitch": "ピッチ調整テスト"},  # Japanese
    "ko": {"rate": "이것은 {speed}배속입니다", "pitch": "음높이 조정 테스트"},  # Korean
    
    # Southeast Asian languages
    "vi": {"rate": "Đây là tốc độ {speed}×", "pitch": "Kiểm tra điều chỉnh cao độ"},  # Vietnamese
    "th": {"rate": "นี่คือความเร็ว {speed}×", "pitch": "ทดสอบการปรับระดับเสียง"},  # Thai
    "id": {"rate": "Ini adalah kecepatan {speed}×", "pitch": "Tes penyesuaian nada"},  # Indonesian
    "ms": {"rate": "Ini adalah kelajuan {speed}×", "pitch": "Ujian pelarasan pic"},  # Malay
    "fil": {"rate": "Ito ay bilis na {speed}×", "pitch": "Pagsubok sa pagsasaayos ng pitch"},  # Filipino
    "km": {"rate": "នេះគឺជាល្បឿន {speed}×", "pitch": "ការសាកល្បងការកែសម្រួលកម្ពស់សំឡេង"},  # Khmer
    "my": {"rate": "ဤသည်မှာ {speed}× အမြန်နှုန်းဖြစ်သည်", "pitch": "အသံအမြင့်အနိမ့်ချိန်ညှိခြင်းစမ်းသပ်မှု"},  # Burmese
    "lo": {"rate": "ນີ້ແມ່ນຄວາມໄວ {speed}×", "pitch": "ທົດສອບການປັບລະດັບສຽງ"},  # Lao
    
    # South Asian languages
    "hi": {"rate": "यह {speed}× गति है", "pitch": "पिच समायोजन परीक्षण"},  # Hindi
    "ta": {"rate": "இது {speed}× வேகம்", "pitch": "பிட்ச் சரிசெய்தல் சோதனை"},  # Tamil
    "te": {"rate": "ఇది {speed}× వేగం", "pitch": "పిచ్ సర్దుబాటు పరీక్ష"},  # Telugu
    "bn": {"rate": "এটি {speed}× গতি", "pitch": "পিচ সমন্বয় পরীক্ষা"},  # Bengali
    "gu": {"rate": "આ {speed}× ગતિ છે", "pitch": "પિચ સમાયોજન પરીક્ષણ"},  # Gujarati
    "kn": {"rate": "ಇದು {speed}× ವೇಗ", "pitch": "ಪಿಚ್ ಹೊಂದಾಣಿಕೆ ಪರೀಕ್ಷೆ"},  # Kannada
    "ml": {"rate": "ഇത് {speed}× വേഗതയാണ്", "pitch": "പിച്ച് ക്രമീകരണ പരിശോധന"},  # Malayalam
    "mr": {"rate": "ही {speed}× गती आहे", "pitch": "पिच समायोजन चाचणी"},  # Marathi
    "pa": {"rate": "ਇਹ {speed}× ਗਤੀ ਹੈ", "pitch": "ਪਿੱਚ ਅਡਜਸਟਮੈਂਟ ਟੈਸਟ"},  # Punjabi
    "si": {"rate": "මෙය {speed}× වේගයයි", "pitch": "පිච් සීරුමාරු කිරීමේ පරීක්ෂණය"},  # Sinhala
    "ne": {"rate": "यो {speed}× गति हो", "pitch": "पिच समायोजन परीक्षण"},  # Nepali
    
    # Semitic languages
    "ar": {"rate": "هذه السرعة {speed}×", "pitch": "اختبار ضبط درجة الصوت"},  # Arabic
    "he": {"rate": "זוהי מהירות {speed}×", "pitch": "בדיקת כוונון גובה צליל"},  # Hebrew
    
    # Uralic languages
    "fi": {"rate": "Tämä on {speed}× nopeus", "pitch": "Sävelkorkeuden säätötesti"},  # Finnish
    "hu": {"rate": "Ez {speed}× sebesség", "pitch": "Hangmagasság beállítás teszt"},  # Hungarian
    
    # Turkic languages
    "tr": {"rate": "Bu {speed}× hız", "pitch": "Perde ayarı testi"},  # Turkish
    "az": {"rate": "Bu {speed}× sürətdir", "pitch": "Yüksəklik tənzimləmə testi"},  # Azerbaijani
    "kk": {"rate": "Бұл {speed}× жылдамдық", "pitch": "Биіктікті реттеу сынағы"},  # Kazakh
    "uz": {"rate": "Bu {speed}× tezlik", "pitch": "Balandlik sozlash testi"},  # Uzbek
    
    # Other European languages
    "sq": {"rate": "Kjo është shpejtësi {speed}×", "pitch": "Test i rregullimit të tonit"},  # Albanian
    "mt": {"rate": "Din hija veloċità {speed}×", "pitch": "Test tal-aġġustament tal-pitch"},  # Maltese
    
    # African languages
    "am": {"rate": "ይህ {speed}× ፍጥነት ነው", "pitch": "የድምጽ ከፍታ ማስተካከያ ሙከራ"},  # Amharic
    "sw": {"rate": "Hii ni kasi ya {speed}×", "pitch": "Jaribio la marekebisho ya sauti"},  # Swahili
    "zu": {"rate": "Lesi isijuqo {speed}×", "pitch": "Ukuhlola ukuhlelwa kwephimbo"},  # Zulu
    
    # Other languages
    "ps": {"rate": "دا {speed}× سرعت دی", "pitch": "د غږ لوړوالي ازموینه"},  # Pashto
    "fa": {"rate": "این سرعت {speed}× است", "pitch": "آزمایش تنظیم زیر و بمی"},  # Persian
    "ur": {"rate": "یہ {speed}× رفتار ہے", "pitch": "پچ ایڈجسٹمنٹ ٹیسٹ"},  # Urdu
    "so": {"rate": "Kani waa xawaaraha {speed}×", "pitch": "Tijaabada hagaajinta codka"},  # Somali
    "su": {"rate": "Ieu kacepetan {speed}×", "pitch": "Tés nyetel pitch"},  # Sundanese
    "jv": {"rate": "Iki kacepetan {speed}×", "pitch": "Tés nyetel pitch"},  # Javanese
    "iu": {"rate": "ᑖᓐᓇ {speed}× ᓱᑲᐃᑎᒋᔪᖅ", "pitch": "ᓂᐱᐅᑉ ᖁᑦᑎᓂᖓᑕ ᐋᖅᑭᒋᐊᕈᑎᖓᑕ ᖃᐅᔨᓴᕈᑎᖓ"},  # Inuktitut
    
    # Default fallback
    "default": {"rate": "Speed test {speed}×", "pitch": "Pitch test"}
}

class ToolTip:
    """Create a tooltip for a given widget with modern styling."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.scheduled_id = None

    def enter(self, event=None):
        # Schedule tooltip to appear after a short delay
        self.scheduled_id = self.widget.after(500, self.show_tooltip)

    def leave(self, event=None):
        # Cancel scheduled tooltip
        if self.scheduled_id:
            self.widget.after_cancel(self.scheduled_id)
            self.scheduled_id = None
        # Destroy tooltip if it exists
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

    def show_tooltip(self):
        # Get screen coordinates of the widget
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        # Create tooltip window
        self.tooltip = tk = tkinter.Toplevel(self.widget)
        tk.wm_overrideredirect(True)
        tk.wm_geometry(f"+{x}+{y}")

        # Create tooltip content with modern styling
        label = tkinter.Label(
            tk,
            text=self.text,
            justify=tkinter.LEFT,
            background="#2B2B2B",
            foreground="white",
            relief=tkinter.SOLID,
            borderwidth=1,
            font=("Segoe UI", "9", "normal"),
            padx=5,
            pady=3
        )
        label.pack()

class EdgeTTSApp(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Try to set application icon
        self._set_app_icon()

        # Initialize pygame mixer
        pygame.mixer.init()
        pygame.mixer.quit()  # Quit immediately to free resources
        
        self.title(WINDOW_TITLE)
        self.geometry(WINDOW_SIZE)
        self.minsize(800, 600)  # Set minimum window size
        
        ctk.set_appearance_mode(DEFAULT_APPEARANCE_MODE)
        ctk.set_default_color_theme(DEFAULT_COLOR_THEME)

        # Setup keyboard shortcuts
        self.bind("<Control-o>", lambda e: self.on_load_file())  # Ctrl+O to open file
        self.bind("<Control-s>", lambda e: self.on_save_as())    # Ctrl+S to save audio
        self.bind("<space>", self.handle_space_key)              # Space to play/pause
        self.bind("<Escape>", lambda e: self.on_stop())          # Esc to stop
        self.bind("<Control-q>", self.on_closing)                # Ctrl+Q to quit
        
        # Function to get header color based on theme
        self.get_header_color = lambda: COLORS["header_dark"] if ctk.get_appearance_mode() == "Dark" else COLORS["header_light"]

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
            text_color=self.get_header_color()
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
            text_color=self.get_header_color()
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
                        
                        # Wait for playback to finish
                        while pygame.mixer.get_init() and (pygame.mixer.music.get_busy() or self.is_paused):
                            if self.stop_requested.is_set():
                                break
                            time.sleep(0.1)
                            
                        if not self.stop_requested.is_set():
                            self.after(0, self.update_detailed_status, "Playback finished. Ready.")
                        else:
                            self.after(0, self.update_detailed_status, "Playback stopped/skipped. Ready.")
                    except Exception as e:
                        if not self.stop_requested.is_set():
                            self.after(0, self.update_detailed_status, f"Error playing audio: {e}")
                    finally:
                        # Only try to remove the file if mixer is not initialized or not busy
                        try:
                            if not pygame.mixer.get_init() or not pygame.mixer.music.get_busy():
                                if os.path.exists(temp_audio_path):
                                    os.remove(temp_audio_path)
                        except Exception as e_del:
                            print(f"Error deleting temp file: {e_del}")
            finally:
                if not self.stop_requested.is_set():
                    self.after(0, lambda: self._set_speaking_state(False))

        threading.Thread(target=synthesis_and_playback_thread, daemon=True).start()

    def play_audio(self, audio_path):
        """Play audio using pygame mixer"""
        try:
            # Make sure mixer is not initialized before initializing
            if pygame.mixer.get_init():
                pygame.mixer.quit()
            
            # Initialize mixer with good settings for speech
            pygame.mixer.init(frequency=44100, size=-16, channels=2, buffer=2048)
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
            if pygame.mixer.get_init():
                pygame.mixer.quit()
            self._set_speaking_state(False)  # Reset state on error

    def update_progress(self):
        """Update progress bar and time display"""
        try:
            if pygame.mixer.get_init() and (pygame.mixer.music.get_busy() or self.is_paused) and not self.stop_requested.is_set():
                current_pos = pygame.mixer.music.get_pos() / 1000  # Convert to seconds
                if current_pos < 0:  # When paused, get_pos returns -1
                    current_pos = float(self.progress_bar.get()) * self.audio_length
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
                # Only reset speaking state if we're not paused and mixer is initialized
                try:
                    if not self.is_paused and pygame.mixer.get_init() and not pygame.mixer.music.get_busy():
                        pygame.mixer.quit()
                        self.after(0, lambda: self._set_speaking_state(False))
                except pygame.error:
                    # If there's an error with the mixer, just reset the speaking state
                    self.after(0, lambda: self._set_speaking_state(False))
        except Exception as e:
            print(f"Error updating progress: {e}")
            self.after(0, lambda: self._set_speaking_state(False))

    def format_time(self, seconds):
        """Format seconds to MM:SS"""
        minutes = int(seconds // 60)
        seconds = int(seconds % 60)
        return f"{minutes}:{seconds:02d}"

    def on_pause_resume(self):
        """Handle pause/resume button click"""
        if not pygame.mixer.get_init():
            return
            
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
        if pygame.mixer.get_init():
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
        if pygame.mixer.get_init():
            pygame.mixer.music.stop()
            pygame.mixer.quit()  # Quit mixer after stopping
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

            # Get rate and pitch values
            rate = self.rate_slider.get()
            pitch = self.pitch_slider.get()

            # Create Communicate instance with rate and pitch adjustments
            # Rate needs to be a percentage string (e.g., "+0%", "+50%", "-50%")
            rate_percent = int((rate - 1.0) * 100)  # Convert multiplier to percentage difference
            communicate = edge_tts.Communicate(
                text,
                voice_short_name,
                rate=f"{rate_percent:+d}%",  # Format: +0%, +50%, -50%
                pitch=f"{int(pitch):+d}Hz"  # Format: +0Hz, +10Hz, etc.
            )
            
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
        """Set the UI state for speaking/not speaking"""
        self.is_speaking = speaking
        self.stop_requested.clear()
        self.is_paused = False

        if speaking:
            # Hide speak and save buttons
            self.speak_button.grid_remove()
            self.save_button.grid_remove()
            
            # Show stop and pause buttons
            self.stop_button.grid(row=0, column=0, sticky="ew", padx=5)
            self.pause_button.grid(row=0, column=1, sticky="ew", padx=5)
            
            # Disable voice selection
            self.voice_combobox.configure(state="disabled")
        else:
            # Hide stop and pause buttons
            self.stop_button.grid_remove()
            self.pause_button.grid_remove()
            
            # Show speak and save buttons
            self.speak_button.grid(row=0, column=0, sticky="ew", padx=5)
            self.save_button.grid(row=0, column=1, sticky="ew", padx=5)
            
            # Enable voice selection
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
        self.voice_selection_header = ctk.CTkLabel(
            voice_frame,
            text=f"{ICONS['VOICE']} Voice Selection",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.get_header_color()
        )
        self.voice_selection_header.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))

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
        self.controls_header = ctk.CTkLabel(
            controls_frame,
            text=f"{ICONS['MICROPHONE']} Controls",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.get_header_color()
        )
        self.controls_header.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 5))

        # Button container
        self.button_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        self.button_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 10))
        self.button_frame.grid_columnconfigure((0, 1), weight=1)

        # Modern button styling
        button_font = ctk.CTkFont(size=14, weight="bold")
        
        # Speak button with hover animation
        self.speak_button = ctk.CTkButton(
            self.button_frame,
            text=f"{ICONS['PLAY']} Speak",
            command=self.on_speak,
            font=button_font,
            height=45,
            fg_color=COLORS["primary"],
            hover_color=COLORS["secondary"],
            corner_radius=8
        )
        self.speak_button.grid(row=0, column=0, sticky="ew", padx=5)
        ToolTip(self.speak_button, "Start speaking (Space)")

        # Stop button
        self.stop_button = ctk.CTkButton(
            self.button_frame,
            text=f"{ICONS['STOP']} Stop",
            command=self.on_stop,
            font=button_font,
            height=45,
            fg_color=COLORS["error"],
            hover_color=COLORS["error_dark"],
            corner_radius=8
        )
        ToolTip(self.stop_button, "Stop speaking (Esc)")

        # Pause/Resume button
        self.pause_button = ctk.CTkButton(
            self.button_frame,
            text=f"{ICONS['PAUSE']} Pause",
            command=self.on_pause_resume,
            font=button_font,
            height=45,
            fg_color=COLORS["warning"],
            hover_color=COLORS["warning_dark"],
            corner_radius=8
        )
        ToolTip(self.pause_button, "Pause/Resume speaking (Space)")

        # Save button with modern styling
        self.save_button = ctk.CTkButton(
            self.button_frame,
            text=f"{ICONS['SAVE']} Save Audio",
            command=self.on_save_as,
            font=button_font,
            height=45,
            fg_color=COLORS["accent"],
            hover_color="#00A080",
            corner_radius=8
        )
        self.save_button.grid(row=0, column=1, sticky="ew", padx=5)
        ToolTip(self.save_button, "Save audio file (Ctrl+S)")

        # Initially hide stop and pause buttons
        self.stop_button.grid_remove()
        self.pause_button.grid_remove()

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

        # Rate control
        rate_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        rate_frame.grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 10))
        rate_frame.grid_columnconfigure(1, weight=1)

        rate_label = ctk.CTkLabel(
            rate_frame,
            text="⚡ Rate",
            font=("Helvetica", 12)
        )
        rate_label.grid(row=0, column=0, padx=(0, 10))

        self.rate_value_label = ctk.CTkLabel(
            rate_frame,
            text="1.0×",
            font=("Helvetica", 10)
        )
        self.rate_value_label.grid(row=0, column=2, padx=(10, 0))

        self.rate_slider = ctk.CTkSlider(
            rate_frame,
            from_=0.5,
            to=2.0,
            number_of_steps=30,
            command=self.on_rate_change
        )
        self.rate_slider.grid(row=0, column=1, sticky="ew")
        self.rate_slider.set(1.0)  # Default rate

        # Preview button for rate
        self.rate_preview_btn = ctk.CTkButton(
            rate_frame,
            text="👂",
            width=30,
            height=24,
            command=lambda: self.preview_settings("rate"),
            font=ctk.CTkFont(size=12)
        )
        self.rate_preview_btn.grid(row=0, column=3, padx=(10, 0))

        # Pitch control
        pitch_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        pitch_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 10))
        pitch_frame.grid_columnconfigure(1, weight=1)

        pitch_label = ctk.CTkLabel(
            pitch_frame,
            text="🎵 Pitch",
            font=("Helvetica", 12)
        )
        pitch_label.grid(row=0, column=0, padx=(0, 10))

        self.pitch_value_label = ctk.CTkLabel(
            pitch_frame,
            text="±0 Hz",
            font=("Helvetica", 10)
        )
        self.pitch_value_label.grid(row=0, column=2, padx=(10, 0))

        self.pitch_slider = ctk.CTkSlider(
            pitch_frame,
            from_=-50,
            to=50,
            number_of_steps=20,
            command=self.on_pitch_change
        )
        self.pitch_slider.grid(row=0, column=1, sticky="ew")
        self.pitch_slider.set(0)  # Default pitch

        # Preview button for pitch
        self.pitch_preview_btn = ctk.CTkButton(
            pitch_frame,
            text="👂",
            width=30,
            height=24,
            command=lambda: self.preview_settings("pitch"),
            font=ctk.CTkFont(size=12)
        )
        self.pitch_preview_btn.grid(row=0, column=3, padx=(10, 0))

        # Bind progress bar click for seeking
        self.progress_bar.bind("<Button-1>", self.on_progress_click)

    def setup_status_section(self):
        """Setup status section with modern visualization"""
        status_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        status_frame.grid(row=3, column=0, sticky="new", padx=10, pady=5)
        status_frame.grid_columnconfigure(0, weight=1)
        
        # Set a maximum width for the status frame
        status_frame.grid_propagate(False)  # Prevent the frame from expanding
        status_frame.configure(width=400, height=200)  # Increased height to fit content better

        # Status header
        self.status_header = ctk.CTkLabel(
            status_frame,
            text=f"{ICONS['STATUS']} Status Information",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.get_header_color()
        )
        self.status_header.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))

        # Status container
        info_frame = ctk.CTkFrame(status_frame, fg_color="transparent")
        info_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        info_frame.grid_columnconfigure(0, weight=1)

        # Status labels with modern styling and text wrapping
        label_font = ctk.CTkFont(size=13)
        
        # Calculate wraplength based on frame width minus padding
        wrap_length = 380  # status frame width (400) minus padding (20)
        
        self.loading_status = ctk.CTkLabel(
            info_frame,
            text=f"{ICONS['SPEAKER']} Status: Initializing...",
            anchor="w",
            font=label_font,
            text_color=("gray20", "gray80"),
            wraplength=wrap_length,
            justify="left"
        )
        self.loading_status.grid(row=0, column=0, sticky="w", pady=2)

        self.cache_status = ctk.CTkLabel(
            info_frame,
            text=f"{ICONS['CACHE']} Cache: Checking...",
            anchor="w",
            font=label_font,
            text_color=("gray20", "gray80"),
            wraplength=wrap_length,
            justify="left"
        )
        self.cache_status.grid(row=1, column=0, sticky="w", pady=2)

        self.voice_count = ctk.CTkLabel(
            info_frame,
            text=f"{ICONS['COUNT']} Voices: -",
            anchor="w",
            font=label_font,
            text_color=("gray20", "gray80"),
            wraplength=wrap_length,
            justify="left"
        )
        self.voice_count.grid(row=2, column=0, sticky="w", pady=2)

        self.last_updated = ctk.CTkLabel(
            info_frame,
            text=f"{ICONS['CLOCK']} Last Updated: -",
            anchor="w",
            font=label_font,
            text_color=("gray20", "gray80"),
            wraplength=wrap_length,
            justify="left"
        )
        self.last_updated.grid(row=3, column=0, sticky="w", pady=2)

    def toggle_theme(self):
        """Toggle between light and dark theme"""
        if ctk.get_appearance_mode() == "Dark":
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("Dark")
        
        # Update header colors
        header_color = self.get_header_color()
        self.logo_label.configure(text_color=header_color)
        
        # Update all section headers
        for widget in [
            self.voice_selection_header,
            self.controls_header,
            self.status_header
        ]:
            if widget:
                widget.configure(text_color=header_color)

    def update_text_stats(self, event):
        """Update both word and character count labels"""
        text = self.text_input.get("1.0", "end-1c")
        char_count = len(text)
        # Split by whitespace and filter out empty strings
        word_count = len([word for word in text.split() if word.strip()])
        
        self.char_count_label.configure(text=f"Characters: {char_count}")
        self.word_count_label.configure(text=f"Words: {word_count}")

    def on_closing(self, event=0):
        """Handle application closing"""
        if pygame.mixer.get_init():
            pygame.mixer.quit()
        self.quit()

    def _set_app_icon(self):
        """Try to set the application icon using available icon files"""
        for icon_file in ICON_FILES:
            if os.path.exists(icon_file):
                try:
                    icon = Image.open(icon_file)
                    photo = ImageTk.PhotoImage(icon)
                    self.wm_iconphoto(True, photo)
                    return  # Successfully set the icon
                except Exception as e:
                    print(f"Could not set icon using {icon_file}: {e}")
        
        print("Could not set any application icon")

    def on_rate_change(self, value):
        """Handle rate slider change"""
        self.rate_value_label.configure(text=f"{value:.1f}×")

    def on_pitch_change(self, value):
        """Handle pitch slider change"""
        self.pitch_value_label.configure(text=f"±{int(value)} Hz")

    def preview_settings(self, setting_type):
        """Preview rate or pitch settings with a short sample."""
        if self.is_speaking and not self.is_paused:
            return  # Don't preview while speaking

        # Store current main playback state
        was_playing = False
        current_pos = 0
        if pygame.mixer.get_init() and pygame.mixer.music.get_busy():
            was_playing = True
            current_pos = pygame.mixer.music.get_pos() / 1000.0
            pygame.mixer.music.pause()

        # Get the selected voice and determine language
        selected_voice = self.get_selected_voice_short_name()
        if not selected_voice:
            return

        # Find the voice in voices list to get its locale
        voice_locale = "default"
        for voice in self.voices_list_full:
            if voice["Name"] == selected_voice:
                # Extract main language code (e.g., "en-US" -> "en")
                voice_locale = voice["Locale"].split("-")[0].lower()
                break

        # Get preview text in the appropriate language
        preview_texts = PREVIEW_TEXTS.get(voice_locale, PREVIEW_TEXTS["default"])
        
        # Sample text based on what's being adjusted
        if setting_type == "rate":
            preview_text = preview_texts["rate"].format(speed=self.rate_slider.get())
        else:  # pitch
            preview_text = preview_texts["pitch"]

        # Create temporary file for preview
        preview_file = os.path.join(tempfile.gettempdir(), "preview_audio_edge_tts.mp3")

        # Synthesize and play preview
        if selected_voice:
            try:
                # Get current settings
                rate = self.rate_slider.get()
                pitch = self.pitch_slider.get()

                # Convert rate to percentage
                rate_percent = int((rate - 1.0) * 100)

                # Create communicate instance for preview
                communicate = edge_tts.Communicate(
                    preview_text,
                    selected_voice,
                    rate=f"{rate_percent:+d}%",
                    pitch=f"{int(pitch):+d}Hz"
                )

                # Save and play preview
                asyncio.run(communicate.save(preview_file))
                
                # Initialize mixer for preview if needed
                if pygame.mixer.get_init():
                    pygame.mixer.quit()
                pygame.mixer.init(frequency=44100, size=-16, channels=2, buffer=2048)
                pygame.mixer.music.load(preview_file)
                pygame.mixer.music.set_volume(self.volume_slider.get() / 100)
                pygame.mixer.music.play()

                # Wait for preview to finish
                while pygame.mixer.get_init() and pygame.mixer.music.get_busy():
                    time.sleep(0.1)

                # Clean up preview file
                pygame.mixer.quit()
                if os.path.exists(preview_file):
                    os.remove(preview_file)

                # Resume main playback if it was playing
                if was_playing and self.current_audio_file:
                    pygame.mixer.init(frequency=44100, size=-16, channels=2, buffer=2048)
                    pygame.mixer.music.load(self.current_audio_file)
                    pygame.mixer.music.set_volume(self.volume_slider.get() / 100)
                    pygame.mixer.music.play()
                    # Try to restore position
                    try:
                        pygame.mixer.music.set_pos(current_pos)
                    except:
                        pass  # Position restoration failed, continue from start

            except Exception as e:
                self.update_detailed_status(f"Preview error: {e}")
                if was_playing:
                    # Attempt to resume main playback on error
                    try:
                        pygame.mixer.init(frequency=44100, size=-16, channels=2, buffer=2048)
                        pygame.mixer.music.load(self.current_audio_file)
                        pygame.mixer.music.set_volume(self.volume_slider.get() / 100)
                        pygame.mixer.music.play()
                        pygame.mixer.music.set_pos(current_pos)
                    except:
                        pass

    def handle_space_key(self, event=None):
        """Handle space key press for play/pause functionality"""
        # Don't trigger if focus is in text input
        if str(self.focus_get()).endswith("ctktextbox"):
            return
            
        if self.is_speaking:
            if pygame.mixer.get_init() and pygame.mixer.music.get_busy():
                self.on_pause_resume()
        else:
            self.on_speak()

if __name__ == "__main__":
    app = EdgeTTSApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)  # Handle window closing
    app.mainloop()