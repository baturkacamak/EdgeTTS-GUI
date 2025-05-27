# EdgeTTS-CustomTkinter-GUI

A user-friendly, cross-platform graphical interface for Microsoft's Edge Text-to-Speech service. Built with Python and CustomTkinter, allowing easy text-to-speech conversion with voice selection, playback, and audio file saving.

## Features (Initial Version)

*   Input text area.
*   Dynamic voice selection (language and voice name).
*   Real-time playback of synthesized speech.
*   Save synthesized speech to an MP3 file.
*   Basic status messages.

## Prerequisites

*   Python 3.8+

## Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/your-username/EdgeTTS-CustomTkinter-GUI.git
    cd EdgeTTS-CustomTkinter-GUI
    ```

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    # On Windows
    venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
    *   **Note for `playsound`:**
        *   On Linux, you might need GStreamer: `sudo apt-get install python3-gst-1.0 gir1.2-gstreamer-1.0 gstreamer1.0-plugins-good`
        *   On macOS, you might need PyObjC: `pip install PyObjC`
        *   Windows usually works out of the box.

## Usage

Run the main application script:

```bash
python main.py