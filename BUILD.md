# Build Instructions

This document explains how to build EdgeTTS-GUI packages for different platforms.

## Prerequisites

- Python 3.8 or higher
- Git
- Platform-specific tools:
  - Linux: `dpkg-dev`, `rpm`
  - Windows: Inno Setup
  - macOS: Xcode Command Line Tools

## Common Setup

1. Clone the repository:
```bash
git clone https://github.com/baturkacamak/EdgeTTS-GUI.git
cd EdgeTTS-GUI
```

2. Create and activate a virtual environment:
```bash
# Linux/macOS
python -m venv venv
source venv/bin/activate

# Windows
python -m venv venv
venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
pip install pyinstaller
```

## Building on Linux

1. Install required tools:
```bash
sudo apt-get update
sudo apt-get install dpkg-dev rpm
```

2. Run the build script:
```bash
chmod +x build_packages.sh
./build_packages.sh
```

The packages will be created in the `releases/linux` directory:
- `EdgeTTS-GUI-1.0.0-linux.tar.gz`: Portable archive
- `edgetts-gui_1.0_amd64.deb`: Debian/Ubuntu package
- `edgetts-gui-1.0.x86_64.rpm`: Fedora/RHEL package

## Building on Windows

1. Install [Inno Setup](https://jrsoftware.org/isdl.php)

2. Run these commands in PowerShell:
```powershell
# Create releases directory
mkdir -p releases\windows

# Build executable
pyinstaller --clean --onefile `
    --add-data "icon.png;." `
    --icon=icon.ico `
    --name EdgeTTS-GUI `
    main.py

# Copy executable
copy dist\EdgeTTS-GUI.exe releases\windows\

# Compile installer
iscc installers\windows_setup.iss
```

The installer will be created as `releases/windows/EdgeTTS-GUI-Setup.exe`

## Building on macOS

1. Run these commands in Terminal:
```bash
# Create releases directory
mkdir -p releases/macos

# Build executable
pyinstaller --clean --onefile \
    --add-data "icon.png:." \
    --icon=icon.icns \
    --name EdgeTTS-GUI \
    main.py

# Create app bundle and DMG
cd installers
chmod +x macos_setup.sh
./macos_setup.sh
cd ..

# Move DMG to releases
mv installers/EdgeTTS-GUI.dmg releases/macos/
```

The DMG will be created as `releases/macos/EdgeTTS-GUI.dmg`

## Creating a Release

1. Build packages on all platforms
2. Collect the following files:
   - `releases/linux/EdgeTTS-GUI-1.0.0-linux.tar.gz`
   - `releases/linux/edgetts-gui_1.0_amd64.deb`
   - `releases/linux/edgetts-gui-1.0.x86_64.rpm`
   - `releases/windows/EdgeTTS-GUI-Setup.exe`
   - `releases/macos/EdgeTTS-GUI.dmg`
3. Create a new release on GitHub and upload these files

## Troubleshooting

### Linux
- If you get permission errors, make sure all scripts are executable:
  ```bash
  chmod +x build_packages.sh installers/create_linux_packages.sh
  ```

### Windows
- If PyInstaller fails, try running in a new command prompt with admin privileges
- Make sure Inno Setup is in your system PATH

### macOS
- If code signing fails, you may need to adjust the signing identity in `macos_setup.sh`
- If DMG creation fails, make sure you have write permissions in the target directory 