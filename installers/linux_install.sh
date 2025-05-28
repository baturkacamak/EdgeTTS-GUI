#!/bin/bash

# Define the installation directory
INSTALL_DIR="$HOME/.local/bin"
DESKTOP_ENTRY_DIR="$HOME/.local/share/applications"

# Create the installation directory if it doesn't exist
mkdir -p "$INSTALL_DIR"
mkdir -p "$DESKTOP_ENTRY_DIR"

# Copy the executable to the installation directory
cp dist/main "$INSTALL_DIR/EdgeTTS-GUI"

# Create a desktop entry
echo "[Desktop Entry]
Name=EdgeTTS-GUI
Exec=$INSTALL_DIR/EdgeTTS-GUI
Icon=utilities-terminal
Type=Application
Categories=Utility;" > "$DESKTOP_ENTRY_DIR/EdgeTTS-GUI.desktop"

# Make the desktop entry executable
chmod +x "$DESKTOP_ENTRY_DIR/EdgeTTS-GUI.desktop"

echo "Installation complete. You can now run EdgeTTS-GUI from your applications menu." 