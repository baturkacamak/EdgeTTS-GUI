#!/bin/bash

# Create the application bundle structure
APP_NAME="EdgeTTS-GUI"
CONTENTS_DIR="$APP_NAME.app/Contents"
MACOS_DIR="$CONTENTS_DIR/MacOS"
RESOURCES_DIR="$CONTENTS_DIR/Resources"

mkdir -p "$MACOS_DIR" "$RESOURCES_DIR"

# Copy the executable
cp ../dist/EdgeTTS-GUI "$MACOS_DIR/"
chmod +x "$MACOS_DIR/EdgeTTS-GUI"

# Create Info.plist
cat > "$CONTENTS_DIR/Info.plist" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>EdgeTTS-GUI</string>
    <key>CFBundleIdentifier</key>
    <string>com.batur.edgettsGUI</string>
    <key>CFBundleName</key>
    <string>EdgeTTS-GUI</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.10</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
EOF

# Create DMG
hdiutil create -volname "EdgeTTS-GUI" -srcfolder "$APP_NAME.app" -ov -format UDZO "EdgeTTS-GUI.dmg"

echo "macOS application bundle and DMG created successfully!" 