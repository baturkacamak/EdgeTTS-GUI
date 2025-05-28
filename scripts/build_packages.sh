#!/bin/bash

# Create release directory
mkdir -p releases

# Function to check if last command was successful
check_status() {
    if [ $? -ne 0 ]; then
        echo "Error: $1 failed"
        exit 1
    fi
}

echo "Building packages for EdgeTTS-GUI..."

# Clean previous builds
rm -rf build/ dist/
check_status "Cleaning previous builds"

# Create version file
VERSION="1.0.0"
echo "VERSION = '$VERSION'" > version.py

# Build Linux executable
echo "Building Linux executable..."
pyinstaller --clean --onefile \
    --add-data "icon.png:." \
    --icon=icon.ico \
    --name EdgeTTS-GUI \
    main.py
check_status "Linux executable build"

# Create releases directory structure
mkdir -p releases/linux
mkdir -p releases/windows
mkdir -p releases/macos

# Copy Linux executable to releases
cp dist/EdgeTTS-GUI "releases/linux/EdgeTTS-GUI"
check_status "Copying Linux executable"

# Build Linux packages (.deb and .rpm)
echo "Building Linux packages..."
cd installers
cp ../dist/EdgeTTS-GUI .
./create_linux_packages.sh
check_status "Linux packages build"
cd ..

# Move Linux packages to releases
mv installers/*.deb "releases/linux/"
mv installers/*.rpm "releases/linux/"
mv "releases/linux/EdgeTTS-GUI" "releases/linux/EdgeTTS-GUI-$VERSION"

# Create Linux tarball
cd releases/linux
tar czf "EdgeTTS-GUI-$VERSION-linux.tar.gz" *
cd ../..

echo "Done! Release packages are in the releases directory."
echo "Don't forget to create Windows and macOS packages on their respective platforms."

# List created files
echo -e "\nCreated files:"
ls -lR releases/ 