#!/bin/bash

# Install required tools
sudo apt-get update
sudo apt-get install -y dpkg-dev rpm

# Package name and version
PACKAGE_NAME="edgetts-gui"
VERSION="1.0"
MAINTAINER="Batur Kacamak <hello@batur.info>"
ARCHITECTURE="amd64"

# Create directory structure for .deb
mkdir -p debian-package/DEBIAN
mkdir -p debian-package/usr/local/bin
mkdir -p debian-package/usr/share/applications
mkdir -p debian-package/usr/share/icons/hicolor/256x256/apps

# Copy executable
cp ../dist/EdgeTTS-GUI debian-package/usr/local/bin/

# Create control file
cat > debian-package/DEBIAN/control << EOF
Package: $PACKAGE_NAME
Version: $VERSION
Architecture: $ARCHITECTURE
Maintainer: $MAINTAINER
Description: EdgeTTS GUI Application
 A user-friendly GUI for Microsoft's Edge Text-to-Speech service.
EOF

# Create desktop entry
cat > debian-package/usr/share/applications/edgetts-gui.desktop << EOF
[Desktop Entry]
Name=EdgeTTS GUI
Exec=/usr/local/bin/EdgeTTS-GUI
Icon=edgetts-gui
Type=Application
Categories=Utility;
EOF

# Build .deb package
dpkg-deb --build debian-package edgetts-gui_${VERSION}_${ARCHITECTURE}.deb

# Create directory structure for .rpm
mkdir -p rpm-package/BUILD
mkdir -p rpm-package/RPMS
mkdir -p rpm-package/SOURCES
mkdir -p rpm-package/SPECS

# Create .spec file for RPM
cat > rpm-package/SPECS/edgetts-gui.spec << EOF
Name: $PACKAGE_NAME
Version: $VERSION
Release: 1
Summary: EdgeTTS GUI Application
License: MIT
BuildArch: x86_64

%description
A user-friendly GUI for Microsoft's Edge Text-to-Speech service.

%install
mkdir -p %{buildroot}/usr/local/bin
mkdir -p %{buildroot}/usr/share/applications
cp ../../dist/EdgeTTS-GUI %{buildroot}/usr/local/bin/
cp ../../debian-package/usr/share/applications/edgetts-gui.desktop %{buildroot}/usr/share/applications/

%files
/usr/local/bin/EdgeTTS-GUI
/usr/share/applications/edgetts-gui.desktop
EOF

# Build .rpm package
rpmbuild -bb --define "_topdir $(pwd)/rpm-package" rpm-package/SPECS/edgetts-gui.spec

echo "Linux packages created successfully!" 