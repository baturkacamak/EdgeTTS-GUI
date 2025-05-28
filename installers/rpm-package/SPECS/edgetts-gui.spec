Name: edgetts-gui
Version: 1.0
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
