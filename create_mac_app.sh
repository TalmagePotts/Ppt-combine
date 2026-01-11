#!/bin/bash

# Script to create a Mac .app bundle for PowerPoint Combiner

APP_NAME="PowerPoint Combiner"
APP_DIR="${APP_NAME}.app"
CONTENTS_DIR="${APP_DIR}/Contents"
MACOS_DIR="${CONTENTS_DIR}/MacOS"
RESOURCES_DIR="${CONTENTS_DIR}/Resources"

echo "Creating Mac application bundle..."

# Create directory structure
mkdir -p "$MACOS_DIR"
mkdir -p "$RESOURCES_DIR"

# Create Info.plist
cat > "${CONTENTS_DIR}/Info.plist" << 'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>launch</string>
    <key>CFBundleName</key>
    <string>PowerPoint Combiner</string>
    <key>CFBundleIdentifier</key>
    <string>com.pptcombiner.app</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleSignature</key>
    <string>????</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.13</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
EOF

# Create launcher script
cat > "${MACOS_DIR}/launch" << 'EOF'
#!/bin/bash

# Get the Resources directory
RESOURCES_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )/../Resources" && pwd )"
cd "$RESOURCES_DIR"

# Name of virtual environment
VENV_DIR="venv"

# Create venv if it doesn't exist
if [ ! -d "$VENV_DIR" ]; then
    /usr/bin/python3 -m venv "$VENV_DIR"
fi

# Activate venv
source "$VENV_DIR/bin/activate"

# Install requirements
pip install -q -r requirements.txt

# Launch GUI
python combine_powerpoints_gui.py

# Deactivate
deactivate
EOF

# Make launcher executable
chmod +x "${MACOS_DIR}/launch"

# Copy project files to Resources
cp combine_powerpoints_gui.py "${RESOURCES_DIR}/"
cp requirements.txt "${RESOURCES_DIR}/"

echo "App bundle created successfully!"
echo ""
echo "You can now:"
echo "1. Double-click '${APP_DIR}' to launch the application"
echo "2. Drag '${APP_DIR}' to your Applications folder"
echo "3. Drag '${APP_DIR}' to your Desktop for easy access"
echo ""
echo "Note: On first launch, macOS may ask for permission to run the app."
echo "      Go to System Preferences > Security & Privacy if needed."
