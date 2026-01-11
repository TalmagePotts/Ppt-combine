#!/bin/bash

# Build script for creating standalone Mac app (no Python required)

echo "Building standalone PowerPoint Combiner for Mac..."

# Check if PyInstaller is installed
if ! command -v pyinstaller &> /dev/null; then
    echo "PyInstaller not found. Installing..."
    pip install pyinstaller
fi

# Clean previous builds
echo "Cleaning previous builds..."
rm -rf build dist "PowerPoint Combiner.app"

# Build the standalone app
echo "Building app bundle..."
pyinstaller --name "PowerPoint Combiner" \
    --windowed \
    --onefile \
    --clean \
    --icon=NONE \
    --osx-bundle-identifier "com.pptcombiner.app" \
    combine_powerpoints_gui.py

# Move the app to the current directory for easy access
if [ -d "dist/PowerPoint Combiner.app" ]; then
    echo "Moving app to current directory..."
    mv "dist/PowerPoint Combiner.app" .

    # Clean up build artifacts
    echo "Cleaning up build files..."
    rm -rf build dist *.spec

    echo ""
    echo "✓ SUCCESS! Your standalone app is ready."
    echo ""
    echo "You can now:"
    echo "  1. Double-click 'PowerPoint Combiner.app' to run it"
    echo "  2. Share this app with your friend - no Python needed!"
    echo "  3. Optional: Move it to /Applications folder"
    echo ""
    echo "Note: On first launch, macOS may ask for permission since it's unsigned."
    echo "      Go to System Preferences > Security & Privacy to allow it."
else
    echo "❌ Build failed. Check the error messages above."
    exit 1
fi
