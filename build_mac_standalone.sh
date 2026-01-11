#!/bin/bash

# Build script for creating standalone Mac app (no Python required)

echo "Building standalone PowerPoint Combiner for Mac..."

# Install required dependencies
echo "Installing dependencies..."
pip3 install -r requirements.txt

# Clean previous builds
echo "Cleaning previous builds..."
rm -rf build dist "PowerPoint Combiner.app"

# Initialize PyInstaller args
PYINSTALLER_ARGS=(
    --name "PowerPoint Combiner"
    --windowed
    --onedir
    --clean
    --hidden-import=pptx
    --hidden-import=pptx.presentation
    --hidden-import=pptx.util
    --hidden-import=pptx.enum
    --hidden-import=lxml
    --hidden-import=lxml.etree
    --hidden-import=pdf2image
    --collect-all=pptx
    --osx-bundle-identifier "com.pptcombiner.app"
)

# Check for Poppler (required for PDF support)
if command -v brew &> /dev/null; then
    POPPLER_PREFIX=$(brew --prefix poppler 2>/dev/null)
    if [ ! -z "$POPPLER_PREFIX" ] && [ -d "$POPPLER_PREFIX/bin" ]; then
        echo "Found Poppler at $POPPLER_PREFIX"
        # Add Poppler binaries to the bundle
        # We add the whole bin folder to a 'poppler' directory in the bundle
        PYINSTALLER_ARGS+=(--add-binary "$POPPLER_PREFIX/bin/*:poppler")
        echo "Enabled PDF support (bundled Poppler)."
    else
        echo "⚠️  Poppler not found via Homebrew."
        echo "   PDF support might not work on machines without Poppler installed."
        echo "   Install it with: brew install poppler"
    fi
else
    echo "⚠️  Homebrew not found. Cannot auto-detect Poppler."
fi

# Build the standalone app
echo "Building app bundle..."
pyinstaller "${PYINSTALLER_ARGS[@]}" combine_powerpoints_gui.py

# Check if pyinstaller succeeded
if [ $? -ne 0 ]; then
    echo "❌ PyInstaller build failed. Check the error messages above."
    exit 1
fi

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