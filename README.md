# PowerPoint Combiner

Combine multiple PowerPoint files into one while preserving each file's original theme and formatting.

## Quick Start (Mac Users)

### Option 1: Double-Click App (Easiest)
1. Run the setup script:
   ```bash
   ./create_mac_app.sh
   ```
2. Double-click the "PowerPoint Combiner.app" that was created
3. Use the GUI to select your folder and combine presentations

### Option 2: Simple Launcher
1. Double-click `launch_mac.sh` (it will auto-setup everything)
2. The GUI will open automatically

### Option 3: Manual Installation
Install the required dependency:
```bash
pip install -r requirements.txt
```

Then run the GUI:
```bash
python combine_powerpoints_gui.py
```

## Usage

### GUI Version (Recommended)
The GUI version provides an easy-to-use interface where you can:
- Browse and select the input folder containing PowerPoint files
- Choose where to save the combined output
- Name your output file
- See real-time progress and status updates

Run it with:
```bash
python combine_powerpoints_gui.py
```

Or on Mac, just double-click `launch_mac.sh` or the `.app` bundle.

### Command Line Version
For automation or scripting, use the CLI version:

#### Basic usage (current directory)
```bash
python combine_powerpoints.py
```
This will combine all `.pptx` files in the current directory into `combined_presentation.pptx`

#### Specify input folder
```bash
python combine_powerpoints.py /path/to/folder
```

#### Specify both input folder and output file
```bash
python combine_powerpoints.py /path/to/folder output_name.pptx
```

## Features

- Combines all PowerPoint files in a folder sequentially
- Preserves original themes and formatting from each source file
- Files are processed in alphabetical order
- Ignores temporary PowerPoint files (starting with `~$`)
- Shows progress and slide counts during processing

## Example

If you have a folder with these files:
- presentation1.pptx (10 slides)
- presentation2.pptx (5 slides)
- presentation3.pptx (8 slides)

Running the script will create `combined_presentation.pptx` with 23 slides total, where each section maintains its original theme.
