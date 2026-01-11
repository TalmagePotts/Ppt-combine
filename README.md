# PowerPoint Combiner

Combine multiple PowerPoint files into one while preserving each file's original theme and formatting.

## Installation

Install the required dependency:

```bash
pip install -r requirements.txt
```

## Usage

### Basic usage (current directory)
```bash
python combine_powerpoints.py
```
This will combine all `.pptx` files in the current directory into `combined_presentation.pptx`

### Specify input folder
```bash
python combine_powerpoints.py /path/to/folder
```

### Specify both input folder and output file
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
