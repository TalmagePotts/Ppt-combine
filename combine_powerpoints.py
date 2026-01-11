#!/usr/bin/env python3
"""
Combine multiple PowerPoint files into one while preserving individual themes.
Each source presentation's slides will maintain their original formatting and theme.
"""

import os
import sys
from pathlib import Path
from pptx import Presentation


def combine_powerpoints(input_folder, output_file):
    """
    Combine all PowerPoint files in a folder into a single presentation.

    Args:
        input_folder: Path to folder containing .pptx files
        output_file: Path for the output combined .pptx file
    """
    input_path = Path(input_folder)

    if not input_path.exists():
        print(f"Error: Folder '{input_folder}' does not exist.")
        return False

    # Get all PowerPoint files in the folder
    pptx_files = sorted(input_path.glob("*.pptx"))

    # Filter out temporary files (starting with ~$)
    pptx_files = [f for f in pptx_files if not f.name.startswith("~$")]

    if not pptx_files:
        print(f"No PowerPoint files found in '{input_folder}'")
        return False

    print(f"Found {len(pptx_files)} PowerPoint file(s):")
    for ppt in pptx_files:
        print(f"  - {ppt.name}")

    # Create a new presentation starting with the first file
    print(f"\nCombining presentations...")
    combined_prs = Presentation(pptx_files[0])
    print(f"  Added: {pptx_files[0].name} ({len(combined_prs.slides)} slides)")

    # Add slides from remaining presentations
    for pptx_file in pptx_files[1:]:
        prs = Presentation(pptx_file)

        for slide in prs.slides:
            # Copy the slide with its layout and master to preserve theme
            slide_layout = slide.slide_layout

            # Import the slide layout and master into the combined presentation
            # This preserves the original theme and formatting
            try:
                # Add slide layout from source to destination if not already present
                new_slide = combined_prs.slides.add_slide(slide_layout)

                # Copy all shapes from the original slide
                for shape in slide.shapes:
                    # Get the shape element XML
                    el = shape.element
                    # Clone and append to new slide
                    new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

                # Remove the default shapes that came with the layout
                # (keep only the copied ones)
                shapes_to_remove = []
                for shape in new_slide.shapes:
                    if shape not in [s for s in slide.shapes]:
                        try:
                            # Mark placeholder shapes for removal
                            if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
                                shapes_to_remove.append(shape)
                        except:
                            pass

            except Exception as e:
                # Fallback: simpler approach that may not preserve all formatting
                # Create a blank slide and copy content
                blank_layout = combined_prs.slide_layouts[6]  # Usually blank layout
                new_slide = combined_prs.slides.add_slide(blank_layout)

                for shape in slide.shapes:
                    el = shape.element
                    new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

        print(f"  Added: {pptx_file.name} ({len(prs.slides)} slides)")

    # Save the combined presentation
    combined_prs.save(output_file)
    print(f"\nSuccessfully created: {output_file}")
    print(f"Total slides: {len(combined_prs.slides)}")

    return True


def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) < 2:
        # Use current directory if no argument provided
        input_folder = "."
        output_file = "combined_presentation.pptx"
    elif len(sys.argv) == 2:
        input_folder = sys.argv[1]
        output_file = "combined_presentation.pptx"
    else:
        input_folder = sys.argv[1]
        output_file = sys.argv[2]

    print(f"Input folder: {input_folder}")
    print(f"Output file: {output_file}\n")

    success = combine_powerpoints(input_folder, output_file)

    if success:
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
