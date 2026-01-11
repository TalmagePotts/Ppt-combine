#!/usr/bin/env python3
"""
PowerPoint Combiner - GUI Version
Combine multiple PowerPoint files into one while preserving individual themes.
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from pptx import Presentation
import threading


class PowerPointCombinerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Combiner")
        self.root.geometry("600x400")
        self.root.resizable(True, True)

        # Variables
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_filename = tk.StringVar(value="combined_presentation.pptx")

        self.setup_ui()

    def setup_ui(self):
        """Create the user interface."""
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="PowerPoint Combiner",
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Input folder selection
        ttk.Label(main_frame, text="Input Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_folder, width=50).grid(
            row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_input_folder).grid(
            row=1, column=2, pady=5)

        # Output folder selection
        ttk.Label(main_frame, text="Output Folder:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_folder, width=50).grid(
            row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_output_folder).grid(
            row=2, column=2, pady=5)

        # Output filename
        ttk.Label(main_frame, text="Output Filename:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_filename, width=50).grid(
            row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=20)

        # Status text area
        ttk.Label(main_frame, text="Status:").grid(row=5, column=0, sticky=tk.NW, pady=5)
        self.status_text = tk.Text(main_frame, height=8, width=60, state='disabled')
        self.status_text.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.grid(row=5, column=2, sticky=(tk.N, tk.S, tk.E), pady=5)
        self.status_text['yscrollcommand'] = scrollbar.set

        # Combine button
        self.combine_button = ttk.Button(main_frame, text="Combine PowerPoints",
                                         command=self.combine_powerpoints,
                                         style="Accent.TButton")
        self.combine_button.grid(row=6, column=0, columnspan=3, pady=20)

        # Make rows expandable
        main_frame.rowconfigure(5, weight=1)

    def browse_input_folder(self):
        """Open dialog to select input folder."""
        folder = filedialog.askdirectory(title="Select folder containing PowerPoint files")
        if folder:
            self.input_folder.set(folder)
            # Auto-set output folder to same location if not set
            if not self.output_folder.get():
                self.output_folder.set(folder)

    def browse_output_folder(self):
        """Open dialog to select output folder."""
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_folder.set(folder)

    def log_status(self, message):
        """Add message to status text area."""
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')
        self.root.update()

    def combine_powerpoints(self):
        """Combine PowerPoint files in a separate thread."""
        # Validate inputs
        if not self.input_folder.get():
            messagebox.showerror("Error", "Please select an input folder")
            return

        if not self.output_folder.get():
            messagebox.showerror("Error", "Please select an output folder")
            return

        if not self.output_filename.get():
            messagebox.showerror("Error", "Please enter an output filename")
            return

        # Ensure .pptx extension
        filename = self.output_filename.get()
        if not filename.endswith('.pptx'):
            filename += '.pptx'
            self.output_filename.set(filename)

        # Clear status
        self.status_text.config(state='normal')
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state='disabled')

        # Disable button and start progress
        self.combine_button.config(state='disabled')
        self.progress.start()

        # Run combination in separate thread
        thread = threading.Thread(target=self.do_combine, daemon=True)
        thread.start()

    def do_combine(self):
        """Perform the actual PowerPoint combination."""
        try:
            input_path = Path(self.input_folder.get())
            output_path = Path(self.output_folder.get()) / self.output_filename.get()

            # Get all PowerPoint files
            pptx_files = sorted(input_path.glob("*.pptx"))
            pptx_files = [f for f in pptx_files if not f.name.startswith("~$")]

            if not pptx_files:
                self.root.after(0, lambda: messagebox.showerror("Error",
                    f"No PowerPoint files found in '{input_path}'"))
                return

            self.log_status(f"Found {len(pptx_files)} PowerPoint file(s):")
            for ppt in pptx_files:
                self.log_status(f"  - {ppt.name}")

            self.log_status("\nCombining presentations...")

            # Create combined presentation starting with first file
            combined_prs = Presentation(pptx_files[0])
            self.log_status(f"  Added: {pptx_files[0].name} ({len(combined_prs.slides)} slides)")

            # Add slides from remaining presentations
            for pptx_file in pptx_files[1:]:
                prs = Presentation(pptx_file)

                for slide in prs.slides:
                    slide_layout = slide.slide_layout

                    try:
                        new_slide = combined_prs.slides.add_slide(slide_layout)

                        # Copy all shapes from original slide
                        for shape in slide.shapes:
                            el = shape.element
                            new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

                    except Exception as e:
                        # Fallback to blank layout
                        blank_layout = combined_prs.slide_layouts[6]
                        new_slide = combined_prs.slides.add_slide(blank_layout)

                        for shape in slide.shapes:
                            el = shape.element
                            new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

                self.log_status(f"  Added: {pptx_file.name} ({len(prs.slides)} slides)")

            # Save combined presentation
            combined_prs.save(str(output_path))

            self.log_status(f"\nSuccess! Created: {output_path}")
            self.log_status(f"Total slides: {len(combined_prs.slides)}")

            # Show success message
            self.root.after(0, lambda: messagebox.showinfo("Success",
                f"Successfully combined {len(pptx_files)} presentations!\n\n"
                f"Output: {output_path}\n"
                f"Total slides: {len(combined_prs.slides)}"))

        except Exception as e:
            self.log_status(f"\nError: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error",
                f"An error occurred:\n{str(e)}"))

        finally:
            # Re-enable button and stop progress
            self.root.after(0, lambda: self.combine_button.config(state='normal'))
            self.root.after(0, lambda: self.progress.stop())


def main():
    """Main function to launch the GUI."""
    root = tk.Tk()
    app = PowerPointCombinerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
