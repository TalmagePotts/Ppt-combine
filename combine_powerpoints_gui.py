#!/usr/bin/env python3
"""
PowerPoint Combiner - GUI Version
Combine multiple PowerPoint and PDF files into one while preserving individual themes.
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from pptx import Presentation
import threading
import io

# Try to import pdf2image for PDF support
try:
    from pdf2image import convert_from_path
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False


class PowerPointCombinerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Combiner")
        self.root.geometry("600x450")
        self.root.resizable(True, True)

        # Determine poppler path
        self.poppler_path = self.get_poppler_path()

        # Variables
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_filename = tk.StringVar(value="combined_presentation.pptx")

        self.setup_ui()
        
        if not PDF_SUPPORT:
            self.root.after(100, lambda: messagebox.showwarning(
                "PDF Support Missing", 
                "The 'pdf2image' library was not found.\nPDF files will be ignored."))
    
    def get_poppler_path(self):
        """Get the path to bundled poppler binaries if frozen."""
        if getattr(sys, 'frozen', False):
            # If running as a PyInstaller bundle
            if hasattr(sys, '_MEIPASS'):
                # PyInstaller onefile or onedir
                bundled_path = os.path.join(sys._MEIPASS, 'poppler')
                if os.path.exists(bundled_path):
                    return bundled_path
            
            # Fallback for onedir if not in MEIPASS (sometimes usually in MacOS/Resources or similar depending on config)
            # But with --add-binary or --add-data it usually ends up in MEIPASS or the executable dir.
            # Let's check the executable directory too.
            exe_dir = os.path.dirname(sys.executable)
            bundled_path = os.path.join(exe_dir, 'poppler')
            if os.path.exists(bundled_path):
                return bundled_path
                
        return None  # Rely on PATH

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
        self.status_text = tk.Text(main_frame, height=10, width=60, state='disabled')
        self.status_text.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.grid(row=5, column=2, sticky=(tk.N, tk.S, tk.E), pady=5)
        self.status_text['yscrollcommand'] = scrollbar.set

        # Combine button
        self.combine_button = ttk.Button(main_frame, text="Combine Files",
                                         command=self.combine_powerpoints,
                                         style="Accent.TButton")
        self.combine_button.grid(row=6, column=0, columnspan=3, pady=20)

        # Make rows expandable
        main_frame.rowconfigure(5, weight=1)

    def browse_input_folder(self):
        """Open dialog to select input folder."""
        folder = filedialog.askdirectory(title="Select folder containing PowerPoint/PDF files")
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

    def process_pdf(self, pdf_path, prs):
        """Convert PDF pages to images and add as slides."""
        if not PDF_SUPPORT:
            self.log_status(f"  Skipping {pdf_path.name}: pdf2image not installed.")
            return False

        self.log_status(f"  Converting PDF: {pdf_path.name}...")
        try:
            # Convert PDF to images
            try:
                images = convert_from_path(str(pdf_path), poppler_path=self.poppler_path)
            except Exception as e:
                if "poppler" in str(e).lower() or "pdfinfo" in str(e).lower():
                    self.log_status(f"  Error: Poppler is not installed/found.")
                    return False
                raise e

            slide_width = prs.slide_width
            slide_height = prs.slide_height

            for img in images:
                layout_idx = 6 if len(prs.slide_layouts) > 6 else len(prs.slide_layouts) - 1
                slide_layout = prs.slide_layouts[layout_idx]
                slide = prs.slides.add_slide(slide_layout)

                image_stream = io.BytesIO()
                img.save(image_stream, format='PNG')
                image_stream.seek(0)

                img_w, img_h = img.size
                aspect_ratio = img_w / img_h
                slide_ratio = slide_width / slide_height
                
                if aspect_ratio > slide_ratio:
                    new_w = slide_width
                    new_h = new_w / aspect_ratio
                    left = 0
                    top = (slide_height - new_h) / 2
                else:
                    new_h = slide_height
                    new_w = new_h * aspect_ratio
                    top = 0
                    left = (slide_width - new_w) / 2
                
                slide.shapes.add_picture(image_stream, left, top, new_w, new_h)
            
            self.log_status(f"  Added: {pdf_path.name} ({len(images)} slides)")
            return True

        except Exception as e:
            self.log_status(f"  Error processing PDF {pdf_path.name}: {str(e)}")
            return False

    def combine_powerpoints(self):
        """Combine files in a separate thread."""
        if not self.input_folder.get():
            messagebox.showerror("Error", "Please select an input folder")
            return

        if not self.output_folder.get():
            messagebox.showerror("Error", "Please select an output folder")
            return

        if not self.output_filename.get():
            messagebox.showerror("Error", "Please enter an output filename")
            return

        filename = self.output_filename.get()
        if not filename.endswith('.pptx'):
            filename += '.pptx'
            self.output_filename.set(filename)

        self.status_text.config(state='normal')
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state='disabled')

        self.combine_button.config(state='disabled')
        self.progress.start()

        thread = threading.Thread(target=self.do_combine, daemon=True)
        thread.start()

    def do_combine(self):
        """Perform the combination logic."""
        try:
            input_path = Path(self.input_folder.get())
            output_path = Path(self.output_folder.get()) / self.output_filename.get()

            # Get files
            files = []
            for ext in ["*.pptx", "*.pdf"]:
                files.extend(input_path.glob(ext))
            
            files = sorted(files, key=lambda p: p.name)
            files = [f for f in files if not f.name.startswith("~$")]

            if not files:
                self.root.after(0, lambda: messagebox.showerror("Error",
                    f"No PowerPoint or PDF files found in '{input_path}'"))
                return

            self.log_status(f"Found {len(files)} file(s):")
            for f in files:
                self.log_status(f"  - {f.name}")

            self.log_status("\nCombining presentations...")

            combined_prs = None
            first_file = files[0]

            # Initialize base presentation
            if first_file.suffix.lower() == '.pptx':
                try:
                    combined_prs = Presentation(first_file)
                    self.log_status(f"  Added: {first_file.name} ({len(combined_prs.slides)} slides) [Base]")
                    start_index = 1
                except Exception as e:
                    self.log_status(f"Error loading base file: {str(e)}")
                    raise e
            else:
                combined_prs = Presentation()
                start_index = 0

            # Process files
            for file_path in files[start_index:]:
                if file_path.suffix.lower() == '.pptx':
                    try:
                        prs = Presentation(file_path)
                        for slide in prs.slides:
                            slide_layout = slide.slide_layout
                            try:
                                new_slide = combined_prs.slides.add_slide(slide_layout)
                                for shape in slide.shapes:
                                    el = shape.element
                                    new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')
                            except Exception:
                                blank_layout = combined_prs.slide_layouts[6]
                                new_slide = combined_prs.slides.add_slide(blank_layout)
                                for shape in slide.shapes:
                                    el = shape.element
                                    new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')
                        self.log_status(f"  Added: {file_path.name} ({len(prs.slides)} slides)")
                    except Exception as e:
                        self.log_status(f"  Error adding {file_path.name}: {str(e)}")

                elif file_path.suffix.lower() == '.pdf':
                    self.process_pdf(file_path, combined_prs)

            combined_prs.save(str(output_path))

            self.log_status(f"\nSuccess! Created: {output_path}")
            self.log_status(f"Total slides: {len(combined_prs.slides)}")

            self.root.after(0, lambda: messagebox.showinfo("Success",
                f"Successfully combined {len(files)} files!\n"
                f"Output: {output_path}\n"
                f"Total slides: {len(combined_prs.slides)}"))

        except Exception as e:
            self.log_status(f"\nError: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error",
                f"An error occurred:\n{str(e)}"))

        finally:
            self.root.after(0, lambda: self.combine_button.config(state='normal'))
            self.root.after(0, lambda: self.progress.stop())


def main():
    """Main function to launch the GUI."""
    root = tk.Tk()
    app = PowerPointCombinerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()