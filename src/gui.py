import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
from pathlib import Path
import threading
import queue
import sys
import os

# Import the GUI configuration
from src.gui_config import GUIConfig

class RedirectText:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.queue = queue.Queue()
        self.update_timer = None

    def write(self, string):
        self.queue.put(string)
        if self.update_timer is None:
            self.update_timer = self.text_widget.after(100, self.update_text)

    def update_text(self):
        while not self.queue.empty():
            string = self.queue.get()
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, string)
            self.text_widget.see(tk.END)
            self.text_widget.configure(state='disabled')
        self.update_timer = None

    def flush(self):
        pass

class RTF2PDFGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("RTF to PDF Converter with TOC")
        self.root.geometry("800x600")
        
        # Set theme colors
        self.bg_color = "#f0f0f0"
        self.accent_color = "#007bff"
        self.root.configure(bg=self.bg_color)
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create variables with default values
        self.input_folder = tk.StringVar(value=str(Path.cwd() / "input"))
        self.output_folder = tk.StringVar(value=str(Path.cwd() / "output"))
        self.output_filename = tk.StringVar(value="final_document_with_toc.pdf")
        self.use_section_file = tk.BooleanVar(value=False)
        self.section_file = tk.StringVar(value="")
        
        # Add PDF settings
        self.page_width = tk.StringVar(value="210")
        self.margin = tk.StringVar(value="15")
        self.font_size = tk.StringVar(value="8")
        self.header_font_size = tk.StringVar(value="10")
        
        # Create widgets
        self.create_widgets()
        
        # Set initial UI state for section file controls
        self.toggle_section_file()
        
        # Create progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.main_frame,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        # Create log display
        self.create_log_display()
        
        # Set up logging
        self.setup_logging()
        
        # Create status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(
            self.main_frame,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=5)

    def create_widgets(self):
        # Input folder selection
        input_frame = ttk.LabelFrame(self.main_frame, text="Input Settings", padding="5")
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="Input Folder:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.input_folder, width=70).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=5, pady=5)
        
        # Output folder selection
        ttk.Label(input_frame, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.output_folder, width=70).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_output).grid(row=1, column=2, padx=5, pady=5)
        
        # Output filename
        ttk.Label(input_frame, text="Output Filename:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.output_filename, width=70).grid(row=2, column=1, padx=5, pady=5)
        
        # Section file checkbox and entry
        ttk.Checkbutton(input_frame, text="Use Section File", variable=self.use_section_file, 
                       command=self.toggle_section_file).grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.section_file_entry = ttk.Entry(input_frame, textvariable=self.section_file, width=70, state='disabled')
        self.section_file_entry.grid(row=3, column=1, padx=5, pady=5)
        self.section_file_button = ttk.Button(input_frame, text="Browse", command=self.browse_section_file, state='disabled')
        self.section_file_button.grid(row=3, column=2, padx=5, pady=5)
        
        # PDF Options
        pdf_frame = ttk.LabelFrame(self.main_frame, text="PDF Options", padding="5")
        pdf_frame.pack(fill=tk.X, pady=5)
        
        # Add PDF options here based on pdf_utils.py settings
        ttk.Label(pdf_frame, text="Page Width (mm):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(pdf_frame, textvariable=self.page_width, width=10).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(pdf_frame, text="Margin (mm):").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(pdf_frame, textvariable=self.margin, width=10).grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(pdf_frame, text="Font Size:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(pdf_frame, textvariable=self.font_size, width=10).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(pdf_frame, text="Header Font Size:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(pdf_frame, textvariable=self.header_font_size, width=10).grid(row=1, column=3, sticky=tk.W, padx=5, pady=5)
        
        # Buttons frame
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(pady=10)
        
        # Process button
        self.process_btn = ttk.Button(
            button_frame,
            text="Process Files",
            command=self.start_processing,
            style='Accent.TButton'
        )
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        # Create custom style for accent button
        style = ttk.Style()
        style.configure('Accent.TButton', background=self.accent_color)

    def create_log_display(self):
        log_frame = ttk.LabelFrame(self.main_frame, text="Log Output", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state='disabled')
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def setup_logging(self):
        # Configure logging
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )
        
        # Redirect logging to the text widget
        self.log_handler = RedirectText(self.log_text)
        logging.getLogger().addHandler(logging.StreamHandler(self.log_handler))

    def browse_input(self):
        folder = filedialog.askdirectory(initialdir=self.input_folder.get())
        if folder:
            self.input_folder.set(folder)

    def browse_output(self):
        folder = filedialog.askdirectory(initialdir=self.output_folder.get())
        if folder:
            self.output_folder.set(folder)

    def toggle_section_file(self):
        """Enable/disable section file entry based on checkbox state."""
        state = 'normal' if self.use_section_file.get() else 'disabled'
        self.section_file_entry.configure(state=state)
        self.section_file_button.configure(state=state)

    def browse_section_file(self):
        """Open file dialog to select section file."""
        file_path = filedialog.askopenfilename(
            initialdir=self.input_folder.get(),
            title="Select Section File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.section_file.set(file_path)

    def validate_inputs(self):
        # Check input folder
        input_path = Path(self.input_folder.get())
        if not input_path.exists():
            messagebox.showerror("Error", "Input folder does not exist")
            return False
            
        # Check output folder
        output_path = Path(self.output_folder.get())
        if not output_path.exists():
            try:
                output_path.mkdir(parents=True)
            except Exception as e:
                messagebox.showerror("Error", f"Could not create output folder: {e}")
                return False
        
        # Check section file if enabled
        if self.use_section_file.get():
            if not self.section_file.get():
                messagebox.showerror("Error", "Please select a section file")
                return False
            section_path = Path(self.section_file.get())
            if not section_path.exists():
                messagebox.showerror("Error", "Section file does not exist")
                return False
                
        # Validate numeric inputs
        try:
            float(self.page_width.get())
            float(self.margin.get())
            float(self.font_size.get())
            float(self.header_font_size.get())
        except ValueError:
            messagebox.showerror("Error", "All numeric values must be valid numbers")
            return False
            
        return True

    def start_processing(self):
        if not self.validate_inputs():
            return
            
        # Disable the process button
        self.process_btn.configure(state='disabled')
        self.status_var.set("Processing...")
        
        # Start processing in a separate thread
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()

    def process_files(self):
        try:
            # Import here to avoid circular imports
            from main import main as process_main
            
            # Create configuration from GUI values
            config = GUIConfig(
                input_folder=Path(self.input_folder.get()),
                output_folder=Path(self.output_folder.get()),
                final_output=self.output_filename.get(),
                use_section_file=self.use_section_file.get(),
                section_file_path=Path(self.section_file.get()) if self.section_file.get() else None,
                section_file_name=Path(self.section_file.get()).name if self.section_file.get() else None,
                page_width_mm=float(self.page_width.get()),
                margin_mm=float(self.margin.get()),
                font_size=float(self.font_size.get()),
                header_font_size=float(self.header_font_size.get())
            )
            
            # Log current GUI state
            logging.info(f"GUI Section file checkbox state: {config.use_section_file}")
            if config.use_section_file:
                logging.info(f"GUI Section file path: {config.section_file_path}")
                logging.info(f"Using manual section mode with file: {config.section_file_name}")
            else:
                logging.info("Using automatic section mode based on filename prefixes")
            
            # Get total number of files for progress calculation
            total_files = len(list(config.input_folder.glob("*.rtf")))
            if total_files == 0:
                raise ValueError("No RTF files found in input folder")
            
            # Calculate progress increments
            base_progress = 60  # 60% for file conversion (from 0% to 60%)
            progress_per_file = base_progress / total_files
            
            # Define progress callback
            def update_progress(value, file_progress=None):
                if file_progress is not None:
                    # Calculate progress based on base value and file progress
                    total_progress = value + (file_progress * progress_per_file)
                    self.root.after(0, lambda: self.progress_var.set(total_progress))
                else:
                    self.root.after(0, lambda: self.progress_var.set(value))
            
            # Run the main process with configuration
            process_main(
                config=config,
                progress_callback=update_progress
            )
            
            self.root.after(0, self.processing_complete, True)
        except Exception as e:
            logging.error(f"Error during processing: {e}")
            self.root.after(0, self.processing_complete, False)

    def processing_complete(self, success):
        self.process_btn.configure(state='normal')
        if success:
            self.status_var.set("Processing completed successfully")
            messagebox.showinfo("Success", "Files processed successfully!")
        else:
            self.status_var.set("Processing failed")
            messagebox.showerror("Error", "Processing failed. Check the log for details.")

def main():
    root = tk.Tk()
    app = RTF2PDFGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 