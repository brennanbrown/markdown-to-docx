#!/usr/bin/env python3
"""
Markdown to DOCX Converter - GUI Version

A user-friendly graphical interface for converting markdown files to DOCX format.
No command line knowledge required!

Features:
- Drag and drop or browse for input folder
- Visual progress tracking
- Easy output folder selection
- Real-time conversion status
- Error handling with helpful messages

Author: Brennan Kenneth Brown
License: MIT
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
from pathlib import Path
import pypandoc
from datetime import datetime

class MarkdownConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Markdown to DOCX Converter")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Variables
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.conversion_running = False
        
        # Setup the GUI
        self.setup_gui()
        self.check_dependencies()
    
    def setup_gui(self):
        """Create the main GUI layout"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Markdown to DOCX Converter", 
                               font=("Helvetica", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input folder selection
        ttk.Label(main_frame, text="Select folder with markdown files:").grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        input_frame.columnconfigure(0, weight=1)
        
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_folder, font=("Helvetica", 10))
        self.input_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(input_frame, text="Browse", command=self.browse_input_folder).grid(row=0, column=1)
        
        # Output folder selection
        ttk.Label(main_frame, text="Output folder (optional - leave blank for auto-generated):").grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_folder, font=("Helvetica", 10))
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(output_frame, text="Browse", command=self.browse_output_folder).grid(row=0, column=1)
        
        # Convert button
        self.convert_button = ttk.Button(main_frame, text="Convert Files", 
                                        command=self.start_conversion, style="Accent.TButton")
        self.convert_button.grid(row=5, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
                                          maximum=100, length=400)
        self.progress_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status label
        self.status_var = tk.StringVar(value="Ready to convert files")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var, 
                                     font=("Helvetica", 10))
        self.status_label.grid(row=7, column=0, columnspan=3, pady=(0, 10))
        
        # Results text area
        ttk.Label(main_frame, text="Conversion Results:").grid(row=8, column=0, columnspan=3, sticky=tk.W, pady=(10, 5))
        
        # Frame for text widget and scrollbar
        text_frame = ttk.Frame(main_frame)
        text_frame.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(9, weight=1)
        
        self.results_text = tk.Text(text_frame, height=8, wrap=tk.WORD, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Info label
        info_text = "💡 Tip: Select a folder containing .md files. The converter will preserve your folder structure!"
        ttk.Label(main_frame, text=info_text, font=("Helvetica", 9), 
                 foreground="gray").grid(row=10, column=0, columnspan=3, pady=(10, 0))
    
    def check_dependencies(self):
        """Check if required dependencies are available"""
        try:
            import pypandoc
            pypandoc.get_pandoc_version()
            self.log_message("✅ Dependencies check passed - ready to convert!")
        except ImportError:
            self.log_message("❌ Error: pypandoc is not installed.")
            self.log_message("Please install it with: pip install pypandoc")
            self.convert_button.configure(state="disabled")
        except OSError:
            self.log_message("❌ Error: pandoc is not installed on your system.")
            self.log_message("Please install pandoc from: https://pandoc.org/installing.html")
            self.convert_button.configure(state="disabled")
    
    def browse_input_folder(self):
        """Open file dialog to select input folder"""
        folder = filedialog.askdirectory(title="Select folder containing markdown files")
        if folder:
            self.input_folder.set(folder)
    
    def browse_output_folder(self):
        """Open file dialog to select output folder"""
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_folder.set(folder)
    
    def log_message(self, message):
        """Add message to results text area"""
        self.results_text.insert(tk.END, message + "\n")
        self.results_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_results(self):
        """Clear the results text area"""
        self.results_text.delete(1.0, tk.END)
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        if self.conversion_running:
            return
            
        # Validate input
        if not self.input_folder.get():
            messagebox.showerror("Error", "Please select an input folder containing markdown files.")
            return
        
        if not os.path.exists(self.input_folder.get()):
            messagebox.showerror("Error", "The selected input folder does not exist.")
            return
        
        # Disable convert button and start conversion
        self.conversion_running = True
        self.convert_button.configure(state="disabled", text="Converting...")
        self.clear_results()
        
        # Start conversion in separate thread to prevent GUI freezing
        thread = threading.Thread(target=self.convert_files)
        thread.daemon = True
        thread.start()
    
    def convert_files(self):
        """Main conversion logic (runs in separate thread)"""
        try:
            # Import conversion functions from the original script
            from markdown_to_docx_converter import (
                find_markdown_files, 
                preserve_folder_structure, 
                convert_markdown_to_docx,
                setup_output_directory
            )
            
            source_dir = Path(self.input_folder.get()).resolve()
            output_base = self.output_folder.get() if self.output_folder.get() else None
            
            self.status_var.set("Scanning for markdown files...")
            self.log_message(f"📁 Scanning folder: {source_dir}")
            
            # Find all markdown files
            markdown_files = find_markdown_files(source_dir)
            
            if not markdown_files:
                self.log_message("⚠️  No markdown files found in the selected folder.")
                self.status_var.set("No files to convert")
                return
            
            self.log_message(f"📋 Found {len(markdown_files)} markdown files to convert")
            
            # Setup output directory
            output_dir = setup_output_directory(source_dir, output_base)
            self.log_message(f"📤 Output directory: {output_dir}")
            self.log_message("-" * 50)
            
            # Convert files
            successful_conversions = 0
            failed_conversions = 0
            
            for i, md_file in enumerate(markdown_files):
                # Update progress
                progress = ((i + 1) / len(markdown_files)) * 100
                self.progress_var.set(progress)
                self.status_var.set(f"Converting file {i + 1} of {len(markdown_files)}")
                
                # Preserve folder structure
                output_folder = preserve_folder_structure(md_file, source_dir, output_dir)
                
                # Create output filename
                docx_filename = md_file.stem + ".docx"
                output_path = output_folder / docx_filename
                
                relative_input = md_file.relative_to(source_dir)
                relative_output = output_path.relative_to(output_dir)
                
                self.log_message(f"🔄 Converting: {relative_input}")
                
                # Convert the file
                if convert_markdown_to_docx(md_file, output_path):
                    successful_conversions += 1
                    self.log_message(f"   ✅ Success → {relative_output}")
                else:
                    failed_conversions += 1
                    self.log_message(f"   ❌ Failed")
            
            # Final results
            self.log_message("-" * 50)
            self.log_message(f"🎉 Conversion complete!")
            self.log_message(f"✅ Successfully converted: {successful_conversions} files")
            if failed_conversions > 0:
                self.log_message(f"❌ Failed conversions: {failed_conversions} files")
            self.log_message(f"📁 Output location: {output_dir}")
            
            self.status_var.set(f"Completed! {successful_conversions} files converted successfully")
            
            # Ask if user wants to open output folder
            if successful_conversions > 0:
                self.root.after(100, lambda: self.ask_open_folder(output_dir))
                
        except Exception as e:
            self.log_message(f"❌ Error during conversion: {str(e)}")
            self.status_var.set("Conversion failed")
        finally:
            # Re-enable convert button
            self.root.after(100, self.conversion_complete)
    
    def conversion_complete(self):
        """Called when conversion is complete"""
        self.conversion_running = False
        self.convert_button.configure(state="normal", text="Convert Files")
        self.progress_var.set(100)
    
    def ask_open_folder(self, output_dir):
        """Ask user if they want to open the output folder"""
        result = messagebox.askyesno("Conversion Complete", 
                                   "Conversion completed successfully!\n\n"
                                   "Would you like to open the output folder to see your converted files?")
        if result:
            self.open_folder(output_dir)
    
    def open_folder(self, folder_path):
        """Open folder in file manager"""
        try:
            if sys.platform == "darwin":  # macOS
                os.system(f'open "{folder_path}"')
            elif sys.platform == "win32":  # Windows
                os.system(f'explorer "{folder_path}"')
            else:  # Linux
                os.system(f'xdg-open "{folder_path}"')
        except Exception as e:
            self.log_message(f"Could not open folder: {str(e)}")

def main():
    """Main function to run the GUI application"""
    root = tk.Tk()
    
    # Set up the style
    style = ttk.Style()
    if "Accent.TButton" not in style.theme_names():
        style.configure("Accent.TButton", font=("Helvetica", 11, "bold"))
    
    app = MarkdownConverterGUI(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()
