#main.py
#!/usr/bin/env python3
"""
proEXy - Professional EX Certificate Processing Application
Main Application File
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import threading
import queue
import logging
from pathlib import Path
import subprocess
from datetime import datetime
import webbrowser
import threading, tkinter.messagebox as msgbox
import requests
from packaging import version

# Import our processing modules
from src.certificate_processor import CertificateProcessor
from src.config import AppConfig

def resource_path(rel_path: str) -> Path:
    """
    Return an absolute Path to a bundled resource, whether running from source
    or from a frozen exe built by cx_Freeze.
    """
    if getattr(sys, "frozen", False):                 # frozen by cx_Freeze / PyInstaller
        base_dir = Path(sys.executable).parent        # the folder with proEXy.exe
    else:                                             # running from source
        base_dir = Path(__file__).parent
    return base_dir / rel_path

class ProEXyApp:
    def __init__(self, root):
        self.root = root
        
        # ─── Set window icon ──────────────────────────────
        ico_file = resource_path("assets/EX_logo.ico")
        try:
            self.root.iconbitmap(default=str(ico_file))
        except Exception as e:
            # if something goes wrong we just log it – don’t crash the GUI
            print(f"[ICON] Could not load window icon: {e}")
        
        self.config = AppConfig()
        self.processor = CertificateProcessor()
        self.setup_ui()

        # --- Logging must be wired to a queue before setup_logging() uses it ---
        self.log_queue = queue.Queue()
        self.setup_logging()

        # Threading support
        self.processing_thread = None
        self.check_log_queue()
        
        # Track processing state
        self.is_processing = False
        
    def setup_ui(self):
        """Setup the main user interface"""
        self.root.title(f"proEXy v{self.config.VERSION}")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Set application icon
        try:
            icon_path = self.config.get_asset_path("EX_logo.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not load icon: {e}")
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Create main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title and logo
        self.create_header(main_frame)
        
        # Input section
        self.create_input_section(main_frame)
        
        # Progress section
        self.create_progress_section(main_frame)
        
        # Action buttons
        self.create_action_buttons(main_frame)
        
        # Log section
        self.create_log_section(main_frame)
        
        # Status bar
        self.create_status_bar()
        
        # Menu bar
        self.create_menu_bar()
        
    def create_header(self, parent):
        """Create application header with logo and title"""
        header_frame = ttk.Frame(parent)
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        header_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(header_frame, text="proEXy", 
                               font=('Arial', 24, 'bold'))
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        subtitle_label = ttk.Label(header_frame, text="Professional EX Certificate Processor", 
                                  font=('Arial', 12))
        subtitle_label.grid(row=1, column=0, sticky=tk.W)
        
        # Version info
        version_label = ttk.Label(header_frame, text=f"v{self.config.VERSION}", 
                                 font=('Arial', 10))
        version_label.grid(row=0, column=1, sticky=tk.E)
        
    def create_input_section(self, parent):
        """Create input folder selection section"""
        input_frame = ttk.LabelFrame(parent, text="Input Configuration", padding="15")
        input_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        input_frame.columnconfigure(1, weight=1)
        
        # Folder selection
        ttk.Label(input_frame, text="Certificate Folder:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        folder_frame = ttk.Frame(input_frame)
        folder_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        folder_frame.columnconfigure(0, weight=1)
        
        self.folder_var = tk.StringVar()
        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var, 
                                     font=('Consolas', 10))
        self.folder_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.browse_btn = ttk.Button(folder_frame, text="Browse...", 
                                    command=self.browse_folder)
        self.browse_btn.grid(row=0, column=1)
        
        # File count display
        self.file_count_var = tk.StringVar(value="No folder selected")
        self.file_count_label = ttk.Label(input_frame, textvariable=self.file_count_var,
                                         font=('Arial', 9), foreground='gray')
        self.file_count_label.grid(row=2, column=0, columnspan=2, sticky=tk.W)
        
    def create_progress_section(self, parent):
        """Create progress tracking section"""
        progress_frame = ttk.LabelFrame(parent, text="Processing Progress", padding="15")
        progress_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        progress_frame.columnconfigure(0, weight=1)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, length=400)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Progress label
        self.progress_label_var = tk.StringVar(value="Ready to process certificates")
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_label_var,
                                       font=('Arial', 10))
        self.progress_label.grid(row=1, column=0, sticky=tk.W)
        
        # Statistics frame
        stats_frame = ttk.Frame(progress_frame)
        stats_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Statistics labels
        self.stats_vars = {
            'proserv': tk.StringVar(value="Proserv: 0"),
            'oman': tk.StringVar(value="Oman: 0"),
            'unclassified': tk.StringVar(value="Unclassified: 0"),
            'errors': tk.StringVar(value="Errors: 0")
        }
        
        col = 0
        for key, var in self.stats_vars.items():
            label = ttk.Label(stats_frame, textvariable=var, font=('Arial', 9))
            label.grid(row=0, column=col, padx=(0, 20), sticky=tk.W)
            col += 1
            
    def create_action_buttons(self, parent):
        """Create action buttons section"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(0, 15))
        
        # Process button
        self.process_btn = ttk.Button(button_frame, text="🚀 Process Certificates", 
                                     command=self.start_processing,
                                     style='Accent.TButton')
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Stop button
        self.stop_btn = ttk.Button(button_frame, text="⏹ Stop Processing", 
                                  command=self.stop_processing, state='disabled')
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Open output folder button
        self.open_output_btn = ttk.Button(button_frame, text="📁 Open Output Folder", 
                                         command=self.open_output_folder)
        self.open_output_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear logs button
        self.clear_logs_btn = ttk.Button(button_frame, text="🗑 Clear Logs", 
                                        command=self.clear_logs)
        self.clear_logs_btn.pack(side=tk.LEFT)
        
    def create_log_section(self, parent):
        """Create logging section"""
        log_frame = ttk.LabelFrame(parent, text="Processing Logs", padding="10")
        log_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        parent.rowconfigure(4, weight=1)
        
        # Log text widget
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, 
                                                 font=('Consolas', 9),
                                                 wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure log text tags for different log levels
        self.log_text.tag_configure("INFO", foreground="black")
        self.log_text.tag_configure("WARNING", foreground="orange")
        self.log_text.tag_configure("ERROR", foreground="red")
        self.log_text.tag_configure("SUCCESS", foreground="green")
        
    def create_status_bar(self):
        """Create status bar"""
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, 
                                   relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
    def create_menu_bar(self):
        """Create application menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Select Folder...", command=self.browse_folder, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit, accelerator="Ctrl+Q")
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Clear Logs", command=self.clear_logs)
        view_menu.add_command(label="Open Output Folder", command=self.open_output_folder)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="GitHub Repository", command=self.open_github)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Check for updates", command=self.check_for_updates)
        # Bind keyboard shortcuts
        self.root.bind('<Control-o>', lambda e: self.browse_folder())
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        
    def setup_logging(self):
        """Setup logging configuration"""
        # Create logs directory
        logs_dir = Path("logs")
        logs_dir.mkdir(exist_ok=True)
        
        # Configure logging
        log_filename = logs_dir / f"proEXy_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename),
                logging.StreamHandler()
            ]
        )
        
        # Custom handler for GUI
        self.gui_handler = GUILogHandler(self.log_queue)
        logging.getLogger().addHandler(self.gui_handler)
        
    def browse_folder(self):
        """Open folder browser dialog"""
        folder = filedialog.askdirectory(
            title="Select Certificate Folder",
            initialdir=os.getcwd()
        )
        
        if folder:
            self.folder_var.set(folder)
            self.update_file_count()
            self.status_var.set(f"Selected folder: {os.path.basename(folder)}")
            
    def update_file_count(self):
        """Update file count display"""
        folder = self.folder_var.get()
        if not folder or not os.path.exists(folder):
            self.file_count_var.set("No folder selected")
            return
            
        pdf_files = list(Path(folder).glob("*.pdf"))
        count = len(pdf_files)
        
        if count == 0:
            self.file_count_var.set("⚠ No PDF files found in selected folder")
        else:
            self.file_count_var.set(f"📄 {count} PDF file{'s' if count != 1 else ''} found")
            
    def start_processing(self):
        """Start certificate processing in separate thread"""
        folder = self.folder_var.get()
        
        if not folder:
            messagebox.showerror("Error", "Please select a certificate folder first.")
            return
            
        if not os.path.exists(folder):
            messagebox.showerror("Error", "Selected folder does not exist.")
            return
            
        pdf_files = list(Path(folder).glob("*.pdf"))
        if not pdf_files:
            messagebox.showerror("Error", "No PDF files found in the selected folder.")
            return
            
        # Reset UI for new processing
        self.is_processing = True
        self.progress_var.set(0)
        self.progress_label_var.set("Starting processing...")
        self.process_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.browse_btn.config(state='disabled')
        
        # Reset statistics
        for var in self.stats_vars.values():
            var.set(var.get().split(':')[0] + ": 0")
            
        # Start processing thread
        self.processing_thread = threading.Thread(
            target=self._process_certificates_thread,
            args=(folder,),
            daemon=True
        )
        self.processing_thread.start()
        
    def _process_certificates_thread(self, folder):
        """Certificate processing thread function"""
        try:
            self.processor.process_certificates(
                folder, 
                progress_callback=self.update_progress,
                stats_callback=self.update_stats
            )
            
            if self.is_processing:  # Check if not stopped
                self.log_queue.put(("SUCCESS", "🎉 Processing completed successfully!"))
                self.root.after(0, self.processing_completed)
                
        except Exception as e:
            logging.error(f"Processing failed: {e}")
            self.root.after(0, lambda: self.processing_failed(str(e)))
            
    def update_progress(self, current, total, message=""):
        """Update progress bar and label"""
        if not self.is_processing:
            return False  # Signal to stop processing
            
        progress = (current / total) * 100 if total > 0 else 0
        self.root.after(0, lambda: self.progress_var.set(progress))
        
        if message:
            self.root.after(0, lambda: self.progress_label_var.set(message))
            
        return True  # Continue processing
        
    def update_stats(self, stats):
        """Update statistics display"""
        for key, value in stats.items():
            if key in self.stats_vars:
                self.root.after(0, lambda k=key, v=value: 
                               self.stats_vars[k].set(f"{k.title()}: {v}"))
                
    def stop_processing(self):
        """Stop the processing"""
        self.is_processing = False
        self.progress_label_var.set("⏹ Stopping processing...")
        self.stop_btn.config(state='disabled')
        logging.info("Processing stopped by user")
        
        # Re-enable controls after a short delay
        self.root.after(1000, self.reset_ui_after_stop)
        
    def reset_ui_after_stop(self):
        """Reset UI after stopping"""
        self.process_btn.config(state='normal')
        self.browse_btn.config(state='normal')
        self.progress_label_var.set("Processing stopped by user")
        self.status_var.set("Ready")
        
    def processing_completed(self):
        """Handle successful processing completion"""
        self.is_processing = False
        self.process_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.browse_btn.config(state='normal')
        self.progress_var.set(100)
        self.progress_label_var.set("✅ Processing completed successfully!")
        self.status_var.set("Processing completed")
        
        # Show completion message
        messagebox.showinfo(
            "Success", 
            "Certificate processing completed successfully!\n\n"
            "Output files:\n"
            "• Proserv_Certificates.xlsx\n"
            "• Oman_Certificates.xlsx"
        )
        
    def processing_failed(self, error_message):
        """Handle processing failure"""
        self.is_processing = False
        self.process_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.browse_btn.config(state='normal')
        self.progress_label_var.set("❌ Processing failed")
        self.status_var.set("Processing failed")
        
        messagebox.showerror("Error", f"Processing failed:\n{error_message}")
        
    def open_output_folder(self):
        """Open the output folder in file explorer"""
        # Always open the folder the user chose
        output_folder = self.folder_var.get() or os.getcwd()
        try:
            if sys.platform == "win32":
                os.startfile(output_folder)
            elif sys.platform == "darwin":
                subprocess.run(["open", output_folder])
            else:
                subprocess.run(["xdg-open", output_folder])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open output folder:\n{e}")
            
    def clear_logs(self):
        """Clear the log display"""
        self.log_text.delete(1.0, tk.END)
        self.status_var.set("Logs cleared")
        
    def open_github(self):
        """Open GitHub repository"""
        webbrowser.open("https://github.com/noelmathen/EX_Certificate_Application")
        
    def show_about(self):
        """Show about dialog"""
        about_text = f"""proEXy v{self.config.VERSION}
Professional EX Certificate Processor

A desktop application for processing EX certificates
from Proserv and Oman sources into Excel files.

GitHub: https://github.com/noelmathen/EX_Certificate_Application

© 2025 - Built with Python & tkinter"""
        
        messagebox.showinfo("About proEXy", about_text)
    
    def check_for_updates(self):
        """Fire off the GitHub release check in a thread."""
        threading.Thread(target=self._check_for_updates, daemon=True).start()

    def _check_for_updates(self):
        try:
            api = f"https://api.github.com/repos/{self.config.REPO}/releases/latest"
            latest = requests.get(api, timeout=8).json().get("tag_name", "")
            cur, lat = self.config.VERSION, latest.lstrip("v")
            if latest and version.parse(lat) > version.parse(cur):
                if msgbox.askyesno("Update available",
                                f"New version {lat} is available "
                                f"(you have {cur}).\n\nDownload now?"):
                    webbrowser.open(
                        f"https://github.com/{self.config.REPO}/releases/latest")
            else:
                msgbox.showinfo("Up to date", f"You’re already on {cur}.")
        except Exception as e:
            msgbox.showwarning("Update check failed", str(e))
        
    def check_log_queue(self):
        """Check for new log messages and display them"""
        try:
            while True:
                level, message = self.log_queue.get_nowait()
                timestamp = datetime.now().strftime("%H:%M:%S")
                formatted_message = f"[{timestamp}] {message}\n"
                
                self.log_text.insert(tk.END, formatted_message, level)
                self.log_text.see(tk.END)
                
        except queue.Empty:
            pass
            
        # Schedule next check
        self.root.after(100, self.check_log_queue)


class GUILogHandler(logging.Handler):
    """Custom logging handler for GUI display"""
    
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue
        
    def emit(self, record):
        level = record.levelname
        message = self.format(record)
        self.log_queue.put((level, message))


def main():
    """Main application entry point"""
    root = tk.Tk()
    app = ProEXyApp(root)
    
    # Center the window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    # Handle window close
    def on_closing():
        if app.is_processing:
            if messagebox.askokcancel("Quit", "Processing is in progress. Do you want to quit?"):
                app.is_processing = False
                root.destroy()
        else:
            root.destroy()
            
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # Start the GUI
    root.mainloop()


if __name__ == "__main__":
    main()