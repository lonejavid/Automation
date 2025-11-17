"""
KFC Drive-Thru Automation - Desktop Application
Windows Desktop App with GUI
"""

import sys
import os
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime, timedelta

# Handle both frozen (executable) and script modes
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    BASE_DIR = Path(sys.executable).parent
    SRC_DIR = BASE_DIR / "src"
    SCRIPTS_DIR = BASE_DIR / "scripts"
    sys.path.insert(0, str(SRC_DIR))
else:
    # Running as script
    BASE_DIR = Path(__file__).parent
    SRC_DIR = BASE_DIR / "src"
    SCRIPTS_DIR = BASE_DIR / "scripts"
    sys.path.insert(0, str(SRC_DIR))

# Set up paths
TEMPLATE_DIR = BASE_DIR / "data" / "templates"
DOWNLOADS_DIR = BASE_DIR / "data" / "downloads"
LOGS_DIR = BASE_DIR / "logs"

# Create directories if they don't exist
TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)


class AutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KFC Drive-Thru Automation")
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        
        # Center window on screen
        self.center_window()
        
        # Variables
        self.automation_running = False
        self.process = None
        
        # Create UI
        self.create_ui()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_ui(self):
        """Create the user interface"""
        
        # Header
        header_frame = tk.Frame(self.root, bg="#667eea", height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="🍗 KFC Drive-Thru Automation",
            font=("Arial", 20, "bold"),
            bg="#667eea",
            fg="white"
        )
        title_label.pack(pady=20)
        
        # Main content
        main_frame = tk.Frame(self.root, padx=30, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Info section
        info_frame = tk.LabelFrame(main_frame, text="Information", font=("Arial", 10, "bold"))
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        info_text = """
This application automates the KFC Drive-Thru data processing workflow:
• Downloads reports from HMECloud
• Transforms data using DT macro
• Updates Drive-Thru Optimization template
• Refreshes pivot tables and updates dates
        """
        info_label = tk.Label(
            info_frame,
            text=info_text.strip(),
            font=("Arial", 9),
            justify=tk.LEFT,
            padx=10,
            pady=10
        )
        info_label.pack()
        
        # Options section
        options_frame = tk.LabelFrame(main_frame, text="Options", font=("Arial", 10, "bold"))
        options_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Store selection
        store_frame = tk.Frame(options_frame)
        store_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(store_frame, text="Store Selection:", font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 10))
        self.store_var = tk.StringVar(value="All Stores")
        store_combo = ttk.Combobox(
            store_frame,
            textvariable=self.store_var,
            values=["All Stores", "Single Store"],
            state="readonly",
            width=20
        )
        store_combo.pack(side=tk.LEFT)
        
        # Date selection
        date_frame = tk.Frame(options_frame)
        date_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(date_frame, text="Report Date:", font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 10))
        self.date_var = tk.StringVar(value="Yesterday")
        date_combo = ttk.Combobox(
            date_frame,
            textvariable=self.date_var,
            values=["Yesterday", "Custom Date"],
            state="readonly",
            width=20
        )
        date_combo.pack(side=tk.LEFT)
        
        # Start button
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        self.start_button = tk.Button(
            button_frame,
            text="🚀 Start Automation",
            font=("Arial", 14, "bold"),
            bg="#667eea",
            fg="white",
            padx=30,
            pady=15,
            cursor="hand2",
            command=self.start_automation,
            relief=tk.FLAT,
            bd=0
        )
        self.start_button.pack()
        
        # Progress bar
        self.progress = ttk.Progressbar(
            main_frame,
            mode='indeterminate',
            length=400
        )
        self.progress.pack(pady=10)
        
        # Status label
        self.status_label = tk.Label(
            main_frame,
            text="Ready to start automation",
            font=("Arial", 10),
            fg="gray"
        )
        self.status_label.pack(pady=5)
        
        # Log output
        log_frame = tk.LabelFrame(main_frame, text="Log Output", font=("Arial", 10, "bold"))
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=8,
            font=("Consolas", 8),
            wrap=tk.WORD,
            bg="#f5f5f5"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Footer
        footer_frame = tk.Frame(self.root, bg="#f0f0f0", height=30)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        footer_label = tk.Label(
            footer_frame,
            text="KFC Guyana - Drive-Thru Optimization Platform v1.0",
            font=("Arial", 8),
            bg="#f0f0f0",
            fg="gray"
        )
        footer_label.pack(pady=5)
    
    def log_message(self, message):
        """Add message to log output"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def start_automation(self):
        """Start the automation process"""
        if self.automation_running:
            messagebox.showwarning("Warning", "Automation is already running!")
            return
        
        # Confirm before starting
        response = messagebox.askyesno(
            "Confirm",
            "This will start the automation process.\n\n"
            "A browser window will open for HMECloud login.\n"
            "The process may take several minutes.\n\n"
            "Do you want to continue?",
            icon="question"
        )
        
        if not response:
            return
        
        # Clear log
        self.log_text.delete(1.0, tk.END)
        
        # Update UI
        self.automation_running = True
        self.start_button.config(state=tk.DISABLED, text="⏳ Running...")
        self.progress.start()
        self.status_label.config(text="Automation in progress...", fg="blue")
        
        # Start automation in separate thread
        thread = threading.Thread(target=self.run_automation, daemon=True)
        thread.start()
    
    def run_automation(self):
        """Run the automation script"""
        try:
            self.log_message("=" * 60)
            self.log_message("Starting KFC Drive-Thru Automation")
            self.log_message("=" * 60)
            self.log_message("")
            
            # Determine Python executable and script path
            if getattr(sys, 'frozen', False):
                # Running as executable
                python_exe = sys.executable
                # Try to find script in various locations
                script_paths = [
                    BASE_DIR / "scripts" / "test_store_selection.py",
                    BASE_DIR / "test_store_selection.py",
                    Path(__file__).parent / "scripts" / "test_store_selection.py",
                ]
                script_path = None
                for path in script_paths:
                    if path.exists():
                        script_path = path
                        break
                
                if not script_path:
                    self.log_message("⚠️ Script not found, using direct import...")
                    self.run_automation_direct()
                    return
            else:
                # Running as script
                python_exe = sys.executable
                script_path = SCRIPTS_DIR / "test_store_selection.py"
            
            self.log_message(f"Python: {python_exe}")
            self.log_message(f"Script: {script_path}")
            self.log_message("")
            
            # Run the script
            self.process = subprocess.Popen(
                [python_exe, str(script_path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True,
                cwd=str(BASE_DIR)
            )
            
            # Read output in real-time
            for line in self.process.stdout:
                if line:
                    self.log_message(line.strip())
                    self.root.update_idletasks()
            
            # Wait for process to complete
            return_code = self.process.wait()
            
            if return_code == 0:
                self.on_automation_complete()
            else:
                self.on_automation_error(f"Process exited with code {return_code}")
                
        except Exception as e:
            self.on_automation_error(f"Error: {str(e)}")
            import traceback
            self.log_message(traceback.format_exc())
    
    def run_automation_direct(self):
        """Run automation by directly importing and calling the function"""
        try:
            self.log_message("Running automation via direct import...")
            
            # Import the automation module
            from automation.hmecloud import download_all_stores
            from datetime import datetime, timedelta
            
            # Determine date
            if self.date_var.get() == "Yesterday":
                report_date = datetime.now() - timedelta(days=1)
            else:
                # For custom date, use yesterday for now
                report_date = datetime.now() - timedelta(days=1)
            
            self.log_message(f"Report date: {report_date.strftime('%Y-%m-%d')}")
            self.log_message("")
            
            # Run download
            if self.store_var.get() == "All Stores":
                self.log_message("Downloading all stores...")
                success = download_all_stores(report_date=report_date)
            else:
                self.log_message("Single store mode - selecting first store...")
                from automation.hmecloud import download_single_store, STORES
                success = download_single_store(STORES[0], report_date=report_date)
            
            if success:
                self.on_automation_complete()
            else:
                self.on_automation_error("Download completed with errors")
                
        except Exception as e:
            self.on_automation_error(f"Error: {str(e)}")
            import traceback
            self.log_message(traceback.format_exc())
    
    def on_automation_complete(self):
        """Handle successful completion"""
        self.root.after(0, lambda: self._complete_ui_update(True))
    
    def on_automation_error(self, error_msg):
        """Handle error"""
        self.root.after(0, lambda: self._complete_ui_update(False, error_msg))
    
    def _complete_ui_update(self, success, error_msg=None):
        """Update UI after automation completes"""
        self.automation_running = False
        self.progress.stop()
        self.start_button.config(state=tk.NORMAL, text="🚀 Start Automation")
        
        if success:
            self.status_label.config(text="✅ Automation completed successfully!", fg="green")
            self.log_message("")
            self.log_message("=" * 60)
            self.log_message("✅ AUTOMATION COMPLETED SUCCESSFULLY!")
            self.log_message("=" * 60)
            messagebox.showinfo(
                "Success",
                "Automation completed successfully!\n\n"
                "Please check the Drive-Thru template file for updated data."
            )
        else:
            self.status_label.config(text=f"❌ Error: {error_msg}", fg="red")
            self.log_message("")
            self.log_message("=" * 60)
            self.log_message(f"❌ AUTOMATION FAILED: {error_msg}")
            self.log_message("=" * 60)
            messagebox.showerror("Error", f"Automation failed:\n\n{error_msg}")
    
    def on_closing(self):
        """Handle window close"""
        if self.automation_running:
            response = messagebox.askyesno(
                "Automation Running",
                "Automation is currently running.\n\n"
                "Do you want to stop it and exit?",
                icon="warning"
            )
            if response:
                if self.process:
                    self.process.terminate()
                self.root.destroy()
        else:
            self.root.destroy()


def main():
    """Main entry point"""
    root = tk.Tk()
    app = AutomationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

