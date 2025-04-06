# ui/main_window.py
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import logging

logger = logging.getLogger(__name__)

from ui.inbox_tab import InboxTab
from ui.drafts_tab import DraftsTab
from ui.tasks_tab import TasksTab
from ui.settings_tab import SettingsTab

from services.ollama_service import OllamaService

class EmailAgentUI:
    """Main UI for the email agent application"""
    
    def __init__(self, outlook_service, storage_service, email_processor, 
                priority_engine, response_generator, action_extractor, config):
        self.outlook_service = outlook_service
        self.storage_service = storage_service
        self.email_processor = email_processor
        self.priority_engine = priority_engine
        self.response_generator = response_generator
        self.action_extractor = action_extractor
        self.config = config
        
        # Initialize main window
        self.root = None
        self.tab_control = None
        self.tabs = {}
        
        # Email monitoring
        self.monitoring = False
    
    def run(self):
        """Initialize and run the UI"""
        self.root = tk.Tk()
        self.root.title("Kannan's personal Email Assistant")
        self.root.geometry("1200x800")
        self.root.minsize(800, 600)
        
        # Set up styles
        self._setup_styles()
        
        # Create tab control
        self.tab_control = ttk.Notebook(self.root)
        
        # Create tabs
        self._setup_tabs()
        
        # Pack the tab control
        self.tab_control.pack(expand=1, fill="both")
        
        # Set up status bar
        self._setup_status_bar()
        
        # Set up event handlers
        self._setup_events()
        
        # Initialize Outlook connection
        threading.Thread(target=self._init_outlook, daemon=True).start()
        
        # Start the main loop
        self.root.mainloop()
    
    def _setup_styles(self):
        """Set up TTK styles for the application"""
        style = ttk.Style()
        
        # Try to use a modern theme if available
        try:
            style.theme_use("clam")  # Most compatible enhanced theme
        except:
            pass  # Fall back to default theme
        
        # Configure styles
        style.configure("TButton", padding=6, relief="flat", background="#e1e1e1")
        style.configure("TNotebook", background="#f0f0f0")
        style.configure("TNotebook.Tab", padding=[12, 4], font=('', 10))
        style.configure("TFrame", background="#f0f0f0")
        
        # Configure hover style for buttons
        style.map("TButton",
                foreground=[('pressed', 'black'), ('active', 'black')],
                background=[('pressed', '#d0d0d0'), ('active', '#e9e9e9')])
    
    def _setup_tabs(self):
        """Set up the application tabs"""
        # Inbox tab
        inbox_frame = ttk.Frame(self.tab_control)
        self.tab_control.add(inbox_frame, text="Inbox")
        self.tabs['inbox'] = InboxTab(
            inbox_frame, 
            self.outlook_service,
            self.email_processor,
            self.priority_engine,
            self.response_generator,
            self.action_extractor,
            self.storage_service,
            self.config
        )
        
        # Drafts tab
        drafts_frame = ttk.Frame(self.tab_control)
        self.tab_control.add(drafts_frame, text="Draft Responses")
        self.tabs['drafts'] = DraftsTab(
            drafts_frame,
            self.outlook_service,
            self.storage_service
        )
        
        # Tasks tab
        tasks_frame = ttk.Frame(self.tab_control)
        self.tab_control.add(tasks_frame, text="Action Items")
        self.tabs['tasks'] = TasksTab(
            tasks_frame,
            self.storage_service,
            self.outlook_service
        )
        
        # Settings tab
        settings_frame = ttk.Frame(self.tab_control)
        self.tab_control.add(settings_frame, text="Settings")
        self.tabs['settings'] = SettingsTab(
            settings_frame,
            self.config,
            self.storage_service,
            self.outlook_service,
            self._on_settings_updated,
            OllamaService  # Pass the class itself
            )
    
    def _setup_status_bar(self):
        """Set up the status bar at the bottom of the window"""
        status_frame = ttk.Frame(self.root, relief=tk.SUNKEN, padding=(2, 2))
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Status message
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor=tk.W)
        status_label.pack(side=tk.LEFT, padx=5)
        
        # Connection status
        self.connection_var = tk.StringVar(value="Connecting...")
        connection_label = ttk.Label(status_frame, textvariable=self.connection_var, anchor=tk.E)
        connection_label.pack(side=tk.RIGHT, padx=5)
        
        # Email monitoring status
        self.monitoring_var = tk.StringVar(value="Monitoring: Off")
        monitoring_label = ttk.Label(status_frame, textvariable=self.monitoring_var, anchor=tk.E)
        monitoring_label.pack(side=tk.RIGHT, padx=5)
    
    def _setup_events(self):
        """Set up event handlers"""
        # Handle tab changes
        self.tab_control.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
        # Handle application close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
    
    def _init_outlook(self):
        """Initialize Outlook connection"""
        try:
            self._update_status("Connecting to Outlook...")
            result = self.outlook_service.initialize()
            
            if result:
                self._update_connection_status("Connected")
                self.root.after(0, lambda: self._start_email_monitoring())
                self.root.after(0, lambda: self._refresh_inbox_data())
            else:
                self._update_connection_status("Disconnected")
                self._update_status("Failed to connect to Outlook")
                self.root.after(0, lambda: messagebox.showerror(
                    "Outlook Connection Error",
                    "Failed to connect to Outlook. Please check that Outlook is installed and running."
                ))
                
        except Exception as e:
            logger.error(f"Error initializing Outlook: {str(e)}")
            self._update_connection_status("Error")
            self._update_status(f"Error connecting to Outlook: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror(
                "Outlook Connection Error",
                f"An error occurred while connecting to Outlook: {str(e)}"
            ))
    
    def _start_email_monitoring(self):
        """Start monitoring for new emails - simplified version"""
        if not self.monitoring:
            self._update_status("Starting email processing...")
            # Just call the method but don't rely on its background functionality
            result = self.outlook_service.start_monitoring(self._on_new_emails, self.root)
            if result:
                self.monitoring = True
                self._update_monitoring_status("Manual")
                self._update_status("Email processing initialized")
                # Do an immediate refresh
                self._refresh_inbox_data()
    
    def _stop_email_monitoring(self):
        """Stop monitoring for new emails"""
        if self.monitoring:
            self._update_status("Stopping email monitoring...")
            result = self.outlook_service.stop_monitoring()
            if result:
                self.monitoring = False
                self._update_monitoring_status("Off")
                self._update_status("Email monitoring stopped")
                logger.info("Stopped email monitoring")
    
    def _on_new_emails(self, emails):
        """Handle new emails"""
        try:
            # Update status
            self._update_status(f"Received {len(emails)} new email(s)")
            
            # Refresh inbox tab if active
            if self.tab_control.index(self.tab_control.select()) == 0:
                self.tabs['inbox']._refresh_emails()
            
            # Show notification if enabled
            if self.config.get("notifications", {}).get("new_emails", True):
                self._show_new_email_notification(emails)
                
        except Exception as e:
            logger.error(f"Error handling new emails: {str(e)}")
    
    def _show_new_email_notification(self, emails):
        """Show notification for new emails"""
        # This would normally use a system notification
        # For now, just show a messagebox if there are high priority emails
        try:
            high_priority_count = 0
            for email in emails:
                # Calculate priority
                priority_result = self.priority_engine.prioritize_email(email)
                category = priority_result.get('category', 'Medium')
                
                if category in ['High', 'Urgent']:
                    high_priority_count += 1
            
            if high_priority_count > 0:
                self.root.after(0, lambda: messagebox.showinfo(
                    "New High Priority Emails",
                    f"You have {high_priority_count} new high priority email(s)."
                ))
                
        except Exception as e:
            logger.error(f"Error showing notification: {str(e)}")
    
    def _refresh_inbox_data(self):
        """Refresh inbox data if inbox tab is active"""
        if self.tab_control.index(self.tab_control.select()) == 0:
            self.tabs['inbox']._refresh_emails()
    
    def _on_tab_changed(self, event):
        """Handle tab changed event"""
        try:
            # Get current tab index
            current_tab = self.tab_control.index(self.tab_control.select())
            
            # Update status
            tab_names = ["Inbox", "Draft Responses", "Action Items", "Settings"]
            self._update_status(f"Viewing {tab_names[current_tab]}")
            
            # Refresh data for the selected tab if needed
            if current_tab == 0:  # Inbox tab
                pass  # Already refreshed on startup
            elif current_tab == 1:  # Drafts tab
                self.tabs['drafts']._load_drafts()
            elif current_tab == 2:  # Tasks tab
                self.tabs['tasks']._load_tasks()
                
        except Exception as e:
            logger.error(f"Error handling tab change: {str(e)}")
    
    def _on_settings_updated(self, config):
        """Handle settings updated event"""
        try:
            # Update local config
            self.config = config
            
            # Update components that need the config
            # ...
            
            # Update status
            self._update_status("Settings updated")
            
        except Exception as e:
            logger.error(f"Error handling settings update: {str(e)}")
    
    def _update_status(self, message):
        """Update the status bar message"""
        self.status_var.set(message)
        
    def _update_connection_status(self, status):
        """Update the connection status"""
        self.connection_var.set(f"Outlook: {status}")
        
    def _update_monitoring_status(self, status):
        """Update the monitoring status"""
        self.monitoring_var.set(f"Monitoring: {status}")
    
    def _on_close(self):
        """Handle application close event"""
        try:
            # Stop email monitoring
            if self.monitoring:
                self._stop_email_monitoring()
                
            # Close the window
            self.root.destroy()
            
        except Exception as e:
            logger.error(f"Error closing application: {str(e)}")
            self.root.destroy()