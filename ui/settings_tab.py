import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import logging
import json

from services.ollama_service import OllamaService

logger = logging.getLogger(__name__)

class SettingsTab:
    """Settings tab UI for the email agent application"""
    
    def __init__(self, parent, config, storage_service, outlook_service, callback=None, ollama_service_class=None):
        self.parent = parent
        self.config = config
        self.storage_service = storage_service
        self.outlook_service = outlook_service
        self.callback = callback
        self.OllamaService = ollama_service_class 
        
        self._setup_ui()
        self._load_config_values()
    
    def _setup_ui(self):
        """Set up the settings UI"""
        main_frame = ttk.Frame(self.parent, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook for settings categories
        self.settings_notebook = ttk.Notebook(main_frame)
        self.settings_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # User settings tab
        user_frame = ttk.Frame(self.settings_notebook, padding="10")
        self.settings_notebook.add(user_frame, text="User Profile")
        self._setup_user_settings(user_frame)
        
        # Email settings tab
        email_frame = ttk.Frame(self.settings_notebook, padding="10")
        self.settings_notebook.add(email_frame, text="Email Settings")
        self._setup_email_settings(email_frame)
        
        # Ollama settings tab
        ollama_frame = ttk.Frame(self.settings_notebook, padding="10")
        self.settings_notebook.add(ollama_frame, text="Ollama Settings")
        self._setup_ollama_settings(ollama_frame)
        
        # Style samples tab
        style_frame = ttk.Frame(self.settings_notebook, padding="10")
        self.settings_notebook.add(style_frame, text="Writing Style")
        self._setup_style_settings(style_frame)
        
        # Save button
        save_frame = ttk.Frame(main_frame)
        save_frame.pack(fill=tk.X, pady=10)
        
        save_button = ttk.Button(save_frame, text="Save Settings", command=self._save_settings)
        save_button.pack(side=tk.RIGHT, padx=5)
        
        reset_button = ttk.Button(save_frame, text="Reset to Defaults", command=self._reset_settings)
        reset_button.pack(side=tk.RIGHT, padx=5)
    
    def _setup_user_settings(self, parent):
        """Set up user profile settings section"""
        # Full name
        name_frame = ttk.Frame(parent)
        name_frame.pack(fill=tk.X, pady=5)
        
        name_label = ttk.Label(name_frame, text="Full Name:", width=20)
        name_label.pack(side=tk.LEFT)
        
        self.name_entry = ttk.Entry(name_frame, width=40)
        self.name_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Email address
        email_frame = ttk.Frame(parent)
        email_frame.pack(fill=tk.X, pady=5)
        
        email_label = ttk.Label(email_frame, text="Email Address:", width=20)
        email_label.pack(side=tk.LEFT)
        
        self.email_entry = ttk.Entry(email_frame, width=40)
        self.email_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Professional role
        role_frame = ttk.Frame(parent)
        role_frame.pack(fill=tk.X, pady=5)
        
        role_label = ttk.Label(role_frame, text="Professional Role:", width=20)
        role_label.pack(side=tk.LEFT)
        
        self.role_entry = ttk.Entry(role_frame, width=40)
        self.role_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Communication style
        style_frame = ttk.Frame(parent)
        style_frame.pack(fill=tk.X, pady=5)
        
        style_label = ttk.Label(style_frame, text="Communication Style:", width=20)
        style_label.pack(side=tk.LEFT)
        
        self.style_var = tk.StringVar()
        style_combo = ttk.Combobox(style_frame, textvariable=self.style_var, width=20)
        style_combo['values'] = ('Professional', 'Casual', 'Formal', 'Technical', 'Friendly')
        style_combo.pack(side=tk.LEFT, padx=5)
        
        style_description = ttk.Label(
            parent, 
            text="This affects the tone and style of AI-generated responses.",
            font=("", 9, "italic")
        )
        style_description.pack(anchor=tk.W, pady=(0, 10))
    
    def _setup_email_settings(self, parent):
        """Set up email settings section"""
        # Refresh interval
        refresh_frame = ttk.Frame(parent)
        refresh_frame.pack(fill=tk.X, pady=5)
        
        refresh_label = ttk.Label(refresh_frame, text="Refresh Interval (sec):", width=20)
        refresh_label.pack(side=tk.LEFT)
        
        self.refresh_var = tk.IntVar()
        refresh_spinbox = ttk.Spinbox(refresh_frame, from_=30, to=3600, textvariable=self.refresh_var, width=10)
        refresh_spinbox.pack(side=tk.LEFT, padx=5)
        
        # Auto-archive option
        archive_frame = ttk.Frame(parent)
        archive_frame.pack(fill=tk.X, pady=5)
        
        self.auto_archive_var = tk.BooleanVar()
        archive_check = ttk.Checkbutton(
            archive_frame, 
            text="Auto-archive emails after response is sent", 
            variable=self.auto_archive_var
        )
        archive_check.pack(anchor=tk.W)
        
        # Important contacts
        important_frame = ttk.LabelFrame(parent, text="Important Contacts")
        important_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        contacts_frame = ttk.Frame(important_frame)
        contacts_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Contact lists with scrollbars
        contacts_list_frame = ttk.Frame(contacts_frame)
        contacts_list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.important_contacts_list = tk.Listbox(contacts_list_frame, height=6)
        self.important_contacts_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        contacts_scrollbar = ttk.Scrollbar(contacts_list_frame, orient=tk.VERTICAL, command=self.important_contacts_list.yview)
        contacts_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.important_contacts_list.configure(yscrollcommand=contacts_scrollbar.set)
        
        # Contact buttons
        contacts_button_frame = ttk.Frame(contacts_frame)
        contacts_button_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5)
        
        add_contact_button = ttk.Button(contacts_button_frame, text="Add", command=self._add_important_contact)
        add_contact_button.pack(fill=tk.X, pady=2)
        
        remove_contact_button = ttk.Button(contacts_button_frame, text="Remove", command=self._remove_important_contact)
        remove_contact_button.pack(fill=tk.X, pady=2)
        
        # VIP contacts
        vip_frame = ttk.LabelFrame(parent, text="VIP Contacts (Always Review)")
        vip_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        vip_contacts_frame = ttk.Frame(vip_frame)
        vip_contacts_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # VIP lists with scrollbars
        vip_list_frame = ttk.Frame(vip_contacts_frame)
        vip_list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.vip_contacts_list = tk.Listbox(vip_list_frame, height=6)
        self.vip_contacts_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        vip_scrollbar = ttk.Scrollbar(vip_list_frame, orient=tk.VERTICAL, command=self.vip_contacts_list.yview)
        vip_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.vip_contacts_list.configure(yscrollcommand=vip_scrollbar.set)
        
        # VIP buttons
        vip_button_frame = ttk.Frame(vip_contacts_frame)
        vip_button_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5)
        
        add_vip_button = ttk.Button(vip_button_frame, text="Add", command=self._add_vip_contact)
        add_vip_button.pack(fill=tk.X, pady=2)
        
        remove_vip_button = ttk.Button(vip_button_frame, text="Remove", command=self._remove_vip_contact)
        remove_vip_button.pack(fill=tk.X, pady=2)
    
    def _setup_ollama_settings(self, parent):
        """Set up Ollama settings section"""
        # Ollama host
        host_frame = ttk.Frame(parent)
        host_frame.pack(fill=tk.X, pady=5)
        
        host_label = ttk.Label(host_frame, text="Ollama Host:", width=20)
        host_label.pack(side=tk.LEFT)
        
        self.host_entry = ttk.Entry(host_frame, width=40)
        self.host_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Model selection
        model_frame = ttk.Frame(parent)
        model_frame.pack(fill=tk.X, pady=5)
        
        model_label = ttk.Label(model_frame, text="Model:", width=20)
        model_label.pack(side=tk.LEFT)
        
        self.model_var = tk.StringVar()
        self.model_combo = ttk.Combobox(model_frame, textvariable=self.model_var, width=30)
        self.model_combo.pack(side=tk.LEFT, padx=5)
        
        refresh_models_button = ttk.Button(model_frame, text="Refresh Models", command=self._refresh_models)
        refresh_models_button.pack(side=tk.LEFT, padx=5)
        
        # Always review option
        review_frame = ttk.Frame(parent)
        review_frame.pack(fill=tk.X, pady=5)
        
        self.always_review_var = tk.BooleanVar()
        review_check = ttk.Checkbutton(
            review_frame, 
            text="Always review responses before sending", 
            variable=self.always_review_var
        )
        review_check.pack(anchor=tk.W)
        
        # System prompt
        prompt_frame = ttk.LabelFrame(parent, text="Custom System Prompt (Optional)")
        prompt_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.system_prompt_text = tk.Text(prompt_frame, wrap=tk.WORD, height=8)
        self.system_prompt_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        prompt_scrollbar = ttk.Scrollbar(self.system_prompt_text, orient=tk.VERTICAL, command=self.system_prompt_text.yview)
        prompt_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.system_prompt_text.configure(yscrollcommand=prompt_scrollbar.set)
        
        prompt_description = ttk.Label(
            parent, 
            text="Leave blank to use the default system prompt based on your profile settings.",
            font=("", 9, "italic")
        )
        prompt_description.pack(anchor=tk.W, pady=(0, 10))
        
        # Model info section
        model_info_frame = ttk.LabelFrame(parent, text="Model Information")
        model_info_frame.pack(fill=tk.X, pady=10)
        
        self.model_info_text = tk.Text(model_info_frame, wrap=tk.WORD, height=4, state=tk.DISABLED)
        self.model_info_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Test connection button
        test_button = ttk.Button(parent, text="Test Ollama Connection", command=self._test_ollama_connection)
        test_button.pack(anchor=tk.E, pady=5)
        
        # Initial refresh of models
        self._refresh_models()
    
    def _setup_style_settings(self, parent):
        """Set up writing style settings section"""
        description_label = ttk.Label(
            parent, 
            text="Add examples of your writing style to help the AI match your tone and style."
        )
        description_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Writing samples list
        samples_frame = ttk.Frame(parent)
        samples_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Samples list with scrollbar
        self.samples_list = tk.Listbox(samples_frame, height=6)
        self.samples_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        samples_scrollbar = ttk.Scrollbar(samples_frame, orient=tk.VERTICAL, command=self.samples_list.yview)
        samples_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.samples_list.configure(yscrollcommand=samples_scrollbar.set)
        
        # Sample management buttons
        samples_button_frame = ttk.Frame(parent)
        samples_button_frame.pack(fill=tk.X, pady=5)
        
        add_sample_button = ttk.Button(samples_button_frame, text="Add Sample", command=self._add_style_sample)
        add_sample_button.pack(side=tk.LEFT, padx=5)
        
        edit_sample_button = ttk.Button(samples_button_frame, text="Edit Sample", command=self._edit_style_sample)
        edit_sample_button.pack(side=tk.LEFT, padx=5)
        
        remove_sample_button = ttk.Button(samples_button_frame, text="Remove Sample", command=self._remove_style_sample)
        remove_sample_button.pack(side=tk.LEFT, padx=5)
        
        import_button = ttk.Button(samples_button_frame, text="Import from Email", command=self._import_style_from_email)
        import_button.pack(side=tk.RIGHT, padx=5)
        
        # Sample preview
        preview_frame = ttk.LabelFrame(parent, text="Sample Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.preview_text = tk.Text(preview_frame, wrap=tk.WORD, height=8)
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_text.configure(yscrollcommand=preview_scrollbar.set)
        
        # Bind list selection to preview
        self.samples_list.bind('<<ListboxSelect>>', self._on_sample_selected)
    
    def _load_config_values(self):
        """Load current configuration values into UI elements"""
        # User settings
        self.name_entry.insert(0, self.config["user"].get("name", ""))
        self.email_entry.insert(0, self.config["user"].get("email", ""))
        self.role_entry.insert(0, self.config["user"].get("role", ""))
        self.style_var.set(self.config["user"].get("communication_style", "Professional").title())
        
        # Email settings
        self.refresh_var.set(self.config["email"].get("refresh_interval", 300))
        self.auto_archive_var.set(self.config["email"].get("auto_archive", False))
        
        # Important contacts
        for contact in self.config["email"].get("important_contacts", []):
            self.important_contacts_list.insert(tk.END, contact)
            
        # VIP contacts
        for contact in self.config["email"].get("vip_contacts", []):
            self.vip_contacts_list.insert(tk.END, contact)
        
        # Ollama settings
        self.host_entry.insert(0, self.config["ollama"].get("host", "http://localhost:11434"))
        self.model_var.set(self.config["ollama"].get("model", "llama3"))
        self.always_review_var.set(self.config["ollama"].get("always_review", True))
        
        if "system_prompt" in self.config["ollama"]:
            self.system_prompt_text.insert("1.0", self.config["ollama"]["system_prompt"])
        
        # Style samples
        samples = self.config["ollama"].get("style_samples", [])
        for i, sample in enumerate(samples):
            self.samples_list.insert(tk.END, f"Sample {i+1}")
    
    def _save_settings(self):
        """Save settings to configuration file"""
        try:
            # Update user configuration
            self.config["user"]["name"] = self.name_entry.get()
            self.config["user"]["email"] = self.email_entry.get()
            self.config["user"]["role"] = self.role_entry.get()
            self.config["user"]["communication_style"] = self.style_var.get().lower()
            
            # Update email configuration
            self.config["email"]["refresh_interval"] = self.refresh_var.get()
            self.config["email"]["auto_archive"] = self.auto_archive_var.get()
            
            # Update important contacts
            important_contacts = []
            for i in range(self.important_contacts_list.size()):
                important_contacts.append(self.important_contacts_list.get(i))
            self.config["email"]["important_contacts"] = important_contacts
            
            # Update VIP contacts
            vip_contacts = []
            for i in range(self.vip_contacts_list.size()):
                vip_contacts.append(self.vip_contacts_list.get(i))
            self.config["email"]["vip_contacts"] = vip_contacts
            
            # Update Ollama configuration
            self.config["ollama"]["host"] = self.host_entry.get()
            self.config["ollama"]["model"] = self.model_var.get()
            self.config["ollama"]["always_review"] = self.always_review_var.get()
            
            # Update system prompt
            system_prompt = self.system_prompt_text.get("1.0", tk.END).strip()
            if system_prompt:
                self.config["ollama"]["system_prompt"] = system_prompt
            else:
                # Remove system_prompt if empty
                self.config["ollama"].pop("system_prompt", None)
            
            # Update style samples
            # We need to extract the actual content from our preview text or storage
            samples = []
            for i in range(self.samples_list.size()):
                # This assumes we're storing the samples in storage and can retrieve them
                sample_id = f"sample_{i+1}"
                sample_text = self.storage_service.get_style_sample(sample_id)
                if sample_text:
                    samples.append(sample_text)
            
            self.config["ollama"]["style_samples"] = samples
            
            # Save configuration
            result = self.storage_service.save_config(self.config)
            
            if result:
                messagebox.showinfo("Settings", "Settings saved successfully.")
                
                # Notify callback if provided
                if self.callback:
                    self.callback(self.config)
            else:
                messagebox.showerror("Settings", "Failed to save settings.")
                
        except Exception as e:
            logger.error(f"Error saving settings: {str(e)}")
            messagebox.showerror("Settings", f"An error occurred while saving settings: {str(e)}")
    
    def _reset_settings(self):
        """Reset settings to defaults"""
        if messagebox.askyesno("Reset Settings", "Are you sure you want to reset all settings to defaults?"):
            # Create default configuration
            default_config = {
                "user": {
                    "name": "Your Name",
                    "email": "your.email@example.com",
                    "role": "DevOps Engineer",
                    "communication_style": "professional"
                },
                "email": {
                    "refresh_interval": 300,
                    "auto_archive": False,
                    "important_contacts": [],
                    "vip_contacts": []
                },
                "ollama": {
                    "host": "http://localhost:11434",
                    "model": "llama3",
                    "system_prompt": "",
                    "always_review": True,
                    "style_samples": []
                }
            }
            
            # Update config
            self.config.update(default_config)
            
            # Clear UI
            self._clear_ui()
            
            # Load default values
            self._load_config_values()
            
            messagebox.showinfo("Settings", "Settings reset to defaults.")
    
    def _clear_ui(self):
        """Clear all UI elements"""
        self.name_entry.delete(0, tk.END)
        self.email_entry.delete(0, tk.END)
        self.role_entry.delete(0, tk.END)
        
        self.important_contacts_list.delete(0, tk.END)
        self.vip_contacts_list.delete(0, tk.END)
        
        self.host_entry.delete(0, tk.END)
        self.system_prompt_text.delete("1.0", tk.END)
        
        self.samples_list.delete(0, tk.END)
        self.preview_text.delete("1.0", tk.END)
    
    def _add_important_contact(self):
        """Add an important contact"""
        contact = self._prompt_for_contact("Add Important Contact")
        if contact:
            self.important_contacts_list.insert(tk.END, contact)
    
    def _remove_important_contact(self):
        """Remove an important contact"""
        selected = self.important_contacts_list.curselection()
        if selected:
            self.important_contacts_list.delete(selected)
    
    def _add_vip_contact(self):
        """Add a VIP contact"""
        contact = self._prompt_for_contact("Add VIP Contact")
        if contact:
            self.vip_contacts_list.insert(tk.END, contact)
    
    def _remove_vip_contact(self):
        """Remove a VIP contact"""
        selected = self.vip_contacts_list.curselection()
        if selected:
            self.vip_contacts_list.delete(selected)
    
    def _prompt_for_contact(self, title):
        """Prompt for a contact email address"""
        # Create a dialog window
        dialog = tk.Toplevel(self.parent)
        dialog.title(title)
        dialog.geometry("400x120")
        dialog.resizable(False, False)
        dialog.transient(self.parent)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Create a frame for the content
        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Add a label and entry for the contact
        ttk.Label(frame, text="Email Address:").pack(anchor=tk.W, pady=(0, 5))
        entry = ttk.Entry(frame, width=40)
        entry.pack(fill=tk.X, pady=(0, 10))
        entry.focus_set()
        
        # Add buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X)
        
        # Variable to store the result
        result = [None]
        
        def on_ok():
            result[0] = entry.get().strip()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT)
        
        # Bind enter key to OK button
        entry.bind("<Return>", lambda event: on_ok())
        
        # Wait for the dialog to be closed
        dialog.wait_window()
        
        return result[0]
    
    def _refresh_models(self):
        """Refresh the list of available Ollama models"""
        try:
            # Show busy cursor
            self.parent.config(cursor="wait")
            self.parent.update()
            
            # Get models in a separate thread
            def get_models():
                try:
                    # Initialize temporary Ollama service with current host
                    host = self.host_entry.get()
                    if not host:
                        host = "http://localhost:11434"
                    
                    ollama_service = OllamaService(host, "")
                    
                    # Get models
                    models = ollama_service.list_available_models()
                    
                    # Update UI in main thread
                    self.parent.after(0, lambda: self._update_models_ui(models))
                    
                except Exception as e:
                    logger.error(f"Error getting models: {str(e)}")
                    # Update UI in main thread
                    self.parent.after(0, lambda: self._show_models_error(str(e)))
                
                # Reset cursor in main thread
                self.parent.after(0, lambda: self.parent.config(cursor=""))
            
            # Start thread
            threading.Thread(target=get_models, daemon=True).start()
            
        except Exception as e:
            logger.error(f"Error refreshing models: {str(e)}")
            self.parent.config(cursor="")
            messagebox.showerror("Refresh Models", f"Error refreshing models: {str(e)}")
    
    def _update_models_ui(self, models):
        """Update UI with list of models"""
        if models:
            # Update combobox values
            self.model_combo['values'] = models
            
            # Keep current selection if possible
            current_model = self.model_var.get()
            if current_model in models:
                self.model_var.set(current_model)
            else:
                # Default to first model
                self.model_var.set(models[0])
                
            messagebox.showinfo("Refresh Models", f"Found {len(models)} models.")
        else:
            messagebox.showwarning("Refresh Models", "No models found. Make sure Ollama is running and has models pulled.")
    
    def _show_models_error(self, error_message):
        """Show error when refreshing models fails"""
        messagebox.showerror("Refresh Models", f"Error refreshing models: {error_message}")
    
    def _test_ollama_connection(self):
        """Test connection to Ollama"""
        try:
            # Show busy cursor
            self.parent.config(cursor="wait")
            self.parent.update()
            
            # Test connection in a separate thread
            def test_connection():
                try:
                    # Get host and model from UI
                    host = self.host_entry.get()
                    if not host:
                        host = "http://localhost:11434"
                    
                    model = self.model_var.get()
                    if not model:
                        model = "llama2"
                    
                    # Initialize Ollama service
                    ollama_service = OllamaService(host, model)
                    
                    # Test ping
                    if ollama_service.ping():
                        # Try a simple completion to test model
                        response = ollama_service.generate_completion(
                            "Hello, this is a test. Please respond with one sentence.",
                            temperature=0.7
                        )
                        
                        # Update UI in main thread
                        self.parent.after(0, lambda: self._show_connection_success(response))
                    else:
                        # Update UI in main thread
                        self.parent.after(0, lambda: messagebox.showerror(
                            "Test Connection", 
                            "Could not connect to Ollama server."
                        ))
                    
                except Exception as e:
                    logger.error(f"Error testing connection: {str(e)}")
                    # Update UI in main thread
                    self.parent.after(0, lambda: messagebox.showerror(
                        "Test Connection", 
                        f"Error testing connection: {str(e)}"
                    ))
                
                # Reset cursor in main thread
                self.parent.after(0, lambda: self.parent.config(cursor=""))
            
            # Start thread
            threading.Thread(target=test_connection, daemon=True).start()
            
        except Exception as e:
            logger.error(f"Error testing connection: {str(e)}")
            self.parent.config(cursor="")
            messagebox.showerror("Test Connection", f"Error: {str(e)}")
    
    def _show_connection_success(self, response):
        """Show success message for connection test"""
        # Update model info text
        self.model_info_text.config(state=tk.NORMAL)
        self.model_info_text.delete("1.0", tk.END)
        self.model_info_text.insert("1.0", f"Connection successful. Model response:\n\n{response[:200]}...")
        self.model_info_text.config(state=tk.DISABLED)
        
        messagebox.showinfo("Test Connection", "Successfully connected to Ollama!")
    
    def _add_style_sample(self):
        """Add a writing style sample"""
        # Create a dialog window
        dialog = tk.Toplevel(self.parent)
        dialog.title("Add Writing Style Sample")
        dialog.geometry("600x400")
        dialog.transient(self.parent)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Create a frame for the content
        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Add instructions
        ttk.Label(
            frame, 
            text="Enter an example of your writing style. This could be an email you've written, a document, or any text that reflects how you communicate.",
            wraplength=580
        ).pack(fill=tk.X, pady=(0, 10))
        
        # Add text area for the sample
        sample_frame = ttk.Frame(frame)
        sample_frame.pack(fill=tk.BOTH, expand=True)
        
        sample_text = tk.Text(sample_frame, wrap=tk.WORD)
        sample_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        sample_scrollbar = ttk.Scrollbar(sample_frame, orient=tk.VERTICAL, command=sample_text.yview)
        sample_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        sample_text.configure(yscrollcommand=sample_scrollbar.set)
        
        # Add buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        def on_save():
            # Get sample text
            text = sample_text.get("1.0", tk.END).strip()
            if not text:
                messagebox.showerror("Add Sample", "Please enter a sample text.")
                return
            
            # Add to list and storage
            try:
                # Add to style samples in config
                sample_list_size = self.samples_list.size()
                sample_id = f"sample_{sample_list_size + 1}"
                
                # Save to storage
                self.storage_service.save_style_sample(sample_id, text)
                
                # Add to listbox
                self.samples_list.insert(tk.END, f"Sample {sample_list_size + 1}")
                
                dialog.destroy()
                
            except Exception as e:
                logger.error(f"Error saving style sample: {str(e)}")
                messagebox.showerror("Add Sample", f"Error saving sample: {str(e)}")
        
        ttk.Button(button_frame, text="Save", command=on_save).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT)
        
        # Wait for the dialog to be closed
        dialog.wait_window()
    
    def _edit_style_sample(self):
        """Edit a writing style sample"""
        selected = self.samples_list.curselection()
        if not selected:
            messagebox.showinfo("Edit Sample", "Please select a sample to edit.")
            return
        
        # Get sample index and text
        index = selected[0]
        sample_id = f"sample_{index + 1}"
        sample_text = self.storage_service.get_style_sample(sample_id)
        
        if not sample_text:
            messagebox.showerror("Edit Sample", "Could not retrieve sample text.")
            return
        
        # Create a dialog window (similar to add sample)
        dialog = tk.Toplevel(self.parent)
        dialog.title("Edit Writing Style Sample")
        dialog.geometry("600x400")
        dialog.transient(self.parent)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Create a frame for the content
        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Add text area with existing content
        sample_frame = ttk.Frame(frame)
        sample_frame.pack(fill=tk.BOTH, expand=True)
        
        sample_editor = tk.Text(sample_frame, wrap=tk.WORD)
        sample_editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sample_editor.insert("1.0", sample_text)
        
        sample_scrollbar = ttk.Scrollbar(sample_frame, orient=tk.VERTICAL, command=sample_editor.yview)
        sample_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        sample_editor.configure(yscrollcommand=sample_scrollbar.set)
        
        # Add buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        def on_save():
            # Get updated text
            text = sample_editor.get("1.0", tk.END).strip()
            if not text:
                messagebox.showerror("Edit Sample", "Sample text cannot be empty.")
                return
            
            # Update storage
            try:
                self.storage_service.save_style_sample(sample_id, text)
                dialog.destroy()
                
                # Update preview if this was the selected sample
                if self.samples_list.curselection() and self.samples_list.curselection()[0] == index:
                    self._on_sample_selected(None)
                    
            except Exception as e:
                logger.error(f"Error updating style sample: {str(e)}")
                messagebox.showerror("Edit Sample", f"Error updating sample: {str(e)}")
        
        ttk.Button(button_frame, text="Save", command=on_save).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT)
        
        # Wait for the dialog to be closed
        dialog.wait_window()
    
    def _remove_style_sample(self):
        """Remove a writing style sample"""
        selected = self.samples_list.curselection()
        if not selected:
            messagebox.showinfo("Remove Sample", "Please select a sample to remove.")
            return
        
        # Confirm deletion
        if not messagebox.askyesno("Remove Sample", "Are you sure you want to remove this writing sample?"):
            return
        
        # Get sample index
        index = selected[0]
        sample_id = f"sample_{index + 1}"
        
        # Remove from storage
        try:
            self.storage_service.delete_style_sample(sample_id)
            
            # Remove from listbox
            self.samples_list.delete(index)
            
            # Clear preview
            self.preview_text.delete("1.0", tk.END)
            
            # Renumber remaining samples in storage
            self._renumber_samples()
            
        except Exception as e:
            logger.error(f"Error removing style sample: {str(e)}")
            messagebox.showerror("Remove Sample", f"Error removing sample: {str(e)}")
    
    def _renumber_samples(self):
        """Renumber samples after deletion"""
        try:
            sample_count = self.samples_list.size()
            
            # Update listbox items
            for i in range(sample_count):
                self.samples_list.delete(i)
                self.samples_list.insert(i, f"Sample {i+1}")
                
            # Update storage
            samples = []
            for i in range(sample_count):
                old_id = f"sample_{i+2}"  # +2 because we just deleted one
                new_id = f"sample_{i+1}"
                
                # Get old sample text
                text = self.storage_service.get_style_sample(old_id)
                if text:
                    # Save with new ID
                    self.storage_service.save_style_sample(new_id, text)
                    samples.append(text)
                    
                    # Delete old ID
                    self.storage_service.delete_style_sample(old_id)
            
            # Update config
            self.config["ollama"]["style_samples"] = samples
            
        except Exception as e:
            logger.error(f"Error renumbering samples: {str(e)}")
    
    def _on_sample_selected(self, event):
        """Handle sample selection"""
        selected = self.samples_list.curselection()
        if not selected:
            return
        
        # Get sample index
        index = selected[0]
        sample_id = f"sample_{index + 1}"
        
        # Get sample text
        sample_text = self.storage_service.get_style_sample(sample_id)
        
        # Update preview
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        
        if sample_text:
            self.preview_text.insert("1.0", sample_text)
        else:
            self.preview_text.insert("1.0", "Sample text not found.")
        
        self.preview_text.config(state=tk.NORMAL)
    
    def _import_style_from_email(self):
        """Import writing style from sent emails"""
        # This would connect to Outlook to get sent emails
        # For now, just show a dialog for selecting an email
        messagebox.showinfo(
            "Import from Email", 
            "This feature will allow you to import your writing style from sent emails.\n\n"
            "Implementation in progress."
        )
