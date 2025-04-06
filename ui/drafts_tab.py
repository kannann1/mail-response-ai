# ui/drafts_tab.py
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

class DraftsTab:
    """Drafts tab UI for the email agent application"""
    
    def __init__(self, parent, outlook_service, storage_service):
        self.parent = parent
        self.outlook_service = outlook_service
        self.storage_service = storage_service
        
        # Current data
        self.drafts = []
        self.current_draft = None
        
        self._setup_ui()
        self._load_drafts()
    
    def _setup_ui(self):
        """Set up the drafts UI"""
        # Create main paned window
        self.paned_window = ttk.PanedWindow(self.parent, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create drafts list frame
        drafts_list_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(drafts_list_frame, weight=40)
        
        # Controls for drafts list
        controls_frame = ttk.Frame(drafts_list_frame)
        controls_frame.pack(fill=tk.X, padx=5, pady=5)
        
        refresh_button = ttk.Button(controls_frame, text="Refresh", command=self._load_drafts)
        refresh_button.pack(side=tk.LEFT, padx=5)
        
        clear_button = ttk.Button(controls_frame, text="Clear All", command=self._clear_all_drafts)
        clear_button.pack(side=tk.RIGHT, padx=5)
        
        # Drafts list
        drafts_frame = ttk.Frame(drafts_list_frame)
        drafts_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create listbox with scrollbar
        self.drafts_listbox = tk.Listbox(drafts_frame, width=40, height=20)
        self.drafts_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(drafts_frame, orient=tk.VERTICAL, command=self.drafts_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.drafts_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Bind selection event
        self.drafts_listbox.bind('<<ListboxSelect>>', self._on_draft_selected)
        
        # Create draft view frame
        draft_view_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(draft_view_frame, weight=60)
        
        # Original email section
        original_frame = ttk.LabelFrame(draft_view_frame, text="Original Email")
        original_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # From field
        from_frame = ttk.Frame(original_frame)
        from_frame.pack(fill=tk.X, padx=5, pady=2)
        
        from_label = ttk.Label(from_frame, text="From:", width=10)
        from_label.pack(side=tk.LEFT)
        
        self.from_value = ttk.Label(from_frame, text="")
        self.from_value.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Subject field
        subject_frame = ttk.Frame(original_frame)
        subject_frame.pack(fill=tk.X, padx=5, pady=2)
        
        subject_label = ttk.Label(subject_frame, text="Subject:", width=10)
        subject_label.pack(side=tk.LEFT)
        
        self.subject_value = ttk.Label(subject_frame, text="")
        self.subject_value.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Original message snippet
        original_text_frame = ttk.Frame(original_frame)
        original_text_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.original_text = tk.Text(original_text_frame, wrap=tk.WORD, height=5)
        self.original_text.pack(fill=tk.X, expand=True)
        
        # Response preview section
        response_frame = ttk.LabelFrame(draft_view_frame, text="Generated Response")
        response_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Response editor
        self.response_text = tk.Text(response_frame, wrap=tk.WORD)
        self.response_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        response_scrollbar = ttk.Scrollbar(self.response_text, orient=tk.VERTICAL, command=self.response_text.yview)
        response_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.response_text.configure(yscrollcommand=response_scrollbar.set)
        
        # Actions frame
        actions_frame = ttk.Frame(draft_view_frame)
        actions_frame.pack(fill=tk.X, padx=10, pady=10)
        
        send_button = ttk.Button(actions_frame, text="Send Email", command=self._send_email)
        send_button.pack(side=tk.LEFT, padx=5)
        
        edit_button = ttk.Button(actions_frame, text="Edit", command=self._toggle_edit_mode)
        edit_button.pack(side=tk.LEFT, padx=5)
        
        delete_button = ttk.Button(actions_frame, text="Delete Draft", command=self._delete_draft)
        delete_button.pack(side=tk.LEFT, padx=5)
        
        # Initialize with read-only response
        self.response_text.config(state=tk.DISABLED)
    
    def _load_drafts(self):
        """Load drafts from storage"""
        try:
            # Show busy cursor
            self._set_busy_cursor(True)
            
            # Clear existing items
            self.drafts_listbox.delete(0, tk.END)
            
            # Get drafts from storage
            self.drafts = self.storage_service.get_drafts()
            
            # Add to listbox
            for i, draft in enumerate(self.drafts):
                # Get subject from original email
                subject = draft.get('original_email', {}).get('subject', 'No subject')
                
                # Format subject for display
                if len(subject) > 40:
                    subject = subject[:37] + "..."
                
                # Add to listbox
                self.drafts_listbox.insert(tk.END, subject)
            
            # Clear current selection
            self._clear_draft_display()
            
        except Exception as e:
            logger.error(f"Error loading drafts: {str(e)}")
            messagebox.showerror("Load Drafts", f"Failed to load drafts: {str(e)}")
            
        finally:
            # Reset cursor
            self._set_busy_cursor(False)
    
    def _on_draft_selected(self, event):
        """Handle draft selection"""
        try:
            # Get selected index
            selection = self.drafts_listbox.curselection()
            if not selection:
                return
                
            index = selection[0]
            if index < 0 or index >= len(self.drafts):
                return
                
            # Get draft data
            self.current_draft = self.drafts[index]
            
            # Update display
            self._update_draft_display()
            
        except Exception as e:
            logger.error(f"Error handling draft selection: {str(e)}")
    
    def _update_draft_display(self):
        """Update the draft display"""
        try:
            if not self.current_draft:
                self._clear_draft_display()
                return
                
            # Get original email
            original_email = self.current_draft.get('original_email', {})
            
            # Update original email section
            self.from_value.config(text=original_email.get('from', ''))
            self.subject_value.config(text=original_email.get('subject', ''))
            
            # Update original text
            self.original_text.config(state=tk.NORMAL)
            self.original_text.delete("1.0", tk.END)
            
            body = original_email.get('body', '')
            if body:
                # Show just the first few lines
                lines = body.split('\n')
                preview = '\n'.join(lines[:10])
                if len(lines) > 10:
                    preview += '\n...'
                
                self.original_text.insert(tk.END, preview)
            
            self.original_text.config(state=tk.DISABLED)
            
            # Update response text
            self.response_text.config(state=tk.NORMAL)
            self.response_text.delete("1.0", tk.END)
            self.response_text.insert(tk.END, self.current_draft.get('response_text', ''))
            self.response_text.config(state=tk.DISABLED)
            
        except Exception as e:
            logger.error(f"Error updating draft display: {str(e)}")
    
    def _clear_draft_display(self):
        """Clear the draft display"""
        self.from_value.config(text="")
        self.subject_value.config(text="")
        
        self.original_text.config(state=tk.NORMAL)
        self.original_text.delete("1.0", tk.END)
        self.original_text.config(state=tk.DISABLED)
        
        self.response_text.config(state=tk.NORMAL)
        self.response_text.delete("1.0", tk.END)
        self.response_text.config(state=tk.DISABLED)
        
        self.current_draft = None
    
    def _toggle_edit_mode(self):
        """Toggle edit mode for response text"""
        if not self.current_draft:
            return
            
        # Toggle state
        if str(self.response_text.cget('state')) == 'disabled':
            self.response_text.config(state=tk.NORMAL)
        else:
            # Save changes
            updated_text = self.response_text.get("1.0", tk.END)
            
            # Update draft in memory
            self.current_draft['response_text'] = updated_text
            
            # Update in storage (would need additional function in storage service)
            
            # Set back to read-only
            self.response_text.config(state=tk.DISABLED)
    
    def _send_email(self):
        """Send the current draft as an email"""
        if not self.current_draft:
            messagebox.showinfo("Send Email", "Please select a draft first.")
            return
            
        try:
            # Get original email for reply information
            original_email = self.current_draft.get('original_email', {})
            if not original_email:
                messagebox.showerror("Send Email", "Original email information is missing.")
                return
                
            # Get response text
            response_text = self.response_text.get("1.0", tk.END)
            
            # Show busy cursor
            self._set_busy_cursor(True)
            
            # Confirm send
            if not messagebox.askyesno("Send Email", "Are you sure you want to send this email?"):
                self._set_busy_cursor(False)
                return
                
            # Get recipient
            to_email = original_email.get('from', '').split('<')[-1].split('>')[0]
            if not to_email:
                messagebox.showerror("Send Email", "Could not determine recipient email address.")
                self._set_busy_cursor(False)
                return
                
            # Create subject line
            subject = "Re: " + original_email.get('subject', '')
            
            # Send in a separate thread
            threading.Thread(
                target=self._send_email_thread,
                args=(to_email, subject, response_text),
                daemon=True
            ).start()
            
        except Exception as e:
            logger.error(f"Error preparing to send email: {str(e)}")
            messagebox.showerror("Send Email", f"Error: {str(e)}")
            self._set_busy_cursor(False)
    
    def _send_email_thread(self, to_email, subject, body):
        """Send email in a background thread"""
        try:
            # Send the email
            result = self.outlook_service.send_email(
                to_email=to_email,
                subject=subject,
                body=body
            )
            
            # Update UI in main thread
            if result:
                self.parent.after(0, lambda: self._handle_email_sent())
            else:
                self.parent.after(0, lambda: messagebox.showerror(
                    "Send Email",
                    "Failed to send the email. Please check your Outlook configuration."
                ))
                
        except Exception as e:
            logger.error(f"Error sending email: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror(
                "Send Email",
                f"Error sending email: {str(e)}"
            ))
            
        finally:
            # Reset cursor in main thread
            self.parent.after(0, lambda: self._set_busy_cursor(False))
    
    def _handle_email_sent(self):
        """Handle successful email sending"""
        try:
            messagebox.showinfo("Send Email", "Email sent successfully!")
            
            # Delete draft after sending
            if self.current_draft and 'id' in self.current_draft:
                self.storage_service.delete_draft(self.current_draft['id'])
                
            # Get selected index
            selection = self.drafts_listbox.curselection()
            if selection:
                index = selection[0]
                
                # Remove from listbox
                self.drafts_listbox.delete(index)
                
                # Remove from list
                if index < len(self.drafts):
                    self.drafts.pop(index)
                    
            # Clear display
            self._clear_draft_display()
            
        except Exception as e:
            logger.error(f"Error handling sent email: {str(e)}")
    
    def _delete_draft(self):
        """Delete the current draft"""
        if not self.current_draft:
            messagebox.showinfo("Delete Draft", "Please select a draft first.")
            return
            
        try:
            # Confirm deletion
            if not messagebox.askyesno("Delete Draft", "Are you sure you want to delete this draft?"):
                return
                
            # Delete from storage
            if 'id' in self.current_draft:
                result = self.storage_service.delete_draft(self.current_draft['id'])
                
                if result:
                    # Get selected index
                    selection = self.drafts_listbox.curselection()
                    if selection:
                        index = selection[0]
                        
                        # Remove from listbox
                        self.drafts_listbox.delete(index)
                        
                        # Remove from list
                        if index < len(self.drafts):
                            self.drafts.pop(index)
                            
                    # Clear display
                    self._clear_draft_display()
                    
                    messagebox.showinfo("Delete Draft", "Draft deleted successfully.")
                else:
                    messagebox.showerror("Delete Draft", "Failed to delete the draft.")
            
        except Exception as e:
            logger.error(f"Error deleting draft: {str(e)}")
            messagebox.showerror("Delete Draft", f"Error: {str(e)}")
    
    def _clear_all_drafts(self):
        """Clear all drafts"""
        try:
            # Confirm deletion
            if not messagebox.askyesno(
                "Clear All Drafts", 
                "Are you sure you want to delete ALL drafts? This cannot be undone."
            ):
                return
                
            # TODO: Add method to storage service to clear all drafts
            # For now, delete them one by one
            for draft in self.drafts:
                if 'id' in draft:
                    self.storage_service.delete_draft(draft['id'])
            
            # Clear listbox
            self.drafts_listbox.delete(0, tk.END)
            
            # Clear list
            self.drafts = []
            
            # Clear display
            self._clear_draft_display()
            
            messagebox.showinfo("Clear All Drafts", "All drafts have been deleted.")
            
        except Exception as e:
            logger.error(f"Error clearing all drafts: {str(e)}")
            messagebox.showerror("Clear All Drafts", f"Error: {str(e)}")
    
    def _set_busy_cursor(self, busy):
        """Set or clear busy cursor"""
        if busy:
            self.parent.config(cursor="wait")
        else:
            self.parent.config(cursor="")