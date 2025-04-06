# ui/inbox_tab.py
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

class InboxTab:
    """Inbox tab UI for the email agent application"""
    
    def __init__(self, parent, outlook_service, email_processor, priority_engine, 
                response_generator, action_extractor, storage_service, config):
        self.parent = parent
        self.outlook_service = outlook_service
        self.email_processor = email_processor
        self.priority_engine = priority_engine
        self.response_generator = response_generator
        self.action_extractor = action_extractor
        self.storage_service = storage_service
        self.config = config
        
        # Current data
        self.emails = []
        self.current_email = None
        self.current_thread = []
        
        self._setup_ui()
    
    def _setup_ui(self):
        """Set up the inbox UI"""
        # Create main paned window
        self.paned_window = ttk.PanedWindow(self.parent, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create email list frame
        email_list_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(email_list_frame, weight=40)
        
        # Controls for email list
        controls_frame = ttk.Frame(email_list_frame)
        controls_frame.pack(fill=tk.X, padx=5, pady=5)
        
        refresh_button = ttk.Button(controls_frame, text="Refresh", command=self._refresh_emails)
        refresh_button.pack(side=tk.LEFT, padx=5)
        
        self.filter_var = tk.StringVar(value="All")
        filter_combo = ttk.Combobox(controls_frame, textvariable=self.filter_var, width=15)
        filter_combo['values'] = ('All', 'Unread', 'Today', 'High Priority', 'With Actions')
        filter_combo.current(0)
        filter_combo.pack(side=tk.RIGHT, padx=5)
        
        filter_label = ttk.Label(controls_frame, text="Filter:")
        filter_label.pack(side=tk.RIGHT, padx=2)
        
        # Email list with columns
        email_list_frame_inner = ttk.Frame(email_list_frame)
        email_list_frame_inner.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        columns = ("From", "Subject", "Time", "Priority")
        self.email_list = ttk.Treeview(email_list_frame_inner, columns=columns, show="headings")
        
        # Define headings
        self.email_list.heading("From", text="From")
        self.email_list.heading("Subject", text="Subject")
        self.email_list.heading("Time", text="Time")
        self.email_list.heading("Priority", text="Priority")
        
        # Define columns
        self.email_list.column("From", width=150)
        self.email_list.column("Subject", width=250)
        self.email_list.column("Time", width=100)
        self.email_list.column("Priority", width=80)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(email_list_frame_inner, orient=tk.VERTICAL, command=self.email_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.email_list.configure(yscrollcommand=scrollbar.set)
        self.email_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Bind selection event
        self.email_list.bind('<<TreeviewSelect>>', self._on_email_selected)
        
        # Create email view frame
        email_view_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(email_view_frame, weight=60)
        
        # Email details section
        details_frame = ttk.LabelFrame(email_view_frame, text="Email Details")
        details_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # From field
        from_frame = ttk.Frame(details_frame)
        from_frame.pack(fill=tk.X, padx=5, pady=2)
        
        from_label = ttk.Label(from_frame, text="From:", width=10)
        from_label.pack(side=tk.LEFT)
        
        self.from_value = ttk.Label(from_frame, text="")
        self.from_value.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Subject field
        subject_frame = ttk.Frame(details_frame)
        subject_frame.pack(fill=tk.X, padx=5, pady=2)
        
        subject_label = ttk.Label(subject_frame, text="Subject:", width=10)
        subject_label.pack(side=tk.LEFT)
        
        self.subject_value = ttk.Label(subject_frame, text="")
        self.subject_value.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Date field
        date_frame = ttk.Frame(details_frame)
        date_frame.pack(fill=tk.X, padx=5, pady=2)
        
        date_label = ttk.Label(date_frame, text="Date:", width=10)
        date_label.pack(side=tk.LEFT)
        
        self.date_value = ttk.Label(date_frame, text="")
        self.date_value.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Priority field
        priority_frame = ttk.Frame(details_frame)
        priority_frame.pack(fill=tk.X, padx=5, pady=2)
        
        priority_label = ttk.Label(priority_frame, text="Priority:", width=10)
        priority_label.pack(side=tk.LEFT)
        
        self.priority_value = ttk.Label(priority_frame, text="")
        self.priority_value.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Email content
        content_frame = ttk.Frame(email_view_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Add notebook for content/thread views
        self.content_notebook = ttk.Notebook(content_frame)
        self.content_notebook.pack(fill=tk.BOTH, expand=True)
        
        # Email body tab
        body_frame = ttk.Frame(self.content_notebook)
        self.content_notebook.add(body_frame, text="Email Body")
        
        self.email_body = tk.Text(body_frame, wrap=tk.WORD, height=10)
        self.email_body.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        body_scrollbar = ttk.Scrollbar(self.email_body, orient=tk.VERTICAL, command=self.email_body.yview)
        body_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.email_body.configure(yscrollcommand=body_scrollbar.set)
        
        # Thread view tab
        thread_frame = ttk.Frame(self.content_notebook)
        self.content_notebook.add(thread_frame, text="Email Thread")
        
        self.thread_view = tk.Text(thread_frame, wrap=tk.WORD, height=10)
        self.thread_view.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        thread_scrollbar = ttk.Scrollbar(self.thread_view, orient=tk.VERTICAL, command=self.thread_view.yview)
        thread_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.thread_view.configure(yscrollcommand=thread_scrollbar.set)
        
        # Action items tab
        actions_frame = ttk.Frame(self.content_notebook)
        self.content_notebook.add(actions_frame, text="Action Items")
        
        self.actions_view = tk.Text(actions_frame, wrap=tk.WORD, height=10)
        self.actions_view.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        actions_scrollbar = ttk.Scrollbar(self.actions_view, orient=tk.VERTICAL, command=self.actions_view.yview)
        actions_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.actions_view.configure(yscrollcommand=actions_scrollbar.set)
        
        # Email actions
        actions_frame = ttk.Frame(email_view_frame)
        actions_frame.pack(fill=tk.X, padx=10, pady=10)
        
        reply_button = ttk.Button(actions_frame, text="Generate Reply", command=self._generate_reply)
        reply_button.pack(side=tk.LEFT, padx=5)
        
        extract_button = ttk.Button(actions_frame, text="Extract Actions", command=self._extract_actions)
        extract_button.pack(side=tk.LEFT, padx=5)
        
        mark_read_button = ttk.Button(actions_frame, text="Mark as Read", command=self._mark_as_read)
        mark_read_button.pack(side=tk.LEFT, padx=5)
        
        archive_button = ttk.Button(actions_frame, text="Archive", command=self._archive_email)
        archive_button.pack(side=tk.LEFT, padx=5)
    
    def _refresh_emails(self):
        """Refresh the email list"""
        try:
            # Show loading indicator
            self._set_busy_cursor(True)
            
            # Clear current selection
            self._clear_email_display()
            
            # Get filter value
            filter_type = self.filter_var.get()
            
            # Start refresh in a separate thread
            threading.Thread(target=self._load_emails, args=(filter_type,), daemon=True).start()
            
        except Exception as e:
            logger.error(f"Error refreshing emails: {str(e)}")
            self._set_busy_cursor(False)
            messagebox.showerror("Refresh Error", f"Failed to refresh emails: {str(e)}")
    
    def _load_emails(self, filter_type):
        """Load emails in a background thread"""
        try:
            # Get emails based on filter
            if filter_type == "Unread":
                emails = self.outlook_service.get_unread_emails(limit=50)
            else:
                # Default to recent emails
                emails = self.outlook_service.get_recent_emails(days=3, limit=100)
            
            # Process emails for display
            processed_emails = []
            
            for email in emails:
                # Process priority
                priority_result = self.priority_engine.prioritize_email(email)
                email['priority'] = priority_result.get('category', 'Medium')
                email['priority_score'] = priority_result.get('score', 50)
                
                # Apply filters
                if filter_type == "High Priority" and email['priority'] != "High" and email['priority'] != "Urgent":
                    continue
                elif filter_type == "Today":
                    # Check if email is from today
                    try:
                        email_date = datetime.strptime(email['date'], "%a, %d %b %Y %H:%M:%S")
                        today = datetime.now().date()
                        if email_date.date() != today:
                            continue
                    except:
                        # If date parsing fails, include the email
                        pass
                elif filter_type == "With Actions":
                    # Check if email has potential action items (simple check)
                    action_indicators = ["please", "request", "action", "needed", "required", "task"]
                    has_action = False
                    
                    for indicator in action_indicators:
                        if indicator in email.get('subject', '').lower() or indicator in email.get('body', '').lower():
                            has_action = True
                            break
                    
                    if not has_action:
                        continue
                
                processed_emails.append(email)
            
            # Store emails
            self.emails = processed_emails
            
            # Update UI in main thread
            self.parent.after(0, lambda: self._update_email_list(processed_emails))
            
        except Exception as e:
            logger.error(f"Error loading emails: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror("Load Error", f"Failed to load emails: {str(e)}"))
            self.parent.after(0, lambda: self._set_busy_cursor(False))
    
    def _update_email_list(self, emails):
        """Update the email list in the UI"""
        try:
            # Clear existing items
            for item in self.email_list.get_children():
                self.email_list.delete(item)
            
            # Add emails to the list
            for email in emails:
                # Format values for display
                from_value = email.get('from', '').split('<')[0].strip()
                if len(from_value) > 30:
                    from_value = from_value[:27] + "..."
                
                subject = email.get('subject', '')
                if len(subject) > 50:
                    subject = subject[:47] + "..."
                
                # Format time
                time_str = email.get('date', '')
                try:
                    date_obj = datetime.strptime(time_str, "%a, %d %b %Y %H:%M:%S")
                    time_str = date_obj.strftime("%I:%M %p")
                except:
                    time_str = time_str[-8:] if len(time_str) > 8 else time_str
                
                # Priority
                priority = email.get('priority', 'Medium')
                
                # Add to treeview
                self.email_list.insert("", tk.END, values=(from_value, subject, time_str, priority),
                                     tags=("unread" if email.get('unread', False) else "read",))
            
            # Configure tag appearance
            self.email_list.tag_configure("unread", font=("", 9, "bold"))
            self.email_list.tag_configure("read", font=("", 9, ""))
            
            # Update status
            if len(emails) > 0:
                logger.info(f"Loaded {len(emails)} emails")
            else:
                logger.info("No emails found matching the filter")
                
        except Exception as e:
            logger.error(f"Error updating email list: {str(e)}")
            
        finally:
            # Reset cursor
            self._set_busy_cursor(False)
    
    def _on_email_selected(self, event):
        """Handle email selection"""
        try:
            # Get selected item
            selection = self.email_list.selection()
            if not selection:
                return
            
            # Get index of selected item
            selected_index = self.email_list.index(selection[0])
            
            if selected_index < 0 or selected_index >= len(self.emails):
                return
            
            # Get email data
            self.current_email = self.emails[selected_index]
            
            # Show busy cursor
            self._set_busy_cursor(True)
            
            # Start thread to load email details
            threading.Thread(target=self._load_email_details, daemon=True).start()
            
        except Exception as e:
            logger.error(f"Error handling email selection: {str(e)}")
            self._set_busy_cursor(False)
    
    def _load_email_details(self):
        """Load email details in background thread"""
        try:
            # Load thread if available
            if self.current_email.get('conversation_id'):
                self.current_thread = self.outlook_service.get_thread_emails(
                    self.current_email.get('conversation_id'),
                    limit=10
                )
            else:
                self.current_thread = []
            
            # Extract action items
            action_items = self.action_extractor.extract_action_items(self.current_email)
            
            # Update UI in main thread
            self.parent.after(0, lambda: self._update_email_display())
            
        except Exception as e:
            logger.error(f"Error loading email details: {str(e)}")
            self.parent.after(0, lambda: self._set_busy_cursor(False))
    
    def _update_email_display(self):
        """Update the email display"""
        try:
            if not self.current_email:
                self._clear_email_display()
                return
            
            # Update details
            self.from_value.config(text=self.current_email.get('from', ''))
            self.subject_value.config(text=self.current_email.get('subject', ''))
            self.date_value.config(text=self.current_email.get('date', ''))
            
            # Update priority
            priority_text = self.current_email.get('priority', 'Medium')
            priority_score = self.current_email.get('priority_score', 0)
            self.priority_value.config(text=f"{priority_text} ({priority_score}/100)")
            
            # Update body
            self.email_body.config(state=tk.NORMAL)
            self.email_body.delete("1.0", tk.END)
            self.email_body.insert(tk.END, self.current_email.get('body', ''))
            self.email_body.config(state=tk.DISABLED)
            
            # Update thread view
            self.thread_view.config(state=tk.NORMAL)
            self.thread_view.delete("1.0", tk.END)
            
            if self.current_thread:
                for thread_email in self.current_thread:
                    # Format sender
                    sender = thread_email.get('from', '').split('<')[0].strip()
                    
                    # Format date
                    date_str = thread_email.get('date', '')
                    try:
                        date_obj = datetime.strptime(date_str, "%a, %d %b %Y %H:%M:%S")
                        date_str = date_obj.strftime("%m/%d/%Y %I:%M %p")
                    except:
                        pass
                    
                    # Add header
                    self.thread_view.insert(tk.END, f"From: {sender}\n", "header")
                    self.thread_view.insert(tk.END, f"Date: {date_str}\n\n", "header")
                    
                    # Add body
                    self.thread_view.insert(tk.END, f"{thread_email.get('body', '')}\n\n")
                    
                    # Add separator
                    self.thread_view.insert(tk.END, "-" * 50 + "\n\n")
            else:
                self.thread_view.insert(tk.END, "No thread history available.")
            
            self.thread_view.config(state=tk.DISABLED)
            
            # Configure text tags
            self.thread_view.tag_configure("header", font=("", 9, "bold"))
            
            # Extract and show action items
            self._extract_actions()
            
        except Exception as e:
            logger.error(f"Error updating email display: {str(e)}")
            
        finally:
            # Reset cursor
            self._set_busy_cursor(False)
    
    def _clear_email_display(self):
        """Clear the email display"""
        self.from_value.config(text="")
        self.subject_value.config(text="")
        self.date_value.config(text="")
        self.priority_value.config(text="")
        
        self.email_body.config(state=tk.NORMAL)
        self.email_body.delete("1.0", tk.END)
        self.email_body.config(state=tk.DISABLED)
        
        self.thread_view.config(state=tk.NORMAL)
        self.thread_view.delete("1.0", tk.END)
        self.thread_view.config(state=tk.DISABLED)
        
        self.actions_view.config(state=tk.NORMAL)
        self.actions_view.delete("1.0", tk.END)
        self.actions_view.config(state=tk.DISABLED)
        
        self.current_email = None
        self.current_thread = []
    
    def _generate_reply(self):
        """Generate a reply to the current email"""
        if not self.current_email:
            messagebox.showinfo("Generate Reply", "Please select an email first.")
            return
        
        try:
            # Show busy cursor
            self._set_busy_cursor(True)
            
            # Start generation in a separate thread
            threading.Thread(target=self._generate_reply_thread, daemon=True).start()
            
        except Exception as e:
            logger.error(f"Error generating reply: {str(e)}")
            self._set_busy_cursor(False)
            messagebox.showerror("Generate Reply", f"Failed to generate reply: {str(e)}")
    
    def _generate_reply_thread(self):
        """Generate reply in background thread"""
        try:
            # Generate response
            response = self.response_generator.generate_response(
                self.current_email,
                self.current_thread
            )
            
            # Save draft to storage
            draft_data = {
                'email_id': self.current_email.get('id', ''),
                'original_email': self.current_email,
                'response_text': response.get('response_text', ''),
                'formatted_email': response.get('formatted_email', '')
            }
            
            draft_id = self.storage_service.save_draft(draft_data)
            
            # Update UI in main thread
            self.parent.after(0, lambda: self._show_draft_saved(draft_id))
            
        except Exception as e:
            logger.error(f"Error in reply generation thread: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror(
                "Generate Reply", 
                f"Failed to generate reply: {str(e)}"
            ))
            self.parent.after(0, lambda: self._set_busy_cursor(False))
    
    def _show_draft_saved(self, draft_id):
        """Show confirmation that draft was saved"""
        # Reset cursor
        self._set_busy_cursor(False)
        
        if draft_id:
            messagebox.showinfo(
                "Generate Reply", 
                "Reply generated and saved as draft.\n\nYou can view and edit it in the Draft Responses tab."
            )
        else:
            messagebox.showerror(
                "Generate Reply",
                "Failed to save generated reply."
            )
    
    def _extract_actions(self):
        """Extract action items from the current email"""
        if not self.current_email:
            return
        
        try:
            # Extract action items
            action_items = self.action_extractor.extract_action_items(self.current_email)
            
            # Update actions view
            self.actions_view.config(state=tk.NORMAL)
            self.actions_view.delete("1.0", tk.END)
            
            if action_items:
                self.actions_view.insert(tk.END, "Action Items:\n\n", "header")
                
                for i, action in enumerate(action_items, 1):
                    # Add action text
                    self.actions_view.insert(tk.END, f"{i}. {action.get('text', '')}\n", "action")
                    
                    # Add due date if available
                    if action.get('due_date'):
                        self.actions_view.insert(tk.END, f"   Due: {action.get('due_date')}\n", "due")
                    
                    # Add confidence
                    confidence = int(action.get('confidence', 0) * 100)
                    self.actions_view.insert(tk.END, f"   Confidence: {confidence}%\n\n")
                
                # Add save button
                save_frame = ttk.Frame(self.actions_view)
                save_button = ttk.Button(save_frame, text="Save All to Tasks", 
                                      command=lambda: self._save_actions(action_items))
                save_button.pack(padx=5, pady=5)
                
                self.actions_view.window_create(tk.END, window=save_frame)
            else:
                self.actions_view.insert(tk.END, "No action items detected in this email.")
            
            # Configure text tags
            self.actions_view.tag_configure("header", font=("", 10, "bold"))
            self.actions_view.tag_configure("action", font=("", 9, "bold"))
            self.actions_view.tag_configure("due", font=("", 9, "italic"))
            
            self.actions_view.config(state=tk.DISABLED)
            
        except Exception as e:
            logger.error(f"Error extracting actions: {str(e)}")
    
    def _save_actions(self, action_items):
        """Save action items to tasks"""
        try:
            saved_count = 0
            
            for action in action_items:
                # Create task data
                task_data = {
                    'text': action.get('text', ''),
                    'email_id': self.current_email.get('id', ''),
                    'email_from': self.current_email.get('from', ''),
                    'due_date': action.get('due_date', None),
                    'priority': action.get('priority', 'Medium'),
                    'status': 'Not Started'
                }
                
                # Save to storage
                task_id = self.storage_service.save_task(task_data)
                
                if task_id:
                    saved_count += 1
            
            if saved_count > 0:
                messagebox.showinfo(
                    "Save Actions", 
                    f"Saved {saved_count} action items to tasks.\n\nYou can view them in the Action Items tab."
                )
            else:
                messagebox.showwarning(
                    "Save Actions",
                    "No action items were saved."
                )
                
        except Exception as e:
            logger.error(f"Error saving actions: {str(e)}")
            messagebox.showerror("Save Actions", f"Error saving actions: {str(e)}")
    
    def _mark_as_read(self):
        """Mark the current email as read"""
        if not self.current_email:
            return
        
        try:
            email_id = self.current_email.get('id', '')
            if not email_id:
                return
                
            result = self.outlook_service.mark_as_read(email_id)
            
            if result:
                # Update local data
                self.current_email['unread'] = False
                
                # Update UI
                selection = self.email_list.selection()
                if selection:
                    self.email_list.item(selection[0], tags=("read",))
            
        except Exception as e:
            logger.error(f"Error marking email as read: {str(e)}")
    
    def _archive_email(self):
        """Archive the current email"""
        if not self.current_email:
            messagebox.showinfo("Archive Email", "Please select an email first.")
            return
        
        try:
            email_id = self.current_email.get('id', '')
            if not email_id:
                return
                
            result = self.outlook_service.archive_email(email_id)
            
            if result:
                messagebox.showinfo("Archive Email", "Email archived successfully.")
                
                # Remove from list
                selection = self.email_list.selection()
                if selection:
                    self.email_list.delete(selection[0])
                    
                # Clear display
                self._clear_email_display()
            else:
                messagebox.showwarning("Archive Email", "Failed to archive email.")
            
        except Exception as e:
            logger.error(f"Error archiving email: {str(e)}")
            messagebox.showerror("Archive Email", f"Error archiving email: {str(e)}")
    
    def _set_busy_cursor(self, busy):
        """Set or clear busy cursor"""
        if busy:
            self.parent.config(cursor="wait")
        else:
            self.parent.config(cursor="")