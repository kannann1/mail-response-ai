# ui/tasks_tab.py
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import logging
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)

class TasksTab:
    """Tasks tab UI for the email agent application"""
    
    def __init__(self, parent, storage_service, outlook_service):
        self.parent = parent
        self.storage_service = storage_service
        self.outlook_service = outlook_service
        
        # Current data
        self.tasks = []
        self.current_task = None
        
        self._setup_ui()
        self._load_tasks()
    
    def _setup_ui(self):
        """Set up the tasks UI"""
        # Create main frame
        main_frame = ttk.Frame(self.parent, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top controls
        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Refresh button
        refresh_button = ttk.Button(controls_frame, text="Refresh", command=self._load_tasks)
        refresh_button.pack(side=tk.LEFT, padx=5)
        
        # Add task button
        add_button = ttk.Button(controls_frame, text="Add Task", command=self._add_task)
        add_button.pack(side=tk.LEFT, padx=5)
        
        # Filter dropdown
        self.filter_var = tk.StringVar(value="All Tasks")
        filter_label = ttk.Label(controls_frame, text="View:")
        filter_label.pack(side=tk.LEFT, padx=(20, 5))
        
        filter_combo = ttk.Combobox(controls_frame, textvariable=self.filter_var, width=15)
        filter_combo['values'] = ('All Tasks', 'Today', 'This Week', 'Completed', 'Not Started')
        filter_combo.current(0)
        filter_combo.pack(side=tk.LEFT, padx=5)
        
        # Bind filter change
        filter_combo.bind('<<ComboboxSelected>>', lambda e: self._apply_filter())
        
        # Create paned window for tasks list and details
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)
        
        # Tasks list frame
        tasks_list_frame = ttk.Frame(paned_window)
        paned_window.add(tasks_list_frame, weight=60)
        
        # Create tasks treeview
        self.tasks_tree = ttk.Treeview(
            tasks_list_frame, 
            columns=("Task", "Due Date", "Status", "Priority"),
            show="headings"
        )
        
        # Define headings
        self.tasks_tree.heading("Task", text="Task")
        self.tasks_tree.heading("Due Date", text="Due Date")
        self.tasks_tree.heading("Status", text="Status")
        self.tasks_tree.heading("Priority", text="Priority")
        
        # Define columns
        self.tasks_tree.column("Task", width=300)
        self.tasks_tree.column("Due Date", width=100)
        self.tasks_tree.column("Status", width=100)
        self.tasks_tree.column("Priority", width=80)
        
        # Scrollbar for tasks list
        tasks_scrollbar = ttk.Scrollbar(tasks_list_frame, orient=tk.VERTICAL, command=self.tasks_tree.yview)
        tasks_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tasks_tree.configure(yscrollcommand=tasks_scrollbar.set)
        self.tasks_tree.pack(fill=tk.BOTH, expand=True)
        
        # Bind selection event
        self.tasks_tree.bind('<<TreeviewSelect>>', self._on_task_selected)
        
        # Task details frame
        task_details_frame = ttk.Frame(paned_window)
        paned_window.add(task_details_frame, weight=40)
        
        # Details section
        details_frame = ttk.LabelFrame(task_details_frame, text="Task Details")
        details_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Task description
        description_frame = ttk.Frame(details_frame)
        description_frame.pack(fill=tk.X, padx=10, pady=5)
        
        description_label = ttk.Label(description_frame, text="Description:")
        description_label.pack(anchor=tk.W)
        
        self.description_text = tk.Text(description_frame, wrap=tk.WORD, height=5)
        self.description_text.pack(fill=tk.X, expand=True, pady=5)
        
        # Email source
        source_frame = ttk.Frame(details_frame)
        source_frame.pack(fill=tk.X, padx=10, pady=5)
        
        source_label = ttk.Label(source_frame, text="Source Email:")
        source_label.pack(anchor=tk.W)
        
        self.source_value = ttk.Label(source_frame, text="")
        self.source_value.pack(anchor=tk.W, pady=5)
        
        # Open email button
        self.open_email_button = ttk.Button(source_frame, text="Open Email", command=self._open_source_email)
        self.open_email_button.pack(anchor=tk.W, pady=5)
        
        # Due date and priority
        date_frame = ttk.Frame(details_frame)
        date_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Due date
        due_label = ttk.Label(date_frame, text="Due Date:")
        due_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.due_date_entry = ttk.Entry(date_frame, width=15)
        self.due_date_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Calendar button
        calendar_button = ttk.Button(date_frame, text="📅", width=3, command=self._show_calendar)
        calendar_button.grid(row=0, column=2, sticky=tk.W, padx=2, pady=5)
        
        # Priority
        priority_label = ttk.Label(date_frame, text="Priority:")
        priority_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.priority_var = tk.StringVar()
        priority_combo = ttk.Combobox(date_frame, textvariable=self.priority_var, width=15)
        priority_combo['values'] = ('High', 'Medium', 'Low')
        priority_combo.current(1)  # Default to Medium
        priority_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Status
        status_label = ttk.Label(date_frame, text="Status:")
        status_label.grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.status_var = tk.StringVar()
        status_combo = ttk.Combobox(date_frame, textvariable=self.status_var, width=15)
        status_combo['values'] = ('Not Started', 'In Progress', 'Completed')
        status_combo.current(0)  # Default to Not Started
        status_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        # Action buttons
        actions_frame = ttk.Frame(details_frame)
        actions_frame.pack(fill=tk.X, padx=10, pady=10)
        
        update_button = ttk.Button(actions_frame, text="Update Task", command=self._update_task)
        update_button.pack(side=tk.LEFT, padx=5)
        
        complete_button = ttk.Button(actions_frame, text="Mark Complete", command=self._mark_complete)
        complete_button.pack(side=tk.LEFT, padx=5)
        
        delete_button = ttk.Button(actions_frame, text="Delete Task", command=self._delete_task)
        delete_button.pack(side=tk.LEFT, padx=5)
    
    def _load_tasks(self):
        """Load tasks from storage"""
        try:
            # Show busy cursor
            self._set_busy_cursor(True)
            
            # Clear existing items
            for item in self.tasks_tree.get_children():
                self.tasks_tree.delete(item)
                
            # Get tasks from storage
            self.tasks = self.storage_service.get_tasks()
            
            # Apply current filter
            self._apply_filter()
            
        except Exception as e:
            logger.error(f"Error loading tasks: {str(e)}")
            messagebox.showerror("Load Tasks", f"Failed to load tasks: {str(e)}")
            
        finally:
            # Reset cursor
            self._set_busy_cursor(False)
    
    def _apply_filter(self):
        """Apply the selected filter to tasks"""
        try:
            # Clear existing items
            for item in self.tasks_tree.get_children():
                self.tasks_tree.delete(item)
                
            filter_type = self.filter_var.get()
            today = datetime.now().date()
            
            # Filter tasks
            filtered_tasks = []
            
            for task in self.tasks:
                if filter_type == "All Tasks":
                    filtered_tasks.append(task)
                elif filter_type == "Today":
                    due_date = task.get('due_date')
                    if due_date:
                        try:
                            # Parse date string
                            task_date = self._parse_date(due_date)
                            if task_date and task_date.date() == today:
                                filtered_tasks.append(task)
                        except:
                            # If parsing fails, skip this task
                            pass
                elif filter_type == "This Week":
                    due_date = task.get('due_date')
                    if due_date:
                        try:
                            # Parse date string
                            task_date = self._parse_date(due_date)
                            if task_date:
                                # Calculate days until end of week (Sunday)
                                days_to_end_of_week = 6 - today.weekday()  # 0=Monday, 6=Sunday
                                end_of_week = today + timedelta(days=days_to_end_of_week)
                                
                                if today <= task_date.date() <= end_of_week:
                                    filtered_tasks.append(task)
                        except:
                            # If parsing fails, skip this task
                            pass
                elif filter_type == "Completed":
                    if task.get('status') == "Completed":
                        filtered_tasks.append(task)
                elif filter_type == "Not Started":
                    if task.get('status') == "Not Started":
                        filtered_tasks.append(task)
            
            # Add filtered tasks to treeview
            for task in filtered_tasks:
                due_date = task.get('due_date', '')
                if due_date:
                    # Try to format date for display
                    try:
                        date_obj = self._parse_date(due_date)
                        if date_obj:
                            due_date = date_obj.strftime("%m/%d/%Y")
                    except:
                        # Keep original string if parsing fails
                        pass
                
                self.tasks_tree.insert(
                    "", tk.END, 
                    values=(
                        task.get('text', ''),
                        due_date,
                        task.get('status', 'Not Started'),
                        task.get('priority', 'Medium')
                    ),
                    tags=(task.get('status', '').lower().replace(' ', '_'),)
                )
            
            # Configure tag appearances
            self.tasks_tree.tag_configure('completed', foreground='gray')
            self.tasks_tree.tag_configure('in_progress', foreground='blue')
            self.tasks_tree.tag_configure('not_started', foreground='black')
            
            # Clear details
            self._clear_task_details()
            
        except Exception as e:
            logger.error(f"Error applying filter: {str(e)}")
    
    def _parse_date(self, date_string):
        """Parse date string into datetime object"""
        # Try various date formats
        formats = [
            "%Y-%m-%d",
            "%m/%d/%Y",
            "%d/%m/%Y",
            "%B %d, %Y"
        ]
        
        for fmt in formats:
            try:
                return datetime.strptime(date_string, fmt)
            except:
                continue
        
        return None
    
    def _on_task_selected(self, event):
        """Handle task selection"""
        try:
            selection = self.tasks_tree.selection()
            if not selection:
                return
                
            # Get selected item values
            item = selection[0]
            values = self.tasks_tree.item(item, 'values')
            
            if not values:
                return
                
            # Find the task in our list
            task_text = values[0]
            for task in self.tasks:
                if task.get('text') == task_text:
                    self.current_task = task
                    break
            
            # Update details
            self._update_task_details()
            
        except Exception as e:
            logger.error(f"Error handling task selection: {str(e)}")
    
    def _update_task_details(self):
        """Update the task details display"""
        if not self.current_task:
            self._clear_task_details()
            return
            
        try:
            # Set description
            self.description_text.delete("1.0", tk.END)
            self.description_text.insert("1.0", self.current_task.get('text', ''))
            
            # Set source email
            source_email = self.current_task.get('email_from', '')
            self.source_value.config(text=source_email if source_email else "No source email")
            
            # Enable/disable open email button
            self.open_email_button.config(state=tk.NORMAL if self.current_task.get('email_id') else tk.DISABLED)
            
            # Set due date
            self.due_date_entry.delete(0, tk.END)
            if self.current_task.get('due_date'):
                self.due_date_entry.insert(0, self.current_task.get('due_date'))
            
            # Set priority
            priority = self.current_task.get('priority', 'Medium')
            self.priority_var.set(priority)
            
            # Set status
            status = self.current_task.get('status', 'Not Started')
            self.status_var.set(status)
            
        except Exception as e:
            logger.error(f"Error updating task details: {str(e)}")
    
    def _clear_task_details(self):
        """Clear the task details display"""
        self.description_text.delete("1.0", tk.END)
        self.source_value.config(text="")
        self.open_email_button.config(state=tk.DISABLED)
        self.due_date_entry.delete(0, tk.END)
        self.priority_var.set("Medium")
        self.status_var.set("Not Started")
        
        self.current_task = None
    
    def _add_task(self):
        """Add a new task"""
        try:
            # Create add task dialog
            dialog = tk.Toplevel(self.parent)
            dialog.title("Add Task")
            dialog.geometry("500x300")
            dialog.transient(self.parent)
            dialog.grab_set()
            
            # Center the dialog
            dialog.update_idletasks()
            width = dialog.winfo_width()
            height = dialog.winfo_height()
            x = (dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (dialog.winfo_screenheight() // 2) - (height // 2)
            dialog.geometry(f"+{x}+{y}")
            
            # Create content frame
            frame = ttk.Frame(dialog, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Task description
            ttk.Label(frame, text="Task Description:").pack(anchor=tk.W, pady=(0, 5))
            description = tk.Text(frame, wrap=tk.WORD, height=5)
            description.pack(fill=tk.X, pady=(0, 10))
            
            # Due date
            date_frame = ttk.Frame(frame)
            date_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(date_frame, text="Due Date:").pack(side=tk.LEFT, padx=(0, 5))
            due_date = ttk.Entry(date_frame, width=15)
            due_date.pack(side=tk.LEFT, padx=(0, 5))
            
            # Calendar button
            ttk.Button(date_frame, text="📅", width=3, command=lambda: self._show_calendar_for_entry(due_date)).pack(side=tk.LEFT)
            
            # Priority
            priority_frame = ttk.Frame(frame)
            priority_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(priority_frame, text="Priority:").pack(side=tk.LEFT, padx=(0, 5))
            priority_var = tk.StringVar(value="Medium")
            priority_combo = ttk.Combobox(priority_frame, textvariable=priority_var, width=15)
            priority_combo['values'] = ('High', 'Medium', 'Low')
            priority_combo.current(1)  # Default to Medium
            priority_combo.pack(side=tk.LEFT)
            
            # Button frame
            button_frame = ttk.Frame(frame)
            button_frame.pack(fill=tk.X, pady=(10, 0), anchor=tk.S)
            
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
            
            def save_task():
                # Get values
                task_text = description.get("1.0", tk.END).strip()
                if not task_text:
                    messagebox.showerror("Add Task", "Task description is required.")
                    return
                
                task_data = {
                    'text': task_text,
                    'due_date': due_date.get().strip(),
                    'priority': priority_var.get(),
                    'status': 'Not Started'
                }
                
                # Save to storage
                task_id = self.storage_service.save_task(task_data)
                
                if task_id:
                    # Add to list
                    task_data['id'] = task_id
                    self.tasks.append(task_data)
                    
                    # Refresh display
                    self._apply_filter()
                    
                    dialog.destroy()
                    messagebox.showinfo("Add Task", "Task added successfully.")
                else:
                    messagebox.showerror("Add Task", "Failed to add task.")
            
            ttk.Button(button_frame, text="Save", command=save_task).pack(side=tk.RIGHT, padx=5)
            
            # Focus the description field
            description.focus_set()
            
        except Exception as e:
            logger.error(f"Error adding task: {str(e)}")
            messagebox.showerror("Add Task", f"Error: {str(e)}")
    
    def _update_task(self):
        """Update the current task"""
        if not self.current_task:
            messagebox.showinfo("Update Task", "Please select a task first.")
            return
            
        try:
            # Get updated values
            task_text = self.description_text.get("1.0", tk.END).strip()
            due_date = self.due_date_entry.get().strip()
            priority = self.priority_var.get()
            status = self.status_var.get()
            
            if not task_text:
                messagebox.showerror("Update Task", "Task description is required.")
                return
            
            # Update task data
            self.current_task['text'] = task_text
            self.current_task['due_date'] = due_date
            self.current_task['priority'] = priority
            self.current_task['status'] = status
            
            # Save to storage
            result = self.storage_service.update_task(
                self.current_task['id'],
                {
                    'text': task_text,
                    'due_date': due_date,
                    'priority': priority,
                    'status': status
                }
            )
            
            if result:
                # Refresh display
                self._apply_filter()
                messagebox.showinfo("Update Task", "Task updated successfully.")
            else:
                messagebox.showerror("Update Task", "Failed to update task.")
            
        except Exception as e:
            logger.error(f"Error updating task: {str(e)}")
            messagebox.showerror("Update Task", f"Error: {str(e)}")
    
    def _mark_complete(self):
        """Mark the current task as complete"""
        if not self.current_task:
            messagebox.showinfo("Mark Complete", "Please select a task first.")
            return
            
        try:
            # Update status
            self.current_task['status'] = 'Completed'
            self.status_var.set('Completed')
            
            # Save to storage
            result = self.storage_service.update_task(
                self.current_task['id'],
                {
                    'text': self.current_task['text'],
                    'due_date': self.current_task.get('due_date', ''),
                    'priority': self.current_task.get('priority', 'Medium'),
                    'status': 'Completed'
                }
            )
            
            if result:
                # Refresh display
                self._apply_filter()
                messagebox.showinfo("Mark Complete", "Task marked as complete.")
            else:
                messagebox.showerror("Mark Complete", "Failed to update task status.")
            
        except Exception as e:
            logger.error(f"Error marking task complete: {str(e)}")
            messagebox.showerror("Mark Complete", f"Error: {str(e)}")
    
    def _delete_task(self):
        """Delete the current task"""
        if not self.current_task:
            messagebox.showinfo("Delete Task", "Please select a task first.")
            return
            
        try:
            # Confirm deletion
            if not messagebox.askyesno("Delete Task", "Are you sure you want to delete this task?"):
                return
                
            # Delete from storage
            result = self.storage_service.delete_task(self.current_task['id'])
            
            if result:
                # Remove from list
                self.tasks = [t for t in self.tasks if t.get('id') != self.current_task['id']]
                
                # Refresh display
                self._apply_filter()
                
                # Clear details
                self._clear_task_details()
                
                messagebox.showinfo("Delete Task", "Task deleted successfully.")
            else:
                messagebox.showerror("Delete Task", "Failed to delete task.")
            
        except Exception as e:
            logger.error(f"Error deleting task: {str(e)}")
            messagebox.showerror("Delete Task", f"Error: {str(e)}")
    
    def _open_source_email(self):
        """Open the source email if available"""
        if not self.current_task or not self.current_task.get('email_id'):
            return
            
        try:
            # This would typically open the email in Outlook
            # For now, just show a message
            messagebox.showinfo(
                "Open Email",
                "This would open the source email in Outlook.\n\n"
                f"Email ID: {self.current_task.get('email_id')}"
            )
            
        except Exception as e:
            logger.error(f"Error opening source email: {str(e)}")
            messagebox.showerror("Open Email", f"Error: {str(e)}")
    
    def _show_calendar(self):
        """Show a calendar for date selection"""
        self._show_calendar_for_entry(self.due_date_entry)
    
# In ui/tasks_tab.py, replace the calendar-related code:

    def _show_calendar_for_entry(self, entry_widget):
        """Show a calendar for date selection for the specified entry widget"""
        try:
            # Create calendar dialog
            dialog = tk.Toplevel(self.parent)
            dialog.title("Select Date")
            dialog.geometry("300x250")
            dialog.transient(self.parent)
            dialog.grab_set()
            
            # Center the dialog
            dialog.update_idletasks()
            width = dialog.winfo_width()
            height = dialog.winfo_height()
            x = (dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (dialog.winfo_screenheight() // 2) - (height // 2)
            dialog.geometry(f"+{x}+{y}")
            
            # Create content frame
            frame = ttk.Frame(dialog, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Get current date or parse existing date
            try:
                current_date = entry_widget.get().strip()
                if current_date:
                    date_obj = self._parse_date(current_date)
                    if date_obj:
                        year, month, day = date_obj.year, date_obj.month, date_obj.day
                    else:
                        year, month, day = datetime.now().year, datetime.now().month, datetime.now().day
                else:
                    year, month, day = datetime.now().year, datetime.now().month, datetime.now().day
            except:
                year, month, day = datetime.now().year, datetime.now().month, datetime.now().day
            
            # Calendar variables
            cal_year = tk.IntVar(value=year)
            cal_month = tk.IntVar(value=month)
            
            # Month and year navigation
            nav_frame = ttk.Frame(frame)
            nav_frame.pack(fill=tk.X, pady=(0, 10))
            
            # Month names
            month_names = [
                "January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"
            ]
            
            month_label = ttk.Label(nav_frame, text=month_names[month-1] + " " + str(year), width=20)
            month_label.pack(side=tk.LEFT, padx=5)
            
            # Day labels
            days_frame = ttk.Frame(frame)
            days_frame.pack(fill=tk.X)
            
            day_labels = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]
            for i, day_label in enumerate(day_labels):
                ttk.Label(days_frame, text=day_label, width=3).grid(row=0, column=i, padx=1, pady=1)
            
            # Calendar grid
            cal_frame = ttk.Frame(frame)
            cal_frame.pack(fill=tk.BOTH, expand=True)
            
            # Function to update calendar
            def update_calendar():
                # Clear existing calendar
                for widget in cal_frame.winfo_children():
                    widget.destroy()
                
                # Get month details
                current_month = cal_month.get()
                current_year = cal_year.get()
                
                # Update month label
                month_label.config(text=month_names[current_month-1] + " " + str(current_year))
                
                # Get first day of month and number of days
                first_day = datetime(current_year, current_month, 1).weekday()  # 0 is Monday
                if current_month == 12:
                    next_month = 1
                    next_year = current_year + 1
                else:
                    next_month = current_month + 1
                    next_year = current_year
                    
                last_day = (datetime(next_year, next_month, 1) - timedelta(days=1)).day
                
                # Create calendar buttons
                day = 1
                for week in range(6):  # Maximum 6 weeks in a month view
                    if day > last_day:
                        break
                        
                    for weekday in range(7):  # 7 days in a week
                        if week == 0 and weekday < first_day:
                            # Empty cell before first day
                            ttk.Label(cal_frame, text="", width=3).grid(row=week+1, column=weekday, padx=1, pady=1)
                        elif day <= last_day:
                            # Day button
                            btn = ttk.Button(
                                cal_frame, 
                                text=str(day), 
                                width=3,
                                command=lambda d=day: select_date(d)
                            )
                            btn.grid(row=week+1, column=weekday, padx=1, pady=1)
                            day += 1
                        else:
                            # Empty cell after last day
                            ttk.Label(cal_frame, text="", width=3).grid(row=week+1, column=weekday, padx=1, pady=1)
            
            # Function to change month
            def prev_month_click():
                current = cal_month.get()
                if current <= 1:
                    cal_month.set(12)
                    cal_year.set(cal_year.get() - 1)
                else:
                    cal_month.set(current - 1)
                update_calendar()
                
            def next_month_click():
                current = cal_month.get()
                if current >= 12:
                    cal_month.set(1)
                    cal_year.set(cal_year.get() + 1)
                else:
                    cal_month.set(current + 1)
                update_calendar()
            
            # Previous month button
            prev_month = ttk.Button(
                nav_frame, 
                text="<", 
                width=2, 
                command=prev_month_click
            )
            prev_month.pack(side=tk.LEFT)
            
            # Next month button
            next_month = ttk.Button(
                nav_frame, 
                text=">", 
                width=2, 
                command=next_month_click
            )
            next_month.pack(side=tk.RIGHT)
            
            # Function to select a date
            def select_date(day):
                selected_date = datetime(cal_year.get(), cal_month.get(), day)
                formatted_date = selected_date.strftime("%Y-%m-%d")
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, formatted_date)
                dialog.destroy()
            
            # Quick buttons
            quick_frame = ttk.Frame(frame)
            quick_frame.pack(fill=tk.X, pady=10)
            
            def select_today():
                today = datetime.now()
                cal_year.set(today.year)
                cal_month.set(today.month)
                update_calendar()
                select_date(today.day)
                
            def select_tomorrow():
                tomorrow = datetime.now() + timedelta(days=1)
                cal_year.set(tomorrow.year)
                cal_month.set(tomorrow.month)
                update_calendar()
                select_date(tomorrow.day)
                
            def select_next_week():
                next_week = datetime.now() + timedelta(days=7)
                cal_year.set(next_week.year)
                cal_month.set(next_week.month)
                update_calendar()
                select_date(next_week.day)
            
            ttk.Button(quick_frame, text="Today", command=select_today).pack(side=tk.LEFT, padx=5)
            ttk.Button(quick_frame, text="Tomorrow", command=select_tomorrow).pack(side=tk.LEFT, padx=5)
            ttk.Button(quick_frame, text="Next Week", command=select_next_week).pack(side=tk.LEFT, padx=5)
            
            # Initialize calendar
            update_calendar()
            
        except Exception as e:
            logger.error(f"Error showing calendar: {str(e)}")
            messagebox.showerror("Calendar", f"Error: {str(e)}")    
    
    def _set_busy_cursor(self, busy):
        """Set or clear busy cursor"""
        if busy:
            self.parent.config(cursor="wait")
        else:
            self.parent.config(cursor="")