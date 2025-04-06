# services/outlook_service.py
import win32com.client
import pythoncom
import logging
import time
import threading
from datetime import datetime, timedelta
import re
import os

logger = logging.getLogger(__name__)

class OutlookService:
    """Simplified service for interacting with Outlook"""
    
    def __init__(self, config):
        self.config = config
        self.outlook = None
        self.namespace = None
        self.inbox = None
        self.drafts = None
        self.sent_items = None
        self._initialized = False
        self._monitoring = False
        self._monitor_thread = None
    
    def initialize(self):
        """Initialize connection to Outlook"""
        try:
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Connect to Outlook
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Test connection
            accounts = self.outlook.Session.Accounts
            if accounts.Count == 0:
                logger.error("No email accounts configured in Outlook")
                return False
                
            logger.info(f"Connected to Outlook with {accounts.Count} accounts")
            
            # Try to access folders but don't fail if individual folders fail
            try:
                self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
                # Test access to items
                test_count = self.inbox.Items.Count
                logger.info(f"Accessed inbox with {test_count} items")
            except Exception as e:
                logger.error(f"Error accessing inbox: {str(e)}")
                self.inbox = None
            
            try:
                self.drafts = self.namespace.GetDefaultFolder(16)  # 16 = olFolderDrafts
            except Exception as e:
                logger.error(f"Error accessing drafts folder: {str(e)}")
                self.drafts = None
                
            try:
                self.sent_items = self.namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
            except Exception as e:
                logger.error(f"Error accessing sent items folder: {str(e)}")
                self.sent_items = None
            
            # Success if we can at least access the inbox
            self._initialized = self.inbox is not None
            return self._initialized
            
        except Exception as e:
            logger.error(f"Error connecting to Outlook: {str(e)}")
            return False
    
    def _ensure_connection(self):
        """Ensure we have a connection to Outlook"""
        if not self._initialized:
            return self.initialize()
        return True
    
    def get_unread_emails(self, limit=20):
        """Get unread emails from inbox"""
        if not self._ensure_connection():
            return []
            
        try:
            # Get unread items
            unread_filter = "[Unread]=True"
            unread_items = self.inbox.Items.Restrict(unread_filter)
            unread_items.Sort("[ReceivedTime]", True)  # Descending order
            
            emails = []
            count = 0
            
            for item in unread_items:
                if count >= limit:
                    break
                    
                try:
                    email_data = {
                        'id': item.EntryID,
                        'subject': item.Subject or "(No Subject)",
                        'from': item.SenderName + " <" + item.SenderEmailAddress + ">",
                        'to': item.To,
                        'date': item.ReceivedTime.strftime("%a, %d %b %Y %H:%M:%S"),
                        'body': item.Body,
                        'unread': True,
                        'has_attachments': item.Attachments.Count > 0,
                        'conversation_id': item.ConversationID if hasattr(item, 'ConversationID') else None
                    }
                    
                    emails.append(email_data)
                    count += 1
                    
                except Exception as e:
                    logger.error(f"Error processing email: {str(e)}")
                    continue
            
            return emails
            
        except Exception as e:
            logger.error(f"Error retrieving unread emails: {str(e)}")
            # Reset connection for next attempt
            self._initialized = False
            pythoncom.CoUninitialize()
            return []
    
    def get_recent_emails(self, days=2, limit=50):
        """Get recent emails from inbox"""
        if not self._ensure_connection():
            return []
            
        try:
            # Calculate date filter
            date_filter = datetime.now() - timedelta(days=days)
            date_filter_str = date_filter.strftime("%m/%d/%Y %H:%M %p")
            filter_criteria = f"[ReceivedTime] >= '{date_filter_str}'"
            
            # Get filtered items
            recent_items = self.inbox.Items.Restrict(filter_criteria)
            recent_items.Sort("[ReceivedTime]", True)  # Descending order
            
            emails = []
            count = 0
            
            for item in recent_items:
                if count >= limit:
                    break
                    
                try:
                    email_data = {
                        'id': item.EntryID,
                        'subject': item.Subject or "(No Subject)",
                        'from': item.SenderName + " <" + item.SenderEmailAddress + ">",
                        'to': item.To,
                        'date': item.ReceivedTime.strftime("%a, %d %b %Y %H:%M:%S"),
                        'body': item.Body,
                        'unread': item.UnRead,
                        'has_attachments': item.Attachments.Count > 0,
                        'conversation_id': item.ConversationID if hasattr(item, 'ConversationID') else None
                    }
                    
                    emails.append(email_data)
                    count += 1
                    
                except Exception as e:
                    logger.error(f"Error processing email: {str(e)}")
                    continue
            
            return emails
            
        except Exception as e:
            logger.error(f"Error retrieving recent emails: {str(e)}")
            # Reset connection for next attempt
            self._initialized = False
            return []
    
    def get_thread_emails(self, conversation_id, limit=10):
        """Get all emails in a conversation thread - simpler approach"""
        if not conversation_id:
            return []
            
        # Reinitialize COM for this operation
        self._initialized = False
        if not self._ensure_connection():
            return []
            
        thread_emails = []
        
        # Process a folder by manually checking ConversationID
        def process_folder(folder, folder_name, date_field='ReceivedTime'):
            if not folder:
                return
                
            try:
                # Get recent items (last 30 days)
                items = folder.Items
                items.Sort("[ReceivedTime]", False)  # Ascending order
                
                # Loop through items and check ConversationID
                item_count = min(100, items.Count)  # Limit search to prevent performance issues
                
                for i in range(item_count):
                    try:
                        item = items.Item(i+1)  # Item indexing starts at 1
                        
                        # Check if this item is part of the thread
                        try:
                            item_conv_id = item.ConversationID
                            if item_conv_id != conversation_id:
                                continue
                        except:
                            continue
                        
                        # This email is part of the thread - process it
                        try:
                            subject = item.Subject or "(No Subject)"
                        except:
                            subject = "(No Subject)"
                            
                        try:
                            sender_name = item.SenderName or "Unknown"
                        except:
                            sender_name = "Unknown"
                            
                        try:
                            sender_email = item.SenderEmailAddress or ""
                        except:
                            sender_email = ""
                        
                        # Get date based on folder type
                        try:
                            if date_field == 'SentOn':
                                date_str = item.SentOn.strftime("%a, %d %b %Y %H:%M:%S")
                            else:
                                date_str = item.ReceivedTime.strftime("%a, %d %b %Y %H:%M:%S")
                        except:
                            date_str = "(unknown date)"
                        
                        email_data = {
                            'id': item.EntryID,
                            'subject': subject,
                            'from': f"{sender_name} <{sender_email}>",
                            'date': date_str,
                            'body': item.Body,
                            'folder': folder_name
                        }
                        
                        thread_emails.append(email_data)
                        
                        # Stop if we've reached the limit
                        if len(thread_emails) >= limit:
                            return
                            
                    except Exception as e:
                        continue
                        
            except Exception as e:
                logger.error(f"Error accessing {folder_name} items: {str(e)}")
        
        # Process inbox items
        process_folder(self.inbox, 'inbox')
        
        # Process sent items
        process_folder(self.sent_items, 'sent', 'SentOn')
        
        # Sort by date
        if thread_emails:
            try:
                thread_emails.sort(key=lambda x: datetime.strptime(x['date'], "%a, %d %b %Y %H:%M:%S"))
            except:
                # If date parsing fails, try to maintain the order as-is
                pass
            
        return thread_emails


    def start_monitoring(self, callback, root=None):
        """Simplified monitoring - just do an initial refresh"""
        logger.info("Email monitoring is disabled in this version for stability")
        # You could implement a simpler version here if needed
        return True
    
    def mark_as_read(self, email_id):
        """Mark an email as read"""
        if not self._ensure_connection():
            return False
            
        try:
            item = self.namespace.GetItemFromID(email_id)
            item.UnRead = False
            item.Save()
            return True
            
        except Exception as e:
            logger.error(f"Error marking email as read: {str(e)}")
            return False
    
    def archive_email(self, email_id):
        """Archive an email (move to Archive folder)"""
        if not self._ensure_connection():
            return False
            
        try:
            item = self.namespace.GetItemFromID(email_id)
            
            # Find or create Archive folder
            archive_folder = None
            
            try:
                # Try to find existing Archive folder
                for folder in self.namespace.Folders.Item(1).Folders:
                    if folder.Name == "Archive":
                        archive_folder = folder
                        break
                
                # Create if not found
                if not archive_folder:
                    archive_folder = self.namespace.Folders.Item(1).Folders.Add("Archive")
            except Exception as e:
                logger.error(f"Error finding/creating Archive folder: {str(e)}")
                return False
            
            # Move item
            item.Move(archive_folder)
            return True
            
        except Exception as e:
            logger.error(f"Error archiving email: {str(e)}")
            return False
    
    def send_email(self, draft_id=None, to_email=None, subject=None, body=None):
        """Send an email, either from a draft ID or new information"""
        if not self._ensure_connection():
            return False
            
        try:
            if draft_id:
                # Send existing draft
                mail_item = self.namespace.GetItemFromID(draft_id)
            else:
                # Create new mail item
                mail_item = self.outlook.CreateItem(0)  # 0 = olMailItem
                mail_item.To = to_email
                mail_item.Subject = subject
                mail_item.Body = body
            
            # Send the email
            mail_item.Send()
            return True
            
        except Exception as e:
            logger.error(f"Error sending email: {str(e)}")
            return False
    
    def create_draft(self, to_email, subject, body):
        """Create a draft email"""
        if not self._ensure_connection():
            return None
            
        try:
            # Create new mail item
            mail_item = self.outlook.CreateItem(0)  # 0 = olMailItem
            mail_item.To = to_email
            mail_item.Subject = subject
            mail_item.Body = body
            
            # Save as draft
            mail_item.Save()
            
            return {
                'id': mail_item.EntryID,
                'subject': mail_item.Subject,
                'to': mail_item.To,
                'body': mail_item.Body
            }
            
        except Exception as e:
            logger.error(f"Error creating draft: {str(e)}")
            return None