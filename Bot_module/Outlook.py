import win32com.client
import os
import datetime
from typing import List, Dict, Any, Optional, Union
import time

class OutlookHandler:
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.inbox = None
        self.sent_items = None
        self.drafts = None
        self._connect()
    
    def _connect(self):
        """Connect to Outlook application"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Get common folders
            self.inbox = self.namespace.GetDefaultFolder(6)  # Inbox
            self.sent_items = self.namespace.GetDefaultFolder(5)  # Sent Items
            self.drafts = self.namespace.GetDefaultFolder(16)  # Drafts
            
            print("Successfully connected to Outlook")
            return True
        except Exception as e:
            print(f"Error connecting to Outlook: {str(e)}")
            return False
    
    def send_email(self, to_recipients: Union[str, List[str]], subject: str, body: str,
                   cc_recipients: Union[str, List[str]] = None, 
                   bcc_recipients: Union[str, List[str]] = None,
                   attachments: List[str] = None, 
                   html_body: bool = False,
                   send_immediately: bool = True):
        """
        Send an email
        
        Args:
            to_recipients: Email recipient(s)
            subject: Email subject
            body: Email body content
            cc_recipients: CC recipient(s)
            bcc_recipients: BCC recipient(s)
            attachments: List of file paths to attach
            html_body: Whether body is HTML format
            send_immediately: Send immediately or save to drafts
        """
        try:
            mail = self.outlook.CreateItem(0)  # 0 = Mail item
            
            # Set recipients
            if isinstance(to_recipients, str):
                mail.To = to_recipients
            else:
                mail.To = "; ".join(to_recipients)
            
            if cc_recipients:
                if isinstance(cc_recipients, str):
                    mail.CC = cc_recipients
                else:
                    mail.CC = "; ".join(cc_recipients)
            
            if bcc_recipients:
                if isinstance(bcc_recipients, str):
                    mail.BCC = bcc_recipients
                else:
                    mail.BCC = "; ".join(bcc_recipients)
            
            # Set subject and body
            mail.Subject = subject
            if html_body:
                mail.HTMLBody = body
            else:
                mail.Body = body
            
            # Add attachments
            if attachments:
                for attachment in attachments:
                    if os.path.exists(attachment):
                        mail.Attachments.Add(attachment)
                        print(f"Attached: {attachment}")
                    else:
                        print(f"Warning: Attachment not found: {attachment}")
            
            # Send or save to drafts
            if send_immediately:
                mail.Send()
                print("Email sent successfully")
            else:
                mail.Save()
                print("Email saved to drafts")
            
            return True
        except Exception as e:
            print(f"Error sending email: {str(e)}")
            return False
    
    def get_emails(self, folder_name: str = "Inbox", count: int = 10, 
                   unread_only: bool = False, from_date: datetime.datetime = None,
                   to_date: datetime.datetime = None) -> List[Dict]:
        """
        Retrieve emails from a specific folder
        
        Args:
            folder_name: Name of the folder (Inbox, Sent Items, etc.)
            count: Number of emails to retrieve
            unread_only: Only get unread emails
            from_date: Get emails from this date onwards
            to_date: Get emails up to this date
        
        Returns:
            List of dictionaries containing email information
        """
        try:
            folder = self._get_folder(folder_name)
            if not folder:
                return []
            
            messages = folder.Items
            
            # Sort by received time (newest first)
            messages.Sort("[ReceivedTime]", True)
            
            # Apply filters
            if from_date or to_date or unread_only:
                filter_str = ""
                conditions = []
                
                if unread_only:
                    conditions.append("[Unread] = True")
                
                if from_date:
                    conditions.append(f"[ReceivedTime] >= '{from_date.strftime('%m/%d/%Y')}'")
                
                if to_date:
                    conditions.append(f"[ReceivedTime] <= '{to_date.strftime('%m/%d/%Y')}'")
                
                if conditions:
                    filter_str = " AND ".join(conditions)
                    messages = messages.Restrict(filter_str)
            
            emails = []
            retrieved = 0
            
            for message in messages:
                if retrieved >= count:
                    break
                
                try:
                    email_info = {
                        'subject': message.Subject,
                        'sender': message.SenderName,
                        'sender_email': message.SenderEmailAddress,
                        'received_time': message.ReceivedTime,
                        'body': message.Body,
                        'html_body': message.HTMLBody,
                        'unread': message.Unread,
                        'size': message.Size,
                        'attachments': self._get_attachment_info(message),
                        'entry_id': message.EntryID
                    }
                    emails.append(email_info)
                    retrieved += 1
                except Exception as e:
                    print(f"Error processing email: {str(e)}")
                    continue
            
            print(f"Retrieved {len(emails)} emails from {folder_name}")
            return emails
        except Exception as e:
            print(f"Error retrieving emails: {str(e)}")
            return []
    
    def search_emails(self, search_term: str, folder_name: str = "Inbox", 
                     search_in: str = "subject") -> List[Dict]:
        """
        Search for emails containing specific terms
        
        Args:
            search_term: Term to search for
            folder_name: Folder to search in
            search_in: Where to search ('subject', 'body', 'sender', 'all')
        
        Returns:
            List of matching emails
        """
        try:
            folder = self._get_folder(folder_name)
            if not folder:
                return []
            
            # Build search filter
            if search_in == "subject":
                filter_str = f"[Subject] LIKE '%{search_term}%'"
            elif search_in == "body":
                filter_str = f"[Body] LIKE '%{search_term}%'"
            elif search_in == "sender":
                filter_str = f"[SenderName] LIKE '%{search_term}%'"
            elif search_in == "all":
                filter_str = f"[Subject] LIKE '%{search_term}%' OR [Body] LIKE '%{search_term}%' OR [SenderName] LIKE '%{search_term}%'"
            else:
                filter_str = f"[Subject] LIKE '%{search_term}%'"
            
            messages = folder.Items.Restrict(filter_str)
            
            emails = []
            for message in messages:
                try:
                    email_info = {
                        'subject': message.Subject,
                        'sender': message.SenderName,
                        'sender_email': message.SenderEmailAddress,
                        'received_time': message.ReceivedTime,
                        'body': message.Body[:200] + "..." if len(message.Body) > 200 else message.Body,
                        'unread': message.Unread,
                        'entry_id': message.EntryID
                    }
                    emails.append(email_info)
                except Exception as e:
                    continue
            
            print(f"Found {len(emails)} emails matching '{search_term}'")
            return emails
        except Exception as e:
            print(f"Error searching emails: {str(e)}")
            return []
    
    def mark_as_read(self, entry_id: str):
        """Mark an email as read"""
        try:
            message = self.namespace.GetItemFromID(entry_id)
            message.Unread = False
            message.Save()
            print("Email marked as read")
            return True
        except Exception as e:
            print(f"Error marking email as read: {str(e)}")
            return False
    
    def mark_as_unread(self, entry_id: str):
        """Mark an email as unread"""
        try:
            message = self.namespace.GetItemFromID(entry_id)
            message.Unread = True
            message.Save()
            print("Email marked as unread")
            return True
        except Exception as e:
            print(f"Error marking email as unread: {str(e)}")
            return False
    
    def delete_email(self, entry_id: str):
        """Delete an email"""
        try:
            message = self.namespace.GetItemFromID(entry_id)
            message.Delete()
            print("Email deleted")
            return True
        except Exception as e:
            print(f"Error deleting email: {str(e)}")
            return False
    
    def move_email(self, entry_id: str, destination_folder: str):
        """Move an email to a different folder"""
        try:
            message = self.namespace.GetItemFromID(entry_id)
            dest_folder = self._get_folder(destination_folder)
            if dest_folder:
                message.Move(dest_folder)
                print(f"Email moved to {destination_folder}")
                return True
            else:
                print(f"Destination folder '{destination_folder}' not found")
                return False
        except Exception as e:
            print(f"Error moving email: {str(e)}")
            return False
    
    def save_attachments(self, entry_id: str, save_path: str) -> List[str]:
        """
        Save all attachments from an email
        
        Args:
            entry_id: Email entry ID
            save_path: Directory to save attachments
        
        Returns:
            List of saved file paths
        """
        try:
            message = self.namespace.GetItemFromID(entry_id)
            attachments = message.Attachments
            
            if not os.path.exists(save_path):
                os.makedirs(save_path)
            
            saved_files = []
            for attachment in attachments:
                filename = attachment.FileName
                file_path = os.path.join(save_path, filename)
                attachment.SaveAsFile(file_path)
                saved_files.append(file_path)
                print(f"Saved attachment: {file_path}")
            
            return saved_files
        except Exception as e:
            print(f"Error saving attachments: {str(e)}")
            return []
    
    def create_appointment(self, subject: str, start_time: datetime.datetime,
                          end_time: datetime.datetime, body: str = "",
                          location: str = "", attendees: List[str] = None,
                          reminder_minutes: int = 15):
        """
        Create a calendar appointment
        
        Args:
            subject: Appointment subject
            start_time: Start time
            end_time: End time
            body: Appointment body/description
            location: Meeting location
            attendees: List of attendee email addresses
            reminder_minutes: Reminder time in minutes before appointment
        """
        try:
            appointment = self.outlook.CreateItem(1)  # 1 = Appointment item
            
            appointment.Subject = subject
            appointment.Start = start_time
            appointment.End = end_time
            appointment.Body = body
            appointment.Location = location
            appointment.ReminderMinutesBeforeStart = reminder_minutes
            
            if attendees:
                for attendee in attendees:
                    appointment.Recipients.Add(attendee)
                appointment.Recipients.ResolveAll()
            
            appointment.Save()
            print("Appointment created successfully")
            return True
        except Exception as e:
            print(f"Error creating appointment: {str(e)}")
            return False
    
    def create_task(self, subject: str, body: str = "", due_date: datetime.datetime = None,
                   priority: int = 1, reminder: bool = True):
        """
        Create a task
        
        Args:
            subject: Task subject
            body: Task description
            due_date: Due date for the task
            priority: Priority level (0=Low, 1=Normal, 2=High)
            reminder: Set reminder for the task
        """
        try:
            task = self.outlook.CreateItem(3)  # 3 = Task item
            
            task.Subject = subject
            task.Body = body
            task.Importance = priority
            
            if due_date:
                task.DueDate = due_date
            
            if reminder and due_date:
                task.ReminderSet = True
                task.ReminderTime = due_date
            
            task.Save()
            print("Task created successfully")
            return True
        except Exception as e:
            print(f"Error creating task: {str(e)}")
            return False
    
    def get_contacts(self, name_filter: str = None) -> List[Dict]:
        """
        Get contacts from Outlook
        
        Args:
            name_filter: Filter contacts by name (optional)
        
        Returns:
            List of contact dictionaries
        """
        try:
            contacts_folder = self.namespace.GetDefaultFolder(10)  # 10 = Contacts
            contacts = contacts_folder.Items
            
            if name_filter:
                contacts = contacts.Restrict(f"[FullName] LIKE '%{name_filter}%'")
            
            contact_list = []
            for contact in contacts:
                try:
                    contact_info = {
                        'full_name': contact.FullName,
                        'email1': contact.Email1Address,
                        'email2': contact.Email2Address,
                        'business_phone': contact.BusinessTelephoneNumber,
                        'mobile_phone': contact.MobileTelephoneNumber,
                        'company': contact.CompanyName,
                        'job_title': contact.JobTitle
                    }
                    contact_list.append(contact_info)
                except Exception as e:
                    continue
            
            print(f"Retrieved {len(contact_list)} contacts")
            return contact_list
        except Exception as e:
            print(f"Error retrieving contacts: {str(e)}")
            return []
    
    def get_folder_list(self) -> List[str]:
        """Get list of all available folders"""
        try:
            folders = []
            
            # Default folders
            default_folders = {
                "Inbox": 6,
                "Sent Items": 5,
                "Drafts": 16,
                "Deleted Items": 3,
                "Outbox": 4,
                "Calendar": 9,
                "Contacts": 10,
                "Tasks": 13,
                "Notes": 12
            }
            
            for folder_name in default_folders.keys():
                folders.append(folder_name)
            
            # Custom folders in Inbox
            try:
                for folder in self.inbox.Folders:
                    folders.append(folder.Name)
            except:
                pass
            
            return folders
        except Exception as e:
            print(f"Error getting folder list: {str(e)}")
            return []
    
    def get_unread_count(self, folder_name: str = "Inbox") -> int:
        """Get count of unread emails in a folder"""
        try:
            folder = self._get_folder(folder_name)
            if folder:
                return folder.UnReadItemCount
            return 0
        except Exception as e:
            print(f"Error getting unread count: {str(e)}")
            return 0
    
    def _get_folder(self, folder_name: str):
        """Helper method to get folder by name"""
        folder_map = {
            "Inbox": 6,
            "Sent Items": 5,
            "Drafts": 16,
            "Deleted Items": 3,
            "Outbox": 4,
            "Calendar": 9,
            "Contacts": 10,
            "Tasks": 13,
            "Notes": 12
        }
        
        try:
            if folder_name in folder_map:
                return self.namespace.GetDefaultFolder(folder_map[folder_name])
            else:
                # Try to find custom folder
                for folder in self.inbox.Folders:
                    if folder.Name == folder_name:
                        return folder
                print(f"Folder '{folder_name}' not found")
                return None
        except Exception as e:
            print(f"Error accessing folder '{folder_name}': {str(e)}")
            return None
    
    def _get_attachment_info(self, message) -> List[Dict]:
        """Helper method to get attachment information"""
        attachments = []
        try:
            for attachment in message.Attachments:
                att_info = {
                    'filename': attachment.FileName,
                    'size': attachment.Size,
                    'type': attachment.Type
                }
                attachments.append(att_info)
        except:
            pass
        return attachments
    
    def get_outlook_info(self):
        """Display Outlook connection information"""
        try:
            print("=== Outlook Information ===")
            print(f"Outlook Version: {self.outlook.Version}")
            print(f"Current User: {self.namespace.CurrentUser.Name}")
            print(f"Inbox Unread Count: {self.get_unread_count('Inbox')}")
            print(f"Available Folders: {', '.join(self.get_folder_list())}")
        except Exception as e:
            print(f"Error getting Outlook info: {str(e)}")

# Example usage
if __name__ == "__main__":
    # Create an instance of OutlookHandler
    outlook_handler = OutlookHandler()
    
    # Get Outlook information
    outlook_handler.get_outlook_info()
    
    # Example: Send an email
    # outlook_handler.send_email(
    #     to_recipients="recipient@example.com",
    #     subject="Test Email from Python",
    #     body="This is a test email sent from Python using OutlookHandler class.",
    #     send_immediately=False  # Save to drafts instead of sending
    # )
    
    # Example: Get recent emails
    recent_emails = outlook_handler.get_emails(count=5)
    for email in recent_emails:
        print(f"Subject: {email['subject']}")
        print(f"From: {email['sender']}")
        print(f"Received: {email['received_time']}")
        print("-" * 50)
    
    # Example: Search emails
    # search_results = outlook_handler.search_emails("meeting", search_in="subject")
    # print(f"Found {len(search_results)} emails with 'meeting' in subject")
    
    # Example: Get unread email count
    unread_count = outlook_handler.get_unread_count()
    print(f"Unread emails in Inbox: {unread_count}")