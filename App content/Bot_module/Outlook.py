import win32com.client
import os
import datetime
import pandas as pd
from typing import List, Dict, Any, Optional, Union
import time
from my_lib.shared_context import ExecutionContext as Context

class OutlookHandler:
    def __init__(self, context: Context):
        self.context = context
        self.outlook = None
        self.namespace = None
        self.inbox = None
        self.sent_items = None
        self.drafts = None
    
    def connect(self):
        """Connect to Outlook application"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Get common folders
            self.inbox = self.namespace.GetDefaultFolder(6)  # Inbox
            self.sent_items = self.namespace.GetDefaultFolder(5)  # Sent Items
            self.drafts = self.namespace.GetDefaultFolder(16)  # Drafts
            
            self.context.add_log("Successfully connected to Outlook")
            return self.outlook
        except Exception as e:
            self.context.add_log(f"Error connecting to Outlook: {str(e)}")
            return None
    
    def send_email(self, outlook, to_recipients: Union[str, List[str]], subject: str, body: str,
                   cc_recipients: Union[str, List[str]] = None, 
                   bcc_recipients: Union[str, List[str]] = None,
                   attachments: List[str] = None, 
                   html_body: bool = False,
                   send_immediately: bool = True):
        """
        Send an email
        
        Args:
            outlook: Outlook application object
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
            mail = outlook.CreateItem(0)  # 0 = Mail item
            
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
                        self.context.add_log(f"Attached: {attachment}")
                    else:
                        self.context.add_log(f"Warning: Attachment not found: {attachment}")
            
            # Send or save to drafts
            if send_immediately:
                mail.Send()
                self.context.add_log("Email sent successfully")
            else:
                mail.Save()
                self.context.add_log("Email saved to drafts")
            
            return True
        except Exception as e:
            self.context.add_log(f"Error sending email: {str(e)}")
            return False
    
    def read_unread_emails(self, outlook, folder_name: str = "Inbox", folder_path: Optional[str] = None) -> pd.DataFrame:
        """
        Read unread emails - version that keeps datetime as strings to avoid conversion errors
        
        Args:
            outlook: Outlook application object
            folder_name: Name of the folder (e.g., "Inbox", "Sent Items", or custom folder name)
            folder_path: Full path to folder for nested folders (e.g., "Inbox\\Custom Folder")
                        If provided, folder_name is ignored
        
        Returns:
            DataFrame with email data (datetime fields as strings)
        """
        try:
            if not outlook:
                self.context.add_log("Outlook object is None. Call connect() first.")
                return pd.DataFrame()
            
            namespace = outlook.GetNamespace("MAPI")
            target_folder = self._get_folder(namespace, folder_name, folder_path)
            
            if not target_folder:
                return pd.DataFrame()
            
            messages = target_folder.Items
            messages = messages.Restrict("[Unread] = True")
            
            # Try to sort, but don't fail if it doesn't work
            try:
                messages.Sort("[ReceivedTime]", True)
            except:
                pass
            
            email_data = []
            
            for message in messages:
                try:
                    email_info = {
                        'Subject': str(getattr(message, 'Subject', '') or ''),
                        'Sender': str(getattr(message, 'SenderName', '') or ''),
                        'SenderEmail': str(getattr(message, 'SenderEmailAddress', '') or ''),
                        'ReceivedTime_String': str(getattr(message, 'ReceivedTime', '') or ''),
                        'Body': str(getattr(message, 'Body', '') or ''),
                        'HTMLBody': str(getattr(message, 'HTMLBody', '') or ''),
                        'HasAttachments': self._safe_has_attachments(message),
                        'AttachmentCount': self._safe_attachment_count(message),
                        'Importance': str(self._get_importance_text(getattr(message, 'Importance', 1))),
                        'Size': int(getattr(message, 'Size', 0) or 0),
                        'Categories': str(getattr(message, 'Categories', '') or ''),
                        'EntryID': str(getattr(message, 'EntryID', '') or ''),
                        'To': str(getattr(message, 'To', '') or ''),
                        'CC': str(getattr(message, 'CC', '') or ''),
                        'Folder': str(target_folder.Name),
                        'CreationTime_String': str(getattr(message, 'CreationTime', '') or ''),
                        'MessageClass': str(getattr(message, 'MessageClass', '') or ''),
                        'AttachmentNames': self._safe_attachment_names(message)
                    }
                    
                    email_data.append(email_info)
                    
                except Exception as e:
                    self.context.add_log(f"Error processing message: {str(e)}")
                    continue
            
            df = pd.DataFrame(email_data)
            self.context.add_log(f"Retrieved {len(df)} unread emails from folder: {target_folder.Name}")
            return df
            
        except Exception as e:
            self.context.add_log(f"Error reading unread emails: {str(e)}")
            return pd.DataFrame()
    
    def _get_folder(self, namespace, folder_name: str, folder_path: Optional[str] = None):
        """
        Get folder object by name or path
        """
        try:
            if folder_path:
                return self._get_folder_by_path(namespace, folder_path)
            
            # Check default folders first
            if folder_name.lower() == "inbox":
                return namespace.GetDefaultFolder(6)  # Inbox
            elif folder_name.lower() == "sent items":
                return namespace.GetDefaultFolder(5)  # Sent Items
            elif folder_name.lower() == "drafts":
                return namespace.GetDefaultFolder(16)  # Drafts
            else:
                # Search for custom folder
                return self._find_folder_by_name(namespace, folder_name)
                
        except Exception as e:
            self.context.add_log(f"Error getting folder '{folder_name}': {str(e)}")
            return None
    
    def _get_folder_by_path(self, namespace, folder_path: str):
        """
        Get folder by full path (e.g., "Inbox\\Custom Folder\\Subfolder")
        """
        try:
            folder_parts = folder_path.split('\\')
            
            # If only one part, it's a root-level folder
            if len(folder_parts) == 1:
                folder_name = folder_parts[0]
                
                if folder_name.lower() == "inbox":
                    return namespace.GetDefaultFolder(6)
                elif folder_name.lower() == "sent items":
                    return namespace.GetDefaultFolder(5)
                elif folder_name.lower() == "drafts":
                    return namespace.GetDefaultFolder(16)
                else:
                    return self._find_root_folder_by_name(namespace, folder_name)
            
            # Multiple parts - handle nested folders
            current_folder = None
            
            if folder_parts[0].lower() == "inbox":
                current_folder = namespace.GetDefaultFolder(6)
            elif folder_parts[0].lower() == "sent items":
                current_folder = namespace.GetDefaultFolder(5)
            elif folder_parts[0].lower() == "drafts":
                current_folder = namespace.GetDefaultFolder(16)
            else:
                current_folder = self._find_root_folder_by_name(namespace, folder_parts[0])
            
            if not current_folder:
                self.context.add_log(f"Root folder '{folder_parts[0]}' not found")
                return None
            
            # Navigate through subfolders
            for folder_part in folder_parts[1:]:
                found_subfolder = None
                for subfolder in current_folder.Folders:
                    if subfolder.Name.lower() == folder_part.lower():
                        found_subfolder = subfolder
                        break
                
                if found_subfolder:
                    current_folder = found_subfolder
                else:
                    self.context.add_log(f"Subfolder '{folder_part}' not found in '{current_folder.Name}'")
                    return None
            
            return current_folder
            
        except Exception as e:
            self.context.add_log(f"Error navigating folder path '{folder_path}': {str(e)}")
            return None
    
    def _find_folder_by_name(self, namespace, folder_name: str):
        """
        Recursively search for folder by name
        """
        try:
            # Search in default folders first
            folders_to_search = []
            try:
                folders_to_search.append(namespace.GetDefaultFolder(6))   # Inbox
                folders_to_search.append(namespace.GetDefaultFolder(5))   # Sent Items
                folders_to_search.append(namespace.GetDefaultFolder(16))  # Drafts
            except:
                pass
            
            # Add root folders
            try:
                for folder in namespace.Folders:
                    folders_to_search.append(folder)
            except:
                pass
            
            for folder in folders_to_search:
                try:
                    found_folder = self._search_folder_recursive(folder, folder_name)
                    if found_folder:
                        return found_folder
                except:
                    continue
            
            self.context.add_log(f"Folder '{folder_name}' not found")
            return None
            
        except Exception as e:
            self.context.add_log(f"Error searching for folder '{folder_name}': {str(e)}")
            return None
    
    def _find_root_folder_by_name(self, namespace, folder_name: str):
        """
        Find folder at root level (same level as Inbox)
        """
        try:
            for folder in namespace.Folders:
                if folder.Name.lower() == folder_name.lower():
                    return folder
            
            self.context.add_log(f"Root folder '{folder_name}' not found")
            return None
            
        except Exception as e:
            self.context.add_log(f"Error searching for root folder '{folder_name}': {str(e)}")
            return None
    
    def _search_folder_recursive(self, parent_folder, target_name: str):
        """
        Recursively search for folder in subfolders
        """
        try:
            # Check current folder
            if parent_folder.Name.lower() == target_name.lower():
                return parent_folder
            
            # Search in subfolders
            for subfolder in parent_folder.Folders:
                if subfolder.Name.lower() == target_name.lower():
                    return subfolder
                
                # Recursive search
                found = self._search_folder_recursive(subfolder, target_name)
                if found:
                    return found
            
            return None
            
        except Exception as e:
            return None
    
    def _safe_has_attachments(self, message):
        """
        Safely check if message has attachments
        """
        try:
            attachments = getattr(message, 'Attachments', None)
            if attachments is None:
                return False
            return len(attachments) > 0
        except:
            return False
    
    def _safe_attachment_count(self, message):
        """
        Safely get attachment count
        """
        try:
            attachments = getattr(message, 'Attachments', None)
            if attachments is None:
                return 0
            return len(attachments)
        except:
            return 0
    
    def _safe_attachment_names(self, message):
        """
        Safely get attachment names
        """
        try:
            if not self._safe_has_attachments(message):
                return ''
            
            attachment_names = []
            for attachment in message.Attachments:
                try:
                    attachment_names.append(str(attachment.FileName))
                except:
                    attachment_names.append('[Unknown filename]')
            return '; '.join(attachment_names)
        except:
            return 'Error reading attachments'
    
    def _get_importance_text(self, importance_level):
        """
        Convert importance level to text
        """
        try:
            importance_map = {
                0: "Low",
                1: "Normal", 
                2: "High"
            }
            return importance_map.get(int(importance_level), "Normal")
        except:
            return "Normal"
                
    def save_attachments(self, outlook, entry_id: str, save_folder: str, 
                        attachment_names: Optional[List[str]] = None) -> List[str]:
        """
        Save attachments from a specific email by EntryID
        
        Args:
            outlook: Outlook application object
            entry_id: EntryID of the email (from DataFrame)
            save_folder: Folder path where to save attachments
            attachment_names: List of specific attachment names to save (None = save all)
        
        Returns:
            List of saved file paths
        """
        try:
            if not outlook:
                self.context.add_log("Outlook object is None")
                return []
            
            # Create save folder if it doesn't exist
            if not os.path.exists(save_folder):
                os.makedirs(save_folder)
                self.context.add_log(f"Created folder: {save_folder}")
            
            namespace = outlook.GetNamespace("MAPI")
            
            # Get the email by EntryID
            try:
                message = namespace.GetItemFromID(entry_id)
            except Exception as e:
                self.context.add_log(f"Could not find email with EntryID: {entry_id}")
                return []
            
            if not message.Attachments or len(message.Attachments) == 0:
                self.context.add_log("No attachments found in this email")
                return []
            
            saved_files = []
            
            for attachment in message.Attachments:
                try:
                    filename = attachment.FileName
                    
                    # Skip if specific attachment names are specified and this isn't one of them
                    if attachment_names and filename not in attachment_names:
                        continue
                    
                    # Create safe filename
                    safe_filename = self._make_safe_filename(filename)
                    file_path = os.path.join(save_folder, safe_filename)
                    
                    # Handle duplicate filenames
                    file_path = self._get_unique_filepath(file_path)
                    
                    # Save the attachment
                    attachment.SaveAsFile(file_path)
                    saved_files.append(file_path)
                    
                    self.context.add_log(f"Saved attachment: {filename} -> {file_path}")
                    
                except Exception as e:
                    self.context.add_log(f"Error saving attachment {filename}: {str(e)}")
                    continue
            
            self.context.add_log(f"Successfully saved {len(saved_files)} attachments")
            return saved_files
            
        except Exception as e:
            self.context.add_log(f"Error saving attachments: {str(e)}")
            return []

    def save_attachments_from_dataframe(self, outlook, df: pd.DataFrame, save_folder: str,
                                       filter_extensions: Optional[List[str]] = None,
                                       max_size_mb: Optional[float] = None) -> Dict[str, List[str]]:
        """
        Save attachments from multiple emails in DataFrame
        
        Args:
            outlook: Outlook application object
            df: DataFrame from read_unread_emails
            save_folder: Base folder where to save attachments
            filter_extensions: List of file extensions to save (e.g., ['.pdf', '.docx'])
            max_size_mb: Maximum file size in MB to save
        
        Returns:
            Dictionary with EntryID as key and list of saved files as value
        """
        try:
            if df.empty:
                self.context.add_log("DataFrame is empty")
                return {}
            
            # Create save folder if it doesn't exist
            if not os.path.exists(save_folder):
                os.makedirs(save_folder)
            
            results = {}
            emails_with_attachments = df[df['HasAttachments'] == True]
            
            self.context.add_log(f"Processing {len(emails_with_attachments)} emails with attachments")
            
            for index, row in emails_with_attachments.iterrows():
                try:
                    entry_id = row['EntryID']
                    subject = row['Subject'][:50]  # Truncate for logging
                    
                    self.context.add_log(f"Processing email: {subject}...")
                    
                    # Create subfolder for this email
                    email_folder = os.path.join(save_folder, f"Email_{index}_{self._make_safe_filename(subject)}")
                    
                    saved_files = self._save_attachments_with_filters(
                        outlook, entry_id, email_folder, filter_extensions, max_size_mb
                    )
                    
                    if saved_files:
                        results[entry_id] = saved_files
                    
                except Exception as e:
                    self.context.add_log(f"Error processing email at index {index}: {str(e)}")
                    continue
            
            total_files = sum(len(files) for files in results.values())
            self.context.add_log(f"Total attachments saved: {total_files}")
            
            return results
            
        except Exception as e:
            self.context.add_log(f"Error in save_attachments_from_dataframe: {str(e)}")
            return {}

    def _save_attachments_with_filters(self, outlook, entry_id: str, save_folder: str,
                                      filter_extensions: Optional[List[str]] = None,
                                      max_size_mb: Optional[float] = None) -> List[str]:
        """
        Save attachments with filtering options
        """
        try:
            if not os.path.exists(save_folder):
                os.makedirs(save_folder)
            
            namespace = outlook.GetNamespace("MAPI")
            message = namespace.GetItemFromID(entry_id)
            
            if not message.Attachments or len(message.Attachments) == 0:
                return []
            
            saved_files = []
            
            for attachment in message.Attachments:
                try:
                    filename = attachment.FileName
                    file_ext = os.path.splitext(filename)[1].lower()
                    
                    # Filter by extension
                    if filter_extensions:
                        if file_ext not in [ext.lower() for ext in filter_extensions]:
                            self.context.add_log(f"Skipped {filename} - extension not in filter")
                            continue
                    
                    # Filter by size
                    if max_size_mb:
                        try:
                            # Get attachment size (in bytes)
                            attachment_size = attachment.Size
                            size_mb = attachment_size / (1024 * 1024)
                            
                            if size_mb > max_size_mb:
                                self.context.add_log(f"Skipped {filename} - size {size_mb:.1f}MB exceeds limit")
                                continue
                        except:
                            # If we can't get size, save anyway
                            pass
                    
                    # Save the attachment
                    safe_filename = self._make_safe_filename(filename)
                    file_path = os.path.join(save_folder, safe_filename)
                    file_path = self._get_unique_filepath(file_path)
                    
                    attachment.SaveAsFile(file_path)
                    saved_files.append(file_path)
                    
                    self.context.add_log(f"Saved: {filename}")
                    
                except Exception as e:
                    self.context.add_log(f"Error saving attachment {filename}: {str(e)}")
                    continue
            
            return saved_files
            
        except Exception as e:
            self.context.add_log(f"Error in _save_attachments_with_filters: {str(e)}")
            return []

    def _make_safe_filename(self, filename: str) -> str:
        """
        Make filename safe for Windows filesystem
        """
        try:
            # Remove or replace invalid characters
            invalid_chars = '<>:"/\\|?*'
            safe_filename = filename
            
            for char in invalid_chars:
                safe_filename = safe_filename.replace(char, '_')
            
            # Remove leading/trailing spaces and dots
            safe_filename = safe_filename.strip(' .')
            
            # Ensure filename is not empty
            if not safe_filename:
                safe_filename = "attachment"
            
            # Limit length
            if len(safe_filename) > 200:
                name, ext = os.path.splitext(safe_filename)
                safe_filename = name[:200-len(ext)] + ext
            
            return safe_filename
            
        except Exception:
            return "attachment.txt"

    def _get_unique_filepath(self, file_path: str) -> str:
        """
        Get unique file path if file already exists
        """
        try:
            if not os.path.exists(file_path):
                return file_path
            
            base_path, extension = os.path.splitext(file_path)
            counter = 1
            
            while os.path.exists(f"{base_path}_{counter}{extension}"):
                counter += 1
            
            return f"{base_path}_{counter}{extension}"
            
        except Exception:
            return file_path

    def get_attachment_info(self, outlook, entry_id: str) -> List[Dict[str, Any]]:
        """
        Get detailed information about attachments without saving them
        
        Args:
            outlook: Outlook application object
            entry_id: EntryID of the email
        
        Returns:
            List of dictionaries with attachment information
        """
        try:
            if not outlook:
                return []
            
            namespace = outlook.GetNamespace("MAPI")
            message = namespace.GetItemFromID(entry_id)
            
            if not message.Attachments or len(message.Attachments) == 0:
                return []
            
            attachment_info = []
            
            for i, attachment in enumerate(message.Attachments):
                try:
                    info = {
                        'Index': i + 1,
                        'FileName': getattr(attachment, 'FileName', f'Attachment_{i+1}'),
                        'Size': getattr(attachment, 'Size', 0),
                        'SizeMB': round(getattr(attachment, 'Size', 0) / (1024 * 1024), 2),
                        'Type': getattr(attachment, 'Type', 'Unknown'),
                        'Extension': os.path.splitext(getattr(attachment, 'FileName', ''))[1].lower()
                    }
                    attachment_info.append(info)
                    
                except Exception as e:
                    self.context.add_log(f"Error getting info for attachment {i+1}: {str(e)}")
                    continue
            
            return attachment_info
            
        except Exception as e:
            self.context.add_log(f"Error getting attachment info: {str(e)}")
            return []
            
    def read_emails_by_period(self, outlook, folder_name: str = "Inbox", folder_path: Optional[str] = None,
                             start_date: Optional[Union[str, datetime.datetime]] = None,
                             end_date: Optional[Union[str, datetime.datetime]] = None,
                             days_back: Optional[int] = None,
                             include_read: bool = True,
                             include_unread: bool = True) -> pd.DataFrame:
        """
        Read all emails (read and unread) from a specific folder within a time period
        
        Args:
            outlook: Outlook application object
            folder_name: Name of the folder (e.g., "Inbox", "Sent Items", or custom folder name)
            folder_path: Full path to folder for nested folders (e.g., "Inbox\\Custom Folder")
            start_date: Start date (string 'YYYY-MM-DD' or datetime object)
            end_date: End date (string 'YYYY-MM-DD' or datetime object)
            days_back: Number of days back from today (alternative to start_date/end_date)
            include_read: Include read emails
            include_unread: Include unread emails
        
        Returns:
            DataFrame with email data (datetime fields as strings)
        """
        try:
            if not outlook:
                self.context.add_log("Outlook object is None. Call connect() first.")
                return pd.DataFrame()
            
            namespace = outlook.GetNamespace("MAPI")
            target_folder = self._get_folder(namespace, folder_name, folder_path)
            
            if not target_folder:
                return pd.DataFrame()
            
            # Calculate date range
            start_dt, end_dt = self._calculate_date_range(start_date, end_date, days_back)
            
            if not start_dt or not end_dt:
                self.context.add_log("Invalid date range specified")
                return pd.DataFrame()
            
            self.context.add_log(f"Searching emails from {start_dt.strftime('%Y-%m-%d')} to {end_dt.strftime('%Y-%m-%d')}")
            
            # Build filter string for date range and read status
            filter_parts = []
            
            # Date filter
            start_str = start_dt.strftime('%m/%d/%Y %H:%M %p')
            end_str = end_dt.strftime('%m/%d/%Y %H:%M %p')
            filter_parts.append(f"[ReceivedTime] >= '{start_str}'")
            filter_parts.append(f"[ReceivedTime] <= '{end_str}'")
            
            # Read status filter
            read_filters = []
            if include_read:
                read_filters.append("[Unread] = False")
            if include_unread:
                read_filters.append("[Unread] = True")
            
            if read_filters:
                filter_parts.append(f"({' OR '.join(read_filters)})")
            
            filter_string = ' AND '.join(filter_parts)
            
            self.context.add_log(f"Using filter: {filter_string}")
            
            # Get messages with filter
            messages = target_folder.Items
            messages = messages.Restrict(filter_string)
            
            # Sort by received time (newest first)
            try:
                messages.Sort("[ReceivedTime]", True)
            except:
                pass
            
            email_data = []
            
            for message in messages:
                try:
                    email_info = {
                        'Subject': str(getattr(message, 'Subject', '') or ''),
                        'Sender': str(getattr(message, 'SenderName', '') or ''),
                        'SenderEmail': str(getattr(message, 'SenderEmailAddress', '') or ''),
                        'ReceivedTime_String': str(getattr(message, 'ReceivedTime', '') or ''),
                        'Body': str(getattr(message, 'Body', '') or ''),
                        'HTMLBody': str(getattr(message, 'HTMLBody', '') or ''),
                        'HasAttachments': self._safe_has_attachments(message),
                        'AttachmentCount': self._safe_attachment_count(message),
                        'Importance': str(self._get_importance_text(getattr(message, 'Importance', 1))),
                        'Size': int(getattr(message, 'Size', 0) or 0),
                        'Categories': str(getattr(message, 'Categories', '') or ''),
                        'EntryID': str(getattr(message, 'EntryID', '') or ''),
                        'To': str(getattr(message, 'To', '') or ''),
                        'CC': str(getattr(message, 'CC', '') or ''),
                        'BCC': str(getattr(message, 'BCC', '') or ''),
                        'Folder': str(target_folder.Name),
                        'CreationTime_String': str(getattr(message, 'CreationTime', '') or ''),
                        'MessageClass': str(getattr(message, 'MessageClass', '') or ''),
                        'AttachmentNames': self._safe_attachment_names(message),
                        'IsRead': not bool(getattr(message, 'Unread', False)),
                        'IsUnread': bool(getattr(message, 'Unread', False)),
                        'ConversationTopic': str(getattr(message, 'ConversationTopic', '') or ''),
                        'SentOn_String': str(getattr(message, 'SentOn', '') or '')
                    }
                    
                    email_data.append(email_info)
                    
                except Exception as e:
                    self.context.add_log(f"Error processing message: {str(e)}")
                    continue
            
            df = pd.DataFrame(email_data)
            self.context.add_log(f"Retrieved {len(df)} emails from folder: {target_folder.Name}")
            
            # Add summary info
            if not df.empty:
                read_count = len(df[df['IsRead'] == True])
                unread_count = len(df[df['IsUnread'] == True])
                self.context.add_log(f"  - Read emails: {read_count}")
                self.context.add_log(f"  - Unread emails: {unread_count}")
            
            return df
            
        except Exception as e:
            self.context.add_log(f"Error reading emails by period: {str(e)}")
            return pd.DataFrame()

    def _calculate_date_range(self, start_date, end_date, days_back):
        """
        Calculate start and end datetime objects from various input formats
        """
        try:
            if days_back is not None:
                # Use days_back to calculate range
                end_dt = datetime.datetime.now()
                start_dt = end_dt - datetime.timedelta(days=days_back)
                return start_dt, end_dt
            
            # Parse start_date
            if isinstance(start_date, str):
                start_dt = datetime.datetime.strptime(start_date, '%Y-%m-%d')
            elif isinstance(start_date, datetime.datetime):
                start_dt = start_date
            elif isinstance(start_date, datetime.date):
                start_dt = datetime.datetime.combine(start_date, datetime.time.min)
            else:
                start_dt = datetime.datetime.now() - datetime.timedelta(days=7)  # Default 7 days back
            
            # Parse end_date
            if isinstance(end_date, str):
                end_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d')
                end_dt = end_dt.replace(hour=23, minute=59, second=59)  # End of day
            elif isinstance(end_date, datetime.datetime):
                end_dt = end_date
            elif isinstance(end_date, datetime.date):
                end_dt = datetime.datetime.combine(end_date, datetime.time.max)
            else:
                end_dt = datetime.datetime.now()  # Default to now
            
            return start_dt, end_dt
            
        except Exception as e:
            self.context.add_log(f"Error calculating date range: {str(e)}")
            return None, None

    def read_emails_last_n_days(self, outlook, folder_name: str = "Inbox", folder_path: Optional[str] = None,
                               days: int = 7, include_read: bool = True, include_unread: bool = True) -> pd.DataFrame:
        """
        Convenience method to read emails from last N days
        
        Args:
            outlook: Outlook application object
            folder_name: Name of the folder
            folder_path: Full path to folder for nested folders
            days: Number of days back from today
            include_read: Include read emails
            include_unread: Include unread emails
        
        Returns:
            DataFrame with email data
        """
        return self.read_emails_by_period(
            outlook=outlook,
            folder_name=folder_name,
            folder_path=folder_path,
            days_back=days,
            include_read=include_read,
            include_unread=include_unread
        )

    def read_emails_today(self, outlook, folder_name: str = "Inbox", folder_path: Optional[str] = None,
                         include_read: bool = True, include_unread: bool = True) -> pd.DataFrame:
        """
        Convenience method to read emails from today only
        """
        today = datetime.date.today()
        return self.read_emails_by_period(
            outlook=outlook,
            folder_name=folder_name,
            folder_path=folder_path,
            start_date=today.strftime('%Y-%m-%d'),
            end_date=today.strftime('%Y-%m-%d'),
            include_read=include_read,
            include_unread=include_unread
        )

    def read_emails_this_week(self, outlook, folder_name: str = "Inbox", folder_path: Optional[str] = None,
                             include_read: bool = True, include_unread: bool = True) -> pd.DataFrame:
        """
        Convenience method to read emails from this week (Monday to Sunday)
        """
        today = datetime.date.today()
        start_of_week = today - datetime.timedelta(days=today.weekday())  # Monday
        end_of_week = start_of_week + datetime.timedelta(days=6)  # Sunday
        
        return self.read_emails_by_period(
            outlook=outlook,
            folder_name=folder_name,
            folder_path=folder_path,
            start_date=start_of_week.strftime('%Y-%m-%d'),
            end_date=end_of_week.strftime('%Y-%m-%d'),
            include_read=include_read,
            include_unread=include_unread
        )

    def read_emails_this_month(self, outlook, folder_name: str = "Inbox", folder_path: Optional[str] = None,
                              include_read: bool = True, include_unread: bool = True) -> pd.DataFrame:
        """
        Convenience method to read emails from this month
        """
        today = datetime.date.today()
        start_of_month = today.replace(day=1)
        
        # Get last day of month
        if today.month == 12:
            end_of_month = today.replace(year=today.year + 1, month=1, day=1) - datetime.timedelta(days=1)
        else:
            end_of_month = today.replace(month=today.month + 1, day=1) - datetime.timedelta(days=1)
        
        return self.read_emails_by_period(
            outlook=outlook,
            folder_name=folder_name,
            folder_path=folder_path,
            start_date=start_of_month.strftime('%Y-%m-%d'),
            end_date=end_of_month.strftime('%Y-%m-%d'),
            include_read=include_read,
            include_unread=include_unread
        )