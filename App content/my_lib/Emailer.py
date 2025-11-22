import win32com.client as win32
from pywintypes import com_error
import os
from datetime import datetime

class Emailer:
    """A simple class to send emails using the local Outlook client."""

    def send_outlook_notification(
        self,
        recipient_email: str,
        bot_name: str,
        status: str,
        error_message: str = ""
    ) -> bool:
        """
        Sends a concise notification email for a scheduled bot run.
        This function is now resilient and will not crash if Outlook is not installed.

        Args:
            recipient_email (str): The email address to send the notification to.
            bot_name (str): The name of the bot that was executed.
            status (str): The final status ('Completed Successfully' or 'Failed with Error').
            error_message (str, optional): The specific error message if the bot failed.

        Returns:
            bool: True if the email was sent successfully, False otherwise.
        """
        if not recipient_email:
            print("Email not sent: No recipient email provided.")
            return False

        try:
            # --- ATTEMPT TO CONNECT TO OUTLOOK ---
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            
            mail.To = recipient_email
            
            # --- Email Subject ---
            computer_name = os.environ.get('COMPUTERNAME', 'Unknown PC')
            if status == "Failed with Error":
                mail.Subject = f"❌ BOT FAILED: '{bot_name}' on {computer_name}"
            else:
                mail.Subject = f"✅ BOT COMPLETED: '{bot_name}' on {computer_name}"

            # --- Email Body (HTML for better formatting) ---
            run_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            body = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: Calibri, sans-serif; font-size: 11pt; }}
                    .status-success {{ color: #107C10; font-weight: bold; }}
                    .status-fail {{ color: #D83B01; font-weight: bold; }}
                    .error-box {{
                        background-color: #FFF4F4;
                        border: 1px solid #D83B01;
                        padding: 10px;
                        font-family: Consolas, 'Courier New', monospace;
                        color: #D83B01;
                        font-weight: bold;
                    }}
                </style>
            </head>
            <body>
                <p>The scheduled bot <strong>'{bot_name}'</strong> has finished running.</p>
                <p><strong>Machine:</strong> {computer_name}</p>
                <p><strong>Time:</strong> {run_time}</p>
            """

            if status == "Failed with Error":
                body += f"<p><strong>Status:</strong> <span class='status-fail'>{status}</span></p>"
                if error_message:
                    body += "<h3>Error Details:</h3>"
                    body += f"<div class='error-box'>{error_message}</div>"
            else:
                body += f"<p><strong>Status:</strong> <span class='status-success'>{status}</span></p>"

            body += "<hr><p style='font-size: 9pt; color: #888;'>This is an automated notification from the AutomateTask application.</p>"
            body += "</body></html>"
            
            mail.HTMLBody = body
            mail.Send()
            print(f"Notification email sent to {recipient_email} for bot '{bot_name}'.")
            return True # Indicate success

        except com_error as e:
            # --- THIS IS THE KEY CHANGE ---
            # This specific error occurs if Outlook is not installed or running.
            print("WARNING: Could not send Outlook email. Outlook may not be installed or configured.")
            print(f" (Details: {e})")
            return False # Indicate failure

        except Exception as e:
            # Catch any other unexpected errors during email creation/sending.
            print(f"CRITICAL: An unexpected error occurred while trying to send an Outlook email. Error: {e}")
            return False # Indicate failure
