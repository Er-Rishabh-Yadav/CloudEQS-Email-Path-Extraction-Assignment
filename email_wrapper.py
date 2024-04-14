import email
import imaplib
import os
import re

class EmailWrapper:
    def __init__(self, email, password):
        self.email = email
        self.password = password
        self.imap_server = "imap.gmail.com"
        self.imap = imaplib.IMAP4_SSL(self.imap_server)
        self.logged_in = False

    def login(self):
        if self.password is None:
            print("Error: Password is None. Please provide a valid password.")
            return

        try:
            # authenticate
            self.imap.login(self.email, self.password)
            self.logged_in = True
            print("Logged in successfully!")
        except imaplib.IMAP4.error as e:
            print("Login failed. Check your email and password.")

    def fetch_emails(self, criteria=None, num_emails=1):
        """
        Fetch emails based on the provided criteria.

        :param criteria: List of IMAP search criteria as strings.
        :param num_emails: Number of emails to fetch.
        :return: List of dictionaries with email data.
        """
        if not self.logged_in:
            print("Please login first.")
            return []

        emails = []

        try:
            # select INBOX
            self.imap.select("INBOX")

            if criteria is None:
                criteria = ['ALL']  # Default criteria is to fetch all emails

            # Combine the criteria into a single string
            search_criteria = ' '.join(criteria)

            # search for emails based on the provided criteria
            typ, data = self.imap.search(None, search_criteria)
            for num in data[0].split()[-num_emails:]:
                typ, msg_data = self.imap.fetch(num, '(RFC822)')
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        email_data = {}
                        msg = email.message_from_bytes(response_part[1])
                        email_data["From"] = msg["From"]
                        email_data["Subject"] = msg["Subject"]
                        # Get the email body
                        body = self.get_email_body(msg)
                        email_data["Body"] = body

                        # Extract attachments
                        attachments = self.extract_attachments(msg)
                        email_data["Attachments"] = attachments

                        emails.append(email_data)
        except Exception as e:
            print("Error fetching emails:", e)

        return emails

    def get_email_body(self, msg):
        """
        Get the email body as plain text
        """
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    body = part.get_payload(decode=True).decode()
                    break
        else:
            body = msg.get_payload(decode=True).decode()

        return body
   

    def extract_paths_from_body(self, emails):
        """
        Get the email body as plain text and extract paths from the body.

        Args:
            emails (list): List of email message objects or dictionaries.

        Returns:
            list: List of paths extracted from email bodies.
        """
        paths = []

        for item in emails:
            if isinstance(item, dict):
                body = item.get("Body", "")
            elif isinstance(item, email.message.Message):
                body = ""
                if item.is_multipart():
                    for part in item.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            body = part.get_payload(decode=True).decode()
                            break
                else:
                    body = item.get_payload(decode=True).decode()
            else:
                continue

            # Extract paths from the body
            
            extracted_paths = re.findall(r'(PALSFTPHOME/\w+/CLOUDEQS_[\d-]+\.txt)', body)

            paths.extend(extracted_paths)

        return paths


    def extract_attachments(self, msg):
        """
        Extract attachments from the email.
        """
        attachments = []

        for part in msg.walk():
            content_type = part.get_content_type()
            if "text/plain" in content_type or "application" in content_type:
                # This part is an attachment
                filename = part.get_filename()
                if filename:
                    attachment = {
                        "filename": filename,
                        "data": part.get_payload(decode=True)
                    }
                    attachments.append(attachment)

        return attachments

    def get_attachment_file_paths(self, emails):
        """
        Get file paths of attachments from a list of email data.
        """
        file_paths = []

        for email_data in emails:
            for attachment in email_data.get("Attachments", []):
                filename = attachment["filename"]
                data = attachment["data"]
                file_path = os.path.join("attachments", filename)

                with open(file_path, "wb") as f:
                    f.write(data)
                    print(f"Saved attachment '{filename}' to '{file_path}'")

                file_paths.append(file_path)

        return file_paths

    def logout(self):
        if self.logged_in:
            try:
                self.imap.logout()
                self.logged_in = False
                print("Logged out successfully.")
            except Exception as e:
                print("Error logging out:", e)
        else:
            print("Not logged in.")
