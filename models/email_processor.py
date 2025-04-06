# models/email_processor.py
import re
from email.parser import BytesParser
from email.policy import default
from bs4 import BeautifulSoup
import logging

logger = logging.getLogger(__name__)

class EmailProcessor:
    """Process and parse email data"""
    
    def __init__(self):
        self.parser = BytesParser(policy=default)
    
    def parse_email(self, raw_email):
        """Parse raw email into structured format"""
        try:
            email_message = self.parser.parsebytes(raw_email)
            
            # Extract basic headers
            headers = {
                'subject': email_message['subject'] or "",
                'from': email_message['from'] or "",
                'to': email_message['to'] or "",
                'date': email_message['date'] or "",
                'thread_id': self._extract_thread_id(email_message),
                'has_attachments': len(email_message.get_payload()) > 1 if email_message.is_multipart() else False
            }
            
            # Extract body
            body = self._extract_body(email_message)
            
            # Combine into result
            result = {**headers, 'body': body}
            
            return result
        
        except Exception as e:
            logger.error(f"Error parsing email: {str(e)}")
            return {
                'subject': 'Error parsing email',
                'from': '',
                'to': '',
                'date': '',
                'thread_id': None,
                'has_attachments': False,
                'body': 'There was an error parsing this email.'
            }
    
    def _extract_body(self, email_message):
        """Extract text content from email body, handling HTML and plaintext"""
        body = ""
        
        if email_message.is_multipart():
            for part in email_message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                # Skip attachments
                if "attachment" in content_disposition:
                    continue
                
                if content_type == "text/plain":
                    body = part.get_payload(decode=True).decode()
                    break
                elif content_type == "text/html":
                    html = part.get_payload(decode=True).decode()
                    body = self._html_to_text(html)
        else:
            content_type = email_message.get_content_type()
            if content_type == "text/plain":
                body = email_message.get_payload(decode=True).decode()
            elif content_type == "text/html":
                html = email_message.get_payload(decode=True).decode()
                body = self._html_to_text(html)
        
        # Clean up the body text
        body = self._clean_text(body)
        return body
        
    def _html_to_text(self, html):
        """Convert HTML to plaintext"""
        try:
            soup = BeautifulSoup(html, features="html.parser")
            return soup.get_text(separator='\n')
        except Exception as e:
            logger.error(f"Error converting HTML to text: {str(e)}")
            return "Error extracting HTML content"
    
    def _clean_text(self, text):
        """Clean up extracted text"""
        # Remove extra whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        # Remove lengthy email signatures
        text = re.sub(r'--+\s*.*$', '', text, flags=re.MULTILINE)
        # Remove forwarded message markers
        text = re.sub(r'-{3,}.*Forwarded message.*-{3,}', '', text)
        return text
            
    def _extract_thread_id(self, email_message):
        """Extract thread ID from email headers"""
        # Check headers used by various email systems
        if 'Thread-Index' in email_message:
            return email_message['Thread-Index']
        elif 'References' in email_message:
            return email_message['References'].split()[-1]
        elif 'In-Reply-To' in email_message:
            return email_message['In-Reply-To']
        elif 'Message-ID' in email_message:
            return email_message['Message-ID']
        return None
    
    def extract_recipients(self, email_data):
        """Extract recipient email addresses"""
        to = email_data.get('to', '')
        recipients = []
        
        # Match email addresses enclosed in angle brackets
        for match in re.finditer(r'<([^>]+)>', to):
            recipients.append(match.group(1))
        
        # If no matches, try splitting by commas
        if not recipients:
            for part in to.split(','):
                part = part.strip()
                if '@' in part:
                    recipients.append(part)
        
        return recipients
    
    def extract_sender_info(self, email_data):
        """Extract sender name and email address"""
        from_field = email_data.get('from', '')
        
        # Try to extract name and email
        match = re.search(r'(.*?)\s*<([^>]+)>', from_field)
        if match:
            name = match.group(1).strip()
            email = match.group(2).strip()
            return {'name': name, 'email': email}
        
        # If no match, the entire field is likely just an email
        if '@' in from_field:
            return {'name': '', 'email': from_field.strip()}
        
        return {'name': from_field.strip(), 'email': ''}
    
    def summarize_email(self, email_data, max_length=150):
        """Create a short summary of the email"""
        body = email_data.get('body', '')
        if not body:
            return ''
        
        # Get first paragraph
        paragraphs = body.split('\n\n')
        first_para = paragraphs[0] if paragraphs else ''
        
        # Truncate to desired length
        if len(first_para) > max_length:
            summary = first_para[:max_length] + '...'
        else:
            summary = first_para
        
        return summary