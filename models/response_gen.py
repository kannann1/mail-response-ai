# models/response_gen.py
import logging
import json
import datetime
import re

logger = logging.getLogger(__name__)

class ResponseGenerator:
    """Generates email responses using AI"""
    
    def __init__(self, ollama_service, user_config, ollama_config):
        self.ollama_service = ollama_service
        self.ai_service = ollama_service
        self.user_config = user_config
        self.ollama_config = ollama_config
        self.user_name = user_config.get('name', 'User')
        self.user_role = user_config.get('role', 'Professional')
        self.communication_style = user_config.get('communication_style', 'professional')
        
        # Load style samples if available
        self.style_samples = ollama_config.get('style_samples', [])
        
    def generate_response(self, email_data, email_history=None):
        """Generate appropriate response based on email content and history"""
        
        if not self.ai_service:
            return {
                'response_text': "AI service not configured. Please provide an API key.",
                'formatted_email': "",
                'confidence_score': 0.0,
                'needs_review': True
            }
        
        try:
            # Prepare context for the AI model
            prompt = self._build_prompt(email_data, email_history)
            
            # Generate response using AI service
            response = self.ai_service.generate_completion(
                prompt=prompt,
                system_prompt=self._get_system_prompt()
            )
            
            # Extract and format response
            reply_text = response.strip()
            formatted_reply = self._format_email_reply(email_data, reply_text)
            
            return {
                'response_text': reply_text,
                'formatted_email': formatted_reply,
                'confidence_score': self._calculate_confidence(reply_text),
                'needs_review': self._determine_if_needs_review(reply_text, email_data)
            }
            
        except Exception as e:
            logger.error(f"Error generating response: {str(e)}")
            return {
                'response_text': "Error generating response. Please try again.",
                'formatted_email': "",
                'confidence_score': 0.0,
                'needs_review': True
            }
    
    def _build_prompt(self, email_data, email_history=None):
        """Build a simplified prompt for Ollama"""
        prompt = f"""
        Write a reply to this email thread, i need response on top of last email:

        FROM: {email_data.get('from', '')}
        SUBJECT: {email_data.get('subject', '')}

        EMAIL CONTENT:
        {" ".join(email_data.get('body', '').splitlines())}

        Keep the reply short and professional.
        """

        return prompt
    
    def _get_system_prompt(self):
        """Generate system prompt based on user preferences"""
        return f"""
        You are an AI email assistant for {self.user_name}, who works as a {self.user_role}.
        
        IMPORTANT GUIDELINES:
        1. Write in {self.communication_style} style that sounds authentic to a DevOps engineer with 10 years of experience.
        2. Use straightforward, decent English. Avoid overly complex sentences or vocabulary.
        3. Be technically accurate when discussing DevOps-related topics.
        4. Keep responses concise and to the point.
        5. If there are any questions that you cannot answer confidently, indicate those in [NEEDS INPUT: question] format.
        6. If you need more information before drafting a complete response, specify what information is needed.
        
        Generate an email response that sounds authentically like it was written by the user.
        """
    
    def _format_email_reply(self, original_email, reply_text):
        """Format the response as an email reply"""
        current_date = datetime.datetime.now().strftime("%a, %d %b %Y %H:%M:%S")
        
        formatted_reply = f"""
        From: {self.user_config.get('email', 'user@example.com')}
        To: {original_email.get('from', '')}
        Subject: Re: {original_email.get('subject', '')}
        Date: {current_date}
        
        {reply_text}
        
        {self._get_signature()}
        """
        
        return formatted_reply
    
    def _get_signature(self):
        """Generate email signature based on user config"""
        return f"""
        Best regards,
        {self.user_name}
        {self.user_role}
        """
    
    def _calculate_confidence(self, response_text):
        """Calculate confidence score for the generated response"""
        # Lower confidence if response contains indicators of uncertainty
        uncertainty_phrases = ["I'm not sure", "I don't know", "NEEDS INPUT", "cannot determine", 
                               "unclear", "don't have enough information"]
        
        base_confidence = 0.85  # Base confidence score
        uncertainty_penalty = sum(0.15 for phrase in uncertainty_phrases if phrase.lower() in response_text.lower())
        
        # Lower confidence for very short or very long responses
        word_count = len(response_text.split())
        if word_count < 20 or word_count > 500:
            uncertainty_penalty += 0.1
        
        return max(0.3, base_confidence - uncertainty_penalty)  # Minimum confidence of 30%
    
    def _determine_if_needs_review(self, response_text, email_data):
        """Determine if the response needs human review"""
        # Always review if confidence is below threshold
        if self._calculate_confidence(response_text) < 0.7:
            return True
            
        # Always review emails from VIPs
        from_address = email_data.get('from', '').lower()
        if any(vip.lower() in from_address for vip in self.user_config.get('vip_contacts', [])):
            return True
            
        # Review if the email appears to be for high-stakes communication
        high_stakes_keywords = ["contract", "legal", "agreement", "offer", "confidential", 
                                "urgent", "critical", "emergency", "security", "breach"]
                                
        email_content = f"{email_data.get('subject', '')} {email_data.get('body', '')}".lower()
        if any(keyword in email_content for keyword in high_stakes_keywords):
            return True
            
        # Review if response contains explicit markers
        if "[NEEDS INPUT:" in response_text:
            return True
            
        # Review if specifically configured to always review
        if self.user_config.get('always_review', True):
            return True
            
        return False