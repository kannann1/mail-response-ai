# models/action_extractor.py
import re
import logging
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)

class ActionItemExtractor:
    """Extracts action items from email content"""
    
    def __init__(self, ollama_service=None):
        # Initialize with optional Ollama service
        self.ollama_service = ollama_service
        
        # Initialize patterns for action item extraction
        self.request_patterns = [
            r"(?:can|could|would) you (?:please)?\s+([^?\.]*)\??",
            r"(?:please|kindly)\s+([^?\.]*)",
            r"(?:need|require|want)(?:ed)? you to\s+([^?\.]*)",
            r"(?:don't forget|remember) to\s+([^?\.]*)",
            r"(?:must|should|have to)\s+([^?\.]*)",
            r"(?:I'm waiting for|I'll be waiting for|I await|expecting) you to\s+([^?\.]*)"
        ]
        
        # Patterns for due dates
        self.date_patterns = [
            r"by\s+(\w+day,?\s+\w+\s+\d{1,2}(?:st|nd|rd|th)?)",
            r"due\s+(?:by|on)\s+(\w+day,?\s+\w+\s+\d{1,2}(?:st|nd|rd|th)?)",
            r"(?:before|until)\s+(\w+day,?\s+\w+\s+\d{1,2}(?:st|nd|rd|th)?)",
            r"(?:by|on)\s+(tomorrow|today|monday|tuesday|wednesday|thursday|friday|saturday|sunday)",
            r"by\s+(?:the\s+)?(?:end\s+of|close\s+of|)\s+(\w+day|this week|next week|this month)",
            r"(?:in|within)\s+the\s+next\s+(\d+)\s+(day|days|week|weeks)"
        ]
    
    def extract_action_items(self, email_data):
        """Extract action items from email content"""
        action_items = []
        
        try:
            body = email_data.get('body', '')
            subject = email_data.get('subject', '')
            
            # Use Ollama for extraction if available
            if self.ollama_service:
                try:
                    ollama_items = self._extract_with_ollama(subject, body)
                    action_items.extend(ollama_items)
                except Exception as e:
                    logger.error(f"Error with Ollama extraction: {str(e)}")
                    # Fall back to rule-based extraction
            
            # Rule-based extraction as fallback or complement
            # Check subject line for action items
            subject_actions = self._extract_from_text(subject)
            for action in subject_actions:
                action['source'] = 'subject'
                action['confidence'] += 0.1  # Higher confidence for subject items
                action_items.append(action)
            
            # Process email body
            body_actions = self._extract_from_text(body)
            for action in body_actions:
                action['source'] = 'body'
                action_items.append(action)
            
            # Deduplicate action items
            action_items = self._deduplicate_actions(action_items)
            
            return action_items
            
        except Exception as e:
            logger.error(f"Error extracting action items: {str(e)}")
            return []
    
    def _extract_with_ollama(self, subject, body):
        """Extract action items using Ollama"""
        if not self.ollama_service:
            return []
            
        prompt = f"""
        Extract actionable items from this email:
        
        SUBJECT: {subject}
        
        BODY:
        {body}
        
        For each action item, provide:
        1. The specific task that needs to be done
        2. Any deadline mentioned (or "None" if no deadline)
        3. The priority level (High, Medium, Low)
        
        Format your response as a JSON list of objects with these fields:
        text, due_date, priority
        
        Only include real action items, not general information.
        """
        
        system_prompt = """
        You are an AI assistant specialized in extracting action items from emails.
        Your task is to identify specific actions the recipient needs to take.
        Return your analysis as valid JSON that can be parsed programmatically.
        """
        
        try:
            result = self.ollama_service.generate_completion(prompt, system_prompt)
            
            # Try to extract JSON from the response
            json_pattern = r'```json\s*([\s\S]*?)\s*```'
            match = re.search(json_pattern, result)
            
            if match:
                json_str = match.group(1)
            else:
                # If no JSON code block found, try to find JSON directly
                json_pattern = r'\[\s*{[\s\S]*}\s*\]'
                match = re.search(json_pattern, result)
                if match:
                    json_str = match.group(0)
                else:
                    json_str = result
            
            try:
                items = json.loads(json_str)
                
                # Convert to our format
                actions = []
                for item in items:
                    actions.append({
                        'text': item.get('text', ''),
                        'due_date': item.get('due_date', None),
                        'confidence': 0.85,  # Higher confidence for LLM-extracted items
                        'priority': item.get('priority', 'Medium')
                    })
                
                return actions
            except json.JSONDecodeError:
                logger.error("Failed to parse JSON from Ollama response")
                return []
                
        except Exception as e:
            logger.error(f"Error using Ollama for action extraction: {str(e)}")
            return []
    
    def _extract_from_text(self, text):
        """Extract action items from text using rule-based approach"""
        actions = []
        
        # Extract using request patterns
        for pattern in self.request_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                action = match.group(1).strip()
                if self._is_valid_action(action):
                    due_date = self._extract_due_date(text)
                    actions.append({
                        'text': action,
                        'confidence': 0.8,
                        'due_date': due_date,
                        'priority': 'Medium'
                    })
        
        # Extract potential action items from lines with bullet points or numbers
        bullet_pattern = r"(?:^|\n)(?:\*|\-|\d+\.)\s+([^\n]*)"
        bullet_matches = re.finditer(bullet_pattern, text)
        for match in bullet_matches:
            bullet_text = match.group(1).strip()
            if self._is_likely_action(bullet_text):
                due_date = self._extract_due_date(text)
                actions.append({
                    'text': bullet_text,
                    'confidence': 0.7,
                    'due_date': due_date,
                    'priority': 'Medium'
                })
        
        return actions
    
    def _is_valid_action(self, action):
        """Validate extracted action item"""
        # Must be at least 3 words
        if len(action.split()) < 3:
            return False
            
        # Must not be too long
        if len(action) > 200:
            return False
            
        # Should not be a question
        if action.endswith('?'):
            return False
            
        return True
    
    def _is_likely_action(self, text):
        """Check if a text is likely an action item"""
        # Action items often start with verbs
        first_word = text.split()[0].lower() if text.split() else ""
        common_action_verbs = ["create", "update", "review", "prepare", "send", "check", 
                              "complete", "implement", "develop", "fix", "test", "deploy"]
                              
        starts_with_verb = first_word in common_action_verbs
        
        # Check for action-oriented phrases
        action_phrases = ["need to", "should", "must", "have to", "required"]
        contains_action_phrase = any(phrase in text.lower() for phrase in action_phrases)
        
        return starts_with_verb or contains_action_phrase
    
    def _extract_due_date(self, text):
        """Extract due date from text"""
        for pattern in self.date_patterns:
            matches = re.search(pattern, text, re.IGNORECASE)
            if matches:
                date_text = matches.group(1)
                try:
                    return self._parse_date_text(date_text)
                except:
                    return date_text  # Return as text if parsing fails
        
        return None
    
    def _parse_date_text(self, date_text):
        """Parse date text into a date object"""
        today = datetime.now()
        
        # Handle relative dates
        date_text = date_text.lower()
        
        if "today" in date_text:
            return today.strftime("%Y-%m-%d")
        elif "tomorrow" in date_text:
            return (today + timedelta(days=1)).strftime("%Y-%m-%d")
        elif "this week" in date_text:
            # End of this week (Friday)
            days_until_friday = (4 - today.weekday()) % 7
            return (today + timedelta(days=days_until_friday)).strftime("%Y-%m-%d")
        elif "next week" in date_text:
            # Middle of next week (Wednesday)
            days_until_next_wednesday = (9 - today.weekday()) % 7
            return (today + timedelta(days=days_until_next_wednesday)).strftime("%Y-%m-%d")
        
        # Handle day names
        days = {
            "monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3,
            "friday": 4, "saturday": 5, "sunday": 6
        }
        
        for day, day_num in days.items():
            if day in date_text:
                days_ahead = (day_num - today.weekday()) % 7
                if days_ahead == 0:  # Today
                    days_ahead = 7  # Next week
                return (today + timedelta(days=days_ahead)).strftime("%Y-%m-%d")
        
        # Return original text if we can't parse it
        return date_text
    
    def _deduplicate_actions(self, actions):
        """Remove duplicate action items"""
        seen = set()
        unique_actions = []
        
        for action in actions:
            text_normalized = action['text'].lower()
            if text_normalized not in seen:
                seen.add(text_normalized)
                unique_actions.append(action)
        
        return unique_actions