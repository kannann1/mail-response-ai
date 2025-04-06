# models/priority_engine.py
import re
import logging
from datetime import datetime, timezone

logger = logging.getLogger(__name__)

class PriorityEngine:
    """Engine to determine email priority based on various factors"""
    
    def __init__(self, email_config):
        self.config = email_config
        self.important_contacts = self.config.get("important_contacts", [])
        self.vip_contacts = self.config.get("vip_contacts", [])
        self.urgency_keywords = [
            "urgent", "asap", "immediately", "deadline", "important",
            "critical", "emergency", "priority", "attention", "needed"
        ]
        
    def prioritize_email(self, email_data):
        """Calculate priority score for an email (0-100)"""
        try:
            score = 0
            explanations = []
            
            # 1. Sender importance (0-30 points)
            from_address = self._extract_email_address(email_data.get('from', ''))
            
            if from_address in self.vip_contacts:
                score += 30
                explanations.append("VIP sender (+30)")
            elif from_address in self.important_contacts:
                score += 20
                explanations.append("Important contact (+20)")
                
            # 2. Direct addressing (0-15 points)
            if self._is_directly_addressed(email_data):
                score += 15
                explanations.append("Directly addressed (+15)")
                
            # 3. Urgency keywords (0-20 points)
            urgency_score = self._calculate_urgency_score(
                email_data.get('subject', ''), 
                email_data.get('body', '')
            )
            if urgency_score > 0:
                score += urgency_score
                explanations.append(f"Urgency keywords (+{urgency_score})")
                
            # 4. Recency (0-15 points)
            recency_score = self._calculate_recency_score(email_data.get('date', ''))
            if recency_score > 0:
                score += recency_score
                explanations.append(f"Recent email (+{recency_score})")
                
            # 5. Thread continuation (0-10 points)
            if email_data.get('thread_id') or email_data.get('conversation_id'):
                score += 10
                explanations.append("Active thread (+10)")
                
            # 6. Question detection (0-10 points)
            if self._contains_questions(email_data.get('body', '')):
                score += 10
                explanations.append("Contains questions (+10)")
                
            return {
                'score': min(score, 100),  # Cap at 100
                'explanations': explanations,
                'category': self._determine_category(score)
            }
            
        except Exception as e:
            logger.error(f"Error calculating priority: {str(e)}")
            return {
                'score': 50,  # Default medium priority
                'explanations': ["Error in priority calculation"],
                'category': 'Medium'
            }
    
    def _extract_email_address(self, from_str):
        """Extract just the email address from a From string"""
        match = re.search(r'<([^>]+)>', from_str)
        if match:
            return match.group(1).lower()
        return from_str.lower()
    
    def _is_directly_addressed(self, email_data):
        """Check if email directly addresses the user by name"""
        # This would check for the user's name in the email body
        # For now, simplified to always return True
        return True
    
    def _calculate_urgency_score(self, subject, body):
        """Calculate urgency based on keywords"""
        combined = (subject + " " + body).lower()
        
        # Count occurrences of urgency keywords
        count = sum(1 for keyword in self.urgency_keywords if keyword in combined)
        
        # Cap at 20 points
        return min(count * 5, 20)
    
    def _calculate_recency_score(self, date_str):
        """Calculate score based on how recent the email is"""
        if not date_str:
            return 0
            
        try:
            # Try to parse the date string
            # The format can vary depending on the email system
            try:
                # Format: "Wed, 12 Feb 2023 14:30:45 +0000"
                date_formats = [
                    "%a, %d %b %Y %H:%M:%S %z",
                    "%a, %d %b %Y %H:%M:%S",
                    "%d %b %Y %H:%M:%S %z",
                    "%Y-%m-%d %H:%M:%S",
                ]
                
                email_date = None
                for fmt in date_formats:
                    try:
                        email_date = datetime.strptime(date_str, fmt)
                        break
                    except:
                        continue
                
                if not email_date:
                    return 0
                    
            except Exception as e:
                logger.error(f"Error parsing date '{date_str}': {str(e)}")
                return 0
            
            # Make sure date has timezone info
            if email_date.tzinfo is None:
                email_date = email_date.replace(tzinfo=timezone.utc)
                
            now = datetime.now(timezone.utc)
            
            # Calculate difference in hours
            diff_hours = (now - email_date).total_seconds() / 3600
            
            # Score based on recency
            if diff_hours < 1:  # Less than 1 hour
                return 15
            elif diff_hours < 4:  # Less than 4 hours
                return 10
            elif diff_hours < 24:  # Less than 24 hours
                return 5
            return 0
            
        except Exception as e:
            logger.error(f"Error calculating recency: {str(e)}")
            return 0
    
    def _contains_questions(self, body):
        """Check if the email contains questions"""
        question_marks = body.count('?')
        question_phrases = ["could you", "would you", "please let me know", 
                           "can you", "do you know", "what is", "when will"]
                           
        # Check for at least one question mark or phrase
        has_question_mark = question_marks > 0
        has_question_phrase = any(phrase in body.lower() for phrase in question_phrases)
        
        return has_question_mark or has_question_phrase
    
    def _determine_category(self, score):
        """Convert numerical score to priority category"""
        if score >= 80:
            return "Urgent"
        elif score >= 60:
            return "High"
        elif score >= 40:
            return "Medium"
        else:
            return "Low"