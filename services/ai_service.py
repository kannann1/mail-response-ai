# services/ai_service.py
import openai
import logging
import time

logger = logging.getLogger(__name__)

class AIService:
    """Service for interacting with OpenAI API"""
    
    def __init__(self, api_key, model="gpt-4"):
        self.api_key = api_key
        self.model = model
        self.max_retries = 3
        self.retry_delay = 2  # seconds
        
        # Set up OpenAI client
        openai.api_key = api_key
    
    def generate_completion(self, prompt, system_prompt=None, temperature=0.7, max_tokens=1500):
        """Generate completion using OpenAI API"""
        messages = []
        
        # Add system prompt if provided
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        
        # Add user prompt
        messages.append({"role": "user", "content": prompt})
        
        # Try to generate completion with retries
        for attempt in range(self.max_retries):
            try:
                response = openai.ChatCompletion.create(
                    model=self.model,
                    messages=messages,
                    temperature=temperature,
                    max_tokens=max_tokens
                )
                
                return response.choices[0].message.content
                
            except Exception as e:
                logger.error(f"Error generating completion (attempt {attempt+1}): {str(e)}")
                
                if attempt < self.max_retries - 1:
                    # Wait before retrying
                    time.sleep(self.retry_delay * (2 ** attempt))  # Exponential backoff
                else:
                    # Max retries reached
                    raise Exception(f"Failed to generate completion after {self.max_retries} attempts: {str(e)}")
    
    def analyze_sentiment(self, text):
        """Analyze sentiment of text"""
        try:
            prompt = f"Analyze the sentiment of the following text. Rate it as positive, neutral, or negative, and provide a brief explanation why: {text}"
            
            response = self.generate_completion(prompt)
            
            return response
            
        except Exception as e:
            logger.error(f"Error analyzing sentiment: {str(e)}")
            return "Error analyzing sentiment"
    
    def extract_key_points(self, text):
        """Extract key points from text"""
        try:
            prompt = f"Extract the key points from the following text. Focus on actionable items, requests, and important information: {text}"
            
            response = self.generate_completion(prompt)
            
            return response
            
        except Exception as e:
            logger.error(f"Error extracting key points: {str(e)}")
            return "Error extracting key points"