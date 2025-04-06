# services/ollama_service.py
import requests
import json
import logging
import time

logger = logging.getLogger(__name__)

class OllamaService:
    """Service for interacting with Ollama API"""
    
    def __init__(self, host, model):
        self.host = host.rstrip('/')
        self.model = model
        self.max_retries = 3
        self.retry_delay = 2  # seconds
    
    def ping(self):
        """Ping Ollama to check if it's running"""
        try:
            response = requests.get(f"{self.host}/api/tags")
            if response.status_code == 200:
                return True
            return False
        except Exception as e:
            logger.error(f"Error pinging Ollama: {str(e)}")
            raise Exception(f"Ollama not running: {str(e)}")

    def generate_completion(self, prompt, system_prompt=None, temperature=0.7, max_tokens=None):
        """Generate completion using Ollama API - handling streaming response"""
        url = f"{self.host}/api/generate"

        # Combine system prompt and user prompt
        full_prompt = prompt
        if system_prompt:
            full_prompt = f"{system_prompt}\n\n{prompt}"

        payload = {
            "model": self.model,
            "prompt": full_prompt,
            "stream": False  # This might be ignored by some Ollama versions
        }

        logger.info(f"Sending request to Ollama API")

        try:
            response = requests.post(url, json=payload)

            # Check response content type
            if 'application/x-ndjson' in response.headers.get('Content-Type', ''):
                logger.info("Received streaming response, concatenating chunks")

                # Split response into lines and parse each as JSON
                combined_response = ""
                for line in response.text.strip().split('\n'):
                    try:
                        data = json.loads(line)
                        chunk = data.get("response", "")
                        combined_response += chunk
                    except json.JSONDecodeError:
                        logger.warning(f"Could not parse line: {line[:50]}...")
                        continue
                    
                logger.info(f"Combined response length: {len(combined_response)}")
                return combined_response

            else:
                # Try normal JSON parsing
                data = response.json()
                return data.get("response", "")

        except Exception as e:
            logger.error(f"Error generating completion: {str(e)}")
            return f"Error: {str(e)}"    
        
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
    
    def list_available_models(self):
        """List available models in Ollama"""
        try:
            response = requests.get(f"{self.host}/api/tags")
            response.raise_for_status()
            
            data = response.json()
            models = [model["name"] for model in data.get("models", [])]
            return models
            
        except Exception as e:
            logger.error(f"Error listing Ollama models: {str(e)}")
            return []