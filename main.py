import os
import sys
import logging
import json
from pathlib import Path
import threading
import tkinter as tk
from tkinter import messagebox

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("email_agent.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def create_default_config():
    """Create default configuration if none exists"""
    default_config = {
        "user": {
            "name": "Your Name",
            "email": "your.email@example.com",
            "role": "DevOps Engineer",
            "communication_style": "professional"
        },
        "email": {
            "refresh_interval": 300,  # seconds
            "auto_archive": False,
            "important_contacts": [],
            "vip_contacts": []
        },
        "ollama": {
            "host": "http://localhost:11434",
            "model": "llama3",
            "system_prompt": "",
            "always_review": True,
            "style_samples": []
        }
    }
    
    config_dir = Path.home() / ".email_agent_ollama"
    config_dir.mkdir(exist_ok=True)
    
    config_file = config_dir / "settings.json"
    
    with open(config_file, 'w') as f:
        json.dump(default_config, f, indent=4)
    
    logger.info(f"Created default configuration at {config_file}")
    return default_config

def check_ollama_running():
    """Check if Ollama is running"""
    from services.ollama_service import OllamaService
    
    try:
        ollama = OllamaService("http://localhost:11434", "llama3")
        ollama.ping()
        return True
    except Exception as e:
        logger.error(f"Ollama not running: {str(e)}")
        return False

def main():
    """Main entry point for the application"""
    logger.info("Starting Email Agent with Ollama")
    
    # Check if Ollama is running
    if not check_ollama_running():
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Ollama Not Running",
            "Ollama is not running. Please start Ollama before running this application."
        )
        root.destroy()
        logger.error("Exiting because Ollama is not running")
        return
    
    # Import modules here to avoid loading them before Ollama check
    from services.storage_service import StorageService
    from services.outlook_service import OutlookService
    from services.ollama_service import OllamaService
    from models.email_processor import EmailProcessor
    from models.priority_engine import PriorityEngine
    from models.response_gen import ResponseGenerator
    from models.action_extractor import ActionItemExtractor
    from ui.main_window import EmailAgentUI
    
    # Initialize storage service
    storage = StorageService()
    
    # Load or create configuration
    if not storage.config_exists():
        config = create_default_config()
        logger.info("Using default configuration")
    else:
        config = storage.load_config()
        logger.info("Loaded existing configuration")
    
    # Initialize services
    outlook_service = OutlookService(config)  # We'll set the root later
    ollama_service = OllamaService(
        config["ollama"]["host"], 
        config["ollama"]["model"]
    )
    
    # Initialize models
    email_processor = EmailProcessor()
    priority_engine = PriorityEngine(config["email"])
    response_generator = ResponseGenerator(ollama_service, config["user"], config["ollama"])
    action_extractor = ActionItemExtractor()
    
    # Start the UI
    app = EmailAgentUI(
        outlook_service=outlook_service,
        storage_service=storage,
        email_processor=email_processor,
        priority_engine=priority_engine,
        response_generator=response_generator,
        action_extractor=action_extractor,
        config=config
    )
    outlook_service.app_root = app.root
    app.run()

if __name__ == "__main__":
    main()