import json
import os
import logging
from pathlib import Path
import sqlite3
import pickle

logger = logging.getLogger(__name__)

class StorageService:
    """Service for handling local storage of settings and data"""
    
    def __init__(self, config_dir=None):
        if config_dir is None:
            self.config_dir = Path.home() / ".email_agent_ollama"
        else:
            self.config_dir = Path(config_dir)
        
        self.config_file = self.config_dir / "settings.json"
        self.db_file = self.config_dir / "email_data.db"
        
        # Create config directory if it doesn't exist
        self.config_dir.mkdir(exist_ok=True)
        
        # Initialize the database
        self._init_database()
    
    def config_exists(self):
        """Check if configuration file exists"""
        return self.config_file.exists()
    
    def load_config(self):
        """Load configuration from file"""
        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)
            return config
        except Exception as e:
            logger.error(f"Error loading configuration: {str(e)}")
            return None
    
    def save_config(self, config):
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=4)
            return True
        except Exception as e:
            logger.error(f"Error saving configuration: {str(e)}")
            return False
    
    def _init_database(self):
        """Initialize the SQLite database"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            # Create tasks table
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY,
                text TEXT NOT NULL,
                email_id TEXT,
                email_from TEXT,
                due_date TEXT,
                priority TEXT DEFAULT 'Medium',
                status TEXT DEFAULT 'Not Started',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            ''')
            
            # Create drafts table
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS drafts (
                id INTEGER PRIMARY KEY,
                email_id TEXT,
                original_email BLOB,
                response_text TEXT,
                formatted_email TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status TEXT DEFAULT 'Draft'
            )
            ''')
            
            # Create style samples table
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS style_samples (
                id TEXT PRIMARY KEY,
                text TEXT NOT NULL,
                added_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            ''')
            
            conn.commit()
            conn.close()
            
            logger.info("Database initialized")
            return True
            
        except Exception as e:
            logger.error(f"Error initializing database: {str(e)}")
            return False
    
    def save_task(self, task_data):
        """Save a task to the database"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
            INSERT INTO tasks (text, email_id, email_from, due_date, priority, status)
            VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                task_data.get('text', ''),
                task_data.get('email_id', ''),
                task_data.get('email_from', ''),
                task_data.get('due_date', ''),
                task_data.get('priority', 'Medium'),
                task_data.get('status', 'Not Started')
            ))
            
            task_id = cursor.lastrowid
            
            conn.commit()
            conn.close()
            
            return task_id
            
        except Exception as e:
            logger.error(f"Error saving task: {str(e)}")
            return None
    
    def get_tasks(self, status=None):
        """Get tasks from the database, optionally filtered by status"""
        try:
            conn = sqlite3.connect(self.db_file)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            if status:
                cursor.execute('SELECT * FROM tasks WHERE status = ? ORDER BY due_date', (status,))
            else:
                cursor.execute('SELECT * FROM tasks ORDER BY due_date')
            
            rows = cursor.fetchall()
            
            tasks = []
            for row in rows:
                task = dict(row)
                tasks.append(task)
            
            conn.close()
            
            return tasks
            
        except Exception as e:
            logger.error(f"Error getting tasks: {str(e)}")
            return []
    def delete_draft(self, draft_id):
        """Delete a draft response from the database"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
        
            cursor.execute('DELETE FROM drafts WHERE id = ?', (draft_id,))
        
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            logger.error(f"Error deleting draft: {str(e)}")    
            return False
        
    def save_draft(self, draft_data):
        """Save a draft response to the database"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
            INSERT INTO drafts (email_id, original_email, response_text, formatted_email)
            VALUES (?, ?, ?, ?)
            ''', (
                draft_data.get('email_id', ''),
                pickle.dumps(draft_data.get('original_email', {})),
                draft_data.get('response_text', ''),
                draft_data.get('formatted_email', '')
            ))
            
            draft_id = cursor.lastrowid
            
            conn.commit()
            conn.close()
            
            return draft_id
            
        except Exception as e:
            logger.error(f"Error saving draft: {str(e)}")
            return None
    
    def get_drafts(self):
        """Get all draft responses from the database"""
        try:
            conn = sqlite3.connect(self.db_file)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute('SELECT * FROM drafts ORDER BY created_at DESC')
            
            rows = cursor.fetchall()
            
            drafts = []
            for row in rows:
                draft = dict(row)
                try:
                    draft['original_email'] = pickle.loads(draft['original_email'])
                except:
                    draft['original_email'] = {}
                drafts.append(draft)
            
            conn.close()
            
            return drafts
            
        except Exception as e:
            logger.error(f"Error getting drafts: {str(e)}")
            return []
    
    def save_style_sample(self, sample_id, text):
        """Save a style sample to the database"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            # Check if sample exists
            cursor.execute('SELECT id FROM style_samples WHERE id = ?', (sample_id,))
            exists = cursor.fetchone()
            
            if exists:
                # Update existing sample
                cursor.execute('UPDATE style_samples SET text = ? WHERE id = ?', (text, sample_id))
            else:
                # Insert new sample
                cursor.execute('INSERT INTO style_samples (id, text) VALUES (?, ?)', (sample_id, text))
            
            conn.commit()
            conn.close()
            
            return True
            
        except Exception as e:
            logger.error(f"Error saving style sample: {str(e)}")
            return False
    
    def get_style_sample(self, sample_id):
        """Get a style sample from the database"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('SELECT text FROM style_samples WHERE id = ?', (sample_id,))
            
            result = cursor.fetchone()
            
            conn.close()
            
            if result:
                return result[0]
            return None
            
        except Exception as e:
            logger.error(f"Error getting style sample: {str(e)}")
            return None
    
    def get_style_samples(self):
        """Get all style samples from the database"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('SELECT id, text FROM style_samples ORDER BY id')
            
            rows = cursor.fetchall()
            
            samples = {}
            for row in rows:
                sample_id, text = row
                samples[sample_id] = text
            
            conn.close()
            
            return samples
            
        except Exception as e:
            logger.error(f"Error getting style samples: {str(e)}")
            return {}
    def update_task(self, task_id, task_data):
        """Update a task in the database"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()

            cursor.execute('''
            UPDATE tasks
            SET text = ?, due_date = ?, priority = ?, status = ?
            WHERE id = ?
            ''', (
                task_data.get('text', ''),
                task_data.get('due_date', ''),
                task_data.get('priority', 'Medium'),
                task_data.get('status', ''),
                task_id
            ))

            conn.commit()
            conn.close()

            return True

        except Exception as e:
            logger.error(f"Error updating task: {str(e)}")
            return False
        
    def delete_task(self, task_id):
        """Delete a task from the database"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
        
            cursor.execute('DELETE FROM tasks WHERE id = ?', (task_id,))
        
            conn.commit()
            conn.close()
        
            return True
        
        except Exception as e:
            logger.error(f"Error deleting task: {str(e)}")
            return False
    def delete_style_sample(self, sample_id):
        """Delete a style sample from the database"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('DELETE FROM style_samples WHERE id = ?', (sample_id,))
            
            conn.commit()
            conn.close()
            
            return True
            
        except Exception as e:
            logger.error(f"Error deleting style sample: {str(e)}")
            return False