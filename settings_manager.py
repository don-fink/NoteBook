"""
settings_manager.py
Manages loading and saving application settings, such as the last opened database, in a JSON file.
"""
import json
import os

SETTINGS_FILE = 'settings.json'

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_settings(settings):
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings, f, indent=2)

def get_last_db():
    settings = load_settings()
    return settings.get('last_db')

def set_last_db(db_path):
    settings = load_settings()
    settings['last_db'] = db_path
    save_settings(settings)
