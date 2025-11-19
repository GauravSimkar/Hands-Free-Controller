import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    WAKE_WORD = "arise"
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
    GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"
    #GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro-exp-03-25:generateContent"
    GOOGLE_CREDS = r"./voiceKey.json"
    AUDIO_SETTINGS = {
    'energy_threshold': 200,  # Lower for sensitive mics
    'pause_threshold': 1.0,   # Longer pause allowance
    'dynamic_energy': False,   # Disable dynamic adjustment
    'adjust_for_ambient_noise': True
    }