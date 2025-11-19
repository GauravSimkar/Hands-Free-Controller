import speech_recognition as sr
from google.cloud import speech
import pyautogui
from config import Config

class SpeechProcessor:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.client = speech.SpeechClient()
        
        # Configure audio settings from config
        from config import Config
        self.recognizer.energy_threshold = Config.AUDIO_SETTINGS['energy_threshold']
        self.recognizer.pause_threshold = Config.AUDIO_SETTINGS['pause_threshold']
        self.recognizer.dynamic_energy_threshold = Config.AUDIO_SETTINGS['dynamic_energy']

    def calibrate(self):
        """Calibrate microphone without logging"""
        with self.microphone as source:
            # Silent calibration - logging handled by VoiceAssistant
            self.recognizer.adjust_for_ambient_noise(source, duration=2)

    def transcribe(self, audio_data):
        config = speech.RecognitionConfig(
            encoding=speech.RecognitionConfig.AudioEncoding.LINEAR16,
            language_code="en-US",  # Set primary language to Hindi
            #alternative_language_codes=["hi-IN", "en-IN"],
            enable_automatic_punctuation=True,
            model="latest_long"
        )
        audio = speech.RecognitionAudio(content=audio_data)
        
        try:
            response = self.client.recognize(config=config, audio=audio)
            return " ".join([result.alternatives[0].transcript 
                           for result in response.results]).strip()
        except Exception as e:
            print(f"Transcription error: {e}")
            return None