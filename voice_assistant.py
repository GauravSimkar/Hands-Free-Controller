# voice_assistant.py
import time
import os
import speech_recognition as sr
from speech_processor import SpeechProcessor
from command_handler import CommandHandler
from config import Config
from window_utils import get_active_window
import threading
from window_utils import get_all_open_windows


class VoiceAssistant:
    is_active = True  # Class variable to control global state

    def __init__(self, gui=None):
        self.speech_processor = SpeechProcessor()
        self.command_handler = CommandHandler(self)
        self.gui = gui
        self.activated = False
        self.should_calibrate = True
        self.is_processing = False  # Add processing state flag

    def run(self):
        """Main execution loop with proper logging"""
        self.log("Calibrating microphone...")
        self.speech_processor.calibrate()
        self.log("Calibration complete")
        
        while VoiceAssistant.is_active:
            self.activated = False
            self.activate()
            self.main_loop()

    def log(self, message):
        """Handle logging through GUI or console"""
        if self.gui:
            self.gui.log(message)
        else:
            print(message)

    def activate(self):
        """Wait for wake word activation with adjusted timeouts"""
        self.log(f"🔈 Say '{Config.WAKE_WORD}' to start...")
        if self.gui:
            self.gui.update_status("waiting")
        
        with self.speech_processor.microphone as source:
            while not self.activated and VoiceAssistant.is_active:
                try:
                    # Increase timeout and phrase limit
                    audio = self.speech_processor.recognizer.listen(
                        source, 
                        timeout=5,  # Increased from 3
                        phrase_time_limit=3  # Increased from 2
                    )
                    transcript = self.speech_processor.transcribe(audio.get_wav_data())
                    print(transcript)
                    if transcript and Config.WAKE_WORD in transcript.lower():
                        self.activated = True
                        self.log("✅ Activated! Listening continuously...")
                        if self.gui:
                            self.gui.update_status("active")
                            
                except sr.WaitTimeoutError:
                    continue  # Silently ignore timeouts during activation
                except Exception as e:
                    self.log(f"Activation Error: {e}")
                time.sleep(0.1)  # Add small delay between attempts

    def main_loop(self):
        """Continuous listening loop with processing state check"""
        with self.speech_processor.microphone as source:
            while self.activated and VoiceAssistant.is_active:
                try:
                    # Check processing state before listening
                    if self.is_processing:
                        time.sleep(0.1)
                        continue

                    # Increased phrase time limit for main listening
                    audio = self.speech_processor.recognizer.listen(
                        source, 
                        timeout=5,
                        phrase_time_limit=10
                    )
                    transcript = self.speech_processor.transcribe(audio.get_wav_data())
                    
                    if transcript:
                        if not self.is_processing:
                            self.log(f"Recognized: {transcript}")
                            self.is_processing = True  # Set processing flag
                            self.log("🔄 Processing commands - new requests will wait...")
                            if self.gui:
                                self.gui.start_loading()
                            
                            processing_thread = threading.Thread(
                                target=self.process_command,
                                args=(transcript,)
                            )
                            processing_thread.start()
                        else:
                            self.log("⚠️ Ignoring new request during command execution")
                        
                except sr.WaitTimeoutError:
                    continue
                except Exception as e:
                    self.log(f"Listening Error: {e}")
                time.sleep(0.2)

    def process_command(self, transcript):
        """Handle command processing and clear processing flag"""
        try:
            if self.gui:
                self.gui.start_loading(transcript)  # Pass transcript here
            commands = self.command_handler.generate_commands(transcript)
            self.execute_commands(commands)
        finally:
            self.is_processing = False
            self.log("✅ Ready for new commands")
            if self.gui:
                self.gui.stop_loading()

    def execute_commands(self, commands):
        """Execute commands with enhanced synchronization"""
        self.log("⚡ Executing commands...")
        index = 0
        #last_window_state = set(get_all_open_windows())
        
        while index < len(commands) and VoiceAssistant.is_active:
            cmd = commands[index]
            command_name = cmd.get("command")
            params = cmd.get("parameters", {})
            
            try:
                pre_windows = set(get_all_open_windows())
                
                self.log(f"Executing: {command_name} {params}")
                handler = self.command_handler.actions.get(command_name)
                
                if handler:
                    start_time = time.time()
                    handler(**params)
                    exec_time = time.time() - start_time
                    
                    # Enhanced synchronization logic
                    if command_name == "run_command":
                        # Extended window detection timeout
                        launch_timeout = 7  # Increased from 3 seconds
                        window_found = False
                        
                        # Detect window creation
                        window_start = time.time()
                        while (time.time() - window_start) < launch_timeout:
                            current_windows = set(get_all_open_windows())
                            new_windows = current_windows - pre_windows
                            
                            if new_windows:
                                window_found = True
                                break
                            time.sleep(0.1)
                        
                        # Additional stabilization delay
                        if window_found:
                            time.sleep(0.5)  # Added post-window detection wait
                            
                            # Skip subsequent long waits
                            if (index + 1 < len(commands)):
                                next_cmd = commands[index+1]
                                if (next_cmd.get("command") == "sleep" and 
                                    next_cmd.get("parameters", {}).get("duration", 0) > 0.5):
                                    index += 1

                    # General inter-command synchronization
                    if command_name != "sleep":
                        # Extended minimum delay between commands
                        remaining_delay = max(0, 0.3 - exec_time)  # Increased from 0.15
                        time.sleep(remaining_delay)

                    # Special handling for browser navigation
                    if command_name == "open_url":
                        time.sleep(1.5)  # Additional stabilization for web pages

                index += 1
                
            except Exception as e:
                self.log(f"Execution error: {e}")
                index += 1
                
        self.log("🗝️ All commands executed")