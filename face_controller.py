# face_controller.py (updated)
import cv2
import mediapipe as mp
import pyautogui
import numpy as np
import time
import keyboard
from threading import Thread
import sys
import os

pyautogui.FAILSAFE = False

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
    os.environ['MEDIAPIPE_MODEL_PATH'] = os.path.join(
        base_path, 
        'mediapipe/modules'
    )

class FaceController(Thread):
    def __init__(self, params):
        super().__init__()
        self.params = params
        self.paused = False
        self.running = True
        self._setup_mediapipe()
        self._initialize_state()
        
    def _setup_mediapipe(self):
        self.mp_face_mesh = mp.solutions.face_mesh
        self.face_mesh = self.mp_face_mesh.FaceMesh(
            static_image_mode=False,
            max_num_faces=1,
            refine_landmarks=True,
            min_detection_confidence=0.8,
            min_tracking_confidence=0.8)
        
    def _initialize_state(self):
        self.prev_smooth_x = 0.5
        self.prev_smooth_y = 0.5
        self.blink_counter = 0
        self.ear_threshold = None
        self.normal_ear = 0
        self.normal_mouth = 0
        self.calibration_eyebrow_left = 0
        self.calibration_eyebrow_right = 0
        self.waiting_for_double = False
        self.first_blink_time = 0
        self.last_right_click_time = 0
        self.scrolling = False
        self.scroll_direction = None
        self.initial_smooth_x = None
        self.initial_smooth_y = None
        self.last_scroll_time = 0
        self.mouth_open_threshold = 0
        self.screen_w, self.screen_h = pyautogui.size()
        
    def run(self):
        self._calibrate()
        cap = cv2.VideoCapture(0)
        try:
            while self.running:
                if not self.paused:
                    self._process_frame(cap)
                time.sleep(0.01)
        finally:
            cap.release()
            cv2.destroyAllWindows()
            
    def _calibrate(self):
        print("Calibrating...")
        cap = cv2.VideoCapture(0)
        calibration_counter = 0
        while calibration_counter < self.params['CALIBRATION_FRAMES'] and self.running:
            ret, frame = cap.read()
            if ret:
                results = self.face_mesh.process(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
                if results.multi_face_landmarks:
                    landmarks = results.multi_face_landmarks[0].landmark
                    self._collect_calibration_data(landmarks)
                    calibration_counter += 1
        cap.release()
        self._finalize_calibration()
        
    def _collect_calibration_data(self, landmarks):
        left_ear = self._calculate_ear(landmarks, [362, 385, 387, 263, 373, 380])
        right_ear = self._calculate_ear(landmarks, [33, 160, 158, 133, 153, 144])
        self.normal_ear += (left_ear + right_ear) / 2
        
        left_eyebrow_y = np.mean([landmarks[i].y for i in [276, 283, 282, 295, 285]])
        self.calibration_eyebrow_left += (landmarks[362].y - left_eyebrow_y)
        
        right_eyebrow_y = np.mean([landmarks[i].y for i in [46, 53, 52, 65, 55]])
        self.calibration_eyebrow_right += (landmarks[133].y - right_eyebrow_y)
        
        self.normal_mouth += self._calculate_mouth_open(landmarks)
        
    def _finalize_calibration(self):
        self.normal_ear /= self.params['CALIBRATION_FRAMES']
        self.ear_threshold = self.normal_ear * self.params['BLINK_THRESHOLD_RATIO']
        self.calibration_eyebrow_left /= self.params['CALIBRATION_FRAMES']
        self.calibration_eyebrow_right /= self.params['CALIBRATION_FRAMES']
        self.normal_mouth /= self.params['CALIBRATION_FRAMES']
        self.mouth_open_threshold = self.normal_mouth * self.params['MOUTH_OPEN_THRESHOLD_RATIO']
        print("Calibration complete")
        
    def _process_frame(self, cap):
        ret, frame = cap.read()
        if not ret: return
        
        results = self.face_mesh.process(cv2.cvtColor(cv2.flip(frame, 1), cv2.COLOR_BGR2RGB))
        if results.multi_face_landmarks:
            self._handle_face_landmarks(results.multi_face_landmarks[0].landmark)
            
    def _handle_face_landmarks(self, landmarks):
        current_time = time.time()
        
        # Calculate EAR early for blink detection
        left_ear = self._calculate_ear(landmarks, [362, 385, 387, 263, 373, 380])
        right_ear = self._calculate_ear(landmarks, [33, 160, 158, 133, 153, 144])
        avg_ear = (left_ear + right_ear) / 2
        is_blink_frame = avg_ear < self.ear_threshold
        
        # Only update cursor when eyes are open
        if not is_blink_frame:
            self._handle_cursor(landmarks)
        
        self._handle_blinks(landmarks, current_time, avg_ear)
        self._handle_eyebrow_gestures(landmarks, current_time)
        self._handle_mouth_gestures(landmarks, current_time)
        
    def _handle_cursor(self, landmarks):
        left_eye = landmarks[362]
        right_eye = landmarks[133]
        new_x = (left_eye.x + right_eye.x) / 2
        new_y = (left_eye.y + right_eye.y) / 2
        
        smooth_x, smooth_y = self._exponential_smoothing(new_x, new_y)
        delta_x = (smooth_x - 0.5) * self.params['CURSOR_SENSITIVITY_X']
        delta_y = (smooth_y - 0.5) * self.params['CURSOR_SENSITIVITY_Y']
        
        if abs(delta_x) > self.params['MOVEMENT_THRESHOLD'] or abs(delta_y) > self.params['MOVEMENT_THRESHOLD']:
            cursor_x = self.screen_w * np.clip(0.5 + delta_x, 0, 1)
            cursor_y = self.screen_h * np.clip(0.5 + delta_y, 0, 1)
            pyautogui.moveTo(cursor_x, cursor_y, _pause=False)
            
    def _exponential_smoothing(self, new_x, new_y):
        alpha = self.params['EMA_ALPHA']
        smooth_x = alpha * new_x + (1 - alpha) * self.prev_smooth_x
        smooth_y = alpha * new_y + (1 - alpha) * self.prev_smooth_y
        self.prev_smooth_x, self.prev_smooth_y = smooth_x, smooth_y
        return smooth_x, smooth_y
    
    def _handle_blinks(self, landmarks, current_time, avg_ear):
        if avg_ear < self.ear_threshold:
            self.blink_counter += 1
        else:
            if self.blink_counter > 3:
                self._handle_possible_click(current_time)
                self.blink_counter = 0
            else:
                self.blink_counter = max(0, self.blink_counter - 1)
            
    def _handle_possible_click(self, current_time):
        if self.waiting_for_double:
            if (current_time - self.first_blink_time) <= self.params['DOUBLE_CLICK_THRESHOLD']:
                pyautogui.doubleClick()
                self.waiting_for_double = False
            else:
                pyautogui.click()
                self.first_blink_time = current_time
        else:
            self.waiting_for_double = True
            self.first_blink_time = current_time
            
    def _handle_eyebrow_gestures(self, landmarks, current_time):
        left_dist = landmarks[362].y - np.mean([landmarks[i].y for i in [276, 283, 282, 295, 285]])
        right_dist = landmarks[133].y - np.mean([landmarks[i].y for i in [46, 53, 52, 65, 55]])
        
        if (left_dist > self.calibration_eyebrow_left * self.params['EYEBROW_RAISE_THRESHOLD_RATIO'] or
            right_dist > self.calibration_eyebrow_right * self.params['EYEBROW_RAISE_THRESHOLD_RATIO']):
            if (current_time - self.last_right_click_time) > self.params['RIGHT_CLICK_COOLDOWN']:
                pyautogui.rightClick()
                self.last_right_click_time = current_time
                
    def _handle_mouth_gestures(self, landmarks, current_time):
        current_mouth = self._calculate_mouth_open(landmarks)
        mouth_open = current_mouth > self.mouth_open_threshold
        
        if mouth_open:
            if not self.scrolling:
                if self.initial_smooth_x is None:
                    self.initial_smooth_x = self.prev_smooth_x
                    self.initial_smooth_y = self.prev_smooth_y
                else:
                    delta_x = (self.prev_smooth_x - self.initial_smooth_x) * self.params['CURSOR_SENSITIVITY_X']
                    delta_y = (self.prev_smooth_y - self.initial_smooth_y) * self.params['CURSOR_SENSITIVITY_Y']
                    
                    if abs(delta_x) > self.params['SCROLL_THRESHOLD'] or abs(delta_y) > self.params['SCROLL_THRESHOLD']:
                        if abs(delta_x) > abs(delta_y):
                            self.scroll_direction = 'right' if delta_x > 0 else 'left'
                        else:
                            self.scroll_direction = 'down' if delta_y > 0 else 'up'
                        self.scrolling = True
                        self.last_scroll_time = current_time
            else:
                if (current_time - self.last_scroll_time) >= self.params['SCROLL_INTERVAL']:
                    self._execute_scroll()
                    self.last_scroll_time = current_time
        else:
            if self.scrolling:
                self.scrolling = False
                self.scroll_direction = None
                self.initial_smooth_x = None
                self.initial_smooth_y = None
            
    def _execute_scroll(self):
        if self.scroll_direction == 'up':
            pyautogui.scroll(self.params['SCROLL_STEP'])
        elif self.scroll_direction == 'down':
            pyautogui.scroll(-self.params['SCROLL_STEP'])
        elif self.scroll_direction == 'left':
            pyautogui.hscroll(-self.params['SCROLL_STEP'])
        elif self.scroll_direction == 'right':
            pyautogui.hscroll(self.params['SCROLL_STEP'])
            
    def _calculate_ear(self, landmarks, eye_indices):
        points = [landmarks[i] for i in eye_indices]
        vertical1 = np.linalg.norm([points[1].x - points[5].x, points[1].y - points[5].y])
        vertical2 = np.linalg.norm([points[2].x - points[4].x, points[2].y - points[4].y])
        horizontal = np.linalg.norm([points[0].x - points[3].x, points[0].y - points[3].y])
        return (vertical1 + vertical2) / (2 * horizontal)
    
    def _calculate_mouth_open(self, landmarks):
        return abs(landmarks[13].y - landmarks[14].y)
    
    def stop(self):
        self.running = False
        self.join()
        
    def toggle_pause(self):
        self.paused = not self.paused
        
    def update_param(self, key, value):
        self.params[key] = value