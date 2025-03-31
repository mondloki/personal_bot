import win32api
import win32gui
import win32con
import time
import winsound
import threading
import pythoncom
import win32com.client
import sys

class WindowsAlarm:
    def __init__(self, duration, message="Your alarm triggered!", sound_frequency=440, sound_duration=1000):
        """
        Initialize the Windows Alarm
        
        :param duration: Timer duration in seconds
        :param message: Notification message
        :param sound_frequency: Frequency of the alarm sound
        :param sound_duration: Duration of the sound in milliseconds
        """
        self.duration = duration
        self.message = message
        self.sound_frequency = sound_frequency
        self.sound_duration = sound_duration

    def show_notification(self):
        """
        Display a Windows pop-up notification
        """
        try:
            # Ensure COM is initialized for this thread
            pythoncom.CoInitialize()
            
            try:
                # Use Windows Shell for notifications
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.Popup(self.message, 0, "Alarm Notification", 0x0 | 0x30)
            except Exception as e:
                print(f"Notification error: {e}")
            finally:
                # Always uninitialize COM
                pythoncom.CoUninitialize()
        except Exception as e:
            print(f"COM initialization error: {e}")

    def play_alarm_sound(self):
        """
        Play the alarm sound
        """
        winsound.Beep(self.sound_frequency, self.sound_duration)

    def start_alarm(self):
        """
        Start the alarm timer and trigger notification
        """
        print(f"Alarm set for {self.duration} seconds")
        
        # Ensure COM is initialized for the main thread
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
        except Exception as e:
            print(f"Main thread COM initialization error: {e}")
        
        try:
            # Wait for the specified duration
            time.sleep(self.duration)
            
            # Create threads for notification and sound
            notification_thread = threading.Thread(target=self.show_notification)
            sound_thread = threading.Thread(target=self.play_alarm_sound)
            
            # Start both threads
            notification_thread.start()
            sound_thread.start()
            
            # Wait for threads to complete
            notification_thread.join()
            sound_thread.join()
        except Exception as e:
            print(f"Alarm error: {e}")
        finally:
            # Uninitialize COM for the main thread
            pythoncom.CoUninitialize()

def set_alarm(duration, message="Your alarm triggered!"):
    """
    Convenience function to set an alarm
    
    :param duration: Timer duration in seconds
    :param message: Optional custom notification message
    """
    alarm = WindowsAlarm(duration, message)
    alarm.start_alarm()

# Example usage
def main():
    # Set an alarm for 5 seconds with a custom message
    set_alarm(5, "Time's up! Your countdown is complete.")

if __name__ == "__main__":
    # Use pythoncom to enable COM threading support
    try:
        main()
    except KeyboardInterrupt:
        print("\nAlarm interrupted by user.")
        sys.exit(0)