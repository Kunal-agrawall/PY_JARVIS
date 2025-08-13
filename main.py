import tkinter as tk
from tkinter import ttk
import threading
import speech_recognition as sr
import pyttsx3
import webbrowser
from datetime import datetime
import os
import platform
import subprocess
import time
import traceback
from PIL import Image, ImageTk
from urllib.parse import quote_plus

# ---------- TTS ----------
tts_engine = pyttsx3.init()
tts_engine.setProperty("rate", 170)
tts_engine.setProperty("volume", 1.0)

def speak(text: str):
    """Speak and block until done."""
    print(f"[JARVIS SPEAKS] {text}")
    try:
        tts_engine.say(text)
        tts_engine.runAndWait()
    except Exception as e:
        print("TTS error:", e)

# ---------- Recognizer helper (fresh recognizer per attempt) ----------
def recognize_from_mic(timeout=None, phrase_time_limit=None, ambient_duration=0.4):
    """
    Listen once from microphone and return:
      - transcript string on success,
      - "" if unintelligible or timeout,
      - None if service unreachable (RequestError).
    Uses a fresh Recognizer each call to avoid stale internal state.
    """
    local_recognizer = sr.Recognizer()
    # Tune pause/energy thresholds (tweak if necessary)
    local_recognizer.pause_threshold = 0.8
    try:
        with sr.Microphone() as source:
            # short ambient noise calibration
            try:
                local_recognizer.adjust_for_ambient_noise(source, duration=ambient_duration)
            except Exception as e:
                # non-fatal: continue
                print("Ambient adjust error:", e)
            try:
                audio = local_recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
            except sr.WaitTimeoutError:
                return ""  # no speech in time window
    except OSError as e:
        print("Microphone error:", e)
        return None

    try:
        return local_recognizer.recognize_google(audio)
    except sr.RequestError:
        # API unreachable / network error
        return None
    except sr.UnknownValueError:
        # couldn't understand
        return ""

# ---------- App resolver (map friendly names to executables) ----------
WINDOWS_APP_MAP = {
    "wordpad": "write.exe",
    "notepad": "notepad.exe",
    "calculator": "calc.exe",
    "calc": "calc.exe",
    "chrome": "chrome.exe",
    "edge": "msedge.exe",

}

def resolve_executable(app_name: str):
    name = app_name.strip().lower()
    if platform.system().lower().startswith("win"):
        return WINDOWS_APP_MAP.get(name, app_name)
    # macOS: use the app display name
    if platform.system() == "Darwin":
        return app_name.title()
    # Linux: return the name (assume it's an executable in PATH)
    return app_name

def open_app(app_name: str):
    exe = resolve_executable(app_name)
    system = platform.system()
    try:
        if system.lower().startswith("win"):
            # try startfile (works for exe or file associations)
            try:
                os.startfile(exe)
            except Exception:
                # fallback: try Popen (may require full path or in PATH)
                subprocess.Popen([exe], shell=False)
        elif system == "Darwin":
            subprocess.Popen(["open", "-a", exe])
        else:
            subprocess.Popen([exe])
        speak(f"Opening {app_name}")
    except Exception as e:
        print("open_app error:", e)
        speak(f"Unable to open {app_name}")

def close_app(app_name: str):
    exe = resolve_executable(app_name)
    system = platform.system()
    try:
        if system.lower().startswith("win"):
            # ensure .exe when calling taskkill
            proc_name = exe if exe.lower().endswith(".exe") else f"{exe}.exe"
            os.system(f'taskkill /f /im "{proc_name}" >nul 2>&1')
        else:
            # pkill by process name (may need exact name)
            os.system(f"pkill -f {exe} > /dev/null 2>&1")
        speak(f"{app_name} has been closed.")
    except Exception as e:
        print("close_app error:", e)
        speak(f"Unable to close {app_name}")

# ---------- Command handling ----------
def take_command(message: str):
    """Process a single command. Exceptions are thrown to caller if needed."""
    if not message:
        speak("Sorry, I did not understand. Please try again.")
        return

    m = message.lower().strip()
    print(f"[JARVIS HEARD] {m}")

    if m in ["open camera", "start camera"]:
        os.system("start microsoft.windows.camera:")
        speak("Opening Camera")
        return
    if m in ["close camera"]:
        os.system("taskkill /f /im WindowsCamera.exe")
        speak("Camera closed")
        return

    if m in ["open snipping tool"]:
        os.system("start ms-screenclip:")
        speak("Opening Snipping Tool")
        return
    if m in ["close snipping tool"]:
        os.system("taskkill /f /im SnippingTool.exe")
        speak("Snipping Tool closed")
        return

    if m in ["open brave"]:
        os.system("start brave")
        speak("Opening Brave Browser")
        return
    if m in ["close brave"]:
        os.system("taskkill /f /im brave.exe")
        speak("Brave Browser closed")
        return

    if m in ["open paint"]:
        os.system("start mspaint")
        speak("Opening Paint")
        return
    if m in ["close paint"]:
        os.system("taskkill /f /im mspaint.exe")
        speak("Paint closed")
        return

    if m in ["open calculator"]:
        os.system("start calculator:")
        speak("Opening Calculator")
        return

    if m in ["close calculator"]:
        # Try killing both possible Calculator process names
        os.system("taskkill /f /im CalculatorApp.exe")
        os.system("taskkill /f /im Win32Calc.exe")
        speak("Calculator closed")
        return

    if m in ["open ms word", "open microsoft word"]:
        os.system("start winword")
        speak("Opening Microsoft Word")
        return
    if m in ["close ms word", "close microsoft word"]:
        os.system("taskkill /f /im WINWORD.EXE")
        os.system("taskkill /f /im winword.exe")
        speak("Microsoft Word closed")
        return


    if m in ["open ms excel", "open microsoft excel"]:
        os.system("start excel")
        speak("Opening Microsoft Excel")
        return
    if m in ["close ms excel", "close microsoft excel"]:
        os.system("taskkill /f /im excel.exe")
        speak("Microsoft Excel closed")
        return

    if m in ["open ms powerpoint", "open microsoft powerpoint"]:
        os.system("start powerpnt")
        speak("Opening Microsoft PowerPoint")
        return
    if m in ["close ms powerpoint", "close microsoft powerpoint"]:
        os.system("taskkill /f /im powerpnt.exe")
        speak("Microsoft PowerPoint closed")
        return

    if m in ["open whatsapp"]:
        os.system("start whatsapp")
        speak("Opening WhatsApp")
        return
    if m in ["close whatsapp"]:
        os.system("taskkill /f /im whatsapp.exe")
        speak("WhatsApp closed")
        return

    if m in ["open microsoft store", "open store"]:
        os.system("start ms-windows-store:")
        speak("Opening Microsoft Store")
        return
    if m in ["close microsoft store", "close store"]:
        os.system("taskkill /f /im WinStore.App.exe")
        speak("Microsoft Store closed")
        return

    if m in ["open windows explorer", "open file explorer"]:
        os.system("start explorer")
        speak("Opening File Explorer")
        return
    if m in ["close windows explorer", "close file explorer"]:
        os.system("taskkill /f /im explorer.exe")
        speak("Closing Windows Explorer")
        os.system("start explorer")
        return

    if m in ["open all browsers"]:
        os.system("start chrome")
        os.system("start brave")
        os.system("start msedge")
        speak("Opening all browsers")
        return
    if m in ["close all browsers"]:
        os.system("taskkill /f /im chrome.exe")
        os.system("taskkill /f /im brave.exe")
        os.system("taskkill /f /im msedge.exe")
        speak("All browsers closed")
        return

    if m in ["close all previous apps", "close all apps"]:
        for proc in [
            "winword.exe", "excel.exe", "powerpnt.exe", "whatsapp.exe",
            "mspaint.exe", "brave.exe", "chrome.exe", "msedge.exe",
            "WindowsCamera.exe", "SnippingTool.exe", "WinStore.App.exe"
        ]:
            os.system(f"taskkill /f /im {proc}")
        speak("All previous apps have been closed") 
        return

    if m in ["bye-bye jarvis","bye bye jarvis", "exit jarvis", "close jarvis", "shutdown jarvis", "stop jarvis", "terminate yourself"]:
        speak("Goodbye Sir. Shutting down JARVIS.")
        os._exit(0)

    # Basic conversational commands
    if "hey" in m or "hello" in m:
        speak("Hello Sir, how may I help you?")
        return
    if "who am i" in m or "who created you" in m or "who is your master" in m:
        speak("You are Kunal Agrawal, my master, who created me.")
        return
    if "open google" in m:
        webbrowser.open("https://google.com")
        speak("Opening Google")
        return
    if "open youtube" in m:
        webbrowser.open("https://youtube.com")
        speak("Opening YouTube")
        return
    if "open facebook" in m:
        webbrowser.open("https://facebook.com")
        speak("Opening Facebook")
        return

    # universal open/close
    if m.startswith("open "):
        app = m.replace("open ", "", 1).strip()
        open_app(app)
        return
    if m.startswith("close "):
        app = m.replace("close ", "", 1).strip()
        close_app(app)
        return

    # searches
    if m.startswith("what is") or m.startswith("who is") or m.startswith("what are") or "search" in m:
        query = m.replace("search", "").strip()
        if not query:
            query = m
        url = "https://www.google.com/search?q=" + quote_plus(query)
        webbrowser.open(url)
        speak(f"This is what I found on the internet regarding {query}")
        return
    if "wikipedia" in m:
        topic = m.replace("wikipedia", "").strip()
        if not topic:
            speak("What would you like me to search on Wikipedia?")
            return
        url = "https://en.wikipedia.org/wiki/" + "_".join(topic.split())
        webbrowser.open(url)
        speak(f"This is what I found on Wikipedia regarding {topic}")
        return
    if "time" in m:
        speak(f"The current time is {datetime.now().strftime('%I:%M %p')}")
        return
    if "date" in m:
        speak(f"Today's date is {datetime.now().strftime('%b %d')}")
        return

    # fallback
    url = "https://www.google.com/search?q=" + quote_plus(m)
    webbrowser.open(url)
    speak(f"I found some information for {m} on Google")

# ---------- GUI + Continuous listener ----------
class JarvisGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("J A R V I S - Virtual Assistant")
        self.configure(bg="black")
        self.geometry("600x750")
        self.resizable(False, False)

        self.image_frame = ttk.Frame(self)
        self.image_frame.pack(pady=12)

        # GIF animation (Pillow)
        self.frames = []
        try:
            gif_path = "giphy.gif"
            pil_image = Image.open(gif_path)
            try:
                while True:
                    frame = ImageTk.PhotoImage(pil_image.copy())
                    self.frames.append(frame)
                    pil_image.seek(len(self.frames))
            except EOFError:
                pass

            if self.frames:
                self.gif_label = tk.Label(self.image_frame, bg="black")
                self.gif_label.pack()
                self.current_frame = 0
                self.animate_gif()
            else:
                raise FileNotFoundError("No frames in GIF")

        except Exception as e:
            print("GIF load failed:", e)
            try:
                self.avatar = ImageTk.PhotoImage(Image.open("avatar.png"))
                self.gif_label = tk.Label(self.image_frame, image=self.avatar, bg="black")
                self.gif_label.pack()
            except Exception as e2:
                print("Avatar load failed:", e2)

        title = tk.Label(self, text="J A R V I S", fg="#00bcd4", bg="black", font=("Helvetica", 28, "bold"))
        title.pack(pady=(8,0))

        subtitle = tk.Label(self, text="I am your Virtual Assistant. How may I help you?", fg="#9fbfbf", bg="black", font=("Helvetica", 10))
        subtitle.pack(pady=(6,18))

        self.content_var = tk.StringVar(value="Starting...")
        self.content_label = tk.Label(self, textvariable=self.content_var, fg="#aed0d0", bg="#000000", font=("Helvetica", 14))
        self.content_label.pack(pady=6)

        self.status_var = tk.StringVar(value="")
        self.status_label = tk.Label(self, textvariable=self.status_var, fg="#9fbfbf", bg="black", font=("Helvetica", 10))
        self.status_label.pack(pady=(6,0))

        # start welcome + continuous listener in separate thread
        self.after(600, lambda: threading.Thread(target=self.initialize_welcome, daemon=True).start())

    def animate_gif(self):
        if self.frames:
            frame = self.frames[self.current_frame]
            self.gif_label.configure(image=frame)
            self.current_frame = (self.current_frame + 1) % len(self.frames)
            self.after(50, self.animate_gif)

    def initialize_welcome(self):
        speak("Initializing JARVIS")
        self.wish_me()
        # Start continuous listening loop (daemon thread)
        threading.Thread(target=self.start_continuous_listening, daemon=True).start()

    def wish_me(self):
        hour = datetime.now().hour
        if 0 <= hour < 12:
            speak("Good Morning Boss. How may I help you?")
        elif 12 <= hour < 17:
            speak("Good Afternoon Sir. How may I help you?")
        else:
            speak("Good Evening Sir. How may I help you?")

    def start_continuous_listening(self):
        """
        Main listening loop. Very defensive: catches exceptions, ensures the loop continues,
        and inserts small pauses after speaking to avoid TTS->mic feedback.
        """
        while True:
            try:
                self.content_var.set("Listening...")
                transcript = recognize_from_mic(timeout=6, phrase_time_limit=7, ambient_duration=0.4)

                if transcript is None:
                    # API / microphone unreachable -> notify and wait a bit
                    print("Recognition service unreachable or microphone error.")
                    self.content_var.set("Speech service unreachable.")
                    speak("Speech recognition service is unavailable. Check your microphone or internet.")
                    time.sleep(2)
                    continue

                if transcript == "":
                    # Could not understand — keep listening
                    print("Unintelligible or timeout; will keep listening.")
                    self.content_var.set("Could not understand. Listening again...")
                    # tiny pause to avoid tight loop
                    time.sleep(0.3)
                    continue

                # Good transcript -> process it
                self.content_var.set(transcript)
                try:
                    take_command(transcript)
                except Exception as e:
                    # If a command fails, report and keep listening
                    print("Error while processing command:", e)
                    traceback.print_exc()
                    speak("I encountered an error while processing the command.")
                # small pause after speaking before next listen (avoid TTS bleed)
                time.sleep(0.35)

            except Exception as e:
                # Catch anything unexpected, keep loop alive
                print("Listener loop unexpected error:", e)
                traceback.print_exc()
                self.content_var.set("Listener error, restarting...")
                time.sleep(1)  # brief backoff and continue

# ---------- run ----------
if __name__ == "__main__":
    app = JarvisGUI()
    app.mainloop()
