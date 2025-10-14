"""Windows Automation Assistant

This module defines a local automation assistant capable of performing
common desktop tasks on a Windows machine. It relies solely on local
Python libraries and automation frameworks such as pyautogui and
selenium.

Before running the script, install the required dependencies:
    pip install pyautogui speechrecognition pyttsx3 vosk selenium
    pip install pillow requests

Some features such as speech recognition require additional setup,
including microphone access.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

try:
    import pyautogui
except ImportError:
    pyautogui = None  # type: ignore

try:
    import speech_recognition as sr
except ImportError:  # pragma: no cover - hardware-dependent
    sr = None  # type: ignore

try:
    import pyttsx3
except ImportError:
    pyttsx3 = None  # type: ignore

try:  # pragma: no cover - optional dependency
    from vosk import Model as VoskModel
    from vosk import KaldiRecognizer
except ImportError:  # pragma: no cover - optional dependency
    VoskModel = None  # type: ignore
    KaldiRecognizer = None  # type: ignore

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.common.by import By
except ImportError:
    webdriver = None  # type: ignore
    ChromeOptions = None  # type: ignore
    By = None  # type: ignore

import requests


BROWSER_SESSION: Optional["webdriver.Chrome"] = None


def log(message: str) -> None:
    """Print a timestamped log message to the console."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


def open_browser(url: str) -> None:
    """Open the default Chrome browser and navigate to the given URL."""
    log(f"Attempting to open browser at URL: {url}")
    try:
        if webdriver is None or ChromeOptions is None:
            log("Selenium is not available. Ensure selenium and a driver are installed.")
            return

        global BROWSER_SESSION
        if BROWSER_SESSION is not None:
            try:
                BROWSER_SESSION.get(url)
                log("Reused existing browser session.")
                return
            except Exception as session_exc:
                log(f"Existing browser session failed: {session_exc}. Reinitializing.")
                try:
                    BROWSER_SESSION.quit()
                except Exception:
                    pass
                BROWSER_SESSION = None

        options = ChromeOptions()
        options.add_argument("--start-maximized")
        BROWSER_SESSION = webdriver.Chrome(options=options)
        BROWSER_SESSION.get(url)
        log("Browser opened successfully.")
    except Exception as exc:  # pragma: no cover - depends on environment
        log(f"Failed to open browser: {exc}")


def launch_telegram() -> None:
    """Launch the Telegram desktop application."""
    log("Attempting to launch Telegram Desktop.")
    try:
        telegram_path = Path(os.getenv("LOCALAPPDATA", "")) / "Telegram Desktop" / "Telegram.exe"
        if telegram_path.exists():
            log(f"Launching Telegram from {telegram_path}")
            os.startfile(str(telegram_path))
        else:
            log("Telegram executable not found in the default location.")
    except Exception as exc:  # pragma: no cover - depends on environment
        log(f"Failed to launch Telegram: {exc}")


def read_folder(path: str) -> None:
    """List the files in the specified folder."""
    log(f"Reading contents of folder: {path}")
    try:
        folder = Path(path)
        if not folder.exists():
            log("Folder does not exist.")
            return

        for item in folder.iterdir():
            log(f" - {item.name}")
    except Exception as exc:
        log(f"Failed to read folder: {exc}")


def take_screenshot(output_dir: str = "screenshots") -> Optional[Path]:
    """Take a screenshot and save it in the specified directory."""
    log("Capturing screenshot.")
    if pyautogui is None:
        log("pyautogui is not available. Install it to enable screenshots.")
        return None

    try:
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        filename = output_path / f"screenshot_{int(time.time())}.png"
        screenshot = pyautogui.screenshot()
        screenshot.save(filename)
        log(f"Screenshot saved to {filename}")
        return filename
    except Exception as exc:  # pragma: no cover - GUI dependent
        log(f"Failed to take screenshot: {exc}")
        return None


def click_button(image_path: str, confidence: float = 0.9) -> None:
    """Locate a button by image and click it using pyautogui."""
    log(f"Attempting to click button using image: {image_path}")
    if pyautogui is None:
        log("pyautogui is not available. Install it to enable GUI automation.")
        return

    try:
        location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
        if location:
            pyautogui.click(location)
            log(f"Clicked button at position: {location}")
        else:
            log("Button image not found on screen.")
    except Exception as exc:  # pragma: no cover - GUI dependent
        log(f"Failed to click button: {exc}")


def fill_form_field(selector: str, value: str) -> None:
    """Fill a form field in the active browser window using Selenium."""
    log(f"Attempting to fill form field '{selector}' with value '{value}'")
    if webdriver is None or By is None:
        log("Selenium WebDriver is not available.")
        return

    try:
        if BROWSER_SESSION is None:
            log("No active browser session. Run 'open browser <url>' first.")
            return

        element = BROWSER_SESSION.find_element(By.CSS_SELECTOR, selector)
        element.clear()
        element.send_keys(value)
        log("Form field filled successfully.")
    except Exception as exc:  # pragma: no cover - depends on driver state
        log(f"Failed to fill form field: {exc}")


def send_http_message(endpoint: str, payload: dict) -> None:
    """Send a HTTP POST request to an API endpoint."""
    log(f"Sending HTTP POST request to {endpoint} with payload {payload}")
    try:
        response = requests.post(endpoint, json=payload, timeout=10)
        log(f"Response status: {response.status_code}")
        log(f"Response body: {response.text}")
    except Exception as exc:
        log(f"Failed to send HTTP request: {exc}")


def recognize_voice(timeout: int = 5) -> Optional[str]:  # pragma: no cover - hardware-dependent
    """Capture audio from the microphone and convert it to text."""
    log("Listening for voice input.")
    if sr is None:
        log("speech_recognition package is not available.")
        return None

    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        try:
            audio = recognizer.listen(source, timeout=timeout)
            text: Optional[str] = None

            # Prefer offline Vosk recognition if a model is configured.
            model_path = os.getenv("VOSK_MODEL_PATH")
            if model_path and Path(model_path).exists() and VoskModel is not None:
                try:
                    text = recognizer.recognize_vosk(audio, model=model_path)
                except Exception as exc:  # pragma: no cover - depends on model availability
                    log(f"Vosk recognition failed: {exc}")
            elif VoskModel is not None:
                log("Set the VOSK_MODEL_PATH environment variable to use offline voice recognition.")

            # Fallback to PocketSphinx if installed for offline recognition.
            if text is None and hasattr(recognizer, "recognize_sphinx"):
                try:
                    text = recognizer.recognize_sphinx(audio)
                except Exception as exc:  # pragma: no cover - depends on pocketsphinx
                    log(f"Sphinx recognition failed: {exc}")

            if text:
                log(f"Recognized voice command: {text}")
                return text

            log("No offline speech-to-text engine is configured. Install Vosk or PocketSphinx for voice commands.")
        except sr.WaitTimeoutError:
            log("No voice input detected in time.")
        except sr.UnknownValueError:
            log("Could not understand audio.")
        except sr.RequestError as exc:
            log(f"Speech recognition request failed: {exc}")
    return None


def speak_text(text: str) -> None:  # pragma: no cover - hardware-dependent
    """Speak text aloud using pyttsx3 if available."""
    if pyttsx3 is None:
        log("pyttsx3 is not installed.")
        return

    try:
        engine = pyttsx3.init()
        engine.say(text)
        engine.runAndWait()
        log(f"Spoke text: {text}")
    except Exception as exc:
        log(f"Failed to speak text: {exc}")


def open_application(executable_path: str, arguments: Optional[str] = None) -> None:
    """Launch any Windows application via subprocess."""
    log(f"Attempting to open application: {executable_path}")
    try:
        if not Path(executable_path).exists():
            log("Executable path does not exist.")
            return

        command = [executable_path]
        if arguments:
            command.extend(arguments.split())

        subprocess.Popen(command)
        log("Application launched successfully.")
    except Exception as exc:  # pragma: no cover - depends on environment
        log(f"Failed to launch application: {exc}")


def execute_command(command: str) -> None:
    """Execute a text command."""
    command = command.strip().lower()
    log(f"Processing command: {command}")

    if command.startswith("open browser"):
        _, _, url = command.partition("open browser")
        open_browser(url.strip() or "https://www.google.com")
    elif command.startswith("launch telegram"):
        launch_telegram()
    elif command.startswith("read folder"):
        _, _, folder = command.partition("read folder")
        read_folder(folder.strip() or str(Path.home()))
    elif command.startswith("take screenshot"):
        take_screenshot()
    elif command.startswith("click button"):
        _, _, image = command.partition("click button")
        click_button(image.strip())
    elif command.startswith("fill form"):
        parts = command.split("::")
        if len(parts) == 3:
            _, selector, value = parts
            fill_form_field(selector.strip(), value.strip())
        else:
            log("Invalid form command format. Use 'fill form::CSS_SELECTOR::VALUE'")
    elif command.startswith("open app"):
        parts = command.split("::")
        if len(parts) >= 2:
            executable = parts[1].strip()
            args = parts[2].strip() if len(parts) == 3 else None
            open_application(executable, args)
        else:
            log("Invalid app command format. Use 'open app::C:/Path/To/App.exe::optional arguments'")
    elif command.startswith("send http"):
        try:
            _, endpoint, payload = command.split("::", 2)
            json_payload = json.loads(payload)
            send_http_message(endpoint.strip(), json_payload)
        except ValueError:
            log("Invalid HTTP command format. Use 'send http::URL::{\"key\": \"value\"}'")
        except Exception as exc:
            log(f"Failed to parse HTTP payload: {exc}")
    elif command in {"exit", "quit"}:
        log("Exiting assistant loop.")
        sys.exit(0)
    else:
        log("Command not recognized.")


def main() -> None:
    """Run the assistant main loop, accepting voice or text commands."""
    log("Starting Windows Automation Assistant.")

    if sr is not None:
        log("Voice recognition available. Say 'stop listening' to end voice mode.")
    else:
        log("Voice recognition not available; defaulting to text commands.")

    while True:
        try:
            if sr is not None:
                voice_command = recognize_voice()
                if voice_command:
                    if voice_command.lower() == "stop listening":
                        log("Voice command loop terminated by user.")
                        break
                    execute_command(voice_command)
                    continue

            command = input("Enter command (or 'exit' to quit): ")
            execute_command(command)
        except KeyboardInterrupt:
            log("Assistant interrupted by user.")
            break
        except Exception as exc:
            log(f"Unexpected error: {exc}")

    if BROWSER_SESSION is not None:
        try:
            BROWSER_SESSION.quit()
        except Exception:
            pass

    log("Assistant stopped.")


if __name__ == "__main__":
    main()
