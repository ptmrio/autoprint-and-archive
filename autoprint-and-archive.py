import os
import sys
import time
import re
import queue
import threading
from threading import Thread
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import win32print
import win32api
from pathlib import Path
import yaml
import pystray
from PIL import Image
import logging
import tempfile
from winotify import Notification, audio
from logging.handlers import RotatingFileHandler
import shutil
import tkinter as tk
from tkinter import messagebox
import json

log_file = os.path.join(tempfile.gettempdir(), "filemonitor.log")
handler = RotatingFileHandler(log_file, maxBytes=100*1024, backupCount=0)
handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))

logging.basicConfig(level=logging.INFO, handlers=[handler])

class Translator:
    """Simple translation class to support multiple languages"""
    
    def __init__(self, language='en', locales_dir=None):
        self.language = language
        self.translations = {}
        self.locales_dir = locales_dir or self._get_locales_dir()
        self._load_translations()
    
    def _get_locales_dir(self):
        """Get the locales directory relative to the executable/script"""
        if getattr(sys, 'frozen', False):
            # Running as compiled executable - PyInstaller extracts to _MEIPASS
            base_path = sys._MEIPASS
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, 'locales')
    
    def _load_translations(self):
        """Load translations for the selected language"""
        translation_file = os.path.join(self.locales_dir, f"{self.language}.json")
        
        try:
            if os.path.exists(translation_file):
                with open(translation_file, 'r', encoding='utf-8') as f:
                    self.translations = json.load(f)
                logging.info(f"Loaded translations for language: {self.language}")
                logging.info(f"Translation file path: {translation_file}")  # Debug info
            else:
                logging.warning(f"Translation file not found: {translation_file}, falling back to English")
                # Try to load English as fallback
                fallback_file = os.path.join(self.locales_dir, "en.json")
                if os.path.exists(fallback_file):
                    with open(fallback_file, 'r', encoding='utf-8') as f:
                        self.translations = json.load(f)
                    logging.info(f"Loaded fallback English translations from: {fallback_file}")
                else:
                    logging.error("No translation files found, using empty translations")
                    logging.error(f"Looked in directory: {self.locales_dir}")
                    # List what's actually in the directory for debugging
                    if os.path.exists(self.locales_dir):
                        files = os.listdir(self.locales_dir)
                        logging.error(f"Files in locales directory: {files}")
                    else:
                        logging.error(f"Locales directory does not exist: {self.locales_dir}")
                    self.translations = {}
        except Exception as e:
            logging.error(f"Error loading translations: {e}")
            self.translations = {}
    
    def t(self, key, **kwargs):
        """Translate a key with optional formatting parameters"""
        text = self.translations.get(key, key)  # Fallback to key if translation not found
        try:
            return text.format(**kwargs) if kwargs else text
        except KeyError as e:
            logging.warning(f"Missing format parameter {e} for translation key '{key}'")
            return text

class PrintPromptDialog:
    def __init__(self, filename, archived_path, translator):
        self.result = None
        self.filename = filename
        self.archived_path = archived_path
        self.translator = translator
        
    def show(self):
        """Show the print prompt dialog and return user choice"""
        try:
            # Create root window (hidden)
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            root.attributes('-topmost', True)  # Bring to front
            
            # Show simple yes/no dialog for printing
            result = messagebox.askyesno(
                self.translator.t("print_document_title"),
                self.translator.t("print_document_message", filename=self.filename),
                icon='question'
            )
            
            root.destroy()
            
            return result  # True for Yes, False for No
                
        except Exception as e:
            logging.error(f"Error showing print prompt: {e}")
            return False  # Default to no printing on error

class FileHandler(FileSystemEventHandler):
    def __init__(self, config, downloads_dir, translator):
        self.patterns = config.get('patterns', [])
        self.default_printer = config.get('default_printer')
        self.downloads_dir = downloads_dir
        self.translator = translator
        self.processing_queue = queue.Queue()
        self.processing_thread = Thread(target=self._process_queue, daemon=True)
        self.processing_thread.start()
        # De-dupe recent events by absolute path (case-insensitive)
        self.recent_events = {}  # {normalized_path: last_seen_ts}
        self.recent_events_ttl = int(config.get('dedupe_ttl_seconds', 30))
        self.processed_files_lock = threading.Lock()

    def _process_queue(self):
        while True:
            try:
                file_info = self.processing_queue.get()
                if file_info is None:
                    break
                self._process_file(*file_info)
            except Exception as e:
                logging.error(f"Error processing queued file: {e}")
            finally:
                self.processing_queue.task_done()

    def _handle_file(self, file_path, event_type):
        filename = os.path.basename(file_path)

        # 1) Pattern-gate first
        matched_pattern = None
        for pattern in self.patterns:
            m = re.match(pattern['pattern'], filename)
            if m:
                matched_pattern = (pattern, m)
                break

        if not matched_pattern:
            # No noise for non-matching files (incl. .crdownload/.tmp)
            logging.debug(f"Ignoring {event_type} (no pattern match): {filename}")
            return

        # 2) De-dupe only for relevant (matched) files
        normalized = os.path.normcase(os.path.abspath(file_path))
        now = time.time()
        with self.processed_files_lock:
            # prune old entries
            for p, t in list(self.recent_events.items()):
                if now - t > self.recent_events_ttl:
                    del self.recent_events[p]
            if normalized in self.recent_events:
                logging.debug(f"Ignoring duplicate {event_type} for {file_path}")
                return
            self.recent_events[normalized] = now

        # 3) Log only for files we will actually process
        logging.info(f"File {event_type}: {file_path}")

        self.processing_queue.put((file_path, matched_pattern[0], matched_pattern[1]))

    def on_created(self, event):
        """Handle file creation events"""
        if event.is_directory:
            return
        self._handle_file(event.src_path, "created")

    def on_moved(self, event):
        """Handle file rename/move events"""
        if event.is_directory:
            return
        self._handle_file(event.dest_path, "moved")

    def stop(self):
        self.processing_queue.put(None)
        self.processing_thread.join()
        
    def notify(self, title: str, message: str, *, success: bool = True) -> None:
        try:
            toast = Notification(
                app_id="AutoPrint and Archive",
                title=title,
                msg=message,
                duration="short"
            )
            toast.set_audio(audio.Default, loop=False)
            toast.add_actions(label=self.translator.t("view_log"), launch=log_file)
            toast.show()
        except Exception as e:
            logging.error(f"Notification error: {e}")
        
    def is_file_locked(self, filepath: str) -> bool:
        try:
            if not os.path.exists(filepath):
                return True
            with open(filepath, 'r+b') as f:
                f.seek(0, os.SEEK_END)
                f.tell()
                return False
        except (IOError, PermissionError):
            return True
        
    def _process_file(self, file_path: str, pattern: dict, match: re.Match) -> None:
        filename = os.path.basename(file_path)
        logging.info(f"File {filename} matches pattern: {pattern['pattern']}")
        
        # Wait for file to be ready and not locked
        retries = 10
        while retries > 0 and self.is_file_locked(file_path):
            time.sleep(0.5)
            retries -= 1

        # Final check - if file disappeared during the wait or is still locked
        if not os.path.exists(file_path):
            logging.info(f"File disappeared during processing, skipping: {file_path}")
            return
            
        if self.is_file_locked(file_path):
            error_msg = self.translator.t("file_locked_error", filepath=file_path)
            logging.error(error_msg)
            self.notify(self.translator.t("file_processing_error"), error_msg, success=False)
            return
            
        dest = pattern['destination'].format(**match.groupdict())
        dest_path = os.path.join(dest, filename)
        
        time.sleep(1)
        
        # Check if file already exists at destination
        if os.path.exists(dest_path):
            msg = self.translator.t("file_already_exists_message", dest_path=dest_path)
            logging.info(msg)
            self.notify(self.translator.t("file_already_exists"), msg, success=False)
            return
            
        # Always archive first
        time.sleep(5)
        archive_success = self._move_file(file_path, dest)
        
        if not archive_success:
            logging.error(f"Failed to archive {filename}, skipping print prompt")
            return
        
        # After successful archiving, check print setting
        print_setting = pattern.get('print', False)
        
        if print_setting == 'prompt':
            # Show print prompt dialog
            dialog = PrintPromptDialog(filename, dest_path, self.translator)
            user_wants_print = dialog.show()
            
            if user_wants_print:
                logging.info(f"User chose to print {filename}")
                self._print_file(dest_path, pattern)
            else:
                logging.info(f"User chose not to print {filename}")
                self.notify(self.translator.t("file_archived"), 
                          self.translator.t("file_archived_without_printing", filename=filename))
        elif print_setting is True or print_setting == 'true':
            # Always print
            logging.info(f"Auto-printing {filename}")
            self._print_file(dest_path, pattern)
        else:
            # print is False or not set - archive only
            logging.info(f"File archived (printing disabled): {filename}")
    
    def _print_file(self, file_path: str, pattern: dict) -> None:
        """Handle the actual printing process"""
        filename = os.path.basename(file_path)
        printer_name = pattern.get('printer', self.default_printer)
        
        if not printer_name:
            logging.info("No printer specified, skipping printing.")
            return
            
        original_printer = win32print.GetDefaultPrinter()
        printer_changed = False
        
        # Only change printer if different from the current default
        if original_printer.lower() != printer_name.lower():
            try:
                win32print.SetDefaultPrinter(printer_name)
                printer_changed = True
            except Exception as e:
                logging.error(f"Could not set printer {printer_name}: {e}")

        try:
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)
            self.notify(self.translator.t("print_job_started"), 
                       self.translator.t("printing_to", filename=filename, printer=printer_name))
            time.sleep(10)
            self._wait_for_print(filename)
        except Exception as e:
            error_msg = self.translator.t("print_error_message", filename=filename, error=str(e))
            logging.error(error_msg)
            self.notify(self.translator.t("print_error"), error_msg, success=False)
        finally:
            if printer_changed:
                # Restore original printer
                try:
                    win32print.SetDefaultPrinter(original_printer)
                except Exception as e:
                    logging.error(f"Could not restore original printer {original_printer}: {e}")

    def _wait_for_print(self, filename: str) -> None:
        logging.info("Waiting for print job...")
        base_filename = os.path.splitext(filename)[0].lower()
        
        start_time = time.time()
        last_jobs_count = 0
        stable_count = 0
        
        while True:
            current_jobs_count = 0
            # If printers fail, just break after timeout
            try:
                printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            except Exception as e:
                logging.debug(f"Error enumerating printers: {e}")
                break
                
            for printer in printers:
                try:
                    jobs = win32print.EnumJobs(printer[2], 0, -1)
                    current_jobs_count += len(jobs)
                    for job in jobs:
                        doc = job.get('pDocument', '').lower()
                        if base_filename in doc:
                            logging.info(f"Found matching job: {doc}")
                            last_jobs_count = current_jobs_count
                            stable_count = 0
                            break
                except Exception as e:
                    logging.debug(f"Printer job enumeration error: {e}")
                    
            if current_jobs_count == last_jobs_count:
                stable_count += 1
            else:
                stable_count = 0
                
            if stable_count >= 3 or time.time() - start_time > 15:
                break
                
            last_jobs_count = current_jobs_count
            time.sleep(1)
            
    def _move_file(self, file_path: str, destination: str) -> bool:
        """Move file to destination and return success status"""
        dest_path = os.path.join(destination, os.path.basename(file_path))
        os.makedirs(destination, exist_ok=True)
        
        retries = 3
        while retries > 0:
            try:
                shutil.move(file_path, dest_path)
                success_msg = self.translator.t("moved_to", filename=os.path.basename(file_path), destination=destination)
                logging.info(success_msg)
                self.notify(self.translator.t("file_archived"), success_msg)
                return True
            except Exception as e:
                retries -= 1
                if retries == 0:
                    error_msg = self.translator.t("move_error_message", filename=os.path.basename(file_path), error=str(e))
                    logging.error(error_msg)
                    self.notify(self.translator.t("move_error"), error_msg, success=False)
                    return False
                else:
                    time.sleep(1)
        return False

class FileMonitor:
    def __init__(self):
        self.observer = None
        self.icon = None
        self.running = True
        
    def create_icon(self):
        try:
            image = Image.new('RGB', (64, 64), color='white')
            self.icon = pystray.Icon(
                "AutoPrintandArchive",
                image,
                "AutoPrint and Archive",
                menu=pystray.Menu(
                    pystray.MenuItem("Open Log", self.open_log),
                    pystray.MenuItem("Exit", self.stop_monitoring)
                )
            )
        except Exception as e:
            logging.error(f"Error creating system tray icon: {e}")
            self.icon = None
    
    def open_log(self):
        try:
            os.startfile(log_file)
        except Exception as e:
            logging.error(f"Cannot open log file: {e}")
    
    def stop_monitoring(self):
        self.running = False
        if self.observer:
            self.observer.stop()
            self.observer.join()
        if self.icon:
            self.icon.stop()
    
    def start_monitoring(self):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        config_path = os.path.join(exe_dir, "config.yaml")
        downloads_dir = os.path.expanduser("~/Downloads")
        
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = yaml.safe_load(f)
            else:
                logging.error("Config file not found. Monitoring not started.")
                return
            
            # Initialize translator with the configured language
            language = config.get('language', 'en')
            translator = Translator(language)
                
            logging.info(f"Monitoring: {downloads_dir}")
            logging.info(f"Default printer: {config.get('default_printer', 'Not set')}")
            logging.info(f"Language: {language}")
            
            event_handler = FileHandler(config, downloads_dir, translator)
            self.observer = Observer()
            self.observer.schedule(event_handler, downloads_dir, recursive=False)
            self.observer.start()
            
            self.create_icon()
            
            # Run icon only if successfully created
            if self.icon:
                try:
                    self.icon.run()
                except Exception as e:
                    logging.error(f"System tray icon error: {e}")
                    # Even if icon fails, keep monitoring
                    while self.running:
                        time.sleep(1)
            else:
                # If no icon, just keep running
                while self.running:
                    time.sleep(1)
            
        except Exception as e:
            logging.error(f"Error: {e}")
            if self.observer:
                self.observer.stop()
                self.observer.join()
            sys.exit(1)

def main():
    monitor = FileMonitor()
    monitor.start_monitoring()

if __name__ == "__main__":
    main()