import os
import sys
import time
import re
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

log_file = os.path.join(tempfile.gettempdir(), "filemonitor.log")
handler = RotatingFileHandler(log_file, maxBytes=100*1024, backupCount=0)
handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))

logging.basicConfig(level=logging.INFO, handlers=[handler])

class FileHandler(FileSystemEventHandler):
    def __init__(self, config, downloads_dir):
        self.patterns = config['patterns']
        self.default_printer = config.get('default_printer')
        self.downloads_dir = downloads_dir
        
    def notify(self, title: str, message: str, *, success: bool = True) -> None:
        toast = Notification(
            app_id="AutoPrint and Archive",
            title=title,
            msg=message,
            duration="short"
        )
        toast.set_audio(audio.Default, loop=False)
        toast.add_actions(label="View Log", launch=log_file)
        toast.show()
        
    def on_created(self, event):
        if event.is_directory:
            return
        logging.info(f"New file detected: {event.src_path}")
        filename = os.path.basename(event.src_path)
        
        matched_pattern = None
        for pattern in self.patterns:
            match = re.match(pattern['pattern'], filename)
            if match:
                matched_pattern = (pattern, match)
                break
                
        if not matched_pattern:
            logging.info(f"No matching pattern found for {filename}")
            return
            
        self._process_file(event.src_path, matched_pattern[0], matched_pattern[1])
        
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
        retries = 10
        while retries > 0 and self.is_file_locked(file_path):
            time.sleep(0.5)
            retries -= 1
        
        if self.is_file_locked(file_path):
            error_msg = f"File {file_path} is locked after 5 seconds"
            logging.error(error_msg)
            self.notify("File Processing Error", error_msg, success=False)
            return
            
        filename = os.path.basename(file_path)
        logging.info(f"File {filename} matches pattern: {pattern['pattern']}")
        dest = pattern['destination'].format(**match.groupdict())
        dest_path = os.path.join(dest, filename)
        
        time.sleep(1)
        
        if pattern.get('print', False):
            printer_name = pattern.get('printer', self.default_printer)
            logging.info(f"Printing file to {printer_name}")
            try:
                if printer_name:
                    win32print.SetDefaultPrinter(printer_name)
                win32api.ShellExecute(0, "print", file_path, None, ".", 0)
                self.notify("Print Job Started", f"Printing {filename} to {printer_name}")
                time.sleep(10)
                self._wait_for_print(filename)
            except Exception as e:
                error_msg = f"Print error for {filename}: {e}"
                logging.error(error_msg)
                self.notify("Print Error", error_msg, success=False)
        
        if os.path.exists(dest_path):
            msg = f"File already exists at destination: {dest_path}"
            logging.info(msg)
            self.notify("File Already Exists", msg, success=False)
            return
            
        time.sleep(5)
        self._move_file(file_path, dest)

    def _wait_for_print(self, filename: str) -> None:
        logging.info("Waiting for print job...")
        base_filename = os.path.splitext(filename)[0].lower()
        
        start_time = time.time()
        last_jobs_count = 0
        stable_count = 0
        
        while True:
            current_jobs_count = 0
            for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL):
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
                    logging.debug(f"Printer error: {e}")
                    
            if current_jobs_count == last_jobs_count:
                stable_count += 1
            else:
                stable_count = 0
                
            if stable_count >= 3 or time.time() - start_time > 60:
                break
                
            last_jobs_count = current_jobs_count
            time.sleep(1)
            
    def _move_file(self, file_path: str, destination: str) -> None:
        dest_path = os.path.join(destination, os.path.basename(file_path))
        os.makedirs(destination, exist_ok=True)
        
        retries = 3
        while retries > 0:
            try:
                os.rename(file_path, dest_path)
                success_msg = f"Moved {os.path.basename(file_path)} to {destination}"
                logging.info(success_msg)
                self.notify("File Moved", success_msg)
                break
            except Exception as e:
                retries -= 1
                if retries == 0:
                    error_msg = f"Failed to move {os.path.basename(file_path)} after 3 attempts: {e}"
                    logging.error(error_msg)
                    self.notify("Move Error", error_msg, success=False)
                else:
                    time.sleep(1)

class FileMonitor:
    def __init__(self):
        self.observer = None
        self.icon = None
        self.running = True
        
    def create_icon(self):
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
    
    def open_log(self):
        os.startfile(log_file)
    
    def stop_monitoring(self):
        self.running = False
        if self.observer:
            self.observer.stop()
        self.icon.stop()
    
    def start_monitoring(self):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        config_path = os.path.join(exe_dir, "config.yaml")
        downloads_dir = os.path.expanduser("~/Downloads")
        
        try:
            with open(config_path, 'r') as f:
                config = yaml.safe_load(f)
                
            logging.info(f"Monitoring: {downloads_dir}")
            logging.info(f"Default printer: {config.get('default_printer', 'Not set')}")
            
            event_handler = FileHandler(config, downloads_dir)
            self.observer = Observer()
            self.observer.schedule(event_handler, downloads_dir, recursive=False)
            self.observer.start()
            
            self.create_icon()
            self.icon.run()
            
        except Exception as e:
            logging.error(f"Error: {e}")
            if self.observer:
                self.observer.stop()
            sys.exit(1)

def main():
    monitor = FileMonitor()
    monitor.start_monitoring()

if __name__ == "__main__":
    main()