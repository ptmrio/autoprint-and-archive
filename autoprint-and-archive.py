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

from logging.handlers import RotatingFileHandler

log_file = os.path.join(tempfile.gettempdir(), "filemonitor.log")
handler = RotatingFileHandler(log_file, maxBytes=100*1024, backupCount=0)  # 100KB, no backups
handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))

logging.basicConfig(level=logging.INFO,
                   handlers=[handler])

class FileHandler(FileSystemEventHandler):
    def __init__(self, config, downloads_dir):
        self.patterns = config['patterns']
        self.default_printer = config.get('default_printer')
        self.downloads_dir = downloads_dir
        
    def on_created(self, event):
        if event.is_directory:
            return
        logging.info(f"New file detected: {event.src_path}")
        self._process_file(event.src_path)
        
    def _process_file(self, file_path):
        filename = os.path.basename(file_path)
        
        for pattern in self.patterns:
            match = re.match(pattern['pattern'], filename)
            if match:
                logging.info(f"File {filename} matches pattern: {pattern['pattern']}")
                dest = pattern['destination'].format(**match.groupdict())
                dest_path = os.path.join(dest, filename)
                
                if pattern.get('print', False):
                    printer_name = pattern.get('printer', self.default_printer)
                    logging.info(f"Printing file to {printer_name}")
                    try:
                        if printer_name:
                            win32print.SetDefaultPrinter(printer_name)
                        win32api.ShellExecute(0, "print", file_path, None, ".", 0)
                        time.sleep(3)
                        self._wait_for_print(filename)
                    except Exception as e:
                        logging.error(f"Print error: {e}")
                
                if os.path.exists(dest_path):
                    logging.info(f"File already exists at destination: {dest_path}")
                    return
                    
                time.sleep(3)
                self._move_file(file_path, dest)
                break
        else:
            logging.info(f"No matching pattern found for {filename}")

    def _wait_for_print(self, filename):
        logging.info("Waiting for print job...")
        while True:
            jobs_exist = False
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            for printer in printers:
                try:
                    jobs = win32print.EnumJobs(printer[2], 0, -1)
                    if any(job.get('pDocument', '').lower() == filename.lower() for job in jobs):
                        jobs_exist = True
                        break
                except Exception:
                    continue
            
            if not jobs_exist:
                logging.info("Print job complete")
                break
            time.sleep(1)
            
    def _move_file(self, file_path, destination):
        dest_path = os.path.join(destination, os.path.basename(file_path))
        os.makedirs(destination, exist_ok=True)
        
        retries = 3
        while retries > 0:
            try:
                os.rename(file_path, dest_path)
                logging.info(f"Moved to {dest_path}")
                break
            except Exception as e:
                retries -= 1
                if retries == 0:
                    logging.error(f"Failed to move file after 3 attempts: {e}")
                else:
                    time.sleep(1)

class FileMonitor:
    def __init__(self):
        self.observer = None
        self.icon = None
        self.running = True
        
    def create_icon(self):
        # Create a simple icon (1x1 pixel)
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
        log_path = os.path.join(tempfile.gettempdir(), "filemonitor.log")
        os.startfile(log_path)
    
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