import os
import time
import re
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import win32print
import win32api
from pathlib import Path
import yaml

class FileHandler(FileSystemEventHandler):
    def __init__(self, config, downloads_dir):
        self.patterns = config['patterns']
        self.default_printer = config.get('default_printer')
        self.downloads_dir = downloads_dir
        
    def on_created(self, event):
        if event.is_directory:
            return
        print(f"\nNew file detected: {event.src_path}")
        self._process_file(event.src_path)
        
    def _process_file(self, file_path):
        filename = os.path.basename(file_path)
        
        for pattern in self.patterns:
            match = re.match(pattern['pattern'], filename)
            if match:
                print(f"File {filename} matches pattern: {pattern['pattern']}")
                # Format destination with named groups
                dest = pattern['destination'].format(**match.groupdict())
                dest_path = os.path.join(dest, filename)
                
                if pattern.get('print', False):
                    printer_name = pattern.get('printer', self.default_printer)
                    print(f"Printing file to {printer_name}")
                    try:
                        if printer_name:
                            win32print.SetDefaultPrinter(printer_name)
                        win32api.ShellExecute(0, "print", file_path, None, ".", 0)
                        print("Print job sent")
                        time.sleep(3)
                        self._wait_for_print(filename)
                    except Exception as e:
                        print(f"Print error: {e}")
                
                if os.path.exists(dest_path):
                    print(f"File already exists at destination: {dest_path}")
                    return
                    
                time.sleep(3)
                
                # self._close_preview(file_path)
                self._move_file(file_path, dest)
                break
        else:
            print(f"No matching pattern found for {filename}")

    def _wait_for_print(self, filename):
        print("Waiting for print job...")
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
                print("Print job complete")
                break
            time.sleep(1)
            
    def _move_file(self, file_path, destination):
        dest_path = os.path.join(destination, os.path.basename(file_path))
        os.makedirs(destination, exist_ok=True)
        
        retries = 3
        while retries > 0:
            try:
                os.rename(file_path, dest_path)
                print(f"Moved to {dest_path}")
                break
            except Exception as e:
                retries -= 1
                if retries == 0:
                    print(f"Failed to move file after 3 attempts: {e}")
                else:
                    time.sleep(1)

def main():
    downloads_dir = os.path.expanduser("~/Downloads")
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yaml")
    
    try:
        with open(config_path, 'r') as f:
            config = yaml.safe_load(f)
            
        print(f"\nMonitoring: {downloads_dir}")
        print("Default printer:", config.get('default_printer', 'Not set'))
        print("Press Ctrl+C to stop")
        
        event_handler = FileHandler(config, downloads_dir)
        observer = Observer()
        observer.schedule(event_handler, downloads_dir, recursive=False)
        observer.start()
        
        while True:
            time.sleep(1)
            
    except KeyboardInterrupt:
        observer.stop()
        observer.join()
        print("\nStopped")

if __name__ == "__main__":
    main()