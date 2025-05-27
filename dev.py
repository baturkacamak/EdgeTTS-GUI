import sys
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import subprocess
import os

class RestartHandler(FileSystemEventHandler):
    def __init__(self):
        self.process = None
        self.start_application()

    def start_application(self):
        if self.process:
            self.process.terminate()
            self.process.wait()
        print("\nStarting application...")
        self.process = subprocess.Popen([sys.executable, 'main.py'])

    def on_modified(self, event):
        if event.src_path.endswith('.py'):
            print(f"\nChange detected in {event.src_path}")
            self.start_application()

if __name__ == "__main__":
    path = '.'
    event_handler = RestartHandler()
    observer = Observer()
    observer.schedule(event_handler, path, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        if event_handler.process:
            event_handler.process.terminate()
        observer.stop()
    observer.join() 