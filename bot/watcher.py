from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
from processor import process_label

class LabelHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.src_path.endswith(".pdf"):
            print(f"[INFO] New label detected: {event.src_path}")
            time.sleep(2)  # wait for file to finish saving
            process_label(event.src_path)

def start_watching(path):
    observer = Observer()
    observer.schedule(LabelHandler(), path=path, recursive=False)
    observer.start()

    print("[INFO] Watching folder:", path)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()