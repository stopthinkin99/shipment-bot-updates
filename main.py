import json
from watcher import start_watching

def main():
    with open("config.json") as f:
        config = json.load(f)

    start_watching(config["watch_folder"])

if __name__ == "__main__":
    main()