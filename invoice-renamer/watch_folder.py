"""inbox 폴더를 감시하다가 PDF가 들어오면 자동으로 rename_invoices.py --apply 를 실행한다.
선택 기능. 창을 닫으면 멈춘다.  사용: python watch_folder.py [--config my.json]
"""
import argparse, subprocess, sys, time
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
sys.path.insert(0, str(Path(__file__).parent))
from rename_invoices import load_config


class Handler(FileSystemEventHandler):
    def __init__(self, config):
        self.config = config
    def on_created(self, e):
        if e.is_directory or not e.src_path.lower().endswith(".pdf"):
            return
        time.sleep(2)  # 다운로드가 끝날 때까지 잠깐 대기
        subprocess.run([sys.executable, str(Path(__file__).with_name("rename_invoices.py")), "--apply", "--config", self.config])


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--config", default=str(Path(__file__).with_name("config.json")))
    a = ap.parse_args()
    inbox = load_config(a.config)["inbox"]; Path(inbox).mkdir(parents=True, exist_ok=True)
    obs = Observer(); obs.schedule(Handler(a.config), str(inbox), recursive=False); obs.start()
    print(f"감시 중: {inbox}  (Ctrl+C 로 종료)")
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt:
        obs.stop()
    obs.join()
