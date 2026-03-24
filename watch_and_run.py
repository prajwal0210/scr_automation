import time
import shutil
from pathlib import Path
from datetime import datetime
import pandas as pd
from watchdog.events import FileSystemEventHandler
from watchdog.observers.polling import PollingObserver

from scr_main import run_scr_automation
from raw_pipeline import run_full_pipeline   # change if your file name is different


RAW_DIR = Path(r"C:\Users\prajw\OneDrive\Documents\Projects\scr_automation\raw_files")
INPUT_DIR = Path(r"C:\Users\prajw\OneDrive\Documents\Projects\scr_automation\input")
PROCESSING_DIR = Path(r"C:\Users\prajw\OneDrive\Documents\Projects\scr_automation\processing")
RAN_ALREADY_DIR = Path(r"C:\Users\prajw\OneDrive\Documents\Projects\scr_automation\ran_already")
LOG_PATH = Path(r"C:\Users\prajw\OneDrive\Documents\Projects\scr_automation\logs\scr_log.txt")

TARGET_OUTPUT_FILE = "Automation data.xlsx"
PROCESSING_FILE = PROCESSING_DIR / TARGET_OUTPUT_FILE

PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
RAN_ALREADY_DIR.mkdir(parents=True, exist_ok=True)
LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
INPUT_DIR.mkdir(parents=True, exist_ok=True)
RAW_DIR.mkdir(parents=True, exist_ok=True)


def wait_until_stable(file_path: Path, checks: int = 3, delay: int = 2) -> bool:
    if not file_path.exists():
        return False

    stable_count = 0
    last_size = -1

    for _ in range(20):
        if not file_path.exists():
            return False

        try:
            current_size = file_path.stat().st_size
        except OSError:
            time.sleep(delay)
            continue

        if current_size > 0 and current_size == last_size:
            stable_count += 1
            if stable_count >= checks:
                return True
        else:
            stable_count = 0

        last_size = current_size
        time.sleep(delay)

    return False


def find_required_raw_files():
    employer_file = None
    trainer_file = None
    unit_result_file = None

    for f in RAW_DIR.iterdir():
        if not f.is_file():
            continue

        name = f.name.lower()

        if name.startswith("employer contact") and f.suffix.lower() == ".csv":
            employer_file = f
        elif name.startswith("trainer_events_report") and f.suffix.lower() == ".csv":
            trainer_file = f
        elif name.startswith("unit result export report") and f.suffix.lower() == ".csv":
            unit_result_file = f

    return employer_file, trainer_file, unit_result_file


def all_raw_files_ready():
    employer_file, trainer_file, unit_result_file = find_required_raw_files()

    if not all([employer_file, trainer_file, unit_result_file]):
        return None

    for f in [employer_file, trainer_file, unit_result_file]:
        if not wait_until_stable(f):
            return None

    return employer_file, trainer_file, unit_result_file


class Handler(FileSystemEventHandler):
    last_run_time = 0
    cooldown_seconds = 10
    is_running = False

    def process(self, path):
        now = time.time()

        if self.is_running:
            print("Another run is already in progress. Ignoring trigger.")
            return

        if now - self.last_run_time < self.cooldown_seconds:
            print("Duplicate trigger ignored.")
            return

        raw_files = all_raw_files_ready()
        if not raw_files:
            print("Waiting for all 3 raw files to arrive and become stable...")
            return

        self.last_run_time = now
        self.is_running = True

        employer_file, trainer_file, unit_result_file = raw_files

        try:
            print("Detected all required raw files.")
            print(f"Employer file: {employer_file}")
            print(f"Trainer events file: {trainer_file}")
            print(f"Unit result file: {unit_result_file}")

            # Clean previous processing file if exists
            if PROCESSING_FILE.exists():
                PROCESSING_FILE.unlink()

            # Remove previous generated input file if exists
            generated_input_file = INPUT_DIR / TARGET_OUTPUT_FILE
            if generated_input_file.exists():
                generated_input_file.unlink()

            # Step 1: Run raw data pipeline
            run_full_pipeline(
                employer_contact_file=employer_file,
                trainer_events_file=trainer_file,
                unit_result_file=unit_result_file,
            )

            if not generated_input_file.exists():
                raise FileNotFoundError(
                    f"Pipeline finished but output file not found: {generated_input_file}"
                )

            if not wait_until_stable(generated_input_file):
                raise RuntimeError("Generated Automation data.xlsx never became stable.")

            # Step 2: Count rows from generated input
            df = pd.read_excel(generated_input_file)
            total_scr = len(df)

            # Step 3: Copy to processing
            shutil.copy2(str(generated_input_file), str(PROCESSING_FILE))
            print(f"Copied to processing: {PROCESSING_FILE}")

            # Step 4: Run SCR automation
            start_time = time.time()
            run_scr_automation(excel_path=PROCESSING_FILE)
            end_time = time.time()
            duration = round(end_time - start_time, 2)

            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_message = f"[{timestamp}] Took {duration} secs for {total_scr} SCR"

            print(log_message)

            with open(LOG_PATH, "w", encoding="utf-8") as f:
                f.write(log_message + "\n")

            # Step 5: Archive generated input file
            archive_timestamp = datetime.now().strftime("%d.%m.%Y_%H.%M.%S")
            archived_input_file = RAN_ALREADY_DIR / f"Automation_data_{archive_timestamp}.xlsx"
            shutil.move(str(generated_input_file), str(archived_input_file))
            print(f"Moved generated input file to: {archived_input_file}")

            # Step 6: Archive raw files
            raw_archive_dir = RAN_ALREADY_DIR / f"raw_files_{archive_timestamp}"
            raw_archive_dir.mkdir(parents=True, exist_ok=True)

            for f in [employer_file, trainer_file, unit_result_file]:
                destination = raw_archive_dir / f.name
                shutil.move(str(f), str(destination))
                print(f"Moved raw file to: {destination}")

            # Step 7: Remove processing copy
            if PROCESSING_FILE.exists():
                PROCESSING_FILE.unlink()
                print("Removed processing copy.")

            print("Full automation completed successfully.")

        except Exception as e:
            print(f"Error: {e}")

        finally:
            self.is_running = False

    def on_created(self, event):
        if not event.is_directory:
            self.process(event.src_path)

    def on_modified(self, event):
        if not event.is_directory:
            self.process(event.src_path)

    def on_moved(self, event):
        if not event.is_directory:
            self.process(event.dest_path)


if __name__ == "__main__":
    observer = PollingObserver(timeout=2)
    observer.schedule(Handler(), str(RAW_DIR), recursive=False)
    observer.start()

    print(f"Monitoring folder: {RAW_DIR}")
    print("Waiting for raw files:")
    print("- Employer contact (primary).csv")
    print("- trainer_events_report_YYYY-MM-DD.csv")
    print("- Unit Result Export Report 1.csv")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()