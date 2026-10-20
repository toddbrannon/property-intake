"""
Folder watcher.  Runs extract.py on every new PDF that appears in the inbox folder.

    python watch.py            # keeps running; checks the folder every POLL_SECONDS
    python watch.py --once     # processes what is there now, then exits (use this with Task Scheduler / cron)

Folders (set in .env):
    INBOX_DIR   where people drop PDFs (e.g. a Nutstore-synced folder)
    DONE_DIR    processed PDFs are moved here
    FAILED_DIR  PDFs the script could not process are moved here, with a .error.txt beside them
    OUTPUT_DIR  master.xlsx, review.csv, log.txt and one .json per document

The script never deletes a PDF. It only moves them.
"""
import os, sys, time, shutil, traceback
from pathlib import Path
import extract

HERE = Path(__file__).resolve().parent
INBOX  = Path(os.environ.get("INBOX_DIR",  HERE / "inbox"))
DONE   = Path(os.environ.get("DONE_DIR",   HERE / "done"))
FAILED = Path(os.environ.get("FAILED_DIR", HERE / "failed"))
POLL   = int(os.environ.get("POLL_SECONDS", "60"))

def settled(p: Path) -> bool:
    """True when the file hasn't changed size for 2 seconds (i.e. Nutstore has finished syncing it)."""
    s1 = p.stat().st_size; time.sleep(2); return p.stat().st_size == s1 and s1 > 0

def run_once(fields):
    for d in (INBOX, DONE, FAILED):
        d.mkdir(parents=True, exist_ok=True)
    pdfs = sorted(p for p in INBOX.iterdir() if p.suffix.lower() == ".pdf" and not p.name.startswith("~"))
    if not pdfs:
        return 0
    for pdf in pdfs:
        if not settled(pdf):
            continue
        try:
            extract.process(pdf, fields)
            target = DONE / pdf.name
            if target.exists():
                target = DONE / f"{pdf.stem}_{int(time.time())}{pdf.suffix}"
            shutil.move(str(pdf), str(target))
            manual = pdf.with_name(pdf.name + ".claude.json")   # reply saved from the Claude app, if any
            if manual.exists():
                shutil.move(str(manual), str(DONE / manual.name))
        except Exception as e:
            extract.log(f"{pdf.name}: FAILED - {e}")
            target = FAILED / pdf.name
            shutil.move(str(pdf), str(target))
            (FAILED / f"{pdf.stem}.error.txt").write_text(traceback.format_exc(), encoding="utf-8")
    return len(pdfs)

def main():
    fields = extract.load_fields()
    once = "--once" in sys.argv
    extract.log(f"watch started (inbox={INBOX}, once={once})")
    while True:
        try:
            run_once(fields)
        except Exception as e:
            extract.log(f"watch loop error: {e}")
        if once:
            break
        time.sleep(POLL)

if __name__ == "__main__":
    main()
