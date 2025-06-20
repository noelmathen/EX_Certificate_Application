#ex_certificates_combined.py

#!/usr/bin/env python3
"""
process_mixed.py
────────────────
One-click runner for a *mixed* folder of EX certificates.

INPUT
    EX Certificates/           # <- mixed PDFs live here

OUTPUT
    Proserv_Certificates.xlsx   # generated by   proserv.py
    Oman_Certificates.xlsx      # generated by   oman.py

HOW IT WORKS
 1) Creates/clears “Proserv Certificates” and “Oman Certificates”.
 2) Reads page-1 text of every PDF in “EX Certificates” and classifies it via
        •  “ELECTRICAL EQUIPMENT IN HAZARDOUS AREAS”  → Proserv
        •  “VISUAL & CLOSE INSPECTION REPORT FOR”     → Oman
 3) Copies the file into the correct sub-folder.
 4) Runs proserv.py and then oman.py exactly as supplied.
"""

import os, sys, glob, shutil, logging, subprocess
import fitz              # PyMuPDF

MIXED_DIR   = "EX Certificates"
PROSERV_DIR = "Proserv Certificates"
OMAN_DIR    = "Oman Certificates"


KEY_PROSERV_A = "ELECTRICAL EQUIPMENT"          # partial match is safer
KEY_PROSERV_B = "HAZARDOUS AREAS"
KEY_OMAN_A    = "VISUAL & CLOSE INSPECTION"
KEY_OMAN_B    = "REPORT FOR"

# —————————————————————————————————————
def classify(pdf_path: str) -> str | None:
    """Return 'proserv', 'oman', or None (unclassified)."""
    try:
        with fitz.open(pdf_path) as doc:
            text = doc[0].get_text().upper()
    except Exception as e:
        logging.error("❌ %s: cannot open (%s)", os.path.basename(pdf_path), e)
        return None

    if KEY_PROSERV_A in text and KEY_PROSERV_B in text:
        return "proserv"
    if KEY_OMAN_A in text and KEY_OMAN_B in text:
        return "oman"
    return None

def reset_dir(path: str) -> None:
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path, exist_ok=True)

# —————————————————————————————————————
def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S"
    )
    log = logging.getLogger(__name__)

    if not os.path.isdir(MIXED_DIR):
        log.error("Input folder '%s' not found.", MIXED_DIR)
        sys.exit(1)

    # fresh target folders
    reset_dir(PROSERV_DIR)
    reset_dir(OMAN_DIR)

    pdfs = sorted(glob.glob(os.path.join(MIXED_DIR, "*.pdf")))
    if not pdfs:
        log.error("No PDFs found in '%s'.", MIXED_DIR)
        sys.exit(1)

    unknown = []
    for pdf in pdfs:
        typ = classify(pdf)
        dest = PROSERV_DIR if typ == "proserv" else \
               OMAN_DIR    if typ == "oman"   else None

        if dest:
            shutil.copy2(pdf, os.path.join(dest, os.path.basename(pdf)))
            log.info("📄 %-40s → %s", os.path.basename(pdf), typ.capitalize())
        else:
            unknown.append(os.path.basename(pdf))

    if unknown:
        log.warning("⚠️ Unclassified PDFs (%d): %s", len(unknown), ", ".join(unknown))

    # — run the original extractors —
    for script in ("proserv.py", "oman.py"):
        if not shutil.which(sys.executable):
            log.error("Python interpreter not found.")
            sys.exit(1)
        try:
            log.info("▶️  Running %s …", script)
            subprocess.run([sys.executable, script], check=True)
        except subprocess.CalledProcessError as e:
            log.error("❌ %s exited with code %s", script, e.returncode)

    log.info("🎉 All done!")

if __name__ == "__main__":
    main()
