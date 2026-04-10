"""
Zefix Firmenbrief-Automation
Täglich: neue GmbH/AG in der Schweiz (aktueller Monat) → PDF-Brief → Gmail
"""

import os
import re
import sys
import shutil
import smtplib
import logging
import datetime
import subprocess
from pathlib import Path
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import requests
from docx import Document

# ── Konfiguration ──────────────────────────────────────────────────────────────
TEMPLATE_PATH   = Path("Brief an neue Firmen.docx")
PROCESSED_FILE  = Path(".processed_firms.txt")
LOG_FILE        = Path(".task_log.txt")
TEMP_DIR        = Path("/tmp/zefix_briefe")
OUTPUT_DIR      = Path("pdfs")          # Temporäre PDF-Ablage im Repo (für E-Mail)
MAX_FIRMS       = 10
EMAIL_TO        = os.environ.get("EMAIL_TO", "fabio.dixon@confidio.ch")
GMAIL_USER      = os.environ.get("GMAIL_USER", "")
GMAIL_PASSWORD  = os.environ.get("GMAIL_APP_PASSWORD", "")

SHAB_API = "https://www.shab.ch/api/v1/publications"
ZEFIX_SEARCH = "https://www.zefix.ch/ZefixREST/api/v1/firm/search.json"
ZEFIX_DETAIL = "https://www.zefix.ch/ZefixREST/api/v1/firm/{ehraid}.json"

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# 1. Tracking
# ══════════════════════════════════════════════════════════════════════════════
def load_processed() -> set:
    if PROCESSED_FILE.exists():
        return {l.strip() for l in PROCESSED_FILE.read_text("utf-8").splitlines() if l.strip()}
    return set()


def mark_processed(name: str):
    with PROCESSED_FILE.open("a", encoding="utf-8") as f:
        f.write(name + "\n")


# ══════════════════════════════════════════════════════════════════════════════
# 2. Neue Firmen via SHAB API
# ══════════════════════════════════════════════════════════════════════════════
def get_new_gmbh_ag_this_month() -> list[dict]:
    """Gibt eine Liste von {'name', 'lang', 'pub_id'} zurück."""
    today = datetime.date.today()
    first = today.replace(day=1)
    results = []

    for page in range(20):  # max 2000 Einträge (sollte für 1 Monat reichen)
        url = (
            f"{SHAB_API}?publicationStates=PUBLISHED&rubricIds=HR"
            f"&dateFrom={first}&dateTo={today}"
            f"&pageRequest.size=100&pageRequest.page={page}"
        )
        try:
            r = requests.get(url, timeout=20)
            r.raise_for_status()
        except Exception as e:
            log.error(f"SHAB API Fehler Seite {page}: {e}")
            break

        pubs = r.json().get("content", [])
        if not pubs:
            break

        for pub in pubs:
            meta = pub["meta"]
            if meta["subRubric"] != "HR01":
                continue  # Nur Neueintragungen

            title_de = meta["title"].get("de", "")
            # Nur GmbH und AG (Schweizer Rechtsformen auf Deutsch)
            has_gmbh = "GmbH" in title_de or "GMBH" in title_de
            has_ag   = bool(re.search(r'\bAG\b', title_de))
            if not (has_gmbh or has_ag):
                continue

            # Firmenname aus Titel extrahieren: "Neueintragung FIRMENNAME, Ort"
            name = title_de.replace("Neueintragung ", "").rsplit(",", 1)[0].strip()
            if not name:
                continue

            results.append({
                "name": name,
                "pub_date": meta["publicationDate"][:10],
            })

    # Deduplizieren (SHAB publiziert manchmal mehrsprachig)
    seen = set()
    unique = []
    for r in results:
        if r["name"] not in seen:
            seen.add(r["name"])
            unique.append(r)

    log.info(f"SHAB: {len(unique)} GmbH/AG-Neueintragungen im Monat gefunden (dedupliziert)")
    return unique


# ══════════════════════════════════════════════════════════════════════════════
# 3. Firmendaten via Zefix API
# ══════════════════════════════════════════════════════════════════════════════
def get_zefix_details(firm_name: str) -> dict | None:
    """Gibt Zefix-Detaildaten zurück oder None bei Fehler."""
    for search_type in ("EXACT", "STARTSWITH"):
        try:
            r = requests.post(ZEFIX_SEARCH, json={
                "languageKey": "de",
                "maxEntries": 5,
                "name": firm_name,
                "searchType": search_type,
                "status": "ACTIVE",
                "deletedFirms": False,
            }, timeout=20)
            firms = r.json().get("list", [])
            if not firms:
                continue

            # Nimm die Firma mit dem aktuellsten shabDate
            firm = sorted(firms, key=lambda f: f.get("shabDate", ""), reverse=True)[0]
            ehraid = firm["ehraid"]
            detail = requests.get(ZEFIX_DETAIL.format(ehraid=ehraid), timeout=20).json()
            return detail

        except Exception as e:
            log.warning(f"Zefix Fehler für '{firm_name}' ({search_type}): {e}")

    return None


def extract_address(detail: dict) -> tuple[str, str]:
    """Gibt (strasse_hausnr, plz_ort) zurück."""
    addr = detail.get("address", {}) or {}
    street = addr.get("street", "")
    number = addr.get("houseNumber", "")
    plz    = addr.get("swissZipCode", "")
    town   = addr.get("town", "")
    strasse  = f"{street} {number}".strip() or "—"
    plz_ort  = f"{plz} {town}".strip() or "—"
    return strasse, plz_ort


def extract_contact_person(detail: dict) -> tuple[str, str, str]:
    """
    Gibt (vorname, nachname, anrede) zurück.
    Priorität: Zefix `roles` API-Feld (strukturiert) → SHAB-Nachrichtentext → Fallback
    """
    # ── Zefix roles-Feld (strukturiert, sprachunabhängig) ──────────────────
    roles = detail.get("roles") or []
    for quality in ("mit Einzelunterschrift", "mit Kollektivunterschrift"):
        for role in roles:
            if quality in (role.get("signatureQuality") or ""):
                fn = (role.get("firstName") or "").strip()
                ln = (role.get("lastName")  or "").strip()
                if fn and ln:
                    return fn, ln, _guess_gender(fn)

    # ── Fallback: SHAB-Nachrichtentext (deutsch) ──────────────────────────
    msg = (detail.get("shabPub") or [{}])[0].get("message", "") or ""
    idx = msg.find("Eingetragene Personen:")
    if idx >= 0:
        persons_block = msg[idx:]
        pattern_einzel = re.compile(
            r"([A-Z\u00c4\u00d6\u00dc][a-z\u00e4\u00f6\u00fc\u00df\-]+),\s+"
            r"([A-Z\u00c4\u00d6\u00dc][a-z\u00e4\u00f6\u00fc\u00df\u00c4\u00d6\u00dc\s\-]+?),\s+"
            r"von\s+.+?,\s+in\s+.+?,\s+"
            r"[^,]+,\s+mit Einzelunterschrift"
        )
        pattern_kollektiv = re.compile(
            r"([A-Z\u00c4\u00d6\u00dc][a-z\u00e4\u00f6\u00fc\u00df\-]+),\s+"
            r"([A-Z\u00c4\u00d6\u00dc][a-z\u00e4\u00f6\u00fc\u00df\u00c4\u00d6\u00dc\s\-]+?),\s+"
            r"von\s+.+?,\s+in\s+.+?,\s+"
            r"(?:Gesch\u00e4ftsf\u00fchrer|Verwaltungsrat)[^,]*,\s+mit Kollektivunterschrift"
        )
        for pattern in (pattern_einzel, pattern_kollektiv):
            m = pattern.search(persons_block)
            if m:
                nachname = m.group(1).strip()
                vorname  = m.group(2).strip().split()[0]
                return vorname, nachname, _guess_gender(vorname)

    return "", "Geschäftsführung", ""


def _guess_gender(vorname: str) -> str:
    """Bestimmt Herr/Frau anhand typischer Vornamen-Endungen."""
    weiblich_endungen = ("a", "e", "i", "ie", "ine", "tte", "lle", "nne", "ise", "ise")
    v = vorname.lower()
    if any(v.endswith(e) for e in weiblich_endungen):
        return "Frau"
    return "Herr"


# ══════════════════════════════════════════════════════════════════════════════
# 4. DOCX befüllen
# ══════════════════════════════════════════════════════════════════════════════
def _replace_in_para(para, old: str, new: str):
    """Ersetzt Text in einem Absatz, auch wenn er über mehrere Runs verteilt ist."""
    full = "".join(r.text for r in para.runs)
    if old not in full:
        return
    new_full = full.replace(old, new)
    # Ersten Run mit dem neuen Text befüllen, Rest leeren
    if para.runs:
        para.runs[0].text = new_full
        for run in para.runs[1:]:
            run.text = ""


def create_personalized_docx(
    firma: str, kontakt: str, strasse: str, plz_ort: str,
    anrede_formal: str, monat: str, out_path: Path
):
    shutil.copy2(TEMPLATE_PATH, out_path)
    doc = Document(out_path)

    # Alle möglichen Monatsangaben aus der Vorlage
    monat_patterns = [
        f"{m} {y}"
        for m in ("Januar","Februar","März","April","Mai","Juni",
                  "Juli","August","September","Oktober","November","Dezember")
        for y in range(2024, 2030)
    ]

    replacements = {
        "Firma":                   firma,
        "Anrede Vorname Nachname": kontakt,
        "Adresse":                 strasse,
        "Postleitzahl Ort":        plz_ort,
        "Formelle Anrede":         anrede_formal,
    }

    for para in doc.paragraphs:
        for old, new in replacements.items():
            _replace_in_para(para, old, new)
        # Datum ersetzen
        for mp in monat_patterns:
            if mp in "".join(r.text for r in para.runs):
                _replace_in_para(para, mp, monat)
                break

    # Auch Tabellen durchsuchen (falls Vorlage Tabellen hat)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for old, new in replacements.items():
                        _replace_in_para(para, old, new)

    doc.save(out_path)


# ══════════════════════════════════════════════════════════════════════════════
# 5. PDF-Konvertierung via LibreOffice
# ══════════════════════════════════════════════════════════════════════════════
def convert_to_pdf(docx_path: Path, out_dir: Path) -> Path:
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf",
        "--outdir", str(out_dir), str(docx_path)
    ], check=True, capture_output=True)
    pdf_path = out_dir / (docx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF nicht erstellt: {pdf_path}")
    return pdf_path


# ══════════════════════════════════════════════════════════════════════════════
# 6. E-Mail versenden
# ══════════════════════════════════════════════════════════════════════════════
def send_email(pdf_paths: list[Path], monat: str):
    if not pdf_paths:
        log.info("Keine PDFs – kein E-Mail")
        return

    n = len(pdf_paths)
    msg = MIMEMultipart()
    msg["From"]    = GMAIL_USER
    msg["To"]      = EMAIL_TO
    msg["Subject"] = f"Neue Firmenbriefe {monat} – {n} Firmen"
    msg.attach(MIMEText(
        f"Guten Tag,\n\n"
        f"Im Anhang befinden sich {n} personalisierte Firmenbriefe für neu eingetragene "
        f"Unternehmen in der Schweiz ({monat}).\n\n"
        f"Diese E-Mail wurde automatisch generiert.\n\n"
        f"Freundliche Grüsse\nConfidio GmbH",
        "plain", "utf-8"
    ))

    for pdf in pdf_paths:
        with pdf.open("rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{pdf.name}"')
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.send_message(msg)

    log.info(f"E-Mail gesendet: {n} PDFs an {EMAIL_TO}")


# ══════════════════════════════════════════════════════════════════════════════
# 7. Hauptprogramm
# ══════════════════════════════════════════════════════════════════════════════
def main():
    today = datetime.date.today()
    # Monat auf Deutsch
    monat_namen = ["Januar","Februar","März","April","Mai","Juni",
                   "Juli","August","September","Oktober","November","Dezember"]
    monat = f"{monat_namen[today.month - 1]} {today.year}"

    log.info(f"=== Start: {today} / {monat} ===")

    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    processed = load_processed()
    log.info(f"Bereits verarbeitet: {len(processed)} Firmen")

    # Neue Firmen suchen
    candidates = get_new_gmbh_ag_this_month()
    new_firms = [f for f in candidates if f["name"] not in processed]
    log.info(f"Neue (noch nicht verarbeitet): {len(new_firms)}")

    to_process = new_firms[:MAX_FIRMS]
    created_pdfs = []

    for entry in to_process:
        firma = entry["name"]
        log.info(f"Verarbeite: {firma}")

        try:
            detail = get_zefix_details(firma)
            if not detail:
                log.warning(f"  Zefix: keine Daten für '{firma}' – überspringe")
                continue

            strasse, plz_ort  = extract_address(detail)
            vorname, nachname, anrede = extract_contact_person(detail)

            if nachname == "Geschäftsführung":
                kontakt       = "Geschäftsführung"
                anrede_formal = "Sehr geehrte Damen und Herren"
            else:
                kontakt       = f"{anrede} {vorname} {nachname}".strip()
                if anrede == "Herr":
                    anrede_formal = f"Sehr geehrter Herr {nachname}"
                else:
                    anrede_formal = f"Sehr geehrte Frau {nachname}"

            log.info(f"  Kontakt: {kontakt} | {strasse}, {plz_ort}")
            log.info(f"  Anrede: {anrede_formal}")

            safe = re.sub(r'[\\/:*?"<>| ]', '_', firma)
            temp_docx = TEMP_DIR / f"{safe}.docx"
            create_personalized_docx(firma, kontakt, strasse, plz_ort, anrede_formal, monat, temp_docx)
            pdf_path  = convert_to_pdf(temp_docx, OUTPUT_DIR)
            created_pdfs.append(pdf_path)
            mark_processed(firma)
            log.info(f"  PDF: {pdf_path.name}")

        except Exception as e:
            log.error(f"  Fehler bei '{firma}': {e}", exc_info=True)
            continue

    # E-Mail
    if created_pdfs:
        send_email(created_pdfs, monat)
    else:
        log.info("Keine neuen Firmen heute.")

    # Log
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.datetime.now().isoformat()}: {len(created_pdfs)} Firmen – {monat}\n")

    log.info(f"=== Ende: {len(created_pdfs)} PDFs erstellt ===")


if __name__ == "__main__":
    main()
