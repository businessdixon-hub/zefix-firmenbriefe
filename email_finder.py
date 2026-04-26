"""
Email-Finder: Findet Kontakt-Email einer Firma via Domain-Rating + Webseiten-Scraping.
"""

import re
import logging
import unicodedata
from urllib.parse import urlparse, urljoin

import requests

log = logging.getLogger(__name__)

USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
TIMEOUT = 8

# Adressen die wir NICHT als Kontakt-Email akzeptieren
EMAIL_BLACKLIST_PREFIXES = (
    "noreply", "no-reply", "donotreply", "do-not-reply", "mailer",
    "postmaster", "abuse", "webmaster", "hostmaster",
)
EMAIL_BLACKLIST_DOMAINS = (
    "wixpress.com", "sentry.io", "example.com", "example.org",
    "sentry-next.wixpress.com", "wix.com",
)
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
MAILTO_RE = re.compile(r'mailto:([^"\'\s?>]+)', re.IGNORECASE)


def _slug(name: str) -> str:
    """Firmenname → URL-tauglicher Slug (ohne GmbH/AG-Suffix, ASCII)."""
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    s = s.lower()
    # Suffixe entfernen
    for suf in (" gmbh", " ag", " sa", " sarl", " ltd", " inc", " holding"):
        if s.endswith(suf):
            s = s[: -len(suf)]
    # Sonderzeichen → Bindestrich
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s


def guess_domain_candidates(firm_name: str) -> list[str]:
    """Generiert plausible Domain-Kandidaten für eine Firma."""
    slug = _slug(firm_name)
    if not slug:
        return []
    candidates = [
        f"{slug}.ch",
        f"{slug.replace('-', '')}.ch",
        f"www.{slug}.ch",
    ]
    # Dedupe
    seen = set()
    return [c for c in candidates if not (c in seen or seen.add(c))]


def _fetch(url: str) -> str | None:
    try:
        r = requests.get(url, timeout=TIMEOUT, headers={"User-Agent": USER_AGENT}, allow_redirects=True)
        if r.status_code == 200 and "text/html" in r.headers.get("Content-Type", ""):
            return r.text
    except Exception as e:
        log.debug(f"Fetch fail {url}: {e}")
    return None


def domain_exists(domain: str) -> bool:
    """Prüft ob Domain via HTTPS erreichbar ist."""
    return _fetch(f"https://{domain}") is not None


def duckduckgo_first_ch_url(firm_name: str) -> str | None:
    """Sucht via DuckDuckGo HTML, gibt erste .ch-URL zurück."""
    try:
        r = requests.post(
            "https://html.duckduckgo.com/html/",
            data={"q": f'"{firm_name}" site:.ch'},
            timeout=TIMEOUT,
            headers={"User-Agent": USER_AGENT},
        )
        # Treffer extrahieren: <a class="result__a" href="...">
        urls = re.findall(r'class="result__a"[^>]*href="([^"]+)"', r.text)
        for u in urls:
            # DDG verlinkt manchmal über /l/?uddg=... → dekodieren
            m = re.search(r"uddg=([^&]+)", u)
            if m:
                from urllib.parse import unquote
                u = unquote(m.group(1))
            host = urlparse(u).netloc.lower()
            if host.endswith(".ch") and not any(b in host for b in ("duckduckgo", "google", "bing")):
                return u
    except Exception as e:
        log.debug(f"DDG fail: {e}")
    return None


def _extract_emails(html: str, allowed_domain: str) -> list[str]:
    """Extrahiert valide Emails (mailto + Klartext) die zur erlaubten Domain passen."""
    found = set()
    # mailto-Links
    for m in MAILTO_RE.findall(html):
        found.add(m.split("?")[0].strip().lower())
    # Klartext (Fallback)
    for e in EMAIL_RE.findall(html):
        found.add(e.lower())

    valid = []
    for e in found:
        if "@" not in e:
            continue
        local, _, dom = e.partition("@")
        if dom in EMAIL_BLACKLIST_DOMAINS:
            continue
        if any(local.startswith(p) for p in EMAIL_BLACKLIST_PREFIXES):
            continue
        # Nur Emails der gleichen Domain (sonst sind es z.B. Footer-Tools)
        if allowed_domain and not (dom == allowed_domain or dom.endswith("." + allowed_domain)):
            continue
        valid.append(e)
    return valid


def _prioritize(emails: list[str], first_name: str, last_name: str) -> str | None:
    """Wählt die beste Email aus."""
    if not emails:
        return None
    fn = (first_name or "").lower().strip()
    ln = (last_name or "").lower().strip()

    # 1. Match Kontaktperson (vorname.nachname@, vornachname@, v.nachname@)
    if fn and ln:
        patterns = [
            f"{fn}.{ln}@",
            f"{fn}{ln}@",
            f"{fn[0]}.{ln}@",
            f"{fn[0]}{ln}@",
            f"{ln}.{fn}@",
            f"{ln}@",
        ]
        for p in patterns:
            for e in emails:
                if e.startswith(p):
                    return e

    # 2. Generische
    for prefix in ("info@", "kontakt@", "contact@", "office@", "hello@", "mail@"):
        for e in emails:
            if e.startswith(prefix):
                return e

    # 3. Erste sonstige
    return emails[0]


def find_email(firm_name: str, first_name: str = "", last_name: str = "") -> tuple[str, str] | None:
    """
    Findet Kontakt-Email für eine Firma.
    Returns: (email, source_url) oder None.
    """
    log.info(f"  Email-Suche: {firm_name}")

    # 1. Domain raten
    target_url = None
    target_domain = None
    for cand in guess_domain_candidates(firm_name):
        host = cand.lstrip("www.")
        if domain_exists(host):
            target_url = f"https://{host}"
            target_domain = host
            log.info(f"    Domain geraten: {host}")
            break

    # 2. Falls nicht gefunden → DuckDuckGo
    if not target_url:
        ddg = duckduckgo_first_ch_url(firm_name)
        if ddg:
            target_url = ddg
            target_domain = urlparse(ddg).netloc.lower().lstrip("www.")
            log.info(f"    Via DDG: {target_domain}")

    if not target_url:
        log.info("    Keine Webseite gefunden.")
        return None

    # 3. Kontakt-Seiten scrapen
    pages_to_check = [
        target_url,
        urljoin(target_url, "/kontakt"),
        urljoin(target_url, "/kontakt/"),
        urljoin(target_url, "/contact"),
        urljoin(target_url, "/impressum"),
        urljoin(target_url, "/impressum/"),
        urljoin(target_url, "/about"),
        urljoin(target_url, "/ueber-uns"),
    ]
    all_emails = []
    for page in pages_to_check:
        html = _fetch(page)
        if html:
            all_emails.extend(_extract_emails(html, target_domain))
    all_emails = list(dict.fromkeys(all_emails))  # dedupe, keep order

    if not all_emails:
        log.info("    Keine Email auf der Seite gefunden.")
        return None

    chosen = _prioritize(all_emails, first_name, last_name)
    log.info(f"    ✓ Gewählt: {chosen}  (von {len(all_emails)} Kandidaten)")
    return (chosen, target_url) if chosen else None
