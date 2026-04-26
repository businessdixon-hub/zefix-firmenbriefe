"""
HTML-Email-Template für Direkt-Versand an neue Firmen.
"""

from pathlib import Path

SIGNATURE_FILE = Path("signature.html")


def _load_signature() -> str:
    """Lädt die Signatur aus signature.html, sonst Fallback."""
    if SIGNATURE_FILE.exists():
        return SIGNATURE_FILE.read_text(encoding="utf-8")
    return """
    <p style="margin: 0;">Freundliche Grüsse</p>
    <p style="margin: 0;"><strong>Ihr Confidio Team</strong></p>
    <p style="margin: 8px 0 0 0; font-size: 13px;">
      <a href="mailto:info@confidio.ch" style="color: #1a73e8; text-decoration: none;">info@confidio.ch</a>
      &nbsp;|&nbsp;
      <a href="https://www.confidio.ch" style="color: #1a73e8; text-decoration: none;">www.confidio.ch</a>
    </p>
    """


SUBJECT = "Pflichtversicherungen für Ihre neue Firma – jetzt in 3 Minuten online anfragen"


def build_html(anrede_formal: str, firm_name: str) -> str:
    """Baut HTML-Email für eine Firma. anrede_formal z.B. 'Sehr geehrter Herr Dixon'."""
    signature = _load_signature()
    return f"""<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="font-family: Arial, Helvetica, sans-serif; color: #222; line-height: 1.5; max-width: 620px; margin: 0 auto; padding: 24px;">

  <p>{anrede_formal}</p>

  <p>Herzlichen Glückwunsch zur Gründung Ihrer Firma! Mit dem Handelsregistereintrag starten Sie Ihr Unternehmen –
  und gleichzeitig beginnen gesetzliche Versicherungspflichten, die Sie als Arbeitgeber einhalten müssen.</p>

  <p style="margin-top: 20px;"><strong>Was für Ihre Firma jetzt gesetzlich gilt:</strong></p>
  <ul style="margin: 8px 0 16px 20px; padding: 0;">
    <li><strong>UVG</strong> – Unfallversicherung: obligatorisch ab dem ersten Mitarbeitenden</li>
    <li><strong>BVG / Pensionskasse</strong> – obligatorisch ab CHF 22'050 Jahreslohn</li>
    <li><strong>KTG</strong> – Krankentaggeld: schützt Ihren Betrieb bei Lohnfortzahlung</li>
    <li><strong>Berufshaftpflicht &amp; Rechtsschutz</strong> – für Ihre unternehmerische Sicherheit</li>
  </ul>

  <p>Wir sind <strong>unabhängig von allen Versicherern</strong> – das bedeutet: Wir vergleichen für Sie den Markt
  und empfehlen nur, was wirklich zu Ihrem Unternehmen passt.</p>

  <p style="margin-top: 20px;"><strong>100 % online – ohne Termin, ohne Papierkram:</strong></p>
  <p>Besuchen Sie uns auf
    <a href="https://www.confidio.ch" style="color: #1a73e8; font-weight: bold; text-decoration: none;">www.confidio.ch</a>,
    füllen Sie unsere Offertanfrage in wenigen Minuten aus und erhalten Sie ein massgeschneidertes Angebot.
    Den Vertrag unterzeichnen Sie anschliessend bequem digital.</p>

  <p style="margin-top: 20px;">Wir freuen uns, Sie auf Ihrem unternehmerischen Weg zu begleiten.</p>

  <div style="margin-top: 24px;">
    {signature}
  </div>

  <hr style="border: 0; border-top: 1px solid #ddd; margin: 32px 0 12px 0;">
  <p style="font-size: 11px; color: #888; line-height: 1.4;">
    Sie möchten keine weiteren Zuschriften von uns?
    Schreiben Sie uns an <a href="mailto:info@confidio.ch?subject=Unsubscribe" style="color: #888;">info@confidio.ch</a> –
    wir entfernen Sie umgehend aus unserem Verteiler.
  </p>

</body>
</html>"""


def build_text(anrede_formal: str) -> str:
    """Plain-Text-Fallback für Mail-Clients ohne HTML."""
    return f"""{anrede_formal}

Herzlichen Glückwunsch zur Gründung Ihrer Firma! Mit dem Handelsregistereintrag starten Sie Ihr Unternehmen – und gleichzeitig beginnen gesetzliche Versicherungspflichten, die Sie als Arbeitgeber einhalten müssen.

Was für Ihre Firma jetzt gesetzlich gilt:
  - UVG – Unfallversicherung: obligatorisch ab dem ersten Mitarbeitenden
  - BVG / Pensionskasse – obligatorisch ab CHF 22'050 Jahreslohn
  - KTG – Krankentaggeld: schützt Ihren Betrieb bei Lohnfortzahlung
  - Berufshaftpflicht & Rechtsschutz – für Ihre unternehmerische Sicherheit

Wir sind unabhängig von allen Versicherern – das bedeutet: Wir vergleichen für Sie den Markt und empfehlen nur, was wirklich zu Ihrem Unternehmen passt.

100 % online – ohne Termin, ohne Papierkram:
Besuchen Sie uns auf www.confidio.ch, füllen Sie unsere Offertanfrage in wenigen Minuten aus und erhalten Sie ein massgeschneidertes Angebot. Den Vertrag unterzeichnen Sie anschliessend bequem digital.

Wir freuen uns, Sie auf Ihrem unternehmerischen Weg zu begleiten.

Freundliche Grüsse
Ihr Confidio Team

info@confidio.ch  |  www.confidio.ch

---
Sie möchten keine weiteren Zuschriften von uns? Schreiben Sie uns an info@confidio.ch – wir entfernen Sie umgehend aus unserem Verteiler.
"""
