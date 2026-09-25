"""Kartei: ausgeschriebener Gesetzesname -> amtliche Abkuerzung (seit 2026-08-29).

WARUM
-----
Nutzerregel 29.08.2026: "'Insbesondere' meint immer den §, Artikel, Ziffer
desselben zuvor genannten Gesetzes." In den Skills steht das Gesetz aber oft
AUSGESCHRIEBEN: 'Betriebsverfassungsgesetz, insbesondere §§ 5, 50'.

Der Bindeglied-Fix darf einen ausgeschriebenen Namen nicht einfach anhaengen -
Stufe 2 matcht gegen `metadata.jurabk`, und '§ 5 Betriebsverfassungsgesetz'
trifft dort nichts. Aus einem weichen 'nicht automatisiert pruefbar' wuerde ein
hartes 'nicht gefunden', also ein NEUER Sperrgrund. Richtig ist deshalb nicht
weglassen, sondern AUFLOESEN: 'Betriebsverfassungsgesetz' -> 'BetrVG'.

ZWEI QUELLEN, UND DIE REIHENFOLGE IST DER PUNKT
-----------------------------------------------
1. **Das CMS selbst** (`metadata.kurzue`/`langue` -> `metadata.amtabk`). Das ist
   die bessere Quelle, weil es **genau der Bestand ist, gegen den Stufe 2
   matcht**: was hier aufgeloest wird, ist danach auch auffindbar.
   Entscheidend ist `kurzue` - das Feld fuehrt den UMGANGSSPRACHLICHEN Kurztitel
   ('Arbeitnehmer-Entsendegesetz'), waehrend `langue` den langen amtlichen Titel
   traegt ('Gesetz ueber zwingende Arbeitsbedingungen fuer grenzueberschreitend
   entsandte ...'). Die Skills zitieren den Kurztitel.
2. **Die Teillisten von gesetze-im-internet.de** als Ergaenzung:
       <abbr title="Betriebsverfassungsgesetz"> BetrVG </abbr>
   26 Seiten (A-Z), **ISO-8859-1** kodiert - wer sie als UTF-8 liest, bekommt
   zerschossene Umlaute und findet 'BAföG' nicht.

**Warum beide gebraucht werden, gemessen am 29.08.2026:** gii fuehrt im
`title`-Attribut den AMTLICHEN Titel. Von sieben im Bestand vorkommenden
ausgeschriebenen Namen loeste gii vier auf (Betriebsverfassungsgesetz,
Buergerliches Gesetzbuch, Einkommensteuergesetz, Arbeitszeitgesetz) und drei
nicht - 'Teilzeit- und Befristungsgesetz', 'Arbeitnehmer-Entsendegesetz' und
'Handwerksordnung' sind Kurztitel, die dort nirgends stehen. Genau die liefert
`kurzue` aus dem CMS.

GRENZE - und sie ist Absicht
----------------------------
Findet die Kartei einen Namen NICHT, bleibt er unaufgeloest und der
Bindeglied-Fix laesst das Zitat ohne Gesetz. Kein Raten, keine Heuristik ueber
Wortstaemme: ein falsch aufgeloester Name erzeugt eine erfundene Fundstelle,
und die ist teurer als eine fehlende. Mehrdeutige Namen (derselbe Titel fuer
zwei Abkuerzungen) werden deshalb ebenfalls verworfen statt geraten.
"""

from __future__ import annotations

import html
import json
import os
import pathlib
import re
import time
import unicodedata
import urllib.request

BASIS = "https://www.gesetze-im-internet.de/Teilliste_{}.html"
BUCHSTABEN = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
KARTEI = pathlib.Path(__file__).with_name("_GESETZESNAMEN.json")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Justinian-Stufe1/1.0"

_ABBR_RE = re.compile(r'<abbr\s+title="([^"]+)">\s*([^<]+?)\s*</abbr>', re.IGNORECASE)
# Der PDF-Link traegt ebenfalls ein <abbr> - das ist kein Gesetz.
_KEIN_GESETZ_TITEL = re.compile(r"PDF-Datei", re.IGNORECASE)


def _schluessel(name: str) -> str:
    """Vergleichsform: NFC, Kleinschreibung, Mehrfach-Leerzeichen zusammen."""
    s = unicodedata.normalize("NFC", html.unescape(name or ""))
    s = s.replace(" ", " ")
    return re.sub(r"\s+", " ", s).strip().lower()


def _aus_cms() -> dict:
    """kurzue/langue -> amtabk, direkt aus dem CMS. Faellt der Zugriff aus,
    ist das kein Fehler - dann traegt die Kartei nur die gii-Namen."""
    import sys
    sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[2]
                           / "stufe_2_pruefung_gegen_cms" / "references"))
    try:
        import pruefung_gegen_cms as p
    except Exception:
        return {}
    pipeline = """aggregate([
{ "$match": { "doctype": "gesetz" } },
{ "$group": { "_id": "$metadata.amtabk",
              "kurz": { "$addToSet": "$metadata.kurzue" },
              "lang": { "$addToSet": "$metadata.langue" } } }
])"""
    namen: dict[str, set] = {}
    for projekt in ("laws", "vdhh"):
        try:
            rows = p.query(pipeline, project=projekt)
        except Exception:
            continue
        for r in rows:
            abk = (r.get("_id") or "").strip()
            if not abk:
                continue
            for feld in ("kurz", "lang"):
                for wert in (r.get(feld) or []):
                    if wert and isinstance(wert, str):
                        namen.setdefault(_schluessel(wert), set()).add(abk)
    return namen


def baue_kartei(ziel: pathlib.Path = KARTEI, pause: float = 0.4) -> dict:
    """Holt CMS + die 26 Teillisten und schreibt die Kartei. Bewusst aufrufen."""
    namen: dict[str, set] = {}
    for k, v in _aus_cms().items():
        namen.setdefault(k, set()).update(v)
    for b in BUCHSTABEN:
        req = urllib.request.Request(BASIS.format(b), headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=60) as r:
            roh = r.read()
        # ISO-8859-1: sonst zerfallen die Umlaute (siehe Docstring).
        text = roh.decode("iso-8859-1", "replace")
        for titel, abk in _ABBR_RE.findall(text):
            if _KEIN_GESETZ_TITEL.search(titel):
                continue
            a = unicodedata.normalize("NFC", html.unescape(abk)).strip()
            if not a:
                continue
            namen.setdefault(_schluessel(titel), set()).add(a)
        time.sleep(pause)

    eindeutig = {k: sorted(v)[0] for k, v in namen.items() if len(v) == 1}
    mehrdeutig = {k: sorted(v) for k, v in namen.items() if len(v) > 1}
    # WEISSLISTE aller bekannten Abkuerzungen. Sie ist der Waechter fuer die
    # Bauform "Gesetz vorn": ohne sie gilt in 'BetrVG-Normanker: § 87 ...' das
    # Wort 'BetrVG-Normanker' als Gesetz (es sieht formal wie eine Abkuerzung
    # aus) und erzeugt eine erfundene Fundstelle. Nur was hier steht, darf als
    # vorangestelltes Gesetz an ein Zitat gehaengt werden.
    alle = sorted({a for v in namen.values() for a in v})
    ziel.write_text(json.dumps(
        {"eintraege": eindeutig, "abkuerzungen": alle, "mehrdeutig": mehrdeutig,
         "quelle": "CMS (metadata.kurzue/langue -> amtabk) + "
                   "gesetze-im-internet.de Teilliste_<A-Z>.html (<abbr title=...>)",
         "stand": time.strftime("%Y-%m-%d")},
        ensure_ascii=False, indent=1), encoding="utf-8")
    return eindeutig


_CACHE = None


def kartei() -> dict:
    """Kartei lesen; fehlt sie, ist das kein Fehler - dann loest nichts auf."""
    global _CACHE
    if _CACHE is None:
        pfad = pathlib.Path(os.environ.get("GESETZESNAMEN_KARTEI", KARTEI))
        try:
            _CACHE = json.loads(pfad.read_text(encoding="utf-8")).get("eintraege", {})
        except Exception:
            _CACHE = {}
    return _CACHE


def abkuerzung(name: str) -> str:
    """'Betriebsverfassungsgesetz' -> 'BetrVG'; unbekannt -> ''.

    Zweiter Versuch als PRAEFIX, aber nur bei EINDEUTIGKEIT: die Skills nennen
    'Grundgesetz', der amtliche Titel lautet 'Grundgesetz fuer die Bundesrepublik
    Deutschland'. Gibt es mehr als einen Titel mit diesem Anfang, wird nichts
    zurueckgegeben - ein falsch aufgeloester Name erzeugt eine erfundene
    Fundstelle, und die ist teurer als eine fehlende."""
    k = _schluessel(name)
    if not k:
        return ""
    d = kartei()
    if k in d:
        return d[k]
    if len(k) < 6:                      # zu kurz fuer eine sichere Praefixsuche
        return ""
    treffer = {v for schl, v in d.items() if schl.startswith(k + " ")}
    return treffer.pop() if len(treffer) == 1 else ""


_ABK_SET = None


def ist_bekannte_abkuerzung(token: str) -> bool:
    """Steht dieses Kuerzel als amtliche Abkuerzung in der Kartei?

    Waechter fuer die Bauform "Gesetz vorn": 'BetrVG-Normanker' sieht formal wie
    eine Abkuerzung aus und wuerde sonst als Gesetz an die Paragraphen gehaengt -
    eine erfundene Fundstelle. 'BetrVG' steht in der Kartei, 'BetrVG-Normanker'
    nicht."""
    global _ABK_SET
    if _ABK_SET is None:
        pfad = pathlib.Path(os.environ.get("GESETZESNAMEN_KARTEI", KARTEI))
        try:
            _ABK_SET = set(json.loads(pfad.read_text(encoding="utf-8"))
                           .get("abkuerzungen", []))
        except Exception:
            _ABK_SET = set()
    return bool(token) and token.strip() in _ABK_SET


if __name__ == "__main__":
    d = baue_kartei()
    print(f"OK: {KARTEI} ({len(d)} eindeutige Namen)")
    for probe in ("Betriebsverfassungsgesetz", "Bürgerliches Gesetzbuch",
                  "Einkommensteuergesetz", "Teilzeit- und Befristungsgesetz",
                  "Arbeitnehmer-Entsendegesetz", "Entgelttransparenzgesetz",
                  "Arbeitszeitgesetz", "Handwerksordnung"):
        print(f"   {probe:<38} -> {abkuerzung(probe)!r}")
