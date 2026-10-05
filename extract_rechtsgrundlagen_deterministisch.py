"""Stufe 1 - deterministische Extraktion der Rechtsgrundlagen aus JUS-Word-Skills.
Endfassung 26.08.2026.

Ersetzt Schritt 5 des Skills /stufe_1_index_rechtsgrundlagen (ein LLM-Subagent
je Datei) durch eine reproduzierbare Tabellen-Extraktion. Gibt dasselbe
JSON-Schema aus und ist gegen 21 LLM-Ergebnisse desselben Docx-Stands kalibriert.

ZWEISTUFIGE ERKENNUNG - beides muss zutreffen:
  1. Abschnitt: der Dokumentkoerper wird in echter Reihenfolge abgelaufen, die
     jeweils letzte Ueberschrift mitgefuehrt. Nur Tabellen in einem
     Rechtsgrundlagen-Abschnitt zaehlen. Ohne diese Eingrenzung schlagen
     Pruefbaum-, Fehlervermeidungs- und Ankertabellen an (gemessen: 10.231
     statt 4.186 Normen).
  2. Inhalt: innerhalb des Abschnitts wird die Spalte gesucht, deren Zellen
     Normzitate bzw. Aktenzeichen SIND.

BELEGTE BAUFORMEN (gemessen ueber alle 314 LABOR-Dateien am 26.08.2026):

Abschnittsueberschriften
  Normen         'Primäre Normen' (226x), 'Ergänzende und parallele Normen',
                 'Normen', 'Normenlandkarte ...', numeriert ('19. Primäre Normen')
  Rechtsprechung 'Tragende Rechtsprechung' (95x), '1. Höchstrichterliche
                 Rechtsprechung (Deutschland und EU)' (36x), '2. Einschlägige
                 Rechtsprechung der Instanzgerichte' (34x), 'Höchstrichterliche
                 Rechtsprechung', '2. Einschlägige Instanzrechtsprechung'

Normzelle
  A) '§ 87 Abs. 1 Nr. 1 BetrVG'          - Standard
  B) 'GewO § 106', 'BGB § 241 Abs. 2'     - Gesetz VORN (wird umgestellt)
  C) '§§ 80, 90, 95 BetrVG'               - Sammelzitat (wird aufgeteilt)
  D) '§§ 16-19 AEntG', '§§ 74 bis 75a SGB III' - Bereich (wird aufgeteilt)
  E) 'Anhang III Nr. 4 KI-VO', 'Art. 26 KI-VO'
  F) 'JArbSchG', 'HinSchG'                - nur Gesetzesname, ohne Paragraph
  G) '[offen]', '[JUS-Fundstelle zu § 90 BetrVG]' - Platzhalter (Befund!)

Aktenzeichen
  A) '8 AZR 679/09'  B) 'C-134/24' (EuGH)  C) 'II ZR 243/13' (roemischer Senat)
  D) '6 P 1.96' (BVerwG, Punktnotation)    E) '5 AZR 143/19 (A)' (Registerzusatz)
  F) '3 Ws 599/14, 3 Ws 600/14' (mehrere in EINER Zelle - alle werden erfasst)
  G) 'BAG, Urteil vom 07.04.2011 - 8 AZR 679/09' (kombinierte Zelle)
  H) Platzhalter - Rohtext wird erhalten, nicht auf '[offen]' normalisiert

BEWUSSTE ABGRENZUNGEN
  - 'Norm-/Fundstellen-Anker für jede Prüfaussage' (171x) und
    'Rechtsprechungs-Kontextprüfung' (26x) zaehlen NICHT: das sind Ankertabellen
    des Pruefteils, keine Quellendeklaration; ihre Normen stehen im Regelfall
    schon unter 'Primäre Normen'. Der LLM-Lauf hat sie ebenfalls nicht gezaehlt.
    Die verworfene Menge wird je Datei mitgezaehlt und ausgewiesen.
  - Aufgeteilt wird nur ueber PARAGRAPHEN-Grenzen ('§§ 14, 15 TzBfG'), nicht
    innerhalb eines Paragraphen ('§ 87 Abs. 1 Nr. 1, 6, 7 BetrVG' bleibt EIN
    Eintrag) - Stufe 2 matcht auf Gesetz+Paragraph, die Nr.-Aufzaehlung wuerde
    denselben Paragraphen mehrfach zaehlen.
"""
import json
import os
import re
import sys

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pruefkriterien_zitate as pk  # noqa: E402  (Fundstellenluecken, seit 2026-08-27)

# ============================================================== Grundfunktionen


def norm_txt(s):
    return re.sub(r"\s+", " ", (s or "").replace("\xa0", " ")).strip()


PLATZHALTER_RE = re.compile(
    r"\[[^\]]*(?:offen|dynamisch|zu ermitteln|unverifiziert|keine spezifische|"
    r"jus-fundstelle|fallbezogen|kein statischer|ausfuehrungsspezifisch|"
    r"ausführungsspezifisch|tbd|n\.n\.)[^\]]*\]",
    re.IGNORECASE,
)


def is_platzhalter(s):
    return bool(PLATZHALTER_RE.search(s or ""))


# ============================================================== Aktenzeichen

AZ_RE = re.compile(
    r"(?:"
    r"\b[CT]-\d{1,4}/\d{2}\b"                                        # EuGH/EuG
    r"|\b[IVXL]{1,4}\s+[A-Z]{1,3}\s?\d{1,4}/\d{2,4}\b"               # II ZR 243/13, V R 53/10
    # BSG: 'B <Senat> <Sachgebiet> <Nr>/<Jahr> <Registerbuchstabe>'. MUSS vor dem
    # allgemeinen Muster darunter stehen - sonst greift dieses zuerst und
    # verstuemmelt das Aktenzeichen zu '12 KR 16/02' (fuehrendes B und
    # Registerbuchstabe fallen weg). Belegt am 29.08.2026 gegen den eigenen
    # CMS-Bestand: von 4000 abgefragten BSG-Entscheidungen tragen 3867 (96,7 %)
    # genau diese Bauform; Registerbuchstaben B (2476), R (1148), BH (114),
    # S (72), C (45), AR (11), D (1). Sachgebiete zweistellig gemischt
    # (KR, AS, KA, AL, SB, SO, EG, LW, AY ...), teils mit Umlaut ('B 10 ÜG 13/13 B').
    # Drei Varianten im Senatsfeld, alle am Bestand belegt: einfach ('B 12 R 23/16 B'),
    # mit Buchstabenzusatz ('B 12a KR 5/20 B', 9x) und Doppelsenat
    # ('B 7/14 AS 197/21 B', 'B 8/7 AY 5/21 B', 8x). Registerzusatz bis DREI
    # Buchstaben ('BH', 'AR', 'KL').
    r"|\bB\s+\d{1,2}[a-z]?(?:/\d{1,2})?\s+[A-ZÄÖÜ][A-ZÄÖÜa-zäöü]{0,3}"
    r"\s+\d{1,4}/\d{2}(?:\s+[A-Z]{1,3})?\b"
    r"|\b\d{1,3}\s+[A-ZÄÖÜ][A-Za-zÄÖÜäöüß]{0,5}\s+\d{1,5}[/.]\d{2,4}\b"  # 8 AZR 679/09, 6 P 1.96
    # Landgerichte fuehren die Kammer fuer Handelssachen als EIGENES Token
    # zwischen Kammer und laufender Nummer: '5 HK O 1387/10' (LG Muenchen I,
    # 10.12.2013 - Neubuerger). Das Muster darueber erlaubt nur EIN Token vor
    # der Zahl, deshalb fiel dieses Aktenzeichen aus dem Index, obwohl es in
    # der Tabelle 'Tragende Rechtsprechung' steht (Befund 28.08.2026, 2 Skills).
    r"|\b\d{1,3}\s+(?:HK\s+)?[OT]\s+\d{1,5}/\d{2,4}\b"                   # 5 HK O 1387/10, 12 O 45/19
    r"|\bGS\s+\d{1,2}/\d{2}\b"
    r")"
    r"(?:\s*\([A-Z]\))?"                                             # Registerzusatz (A)/(F)
)

GERICHTE = [
    "BAG", "BGH", "BFH", "BSG", "BVerfG", "BVerwG", "BPatG", "BAGE",
    "EuGH", "EuG", "EGMR",
    "LAG", "ArbG", "LSG", "SG", "FG", "OLG", "LG", "AG", "OVG", "VGH", "VG",
    "Bundesarbeitsgericht", "Bundesgerichtshof", "Bundesfinanzhof",
    "Bundessozialgericht", "Bundesverfassungsgericht", "Bundesverwaltungsgericht",
    "Landesarbeitsgericht", "Arbeitsgericht", "Gerichtshof",
]
GERICHT_RE = re.compile(r"\b(" + "|".join(re.escape(g) for g in GERICHTE) + r")\b")
DATUM_RE = re.compile(r"\b(\d{1,2})\.\s?(\d{1,2})\.\s?(\d{4})\b")
DATUM_ISO_RE = re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b")

ENTSCHEIDUNGSART = ("Urteil", "Beschluss", "Beschluß", "Vorlage", "Vorabentscheidung",
                    "Anfragebeschluss", "Versäumnisurteil", "Teilurteil")


# Mitteltoken, die ein Aktenzeichen ausschliessen: 'Art. 14 Abs. 1 VO 987/2009'
# passt sonst auf das Aktenzeichenmuster (Fund 26.08.2026 in
# RESEARCH.INTERNATIONAL.COUNTRY_REQUIREMENTS_UPDATE - 2 Falsch-Positive).
#
# `arb` (31.08.2026): `Art. 6 oder 7 ARB 1/80` (Assoziationsratsbeschluss) traf
# als `7 ARB 1/80` das Aktenzeichenmuster - zwei Phantom-Aktenzeichen in
# ANALYSIS.INTERNATIONAL.RESIDENCE_TITLE_SCOPE, im CMS nie zu finden und damit
# dauerhaft sperrend. Am CMS gegengeprueft, mit Positivkontrolle: das Muster
# `\d+ ARB \d+/\d+` liefert 0 Treffer, `\d+ ABR \d+/\d+` (echtes Registerzeichen)
# liefert Treffer. ARB ist also nie ein Registerzeichen, sondern immer der
# Rechtsakt. Die sieben "ARB"-Treffer im Bestand sind die Teilzeichenkette in
# "ArbG Marburg" - kein Mitteltoken.
AZ_BLOCK = {"vo", "rl", "egrl", "eg", "eu", "ewg", "egv", "aeuv", "gg", "abl",
            "richtlinie", "verordnung", "nr", "abs", "satz", "s", "arb"}


def az_ok(s):
    teile = norm_txt(s).split()
    if len(teile) >= 3 and teile[1].lower().rstrip(".") in AZ_BLOCK:
        return False
    return True


def alle_az(text):
    return [norm_txt(m.group(0)) for m in AZ_RE.finditer(norm_txt(text))
            if az_ok(m.group(0))]


def is_az_cell(s):
    s = norm_txt(s)
    if not s:
        return False
    return bool(alle_az(s)) or is_platzhalter(s)


_OHNE_ORT = {"BAG", "BGH", "BFH", "BSG", "BVerfG", "BVerwG", "BPatG", "BAGE", "EuGH", "EuG",
             "EGMR", "Bundesarbeitsgericht", "Bundesgerichtshof", "Bundesfinanzhof",
             "Bundessozialgericht", "Bundesverfassungsgericht", "Bundesverwaltungsgericht"}


def gericht_aus(text):
    text = norm_txt(text)
    mg = GERICHT_RE.search(text)
    if not mg:
        return ""
    gericht = mg.group(1)
    vor = text[: mg.start()].strip(" ,;-–")
    if vor and len(vor.split()) <= 2 and vor[:1].isupper() and "," not in vor:
        gericht = vor + " " + gericht          # 'Hessisches LAG'
    nach = text[mg.end():].lstrip(" ,;-–")
    ort = r"[A-ZÄÖÜ][a-zäöüß]*(?:-[A-ZÄÖÜa-zäöüß]+)*"
    m2 = re.match(r"^(%s(?:\s+%s){0,2})" % (ort, ort), nach)
    # Bundes- und EU-Gerichte haben KEINEN Ortszusatz (seit 25.09.2026): was dahinter
    # steht, ist ein Parteiname ('EuGH Rosenbladt'), das Registerzeichen ('BFH I R',
    # 'EuGH C') oder der Spruchkoerper - nie Teil des Gerichts. Abschaltbar mit
    # STUFE1_OHNE_BUNDESGERICHT_ORT=1.
    if mg.group(1) in _OHNE_ORT and os.environ.get("STUFE1_OHNE_BUNDESGERICHT_ORT") != "1":
        m2 = None
    if m2 and m2.group(1) not in ENTSCHEIDUNGSART:
        gericht = gericht + " " + m2.group(1)  # 'LAG Baden-Württemberg'
    return norm_txt(gericht)


def gericht_zum_az(text, von, a0):
    """Das Gericht, das im Zitatglied UNMITTELBAR vor dem Az an text[a0] steht (seit 25.09.2026).

    Gesucht wird nur zwischen dem vorigen Az (`von`) und diesem; das letzte Gericht dort
    zaehlt, beginnend am Gliedanfang (nach `,` `;` `:`), damit 'Hessisches LAG' erhalten
    bleibt. Kein Gericht im Glied -> "" (Aufrufer faellt auf Zelle/Spalte zurueck)."""
    abschnitt = text[von:a0]
    treffer = list(GERICHT_RE.finditer(abschnitt))
    if not treffer:
        return ""
    ls = treffer[-1].start()
    anfang = max(abschnitt.rfind(t, 0, ls) for t in ",;:") + 1
    return gericht_aus(abschnitt[anfang:])


_ZITAT_TRENNER = re.compile(r";|\.\s+[A-ZÄÖÜ]|\n")
# Abkuerzungspunkt ist kein Satzende (seit 25.09.2026): 'EuGH, 12.09.2019 – verb. Rs. C-64/18,
# C-140/18, C-146/18 und C-148/18' - 'verb. Rs' und 'Rs. C' galten als Satzende, alle vier Az
# verloren das Datum. Vor der Trennerpruefung wird der Punkt dieser Kuerzel maskiert.
_ABK_PUNKT_RE = re.compile(r"\b(verb|Rs|Az|Urt|Beschl|vgl)\.")


def datum_zum_az(text, a0, a1, vor_max=60, nach_max=30):
    """Das Datum, das zum Aktenzeichen an text[a0:a1] GEHOERT (seit 24.09.2026).

    Vorher nahmen alle drei Wege (Tabellenzelle, Fliesstext, Arbeitsschritt) das ERSTE Datum
    im Umfeld. Zwei gemessene Folgen: '4 AZR 521/15, 25.01.2017; 4 AZR 375/16, 25.10.2017'
    gab BEIDEN den 25.01.2017; 'BAG 21.09.2011 - 7 AZR 375/10. ... BVerfG 06.06.2018 -
    7 AZR 452/17' gab 452/17 das Datum des Vorsatzes. Regel: ein Datum gehoert zum Az, wenn
    zwischen beiden KEIN anderes Aktenzeichen und KEIN Zitattrenner (';', Satzende, Zeilen-
    umbruch) steht; vor dem Az hoechstens vor_max, dahinter hoechstens nach_max Zeichen
    Abstand; der naechste Kandidat gewinnt, bei Gleichstand der vor dem Az ('BAG, Urteil vom
    21.11.2013 - 2 AZR 1057/12'). Kein Kandidat -> '' (kein Datum ist besser als ein fremdes).
    """
    text = text or ""
    daten = [(m.start(), m.end(), "%04d-%02d-%02d" % (int(m.group(3)), int(m.group(2)), int(m.group(1))))
             for m in DATUM_RE.finditer(text)]
    daten += [(m.start(), m.end(), m.group(0)) for m in DATUM_ISO_RE.finditer(text)]
    andere = [(m.start(), m.end()) for m in AZ_RE.finditer(text)
              if az_ok(m.group(0)) and not (m.start() < a1 and m.end() > a0)]
    def frei(x0, x1, liste_ok=False):
        if _ZITAT_TRENNER.search(_ABK_PUNKT_RE.sub(r"\1_", text[x0:x1])):
            return False
        innen = [(s, e) for s, e in andere if s >= x0 and e <= x1]
        if not innen:
            return True
        if not liste_ok:
            return False
        # Aktenzeichenliste hinter EINEM Datum teilt sich das Datum - verbundene Verfahren
        # ('EuGH 12.09.2019 - C-64/18, C-140/18'), Parallelverfahren ('27.03.2025 - 8 AZR 63/24 -
        # sowie Parallelverfahren 8 AZR 139/24'), Klammerzusatz ('XI R 35/22 (XI R 14/20)').
        # Gemessen: ohne diese Ausnahme verloren 7 Az ein richtiges Datum.
        rest, pos = [], x0
        for s, e in sorted(innen):
            rest.append(text[pos:s]); pos = e
        rest.append(text[pos:x1])
        return bool(re.fullmatch(r"[\s,–\-:·()\[\]]*(?:(?:und|sowie|bzw\.|u\.|Parallelverfahren|verb\.|Rs\.)[\s,–\-:·()\[\]]*)*",
                                 "".join(rest)))
    best = None
    for s, e, iso in daten:
        if e <= a0 and a0 - e <= vor_max + 40 and frei(e, a0, liste_ok=True) and (a0 - e <= vor_max or not frei(e, a0)):
            k = (a0 - e, 0)
        elif s >= a1 and s - a1 <= nach_max and frei(a1, s):
            k = (s - a1, 1)
        else:
            continue
        if best is None or k < best[0]:
            best = (k, iso)
    return best[1] if best else ""


def datum_aus(text):
    md = DATUM_RE.search(text or "")
    if md:
        return "%04d-%02d-%02d" % (int(md.group(3)), int(md.group(2)), int(md.group(1)))
    mi = DATUM_ISO_RE.search(text or "")
    return mi.group(0) if mi else ""


# ==================================================================== Normen

NORM_MARKER_RE = re.compile(r"(§§?|\bArt\.|\bArtikel\b|\bAnhang\b)")
GESETZ_TOKEN = r"[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*"
# Gesetz am Ende: 'BetrVG', 'SGB IX', 'KI-VO', 'RL 2001/23/EG', 'EGRL 88/2003'
# `\d{1,4}` (bis 2026-08-28: `\d{1,2}`): deutsche Abkuerzungen tragen oft einen
# VIERSTELLIGEN Jahres-Suffix - 'BLV 2026', 'AO 1977', 'MuSchG 2018',
# 'GO-BT 2025'; das CMS fuehrt sie ebenso ('AO1977', 'MuSchG2018'). Mit `\d{1,2}`
# galt 'BLV 2026' als KEIN Gesetz, mit zwei belegten Folgen (28.08.2026):
#   * die ';'-Trennung hielt den Teil fuer gesetzlos und hing das Gesetz des
#     NACHBARTEILS an -> '§§ 10 bis 20 BLV 2026 BBG' (falsches Gesetz),
#   * und weil `ges` leer blieb, wurden die elf Paragraphen nicht aufgeloest.
GESETZ_ENDE_RE = re.compile(
    # Ausgeschriebene EU-Rechtsakte mit Klammerzusatz stehen ebenfalls am Ende:
    # 'Art. 11, 21, 22 Durchfuehrungsverordnung (EU) Nr. 282/2011'. Ohne diesen
    # Zweig fand _gesetz_ende() dort kein Gesetz, und nur der LETZTE Artikel
    # behielt den Rechtsakt - 'Art. 11'/'Art. 21' standen nackt im Index
    # (Befund 28.08.2026, JUS.TAX.ANALYSIS.VAT.DE.FIXED_ESTABLISHMENT_CHECK).
    r"((?:Durchführungsverordnung|Durchfuehrungsverordnung|Durchführungsrichtlinie|"
    r"Richtlinie|Verordnung|DVO)\s*(?:\(\s*(?:EU|EG|EWG)\s*\)\s*)?"
    r"(?:Nr\.?\s*)?\d{1,4}/\d{2,4}|"
    r"(?:RL|EGRL|VO|EGV)\s+[\d/]+(?:/[A-Z]+)?|"
    r"(?:[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*|[A-ZÄÖÜ]{2,})"
    # Das Kuerzel KANN einen Schraegstrich tragen: 'FreizuegG/EU', 'AufenthV/EU'.
    # Ohne diesen Zweig schnitt die Zeichenklasse (kein '/') das Kuerzel ab und
    # `_gesetz_ende` lieferte nur den Rechtsraum 'EU'. Das war lange folgenlos,
    # weil ein nicht leerer Wert genuegte - seit der Rechtsraum verworfen wird,
    # galt '§ 12 FreizuegG/EU' als gesetzlos und erbte das Gesetz des Nachbarn
    # ('§ 12 FreizuegG/EU AEUV', gemessen 29.08.2026 in HIRING_FOREIGN_EMPLOYEE
    # und RIGHT_TO_WORK). Nur Rechtsraum-Suffixe sind erlaubt, damit
    # '2008/104/EG' nicht als Kuerzel durchgeht.
    r"(?:/(?:EU|EG|EWG|EAG))?"
    r"(?:\s+(?:[IVXL]{1,4}|\d{1,4}))?)\s*$"      # 'SGB IX', 'BLV 2026'
)
# Gesetz VORN: 'GewO § 106', 'SGB IX § 168', 'KI-VO Anhang III Nr. 4'
GESETZ_VORN_RE = re.compile(
    r"^((?:[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*|[A-ZÄÖÜ]{2,})"
    r"(?:\s+[IVXL]{1,4})?)\s+(§§?\s*\d|Art\.?\s*\d|Artikel\s*\d|Anhang\s|Anlage\s)"
)
# DEFEKT A (behoben 29.08.2026): MEHRERE vorangestellte Gesetze in EINER Zelle.
# 'GewO § 106, BGB §§ 611a, 241, 280, 823; OWiG § 130' - der ';'-Teil
# 'GewO § 106, BGB §§ 611a, 241, 280, 823' traegt ZWEI vorangestellte Gesetze.
# _umstellen() kennt nur das erste, haengt 'GewO' ans ENDE des ganzen Teils und
# vertauscht die Zuordnung damit ueber Kreuz: '§ 106' bekam BGB, die BGB-
# Paragraphen bekamen GewO. Im Index standen '§ 241 GewO', '§ 280 GewO',
# '§ 611a GewO', '§ 823 GewO' - vier Paragraphen, die es in der GewO nicht gibt,
# waehrend die BGB-Normen (alle im CMS) als fehlend galten.
# Anders als GESETZ_VORN_RE ist dieses Muster NICHT verankert; es findet die
# Trennstellen im Inneren. Gesplittet wird nur bei MEHR ALS EINER Fundstelle -
# bei genau einer bleibt alles wie bisher bei _umstellen().
GESETZ_VORN_TRENN_RE = re.compile(
    r"(?<![A-Za-zÄÖÜäöüß0-9\-])"
    r"(?:[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*|[A-ZÄÖÜ]{2,})"
    r"(?:\s+[IVXL]{1,4})?\s+(?=§§?\s*\d|Art\.?\s*\d|Artikel\s*\d)")
RECHTSSTAND_RE = re.compile(r"\s+(n\.\s?F\.?|a\.\s?F\.?)\s*$", re.IGNORECASE)
NUR_GESETZ_RE = re.compile(
    r"^(?:[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*|[A-ZÄÖÜ]{2,})"
    r"(?:\s+[IVXL]{1,4})?$"
)
ABKUERZUNG = {"abs", "nr", "satz", "s", "ziff", "lit", "halbs", "hs", "buchst", "alt",
              "n.f", "a.f", "ff", "f", "i.v.m", "iVm"}
NORM_MAXLEN = 130


# EU-Rechtsakte sind vollwertige Zitate, tragen aber kein § und kein 'Art.':
# 'Richtlinie 2014/67/EU', 'Verordnung (EU) 2021/821', 'RL 96/71/EG',
# 'Verordnung (EG) Nr. 593/2008'
# 'Durchfuehrungsverordnung' / 'DVO' fehlten bis 28.08.2026: im TAX-Bestand
# steht 'Art. 11, 21, 22 Durchfuehrungsverordnung (EU) Nr. 282/2011'
# ausgeschrieben, und ohne den Treffer bekam nur der LETZTE Artikel das Gesetz
# angehaengt - 'Art. 11' und 'Art. 21' standen ohne Rechtsakt im Index.
EU_AKT_RE = re.compile(
    r"\b(?:Durchführungsverordnung|Durchfuehrungsverordnung|Durchführungsrichtlinie|"
    r"Richtlinie|Verordnung|DVO|RL|VO|Beschluss)\s*(?:\(\s*(?:EU|EG|EWG)\s*\)\s*)?"
    r"(?:Nr\.?\s*)?\d{1,4}/\d{2,4}", re.IGNORECASE)


def _nur_gesetzesnamen(s):
    """'AGG; SGB IX; MuSchG' -> ['AGG', 'SGB IX', 'MuSchG'], sonst []."""
    teile = [t.strip() for t in re.split(r"[;/]", norm_txt(s)) if t.strip()]
    if len(teile) < 2:
        return []
    return teile if all(NUR_GESETZ_RE.match(t) for t in teile) else []


# Stufe 2b konkretisiert eine Normkategorie, indem sie die gefundenen
# Fundstellen ANHAENGT: 'BBiG / JArbSchG, soweit einschlaegig - im Einzelnen:
# § 10 BBiG; § 11 BBiG; …'. Genau dieser Zusatz treibt die Zelle ueber
# NORM_MAXLEN - und dann fallen die NEU EINGETRAGENEN Paragraphen weg, also
# gerade das Ergebnis der Fehlstellenkorrektur (Befund 28.08.2026: 13 Zeilen in
# 2 Skills, u.a. '§ 17 MuSchG; § 168 SGB IX; § 18 BEEG; § 5 PflegeZG').
#
# Die Laengenschranke deshalb NICHT anheben - gemessen ueber alle 348 Skills
# sind die uebrigen Zellen jenseits von 130 Zeichen Erlaeuterungstext, aus dem
# die Zerlegung nur Bruchstuecke macht ('§ 111-Briefing ersetzen.',
# '§ 65 HGB im Arbeitnehmerfall.'). Stattdessen wird der Kategorie-Vorspann
# abgeschnitten und nur die Fundstellenliste dahinter ausgewertet.
KONKRETISIERUNG_RE = re.compile(r"[-–—]\s*im\s+Einzelnen\s*:\s*", re.IGNORECASE)


def _konkretisierung(s):
    """Gibt den Teil hinter '- im Einzelnen:' zurueck, sonst den String selbst."""
    m = KONKRETISIERUNG_RE.search(s or "")
    if not m:
        return s
    rest = s[m.end():].strip()
    return rest if NORM_MARKER_RE.search(rest) else s


def is_norm_cell(s):
    """Innerhalb eines Rechtsgrundlagen-Abschnitts absichtlich weit gefasst -
    die Abschnittseingrenzung haelt Fliesstext-Tabellen schon draussen."""
    s = _konkretisierung(norm_txt(s))
    if not s:
        return False
    if len(s) > NORM_MAXLEN and not _lange_zitatliste(s):
        return False
    if is_platzhalter(s):
        return True
    if NORM_MARKER_RE.search(s) or EU_AKT_RE.search(s):
        return True
    if NUR_GESETZ_RE.match(s):
        return True
    return bool(_nur_gesetzesnamen(s))


def _lange_zitatliste(s):
    """Ausnahme von NORM_MAXLEN (seit 25.09.2026): eine Zelle ueber 130 Zeichen,
    die MIT EINEM ZITAT BEGINNT und nur aus Zitaten besteht, ist kein Fliesstext.
    Gemessen: 'Primaere Normen' in PUBLIC_SECTOR_CLASSIFICATION (148 Zeichen, zehn
    Tarif-§ 1) fiel still aus dem Index, TERMINATION.PROBATION (141) nur deshalb
    nicht, weil Schritt 3 dieselbe Liste traegt. Das Kriterium ist die vorhandene
    _reine_zitatliste() - kein zweites; mehrwortige Tarifnamen am Stueckende
    ('TVA-L Pflege', 'TVA-Forst Hessen') zaehlen dabei als EIN Wort, sonst
    scheitert sie an 'Pflege'/'Hessen'. Anders als dort ist KEIN Prosa-Vorspann
    erlaubt. Abschaltbar mit STUFE1_OHNE_ZITATLISTE_LANG=1 (A/B)."""
    if os.environ.get("STUFE1_OHNE_ZITATLISTE_LANG") == "1":
        return False
    if not re.match(r"(?:§§?\s*\d|Art\.?\s*\d|Artikel\s*\d)", s):
        return False
    teile = []
    for seg in s.split(";"):
        mt = _TARIFNAME_ENDE_RE.search(seg)
        teile.append(seg[:mt.start(1)] + "TVX" if mt else seg)
    return _reine_zitatliste(";".join(teile))


def _eu_akt_ende(s):
    """EU-Rechtsakt am ENDE des Strings - 'VO (EG) 883/2004', 'Richtlinie
    2008/104/EG', 'VO (EG) 593/2008 (Rom I)'.

    GESETZ_ENDE_RE kennt zwar EU-Formen, aber nur zwei enge: 'Richtlinie
    2023/970' (Wort + Zahl/Zahl am Ende) und 'RL 96/71' (Kuerzel + Ziffern
    DIREKT dahinter). Ein **bare 'VO' mit Klammer-Rechtsraum** faellt durch
    beide, und ein angehaengtes '/EG' bzw. '(Rom I)' verschiebt das Ende.

    Folge, gemessen 29.08.2026 - 9 Zellen, 14 verwaiste Eintraege, 8 Skills:
        'Art. 12, 13 VO (EG) 883/2004'  -> 'Art. 12' (ohne Gesetz)
                                           'Art. 13 VO (EG) 883/2004'
    Weil `_gesetz_ende` nichts fand, hatte `_zerlege_paragraphen` kein Gesetz
    zu verteilen, und nur das LETZTE Stueck behielt den Rechtsakt - der stand
    dort ja im Text. Der Fix gehoert deshalb hierher und nicht in die
    Aufzaehlungslogik: sobald das Gesetz erkannt ist, verteilt die vorhandene
    Vererbung es von selbst richtig.
    """
    t = norm_txt(s).rstrip(" .,;")
    treffer = list(EU_AKT_RE.finditer(t))
    if not treffer:
        return ""
    m = treffer[-1]
    rest = t[m.end():]
    # erlaubt sind nur noch Rechtsraum-Suffix und ein Kurzname in Klammern
    zusatz = re.match(r"^(/(?:EU|EG|EWG|EAG))?(\s*\([^)]{1,20}\))?\s*$", rest)
    if not zusatz:
        return ""
    return (t[m.start():m.end()] + (zusatz.group(1) or "")).strip()


def _gesetz_ende(s):
    s = norm_txt(s).rstrip(" .,;")
    s = re.sub(r"\s+(?:n\.\s?F\.?|a\.\s?F\.?)$", "", s, flags=re.IGNORECASE)
    m = GESETZ_ENDE_RE.search(s)
    if not m:
        return _eu_akt_ende(s)
    kand = m.group(1)
    # Verworfener Kandidat heisst NICHT "kein Gesetz". Bei
    # 'Richtlinie 2008/104/EG' trifft GESETZ_ENDE_RE das angehaengte **'EG'** -
    # das ist der RECHTSRAUM, nie ein Gesetz, und der Rechtsakt davor bleibt
    # ungesehen. Ohne dieses Nachfassen blieb 'Art. 3' aus
    # 'Art. 3, 5 Richtlinie 2008/104/EG' ohne Gesetz (gemessen 29.08.2026).
    # `_RECHTSRAUM` steht weiter unten im Modul - Python loest den Namen erst
    # beim Aufruf auf, die Reihenfolge ist also unkritisch.
    if kand.lower().rstrip(".") in ABKUERZUNG or kand in _RECHTSRAUM:
        return _eu_akt_ende(s)
    return kand


def _umstellen(s):
    """'GewO § 106 Abs. 1' -> '§ 106 Abs. 1 GewO'

    Steht dasselbe Gesetz schon am ENDE, wird es NICHT ein zweites Mal
    angehaengt. Bis 2026-08-28 machte diese Funktion aus
    'BGB §§ 305 ff., § 611a BGB' das Zitat '§§ 305 ff., § 611a BGB BGB' - im
    JUS_LABOR-Index belegt (ANALYSIS.WORKS_COUNCIL.AGREEMENT), und der daraus
    gebaute Crawl-Slug lautete '/bgb_bgb/', ein Gesetz, das es nicht gibt.
    """
    m = GESETZ_VORN_RE.match(s)
    if not m:
        return s
    gesetz = m.group(1)
    rest = s[m.end(1):].strip()
    if _gesetz_ende(rest).lower() == gesetz.lower():
        return rest
    return "%s %s" % (rest, gesetz)


UNTERGLIEDERUNG_RE = re.compile(
    r"\b(?:Abs|Nr|S|Satz|Halbs|HS|lit|Buchst|Alt|Unterabs)\b\.?", re.IGNORECASE)

# ============================================================ ff.-Auflösung
# `ff.` reicht bis zum Ende des Gesetzesabschnitts (Nutzerregel 28.08.2026).
# Diese Grenze steht NICHT im Zitat und ist nur aus der amtlichen Gliederung zu
# holen - das macht `abschnittsgrenzen.py` (amtliches XML von
# gesetze-im-internet.de bzw. Kapitel aus der CMS-doc_id fuer EU-Recht).
#
# WARUM EINE KARTEI UND KEIN DIREKTER AUFRUF: dieser Extraktor liest
# ausschliesslich Dokumente und geht NIE ins Netz - so ist ein Lauf
# reproduzierbar und laeuft auch offline. `abschnittsgrenzen.py` schreibt
# deshalb die aufgeloesten Paragraphenlisten in eine JSON-Kartei, und hier wird
# nur diese Kartei gelesen. Fehlt sie, bleibt ein `ff.`-Zitat als EINE Zeile
# stehen und es wird die Warnung `ff_ohne_abschnittsgrenze` gesetzt - es
# verschwindet also nicht still (Nutzervorgabe: weder Geist noch ignoriert).
#
# Expandiert wird entlang der TATSAECHLICHEN Paragraphenliste des amtlichen
# Textes, nie ueber eine Zahlenspanne: `range(850, 864)` verliert 850a-850l,
# 851a-851d, 855a - bibliotheksweit waeren das 52 von 204 Paragraphen, ein
# Viertel (gemessen 28.08.2026).
FF_KARTE_NAME = "_FF_KARTE.json"
_FF_KARTE = None
_FF_RE = re.compile(r"^(§§?|Art\.?|Artikel)\s*(\d+[a-z]?)\s*ff?\.", re.IGNORECASE)


def ff_karte(pfad=None):
    """Laedt die Kartei {"<GESETZ>|<START>": ["305", "306", ...]} einmalig.
    Reihenfolge: ausdruecklicher Pfad, dann Umgebungsvariable FF_KARTE, dann
    die Datei neben diesem Skript."""
    global _FF_KARTE
    if _FF_KARTE is not None and pfad is None:
        return _FF_KARTE
    kandidaten = [pfad, os.environ.get("FF_KARTE"),
                  os.path.join(os.path.dirname(os.path.abspath(__file__)), FF_KARTE_NAME)]
    karte = {}
    for k in kandidaten:
        if k and os.path.exists(k):
            try:
                with open(k, encoding="utf-8") as f:
                    roh = json.load(f)
                karte = {str(a).upper().replace(" ", ""): b
                         for a, b in (roh.get("karte") or roh).items()}
                break
            except Exception:
                karte = {}
    if pfad is None:
        _FF_KARTE = karte
    return karte


def ff_schluessel(eintrag):
    """'§§ 305 ff. BGB' -> ('BGB|305'); '' wenn kein ff.-Zitat oder kein Gesetz."""
    t = norm_txt(eintrag)
    m = _FF_RE.match(t)
    if not m:
        return ""
    gesetz = _gesetz_ende(t)
    if not gesetz:
        return ""
    return ("%s|%s" % (gesetz, m.group(2))).upper().replace(" ", "")


def expandiere_ff(eintraege, karte=None, offen=None):
    """Ersetzt jedes aufloesbare `ff.`-Zitat durch eine Zeile JE PARAGRAPH.

    Nicht aufloesbare bleiben unveraendert stehen und werden - wenn <offen>
    uebergeben wird - dort vermerkt, damit der Lauf sie meldet statt sie zu
    verschlucken."""
    karte = ff_karte() if karte is None else karte
    out = []
    for e in eintraege:
        s = ff_schluessel(e)
        liste = karte.get(s) if s else None
        if not liste:
            if s or _FF_RE.match(norm_txt(e)):
                if offen is not None:
                    offen.append(e)
            out.append(e)
            continue
        m = _FF_RE.match(norm_txt(e))
        marker = "§" if m.group(1).startswith("§") else "Art."
        gesetz = _gesetz_ende(norm_txt(e))
        for nr in liste:
            neu = "%s %s %s" % (marker, nr, gesetz)
            if neu not in out:
                out.append(neu)
    return out
BEREICH_RE = re.compile(r"^(\d+[a-z]?)\s*(?:-|–|—|bis)\s*(\d+)([a-z]?)$")


FUELLWORT = ("ggf.", "ggf", "ff.", "ff", "i.v.m.", "jeweils", "usw.", "u.a.",
             "entsprechend", "analog", "anlagen", "und")


def _mit_gesetz(stueck, gesetz, marker="§"):
    stueck = norm_txt(stueck)
    return "%s %s %s" % (marker, stueck, gesetz) if gesetz else "%s %s" % (marker, stueck)


def _expand_bereich(a_txt, b_num, b_suf, gesetz, marker="§"):
    """'140-147' -> acht Paragraphen. Drei Bauformen:
       '140-147'        rein numerisch
       '74 bis 75a'     Zahlbereich, letzter Paragraph mit Buchstabe
       '138d bis 138k'  BUCHSTABENreihe innerhalb derselben Zahl - bis
                        2026-08-28 falsch: die Funktion las 138 und 138k und
                        liess 138e bis 138j weg (Nutzerhinweis: "§§ heisst
                        MEHRERE Paragraphen").
    [] heisst: Bereich unklar, gar nicht aufteilen."""
    a = int(re.match(r"^(\d+)", a_txt).group(1))
    a_suf = a_txt[len(str(a)):]
    b = int(b_num)
    if b < a or b - a > 80:
        return []
    if a == b:
        if not (len(a_suf) == 1 and len(b_suf) == 1 and a_suf < b_suf):
            return []
        return [_mit_gesetz("%d%s" % (a, chr(c)), gesetz, marker)
                for c in range(ord(a_suf), ord(b_suf) + 1)]
    nums = [a_txt] if a_suf else []
    nums.extend(str(n) for n in range(a + (1 if a_suf else 0), b + 1))
    if b_suf:
        nums.append("%d%s" % (b, b_suf))
    return [_mit_gesetz(n, gesetz, marker) for n in nums]


def _zerlege_paragraphen(rest, gesetz, marker="§", mehrere=False):
    """Inhalt hinter dem §/§§-Marker in einzelne Paragraphen aufteilen.

    '12, 13, 15 Abs. 4, 16'  -> § 12, § 13, § 15 Abs. 4, § 16
    '8–19'                   -> § 8 ... § 19
    '74 bis 75a'             -> § 74, § 75, § 75a
    '87 Abs. 1 Nr. 1, 6, 7'  -> EIN Eintrag (die Kommaliste gehoert zur Nr.,
                                nicht zu einer Paragraphenaufzaehlung)

    `mehrere` = das Zitat begann mit **`§§`**. Dann ist die Kommaliste eine
    PARAGRAPHENaufzaehlung, auch wenn das erste Stueck eine Untergliederung
    traegt: '§§ 95 Abs. 3, 99 bis 101 BetrVG' sind § 95 Abs. 3, § 99, § 100 und
    § 101 - vier Eintraege, nicht einer.

    Warum am Doppelparagraphen und nicht am Inhalt: '§ 87 Abs. 1 Nr. 1, 6, 7'
    und '§§ 241 Abs. 2, 242, 612a' sind strukturell nicht zu unterscheiden -
    beide sind "Zahl mit Untergliederung, dann Kommaliste aus Zahlen". Das
    einzige verlaessliche Signal ist das `§§` (Nutzervorgabe 28.08.2026:
    "Kennzeichen fuer mehrere Paragraphen ist immer §§"). Deshalb MUSS diese
    Unterscheidung hier fallen, wo der Marker noch sichtbar ist - im Index steht
    spaeter nur noch ein einfaches '§'.
    """
    rest = norm_txt(rest)
    # Ein ANGEHAENGTES Bindewort blockiert die Bereichsauflösung: aus
    # 'FVerlV, insbesondere §§ 1 bis 3 und § 8' schneidet der Segment-Split
    # '§§ 1 bis 3 und' heraus; BEREICH_RE passt darauf nicht mehr, und die drei
    # Paragraphen fielen weg. ' und ' mitten im String trennt weiter unten.
    rest = re.sub(r"\s+(?:und|sowie|bzw\.?|oder)\s*$", "", rest, flags=re.IGNORECASE)
    if not rest:
        return []
    stuecke = [s.strip() for s in re.split(r",| und ", rest) if s.strip()]
    if len(stuecke) > 1 and UNTERGLIEDERUNG_RE.search(stuecke[0]) and not mehrere:
        return []                       # Kommaliste gehoert zur Untergliederung
    out = []
    # NUMMERNLISTE HINTER EINER UNTERGLIEDERUNG (seit 25.09.2026). Auch nach
    # '§§' gehoert eine Kommaliste manchmal zur Nr./zum Abs. des vorigen Stuecks:
    # '§§ 13, 14 Abs. 1 Nr. 1, 2, 3 und 5 BBiG' las sich als § 13, § 14 Abs. 1
    # Nr. 1, § 2, § 3, § 5 BBiG - GRUENE SCHEINTREFFER auf nie zitierten
    # Paragraphen. Strukturell gleich gebaut ist aber '§§ 87 Abs. 1 Nr. 6, 94, 95
    # BetrVG', wo 94/95 echte Paragraphen sind. Unterscheidung am WERT, eng
    # gefasst: Fortsetzung nur, wenn die nackte Zahl die letzte Untergliederungs-
    # zahl um 1-3 uebersteigt UND kleiner als der Paragraph selbst ist. So bleiben
    # '§§ 58 Abs. 1, 50' (Sprung 49) und '§§ 241 Abs. 2, 242' (> Paragraph)
    # Paragraphenlisten. Abschaltbar mit STUFE1_OHNE_NR_FORTSETZUNG=1 (A/B).
    fortsetzung = os.environ.get("STUFE1_OHNE_NR_FORTSETZUNG") != "1"
    offen = None                        # (Paragraph, letzte Untergliederungszahl)
    for st in stuecke:
        if fortsetzung and offen and out and re.fullmatch(r"\d+", st):
            n = int(st)
            schwanz = (" " + gesetz) if gesetz else ""
            basis = out[-1][:-len(schwanz)] if schwanz and out[-1].endswith(schwanz) else \
                (out[-1] if not schwanz else None)
            if basis is not None and re.search(r"\d$", basis) \
                    and offen[1] < n <= offen[1] + 3 and n < offen[0]:
                out[-1] = basis + ", " + st + schwanz
                offen = (offen[0], n)
                continue
        offen = None
        mu = re.match(r"^(\d+)[a-z]?\s+.*\b(?:Abs|Nr|Satz|S)\.?\s*(\d+)$", st)
        if mu:
            offen = (int(mu.group(1)), int(mu.group(2)))
        mb = BEREICH_RE.match(st)
        if mb:
            erweitert = _expand_bereich(mb.group(1), mb.group(2), mb.group(3),
                                        gesetz, marker)
            if not erweitert:
                return []
            out.extend(erweitert)
            continue
        if not re.match(r"^\d+[a-z]?\b", st):
            if st.lower().rstrip(".") in FUELLWORT or len(st) <= 6:
                continue                # 'ggf.', 'ff.' - Fuellwort, ueberspringen
            return []                   # kein Paragraphenanfang -> nicht aufteilen
        # Traegt das Stueck sein Gesetz SELBST ('§§ 257 HGB, 147 AO' -> '257 HGB'),
        # darf nicht zusaetzlich das geerbte angehaengt werden - sonst entsteht
        # '§ 257 HGB AO', ein Zitat mit zwei Gesetzen, das Stufe 2 an 'AO' prueft
        # und als fehlend meldet (belegt 29.08.2026, OFFBOARDING.CHECKLIST).
        eigenes = _gesetz_ende(st)
        out.append(_mit_gesetz(st, "", marker) if eigenes
                   else _mit_gesetz(st, gesetz, marker))
    return out if len(out) > 1 else []


# DEFEKT B (behoben 29.08.2026): ein NACHLAUFENDER ZUSATZ verdeckt das Gesetz.
# '§§ 102, 103 BetrVG; §§ 168, 178 SGB IX, soweit einschlaegig' -> _gesetz_ende()
# sah am Ende 'einschlaegig' statt 'SGB IX', fand also KEIN Gesetz; die
# Nachbaruebernahme in _gesetze_je_teil() hing daraufhin das Gesetz des VORIGEN
# Teils an, und im Index stand '§§ 168, 178 SGB IX, soweit einschlaegig BetrVG'.
# Zweiter belegter Fall: '§ 12 AO; §§ 38, 49 EStG als deutsche Routinganker'
# -> '§ 49 EStG als deutsche Routinganker AO'. Beide erzeugten einen
# Phantom-Crawlbedarf UND eine Phantom-Sperre, obwohl die gemeinte Norm
# (§ 168/178 SGB IX, § 49 EStG) laengst im CMS liegt.
#
# DER WAECHTER IST DER KERN: abgeschnitten wird nur ein Schwanz, der KEIN
# gesetzesfoermiges Token enthaelt. Ohne ihn frisst die Regel echte Zitate -
# '§ 1 AStG und FVerlV' wuerde zu '§ 1 AStG', und die FVerlV (ein echtes,
# nicht im CMS liegendes Crawl-Ziel) verschwaende spurlos.
#
# `{2,}` statt `*` beim ersten Wort ist NICHT kosmetisch: der Buchstabe einer
# Untergliederung ist ein EINZELNER Kleinbuchstabe und sieht wie ein Prosa-Anfang
# aus. Mit `*` fraß die Regel ihn weg - aus 'Art. 5 Abs. 1 lit. e, Art. 9 …'
# wurde 'Art. 5 Abs. 1 lit. DSGVO' und aus 'Art. 2 Abs. 1 Buchst. e EG' wurde
# 'Art. 2 Abs. 1 Buchst. EG' (gemessen 29.08.2026 ueber die Bibliothek). Ein
# echter Zusatz ist immer laenger ('als', 'soweit', 'insbesondere').
_PROSA_SCHWANZ_RE = re.compile(
    r",?\s+[a-zäöüß][A-Za-zÄÖÜäöüß\-]{2,}(?:\s+[A-Za-zÄÖÜäöüß\-]+)*\s*$")
_GESETZFOERMIG_RE = re.compile(
    r"[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*[A-ZÄÖÜ]|[A-ZÄÖÜ]{2,}")


def _ohne_prosa_schwanz(s):
    """'§§ 168, 178 SGB IX, soweit einschlaegig' -> '§§ 168, 178 SGB IX'.

    Schneidet einen nachlaufenden Erlaeuterungstext ab, damit das Gesetz wieder
    am Ende steht. Enthaelt der Schwanz ein gesetzesfoermiges Token, bleibt er
    unangetastet - siehe Waechter-Begruendung oben."""
    s = norm_txt(s)
    m = _PROSA_SCHWANZ_RE.search(s)
    if not m or _GESETZFOERMIG_RE.search(m.group(0)):
        return s
    gekuerzt = s[:m.start()].strip(" ,;")
    return gekuerzt if gekuerzt else s


# DEFEKT C (behoben 10.09.2026): ein NACHGESTELLTER NACKTER GESETZESNAME wird
# zum Gesetz der davorstehenden Paragraphen.
#
#   split_normen('§§ 138, 242, 612a BGB; AGG')
#       -> ['§ 138 AGG', '§ 242 AGG', '§ 612a BGB']      (zwei erfundene Normen)
#
# `_gesetz_ende` liest das LETZTE Token, hier also 'AGG'; die Zerlegung verteilt
# es auf alle Paragraphen ausser dem letzten, der sein 'BGB' aus dem Resttext
# behaelt. Der Semikolon-Zweig in `split_normen` greift nicht, weil er ZWEI
# normtragende Teile verlangt - 'AGG' allein traegt kein Normmuster.
#
# WARUM DER FIX VOM 29.08. IHN NICHT MITNAHM: Waechter 1 von
# `_ohne_prosa_schwanz` schneidet nur Schwaenze OHNE gesetzesfoermiges Token ab,
# und 'AGG' IST eines. Derselbe Waechter, der die FVerlV in '§ 1 AStG und FVerlV'
# vor dem Verschwinden bewahrt, haelt damit den Fehlbezug fest. Die Aufloesung
# der Zielkollision: der Name wird nicht weggeschnitten, sondern als EIGENER
# Eintrag gefuehrt. Ein Gesetzesname ohne Paragraph ist ein gueltiger - und
# sperrender - Befund (dieselbe Semantik wie der 'einschlaegiges DBA'-Zweig
# unten), die Sperre bleibt also, benennt aber ab jetzt das Richtige.
#
# GEMESSEN im Index _260908_ (2774 Normzeilen): 21 Klumpen in 20 Skills,
# darunter '§ 103 BetrVG, BDSG, PflegeZG/FPfZG' und '§ 90 Abs. 3 AO und GAufzV'.
#
# VIER WAECHTER - jeder haelt einen Fall, der sonst kaputtgeht:
#  1. TRENNZEICHEN ODER BINDEWORT ist Pflicht, blosser Zwischenraum genuegt
#     NICHT. Sonst zerfaellt 'SGB IX' ('IX' passt auf [A-ZÄÖÜ]{2,}) und
#     'BLV 2026'/'AO 1977' verlieren ihr vierstelliges Jahressuffix.
#  2. Der Name muss in der KARTEI stehen. 'Landesdatenschutzrecht' in
#     '§ 26 BDSG; einschlaegiges Landesdatenschutzrecht' ist keine Abkuerzung -
#     der String bleibt unangetastet, wie bisher.
#  3. Der KOPF muss ein vollstaendiges Zitat MIT eigenem Gesetz sein. Nur dann
#     ist belegt, dass der nachgestellte Name nicht das Gesetz der Paragraphen
#     ist. 'AGG, insbesondere §§ 1, 3, 7' (Gesetz VORN) bleibt damit dem
#     Bindeglied-Fix ueberlassen.
#  4. RECHTSRAUM-SUFFIX ausgenommen: in 'Art. 3, 5 Richtlinie 2008/104/EG'
#     steht hinter dem '/' kein Gesetz, sondern der Rechtsraum. Ohne diesen
#     Waechter entstuende ein Eintrag 'EG' und der Rechtsakt verlore sein
#     Suffix (deshalb zusaetzlich: nach '/' nie hinter einer Ziffer schneiden).
_KONNEKTOR = (r"und|sowie|bzw\.?|oder|ergänzend|ergaenzend|zudem|ggf\.?|"
              r"gegebenenfalls|jeweils|i\.\s*V\.\s*m\.?|"
              r"in\s+Verbindung\s+mit|in\s+Verbindung")
_NACH_SEP_RE = re.compile(
    r"(?:[,;/]\s*(?:(?:%(k)s)\s+)*|\s+(?:%(k)s)\s+(?:(?:%(k)s)\s+)*)$"
    % {"k": _KONNEKTOR}, re.IGNORECASE)
_NACH_NAME_RE = re.compile(
    r"([A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*"
    r"|[A-ZÄÖÜ]{2,})\s*$")
# Ein Bindewort, hinter dem NICHTS mehr steht ('§ 8 Abs. 2 EStG i.V.m.',
# '§ 18 BEEG in Verbindung'). Im Index _260908_ zehn solcher Zeilen. Der Rumpf
# ist ein sauberes Zitat, das Bindewort haengt nur in der Luft - Stufe 2 sucht
# aber nach dem GANZEN String und findet ihn nie.
_LEERER_KONNEKTOR_RE = re.compile(
    r"(?:[,;]\s*)?\s+(?:%s)\s*[.,;]?\s*$" % _KONNEKTOR, re.IGNORECASE)


def _nachgestellte_namen(cell):
    """'§ 103 BetrVG, BDSG, PflegeZG/FPfZG' -> ('§ 103 BetrVG', ['BDSG',
    'PflegeZG', 'FPfZG']). Ohne Treffer: (cell, [])."""
    # Abschalter wie beim Bindeglied-Fix: nur so misst die A/B-Strecke den
    # Erkenner gegen sich selbst, am GLEICHEN Dokumentstand. Ein Vergleich gegen
    # eine aeltere Indexdatei mischt Dokumentaenderungen mit hinein - genau der
    # Fehlalarm, den die SKILL.md bei "Neubau statt Nachtrag" beschreibt.
    if os.environ.get("STUFE1_OHNE_DEFEKT_C") == "1":
        return norm_txt(cell), []
    from gesetzesnamen import ist_bekannte_abkuerzung as _ist_abk
    rest = norm_txt(cell).rstrip(" .")
    namen = []
    while True:
        m = _NACH_NAME_RE.search(rest)
        if not m:
            break
        name = m.group(1)
        if name in _RECHTSRAUM or not _taugt_als_gesetz(name) or not _ist_abk(name):
            break
        kopf = rest[:m.start()]
        s = _NACH_SEP_RE.search(kopf)
        if not s:
            break
        # Waechter 4, zweite Haelfte: '.../104' + '/EG' - nach einem Schraegstrich
        # hinter einer Ziffer steht ein Rechtsraum, kein zweites Gesetz.
        if kopf[s.start():].lstrip().startswith("/") and re.search(r"\d\s*$", kopf[:s.start()]):
            break
        kopf = kopf[:s.start()].rstrip(" .,;/")
        if not (NORM_MARKER_RE.search(kopf) and _gesetz_ende(kopf)):
            break                              # Waechter 3
        namen.append(name)
        rest = kopf
    return rest, list(reversed(namen))


_TARIFNAME_ENDE_RE = re.compile(
    r"(?<![\w-])((?!TVG(?![\w-]))TV[A-Za-zÄÖÜäöüß]*(?:-[A-Za-zÄÖÜäöüß]+)*"
    r"(?:\s+[A-ZÄÖÜ][A-Za-zÄÖÜäöüß]*(?:-[A-Za-zÄÖÜäöüß]+)*){0,2})\s*$")


def _teil_gesetz(teil):
    """Welches Gesetz gehoert zu diesem Semikolon-Teil? Drei Bauformen, in
    dieser Reihenfolge: vorn ('AGG § 12'), hinten ('§ 12 AGG') und als
    Kopfwort mit Fuelltext dazwischen ('JArbSchG insbesondere §§ 8, 9')."""
    # TARIFWERK AM ENDE (seit 25.09.2026): mehrwortige Tarifnamen ('TVA-L Pflege',
    # 'TVA-Forst Hessen', 'TV Prakt-L') liest _gesetz_ende() nicht - der Teil galt als
    # gesetzlos und ERBTE das Gesetz des Nachbarn: '§ 1 TVA-L BBiG; § 1 TVA-L Pflege'
    # wurde '§ 1 TVA-L Pflege BBiG', und Stufe 2 meldete das als § 1 BBiG GEFUNDEN
    # (gruener Scheintreffer im falschen Werk). TVG ist das Gesetz und bleibt aussen vor.
    # Abschaltbar mit STUFE1_OHNE_TARIFNAME=1.
    if os.environ.get("STUFE1_OHNE_TARIFNAME") != "1" and NORM_MARKER_RE.search(teil):
        mt = _TARIFNAME_ENDE_RE.search(norm_txt(teil))
        if mt:
            return mt.group(1)
    m = GESETZ_VORN_RE.match(teil)
    if m:
        return m.group(1)
    g = _gesetz_ende(teil)
    if g:
        return g
    # Nachlaufender Zusatz (Defekt B): erst kuerzen, dann erneut hinten suchen.
    if NORM_MARKER_RE.search(teil):
        g = _gesetz_ende(_ohne_prosa_schwanz(teil))
        if g:
            return g
    # Bindeglied im Teil: 'JArbSchG, insbesondere §§ 8, 9' - das Gesetz steht
    # vorn, das Zitat dahinter. Ohne diesen Zweig fand keine der Bauformen es,
    # der Teil galt als gesetzlos und erbte das Gesetz des NACHBARN: aus
    # 'ArbSchG §§ 3, 12; JArbSchG, insbesondere §§ 8, 9' wurde '§ 8 ArbSchG'
    # statt '§ 8 JArbSchG' (gemessen 29.08.2026). Derselbe Fehler traf
    # '§§ 39, 40 AufenthG; BeschV, insbesondere §§ 30, 32'.
    m = re.search(r"^.*?(?:,|;)?\s*(?:insbesondere|namentlich|vor\s+allem|u\.\s*a\.)\s+"
                  r"(?=§§?\s*\d|Art\.?\s*\d|Artikel\s*\d)", teil, re.IGNORECASE)
    if m:
        g = _vorspann_gesetz(m.group(0))
        if g:
            return g
    kopf = teil.split()[0] if teil.split() else ""
    return kopf if NUR_GESETZ_RE.match(kopf) else ""


def _gesetze_je_teil(teile):
    """Je Semikolon-Teil das Gesetz; fehlt es, wird es vom Nachbarn uebernommen
    (erst vorwaerts, dann rueckwaerts). 'ArbSchG § 2 Abs. 2 Nr. 2; §§ 15, 16'
    meint zweimal ArbSchG - ohne diese Uebernahme verloere der zweite Teil
    sein Gesetz und stuende als '§ 15' ohne Zuordnung im Index."""
    gesetze = [_teil_gesetz(t) for t in teile]
    letztes = ""
    for i, g in enumerate(gesetze):
        if g:
            letztes = g
        elif letztes:
            gesetze[i] = letztes
    naechstes = ""
    for i in range(len(gesetze) - 1, -1, -1):
        if gesetze[i]:
            naechstes = gesetze[i]
        elif naechstes:
            gesetze[i] = naechstes
    return gesetze


# BINDEGLIED-FIX - gebaut 29.08.2026 auf die Nutzerregel:
#   "'insbesondere' meint immer den §, Artikel, Ziffer DESSELBEN zuvor
#    genannten Gesetzes."
#
# 'AGG, insbesondere §§ 1, 3, 7' - das Gesetz steht am ANFANG, zwischen ihm und
# dem § steht ein Bindeglied. GESETZ_VORN_RE verlangt das § direkt hinter dem
# Gesetz, deshalb standen die Paragraphen bis hierher OHNE Gesetz im Index.
# Gemessen am 29.08.2026: 95 verschiedene Eintraege, 248 Vorkommen, 45 Skills -
# und ein gesetzloser Paragraph ist im CMS nie auffindbar, erzeugt also Sperren
# und unbrauchbare Crawl-Ziele fuer Zitate, die im Dokument voellig korrekt sind.
#
# WARUM DER VORGAENGERVERSUCH (28.08.2026) SCHEITERTE - das ist der Kern:
# er entschied JE ZELLE. Seine Waechter sahen nur den unmittelbaren Rest; trug
# der ERSTE Paragraph kein Gesetz, ein spaeterer aber schon, griff die Uebernahme
# trotzdem und ueberschrieb das richtige Gesetz. 112 falsche Zuordnungen:
#     '§ 12 ArbSchG'  -> '§ 12 AGG'          '§ 22 JArbSchG' -> '§ 22 ArbSchG'
#     '§ 130 OWiG'    -> '§ 130 ArbSchG OWiG'  '§ 43 GmbHG'  -> '§ 43 AktG GmbHG'
#     '§ 1 KStG'      -> '§ 1 GewStG'        'Art. 35 DSGVO' -> 'Art. 35 BDSG DSGVO'
#
# DIESER FIX ENTSCHEIDET JE SEGMENT, ALSO JE ZITAT. Das Vorspann-Gesetz wird
# NUR in Segmente gesetzt, die selbst keines tragen (`gesetze[i]` leer). Alle
# sechs Faelle oben tragen ihr Gesetz selbst und bleiben damit unberuehrt; sie
# stehen als Regressionsfaelle im Selbsttest.
#
# ZWEITER WAECHTER - nur ABKUERZUNGEN werden vererbt, keine ausgeschriebenen
# Namen. 'Betriebsverfassungsgesetz, insbesondere §§ 5, 50 ...' bekommt NICHTS,
# obwohl die Regel es hergaebe: Stufe 2 matcht gegen `metadata.jurabk`, und ein
# ausgeschriebener Name trifft dort nicht. Aus '§ 5' (nicht automatisiert
# pruefbar) wuerde '§ 5 Betriebsverfassungsgesetz' = nicht gefunden - ein
# NEUER, harter Sperrgrund statt einer Verbesserung. Ein Klammerzusatz zaehlt
# dagegen als Abkuerzung: 'Arbeitszeitgesetz (ArbZG), insbesondere §§ 5, 7'
# vererbt 'ArbZG'.
_BINDEGLIED_RE = re.compile(
    r"(?:,|;)?\s*(?:insbesondere|namentlich|vor\s+allem|u\.\s*a\.|hier|dort)\s*$|:\s*$",
    re.IGNORECASE)
# Abkuerzung in Klammern: 'Arbeitszeitgesetz (ArbZG)' -> 'ArbZG'.
_KLAMMER_ABK_RE = re.compile(
    r"\(((?:[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*|[A-ZÄÖÜ]{2,})"
    r"(?:\s+[IVXL]{1,4})?)\)")
# Rechtsraum, nicht Gesetz: '(EU)' in 'Richtlinie (EU) 2023/970' ist kein
# Gesetzeskuerzel. Ohne diese Sperre erbt der Artikel 'EU' statt des Rechtsakts.
_RECHTSRAUM = {"EU", "EG", "EWG", "EAG", "EGKS"}
# Eine blosse roemische Zahl ist NIE ein Gesetz. 'EU-KI-VO …: Art. 3, 4, 14, 26,
# 50, Anhang III' liess `_gesetz_ende` das 'III' als Gesetz lesen (es passt auf
# [A-ZÄÖÜ]{2,}) und machte daraus 'Art. 3 III' … 'Art. 50 III' - fuenf erfundene
# Fundstellen in einer Zelle (gemessen 29.08.2026).
_ROEMISCH_RE = re.compile(r"^[IVXLCDM]+$")


def _taugt_als_gesetz(g):
    """Waechter fuer JEDE vorangestellte Gesetzesuebernahme."""
    g = (g or "").strip()
    return bool(g) and g not in _RECHTSRAUM and not _ROEMISCH_RE.match(g)


def _vorn_gesetz(seg):
    """Zitatloses Segment am Zellanfang -> Gesetz, falls es eines IST.

    Deckt die Bauform 'Abgabenordnung §§ 12, 13' / 'Grundgesetz Art. 9 Abs. 3' /
    'RL 2003/88/EG, Art. 3, 5, 6' ab: das Gesetz steht vorn, aber weder direkt
    am § (dann griffe GESETZ_VORN_RE) noch hinter einem Bindeglied.

    ZWEI WAECHTER, sonst entstehen erfundene Fundstellen:
      * ein ausgeschriebener Name wird ueber die Kartei AUFGELOEST, nie roh
        angehaengt (Stufe 2 matcht `metadata.jurabk`);
      * ein Kuerzel wird nur uebernommen, wenn es in der Kartei als amtliche
        Abkuerzung steht - 'BetrVG-Normanker:' sieht formal wie eine Abkuerzung
        aus und wuerde sonst als Gesetz gelten.
    """
    # Derselbe Abschalter wie beim Bindeglied-Fix: nur so misst die A/B-Strecke
    # den GANZEN Aenderungssatz vom 29.08.2026 statt nur die Haelfte.
    if os.environ.get("STUFE1_OHNE_BINDEGLIED") == "1":
        return ""
    from gesetzesnamen import abkuerzung as _abk, ist_bekannte_abkuerzung as _ist_abk
    rumpf = norm_txt(seg).strip(" ,;:-–—")
    if not rumpf or NORM_MARKER_RE.search(rumpf):
        return ""
    eu = list(EU_AKT_RE.finditer(rumpf))
    if eu:
        treffer = eu[-1].group(0).strip()
        rest = rumpf[eu[-1].start() + len(treffer):]
        m = re.match(r"^/(EU|EG|EWG|EAG)\b", rest)
        return treffer + (m.group(0) if m else "")
    a = _abk(rumpf)
    if a:
        return a
    g = _gesetz_ende(rumpf)
    if _taugt_als_gesetz(g):
        a = _abk(g)
        if a:
            return a
        if _ist_abk(g):
            return g
    return ""


def _vorspann_gesetz(vorspann):
    """Gesetz aus dem Text VOR dem ersten Zitat - nur mit Bindeglied.

    Gibt '' zurueck, wenn kein Bindeglied dasteht (dann ist die Nutzerregel
    nicht einschlaegig) oder wenn nur ein ausgeschriebener Name gefunden wird
    (zweiter Waechter, s.o.)."""
    # Abschalter fuer die A/B-Messung: nur so laesst sich die Wirkung DIESES
    # Fixes von anderen Erkenneraenderungen desselben Tages trennen. Am
    # 29.08.2026 war genau das noetig - eine Messung gegen einen aelteren
    # JSON-Stand schrieb dem Fix 59 fremde Aenderungen zu.
    if os.environ.get("STUFE1_OHNE_BINDEGLIED") == "1":
        return ""
    v = norm_txt(vorspann)
    if not v or not _BINDEGLIED_RE.search(v):
        return ""
    rumpf = _BINDEGLIED_RE.sub("", v).strip(" ,;:")
    if not rumpf:
        return ""
    # Traegt der Vorspann SELBST ein Zitat, ist das Gesetz trotzdem das zuvor
    # genannte: '§§ 305 bis 310 BGB, insbesondere § 307 und § 309 Nr. 13' meint
    # zweimal BGB. `_gesetz_ende` liest es vom Ende des Vorspanns.
    # (Bis 29.08.2026 stand hier ein `return ""` - das liess 7 von 13
    # Restfaellen unaufgeloest, darunter 'ArbSchG §§ 3, 12; JArbSchG',
    # '§§ 1-4 NachwG' und 'Art. 76 bis 82 GG'.)
    # Der Fall '§ 14 TzBfG, insbesondere Abs. 4' ist davon NICHT betroffen:
    # dort folgt kein neues §-Segment, die Zelle bleibt ein Eintrag.
    # EU-RECHTSAKTE ZUERST - sonst erbt der Artikel den RECHTSRAUM statt des
    # Rechtsakts. Gemessen am 29.08.2026, bevor dieser Zweig da war:
    #   'VO (EG) 987/2009, insbesondere Art. 14-16' -> 'Art. 14 EG'      (falsch)
    #   'Richtlinie (EU) 2023/970, insbesondere Art. 4' -> 'Art. 4 EU'   (falsch)
    #   'Richtlinie 2008/104/EG, insbesondere Art. 5'   -> 'Art. 5 EG'   (falsch)
    # 28 solcher Eintraege waeren entstanden - jeder eine erfundene Fundstelle,
    # die im CMS als Gesetz 'EU' bzw. 'EG' gesucht und nie gefunden wuerde.
    # Der ganze Rechtsakt ist die richtige Antwort, nicht sein Klammerzusatz.
    eu = list(EU_AKT_RE.finditer(rumpf))
    if eu:
        treffer = eu[-1].group(0).strip()
        # '...2008/104/EG': das angehaengte Rechtsraumkuerzel gehoert zur Nummer.
        m = re.match(r"^(.*\d{1,4}/\d{2,4})", rumpf[eu[-1].start():])
        rest = rumpf[eu[-1].start() + len(treffer):]
        m2 = re.match(r"^/(EU|EG|EWG|EAG)\b", rest)
        return treffer + (m2.group(0) if m2 else "")

    # Klammerzusatz gewinnt vor dem ausgeschriebenen Namen davor - aber ein
    # Rechtsraumkuerzel ist KEINE Gesetzesabkuerzung.
    treffer = [t for t in _KLAMMER_ABK_RE.findall(rumpf) if t not in _RECHTSRAUM]
    if treffer:
        return treffer[-1]

    # Das Gesetz muss am Ende des Rumpfs stehen und Abkuerzungsform haben.
    # 'Arbeitszeitgesetz (ArbZG), Gesetze im Internet' faellt hier durch -
    # am Ende steht Prosa, nicht das Gesetz.
    g = _gesetz_ende(rumpf)
    if g in _RECHTSRAUM:
        g = ""

    # AUSGESCHRIEBENE NAMEN AUFLOESEN (Nutzerklarstellung 29.08.2026:
    # 'Betriebsverfassungsgesetz, insbesondere §§ 5, 50' meint § 5 und § 50
    # DIESES Gesetzes). Nicht der Name wird angehaengt, sondern seine amtliche
    # Abkuerzung - sonst trifft Stufe 2 kein `metadata.jurabk` und aus einem
    # weichen 'nicht pruefbar' wuerde ein hartes 'nicht gefunden'.
    # Reihenfolge: der ganze Rumpf zuerst (er IST oft der Name), dann das, was
    # `_gesetz_ende` gefunden hat - Letzteres faengt die Bindestrich-Namen wie
    # 'Arbeitnehmer-Entsendegesetz', die formal wie eine Abkuerzung aussehen.
    from gesetzesnamen import abkuerzung as _abk, ist_bekannte_abkuerzung as _ist_abk
    for kandidat in (rumpf, g):
        if kandidat:
            a = _abk(kandidat)
            if a:
                return a
    # WEISSLISTE: nur ein Kuerzel, das die Kartei als amtliche Abkuerzung kennt,
    # darf uebernommen werden. 'BetrVG-Normanker:' sieht formal wie eine
    # Abkuerzung aus - `_gesetz_ende` liefert es, und ohne diesen Waechter stand
    # '§ 87 Abs. 1 Nr. 1 BetrVG-Normanker' im Index (gemessen 29.08.2026).
    # Ein unbekanntes Kuerzel laesst das Zitat lieber ohne Gesetz: das ist ein
    # weicher Befund, eine erfundene Fundstelle waere ein harter.
    return g if _taugt_als_gesetz(g) and _ist_abk(g) else ""
_NACKTER_EU_AKT_RE = re.compile(
    r"^(?P<akt>(?:Richtlinie|Verordnung|RL|VO)\s*(?:\((?:EU|EG|EWG)\)\s*)?(?:Nr\.\s*)?"
    r"\d{2,4}/\d{1,4}(?:/(?:EU|EG|EWG))?)\s*,?\s*(?:insbesondere)?\s*$")
_ARTIKEL_BRUCHSTUECK_RE = re.compile(r"^Art(?:ikel|\.)\s*\d")


def _eu_akt_vorn_zusammenfuehren(teile):
    """EU-Rechtsakt VORN, Artikel dahinter (05.10.2026): 'Richtlinie 2005/36/EG,
    insbesondere Art. 3, 4 ...' und 'Verordnung (EU) 2024/1689 Art. 4 in
    Verbindung mit ...' zerfielen in den nackten Rechtsakt und ein Artikel-
    Bruchstueck OHNE Gesetz - in der Crawl-Liste stand dann 'Art. 3, 4, 5 bis 7,
    10 bis 14 und Titel III Kapitel III' als eigenes Werk.

    Bewusst NACH der regulaeren Zerlegung und eng: nur ein Paar aus nacktem
    Rechtsakt und unmittelbar folgendem Artikelteil, der selbst KEINEN Rechtsakt
    traegt. Ein frueher Ansatz vor der Zerlegung (A/B 05.10.: 17 Skills geaendert)
    fasste sauber einzeln erkannte Artikel zusammen und verlor Nachbar-Rechtsakte."""
    if os.environ.get("STUFE1_OHNE_EU_AKT_VORN") == "1":
        return teile
    out, i = [], 0
    while i < len(teile):
        m = _NACKTER_EU_AKT_RE.match(teile[i])
        if (m and i + 1 < len(teile) and _ARTIKEL_BRUCHSTUECK_RE.match(teile[i + 1])
                and not EU_AKT_RE.search(teile[i + 1])):
            out.append(f"{teile[i + 1].strip()} {m.group('akt').strip()}")
            i += 2
            continue
        out.append(teile[i])
        i += 1
    return out


def split_normen(cell):
    return _eu_akt_vorn_zusammenfuehren(_split_normen_roh(cell))


def _split_normen_roh(cell):
    cell = _konkretisierung(norm_txt(cell))
    if not cell:
        return []
    if is_platzhalter(cell):
        return [cell]
    namen = _nur_gesetzesnamen(cell)
    if namen:
        return namen
    if EU_AKT_RE.search(cell) and not NORM_MARKER_RE.search(cell):
        return [cell]

    # DEFEKT C: nachgestellte nackte Gesetzesnamen abtrennen (Begruendung und
    # die vier Waechter bei `_nachgestellte_namen`). Das muss VOR der
    # Semikolon-Trennung stehen, sonst ist der Name schon Gesetzeskandidat.
    kopf, nachnamen = _nachgestellte_namen(cell)
    if nachnamen:
        eigenes = _gesetz_ende(kopf)
        # Ein Name, der das Gesetz des Zitats WIEDERHOLT ('§ 33 GmbHG i.V.m.
        # GmbHG'), ist keine zweite Fundstelle, sondern eine Dopplung.
        weiter = [n for n in nachnamen if n != eigenes]
        return split_normen(kopf) + weiter

    # Bindewort ohne Anschluss ('§ 8 Abs. 2 EStG i.V.m.'): der Rumpf ist das
    # Zitat. Nur kuerzen, wenn danach noch ein vollstaendiges Zitat MIT Gesetz
    # steht - sonst wuerde aus einem Bruchstueck ein scheinbar sauberer Eintrag.
    m_leer = (None if os.environ.get("STUFE1_OHNE_DEFEKT_C") == "1"
              else _LEERER_KONNEKTOR_RE.search(cell))
    if m_leer:
        rumpf = cell[:m_leer.start()].rstrip(" .,;")
        if rumpf != cell and NORM_MARKER_RE.search(rumpf) and _gesetz_ende(rumpf):
            return split_normen(rumpf)

    # ';' trennt im Bestand zuverlaessig VERSCHIEDENE Zitate. Ohne diese
    # Trennung zieht _umstellen() das vorn stehende Gesetz ans Ende des GANZEN
    # Strings und die Paragraphen bekommen das FALSCHE Gesetz: aus
    # 'AktG §§91,93; GmbHG §43' wurde '§ 91 GmbHG, § 93 GmbHG, § 43 AktG'
    # (Befund 28.08.2026). Nur greifen, wenn MEHRERE Teile ein eigenes
    # Normmuster tragen - sonst zerfaellt '§ 26 BDSG; einschlaegiges
    # Landesdatenschutzrecht' in ein Zitat und ein Bruchstueck.
    if ";" in cell:
        teile = [t.strip() for t in cell.split(";") if t.strip()]
        # Ein Teil, der eine Gerichtsentscheidung zitiert ('BAG 9 AZR 78/14'), ist keine
        # Norm und wird hier verworfen (seit 25.09.2026) - die Rechtsprechung liest
        # rspr_aus() aus derselben Zelle. Vorher zog '; BAG 9 AZR 78/14' das Gesetz vom
        # ersten Zitat weg: '§§ 26, 17 BBiG; BAG …' -> '§ 26' ohne Gesetz (INTERNSHIP).
        if os.environ.get("STUFE1_OHNE_BRUCHSTUECK") != "1":
            ohne_rspr = [t for t in teile
                         if not (alle_az(t) and GERICHT_RE.search(t) and not NORM_MARKER_RE.search(t))]
            if ohne_rspr and len(ohne_rspr) < len(teile):
                return split_normen("; ".join(ohne_rspr))
        mit_marker = [t for t in teile if NORM_MARKER_RE.search(t)]
        # EIN Zitat MIT EIGENEM GESETZ plus nackte Gesetzesnennung (seit
        # 2026-09-24): '§§ 10, 12, 13 AO; jeweiliges DBA' wurde nicht getrennt,
        # weil nur EIN Teil ein Normmuster traegt - das Gesetz kam dann vom
        # ZELLENENDE, und daraus wurden '§ 10 DBA', '§ 12 DBA' (Crawl-Ziele fuer
        # Paragraphen, die es nicht gibt; gemeint und im CMS: §§ 10, 12 AO).
        # Getrennt wird nur, wenn (a) der Zitatteil VORN steht und sein Gesetz
        # SELBST traegt - dann kann ihm der Rest nichts mehr wegnehmen - und
        # (b) die uebrigen Teile KEINE ZIFFER tragen, also reine Gesetzes-
        # nennungen sind. (b) ist aus dem A/B vom 24.09.2026: ohne ihn wurde aus
        # '§ 165 SGB IX; BAG 8 AZR 143/23' die "Norm" 'BAG 8 AZR 143/23 BAG'
        # und § 165 SGB IX fiel weg. Die Namen des Rests werden EINZELN
        # eingetragen ('MiLoG, TVG/AEntG' -> MiLoG, TVG, AEntG), nicht
        # verworfen - eine blosse Gesetzesnennung ist ein gueltiger Befund
        # (siehe den 'einschlaegiges DBA'-Zweig unten).
        # Abschaltbar mit STUFE1_OHNE_SEMIKOLON_EINZEL=1 (A/B).
        if (os.environ.get("STUFE1_OHNE_SEMIKOLON_EINZEL") != "1"
                and len(mit_marker) == 1 and len(teile) > 1
                and NORM_MARKER_RE.search(teile[0]) is not None
                # Gesetz am Ende ODER vorangestellt ('AStG § 1 Abs. 5; BsGaV', seit
                # 25.09.2026 - vorher '§ 1 Abs. 5; BsGaV AStG')
                and (_gesetz_ende(_ohne_prosa_schwanz(teile[0]))
                     or (GESETZ_VORN_RE.match(teile[0])
                         and os.environ.get("STUFE1_OHNE_BRUCHSTUECK") != "1"))
                and not any(re.search(r"\d", t) for t in teile[1:])):
            kopf_out = split_normen(teile[0])
            if kopf_out:
                out = list(kopf_out)
                for rest in teile[1:]:
                    # Auch an Bindewoertern trennen: 'BPersVG und Landes-
                    # personalvertretungsrecht' -> BPersVG (der Rest ist eine
                    # Kategorie und faellt an NUR_GESETZ_RE durch).
                    for stueck in re.split(r"[,/]|\s+(?:und|sowie|oder|bzw\.?)\s+", rest):
                        name = _gesetz_ende(stueck.strip())
                        # Gesetz VORN, Zusatz dahinter ('GewStG je nach Sachverhalt', seit
                        # 25.09.2026): sonst fiel GewStG aus 'KStG/GewStG je nach Sachverhalt'.
                        # Abschaltbar mit STUFE1_OHNE_GESETZ_VORN_REST=1.
                        if not name and os.environ.get("STUFE1_OHNE_GESETZ_VORN_REST") != "1":
                            erstes = (stueck.strip().split() or [""])[0]
                            if NUR_GESETZ_RE.match(erstes) and _gesetz_ende(erstes) == erstes:
                                name = erstes
                        if name and NUR_GESETZ_RE.match(name) and name not in out:
                            out.append(name)
                return out
        if len(mit_marker) > 1:
            out = []
            for teil, gesetz in zip(teile, _gesetze_je_teil(teile)):
                # `_ohne_prosa_schwanz` verhindert die DOPPELUNG: traegt der Teil
                # sein Gesetz schon, nur verdeckt von einem Zusatz, wuerde es hier
                # ein zweites Mal angehaengt ('… SGB IX, soweit einschlaegig SGB IX').
                # Traegt der Teil sein Tarifwerk selbst ('§ 1 TVA-L Pflege'), ist `gesetz`
                # genau dieser Name - anhaengen wuerde ihn verdoppeln (seit 25.09.2026).
                if gesetz and not _gesetz_ende(_ohne_prosa_schwanz(teil)) \
                        and not GESETZ_VORN_RE.match(teil) \
                        and not norm_txt(teil).endswith(gesetz):
                    teil = "%s %s" % (teil, gesetz)   # 'ArbSchG …; §§ 15, 16'
                teil_out = split_normen(teil)
                if not teil_out and is_norm_cell(teil):
                    teil_out = [teil]                 # z.B. 'Anhang IV KI-VO'
                if not teil_out:
                    # Nennt der Teil ein Gesetz OHNE Paragraph
                    # ('einschlaegiges DBA'), ist das ein gueltiger - und
                    # sperrender - Befund, kein Nichts. Ohne diesen Zweig fiel
                    # 'DBA' aus '§ 12 AO; §§ 1, 38 EStG; einschlaegiges DBA'
                    # ganz weg (belegt 28.08.2026). Der bloße Gesetzesname wird
                    # eingetragen, nicht der Fliesstext drumherum.
                    nur_gesetz = _gesetz_ende(teil)
                    if nur_gesetz:
                        teil_out = [nur_gesetz]
                out.extend(teil_out)
            if out:
                return out

    # Mehrere VORANGESTELLTE Gesetze in einem Teil (Defekt A, siehe
    # GESETZ_VORN_TRENN_RE): an jeder Trennstelle schneiden, damit _umstellen()
    # je Stueck genau ein Gesetz sieht. Nur ab ZWEI Fundstellen - sonst bleibt
    # das bisherige Verhalten unveraendert.
    starts = [m.start() for m in GESETZ_VORN_TRENN_RE.finditer(cell)]
    if len(starts) > 1:
        grenzen = starts + [len(cell)]
        out = []
        if starts[0] > 0:                      # Vorspann vor dem ersten Gesetz
            grenzen = [0] + grenzen
        for a, b in zip(grenzen, grenzen[1:]):
            # Das TRENNZEICHEN bleibt am Ende des vorigen Stuecks stehen und
            # wandert sonst mit _umstellen() mitten ins Zitat: aus
            # 'BDSG § 26 / DSGVO Art. 5' wurde '§ 26 / BDSG', aus
            # 'AO § 147 und HGB § 257' wurde '§ 147 und AO'.
            stueck = re.sub(r"\s+(?:und|sowie|bzw\.?|oder)\s*$", "",
                            cell[a:b].strip(" ,;/"), flags=re.IGNORECASE).strip(" ,;/")
            if stueck:
                out.extend(split_normen(stueck))
        if out:
            return out

    # ' / ' trennen; Teile ohne Zitat (z.B. '§ 77 BetrVG / bestehende
    # Betriebsvereinbarungen') fallen weg, mehrere Zitate werden getrennt
    teile = [cell]
    if " / " in cell:
        kand = [t.strip() for t in cell.split(" / ")]
        mit_zitat = [k for k in kand if NORM_MARKER_RE.search(k)]
        if mit_zitat and len(mit_zitat) < len(kand):
            teile = mit_zitat
        elif len(mit_zitat) > 1:
            teile = mit_zitat

    out = []
    for teil in teile:
        teil = _umstellen(teil)
        # `(?<!§)` ist entscheidend: ohne die Sperre trennt der Lookahead
        # AUCH zwischen den beiden Zeichen von '§§' (denn '§ 241' passt auf
        # '§§?\s*\d'), und aus '§§ 241 Abs. 2, 305 bis 310 BGB' wird
        # ['§', '§ 241 Abs. 2, 305 bis 310 BGB']. Das einsame '§' fiel bisher nur
        # zufaellig durch den Endfilter - der Doppelparagraph war damit aber
        # verloren, bevor unten `mehrere` ihn lesen konnte, und genau daran
        # scheiterte die Aufloesung solcher Sammelzitate (Befund 28.08.2026).
        segs = [s.strip(" ,;") for s in re.split(
            r"(?<!§)(?=§§?\s*\d|\bArt\.?\s*\d|\bArtikel\s*\d)", teil) if s.strip(" ,;")]
        # UNGEKUERZT merken: `_ohne_prosa_schwanz` schneidet gleich ', insbesondere'
        # weg, und damit waere das Bindeglied verschwunden, bevor es gelesen ist.
        # Genau daran blieben '§ 307' und '§ 309 Nr. 13' aus
        # '§§ 305 bis 310 BGB, insbesondere § 307 und § 309 Nr. 13' gesetzlos.
        segs_roh = list(segs)
        # Nachlaufenden Zusatz je Segment abschneiden (Defekt B). Ohne das steht
        # das Gesetz nicht mehr am Ende: `_gesetz_ende` liefert '', das Segment
        # erbt ein fremdes Gesetz, UND `_zerlege_paragraphen` bricht am Stueck
        # 'soweit einschlaegig' ab (kein Paragraphenanfang) - die Paragraphen
        # blieben als ein einziger Klumpen stehen.
        #
        # NUR Segmente MIT Zitat: ein zitatloses Segment ('FVerlV, insbesondere')
        # wuerde sonst auf den nackten Gesetzesnamen eingedampft und damit zu
        # einem NEUEN, sperrenden Befund - genau der Fall, der oben bewusst
        # offen gelassen ist (Bindeglied-Fix, verworfen 28.08.2026).
        segs = [_ohne_prosa_schwanz(s) if re.match(r"^(?:§§?|Art\.?|Artikel)\s*\d", s)
                else s for s in segs]
        if not segs:
            out.append(teil)
            continue
        gesetze = [_gesetz_ende(s) for s in segs]

        # BINDEGLIED-FIX (Nutzerregel 29.08.2026, Begruendung oben bei
        # `_vorspann_gesetz`): traegt das ERSTE Segment kein Zitat, sondern den
        # Gesetzesnamen samt Bindeglied ('AGG, insbesondere'), erben die
        # nachfolgenden Zitate dieses Gesetz - aber NUR, wenn sie selbst keines
        # tragen. Das laeuft VOR der Rueckwaerts-Uebernahme, weil die
        # Nutzerregel Vorrang hat: in 'AGG, insbesondere §§ 1, 3, 7 und § 12
        # ArbSchG' gehoeren §§ 1, 3, 7 zum AGG, nicht zum spaeter genannten
        # ArbSchG - rueckwaerts wuerden sie sonst ArbSchG erben.
        # Der Bindeglied-Traeger ist NICHT immer das erste Segment. In
        # '§§ 305 bis 310 BGB, insbesondere § 307 und § 309 Nr. 13' traegt ihn
        # das erste ZITAT-Segment; eine Beschraenkung auf segs[0] ohne Zitat
        # liess '§ 307' und '§ 309 Nr. 13' gesetzlos stehen (gemessen 29.08.2026).
        # Deshalb: jedes Segment, das auf ein Bindeglied endet, setzt das Gesetz
        # fuer die FOLGENDEN Segmente - und wieder nur dort, wo keines steht.
        bg = ""
        # Bauform "Gesetz vorn ohne Bindeglied" ('Abgabenordnung §§ 12, 13',
        # 'RL 2003/88/EG, Art. 3, 5'): das erste, zitatlose Segment IST das
        # Gesetz. Waechter stecken in `_vorn_gesetz`.
        # NUR wenn KEIN Zitat der Zelle ein eigenes Gesetz traegt. Das ist der
        # Waechter, ohne den dieser Zweig denselben Fehler macht wie der
        # verworfene Fix vom 28.08.2026: in 'AGG, § 612a und § 242 BGB als
        # Gegenpruefung' meint das fuehrende AGG sich selbst, waehrend § 612a
        # zum BGB dahinter gehoert - mit Uebernahme wurde daraus '§ 612a AGG'.
        # Traegt dagegen KEIN Zitat ein Gesetz ('Abgabenordnung §§ 12, 13'),
        # kann das fuehrende Wort nur ihr gemeinsames Gesetz sein.
        # Gezaehlt werden nur ZITAT-Segmente. Das Vorspann-Segment traegt sein
        # Gesetz naturgemaess selbst ('RL 2003/88/EG, Art. 3, 5, 6') - wer es
        # mitzaehlt, blockiert genau die Faelle, die der Zweig loesen soll.
        _ist_zitat = lambda s: bool(re.match(r"^(?:§§?|Art\.?|Artikel)\s*\d", s))
        zitat_hat_gesetz = any(
            g for g, s in zip(gesetze, segs_roh) if _ist_zitat(s))
        if segs_roh and not zitat_hat_gesetz and not _ist_zitat(segs_roh[0]):
            bg = _vorn_gesetz(segs_roh[0])
        for i, s in enumerate(segs_roh):
            if _BINDEGLIED_RE.search(s):
                neu = _vorspann_gesetz(s)
                if neu:
                    bg = neu
                continue
            if bg and not gesetze[i]:
                gesetze[i] = bg
        # Eine blosse roemische Zahl ist nie ein Gesetz (s. _ROEMISCH_RE).
        gesetze = [g if _taugt_als_gesetz(g) else "" for g in gesetze]

        letztes = ""
        for i in range(len(segs) - 1, -1, -1):
            if gesetze[i]:
                letztes = gesetze[i]
            gesetze[i] = gesetze[i] or letztes
        for seg, ges in zip(segs, gesetze):
            m = re.match(r"^(§§?|Art\.?|Artikel)\s*(.+)$", seg)
            if m:
                marker = "§" if m.group(1).startswith("§") else "Art."
                # `§§` heisst: MEHRERE Paragraphen (Nutzervorgabe 28.08.2026).
                # Nur hier ist der Doppelparagraph noch sichtbar.
                mehrere = m.group(1) == "§§"
                rest = m.group(2)
                if ges:
                    rest = re.sub(re.escape(ges) + r"\s*$", "", rest).strip(" ,;")
                einzeln = _zerlege_paragraphen(rest, ges, marker, mehrere=mehrere)
                if einzeln:
                    out.extend(einzeln)
                    continue
            if not re.search(r"\d", seg) and not NUR_GESETZ_RE.match(seg):
                continue                      # '§ BetrVG' o.ä. verwerfen
            # Das aus dem Segmentkontext geerbte Gesetz gehoert an JEDES Segment,
            # nicht nur an §-Segmente. Bis 2026-08-28 stand hier
            # `seg.startswith("§")` - dadurch ergab
            # 'Art. 3, Art. 4, Art. 19 Abs. 4, Art. 33 Abs. 2 GG' die Eintraege
            # 'Art. 3' und 'Art. 4' OHNE Gesetz. Ein Artikel ohne Gesetz ist in
            # Stufe 2 nicht pruefbar ("kein §/Art.-Muster erkannt") und war damit
            # weder Befund noch Bedarf - genau die Klasse, die weder Geist noch
            # ignoriert sein darf (Nutzervorgabe 28.08.2026).
            if ges and ges.lower() not in seg.lower() \
                    and re.match(r"^(?:§§?|Art\.?|Artikel)\s*\d", seg):
                seg = seg + " " + ges
            out.append(seg)
    return [x for x in (norm_txt(o) for o in out) if x and x not in ("§", "§§")]


# ===================================================== Abschnitte / Dokumentfluss

NUMMER_RE = re.compile(r"^\s*(?:\d+(?:\.\d+)*\.?|[A-Z]\.|[IVXL]+\.)\s*")
NORM_ABSCHNITT = (
    "primäre normen", "primaere normen", "normen", "ergänzende und parallele normen",
    "normenlandkarte", "primäre rechtsgrundlagen", "normenkatalog",
    "primäre normen und rechtsquellen", "ergänzende normen",
)
RSPR_ABSCHNITT_WORT = ("rechtsprechung", "leitentscheidung", "entscheidungen")
BEIDE_ABSCHNITT = ("rechtsgrundlagen", "quellenhierarchie", "primärquellen")
# Amtliche LINKLISTEN sind Quellendeklarationen - und zwar die einzigen, die die
# amtliche URL mitfuehren ('§ 15 BBiG - ... - https://www.gesetze-im-internet.de/
# bbig_2005/__15.html', 'BAG 09.06.2016 - 6 AZR 396/15 - https://www.bundes
# arbeitsgericht.de/entscheidung/6-azr-396-15/'). Bis 31.08.2026 war KEINER
# dieser Abschnittsnamen erfasst: 469 Absaetze in rund 40 Skills lagen ausserhalb
# jedes Abschnitts, und weil sie Normen UND Entscheidungen mischen, fehlten sie
# in beiden Spalten. Bewusst eine WEISSE LISTE mit Praefixvergleich, kein
# Teilzeichenketten-Kriterium auf 'amtlich': Ueberschriften wie 'Schritt 2 -
# amtlichen Quellenraum festlegen' oder 'AGG - amtliche Gesetzesfassung' (eine
# Quellenzeile, keine Ueberschrift) wuerden sonst mitkommen.
QUELLEN_LINK_ABSCHNITT = (
    "amtliche quellenlinks", "amtliche arbeitslinks", "amtliche links",
    "amtliche normlinks", "amtliche normenlinks", "amtliche normquellen",
    "amtliche entscheidungslinks", "amtliche entscheidungsquellen",
    "amtliche online-quellen", "amtliche onlinequellen",
    "amtlich verifizierte quellenlinks", "verifizierte fundstellen",
)
AUSSCHLUSS = (
    "kontextprüfung", "vollständigkeitscheck", "anker", "fehlervermeidung",
    "korrektur", "prüfbaum", "subskill", "nachbar-skill", "fachreview",
    "quellenbasis für die erstellung", "testfall", "testfälle", "decision-tree",
    "spezial-output", "rechtsstandsmanagement", "referenzdatei",
    "einordnung im jus", "freigabeprotokoll", "pflicht-eingabe",
    "rückfragenlogik", "negativprüfung", "gegenprüfung", "effect-register",
    "adversarius", "eskalation", "sichtbarkeitslogik", "feldkatalog",
    # Fund 26.08.2026 (Phase 3): 'Schritt 2 - Amtliche Quellenhierarchie und
    # Suchplan bestimmen' der RESEARCH-Skills traf ueber 'quellenhierarchie' in
    # BEIDE_ABSCHNITT an. Die Tabelle dort ist ein SUCHPLAN
    # ('Prioritaet | Amtliche Quelle / Register | Suchziel'), keine
    # Quellendeklaration - sie hat Behoerden ('Bundesrat', 'BMAS und BMG',
    # 'G-BA') als Normen bzw. als 'Normen ohne Zitat' in den Index gebracht.
    # Gemessen ueber 305 Dateien: -6 falsche 'ohne Zitat', -2 falsche Normen
    # (G-BA, GKV-Spitzenverband), 0 echte Normzitate verloren - die 11 echten
    # Normen der betroffenen Datei stehen unter 'Primaere Normen'.
    "suchplan",
)
NORM_HEADER_WORDS = ("norm", "vorschrift", "rechtsquelle", "regelwerk", "normgruppe")
RSPR_HEADER_WORDS = ("aktenzeichen", "entscheidung", "rechtsprechung", "urteil",
                     "fundstelle", "gericht")
QUELLEN_HEADER_WORDS = ("quelle", "quellen", "verwaltungsauffassung", "quellenart",
                        "quellenrang", "rangfolge")


def clean_heading(t):
    return NUMMER_RE.sub("", norm_txt(t)).strip().lower()


def absatz_stil(par):
    return ((par.style.name or "") if par.style is not None else "").strip()


def ueberschrift_art(par):
    """Wie wurde die Ueberschrift erkannt: 'stil' (Word-Ueberschriftenformat),
    'heuristik' (kurzer Absatz ohne Schlusspunkt) oder None (keine Ueberschrift).

    Die Unterscheidung ist noetig, weil nur die HEURISTIK irren kann - und sie
    irrt an einer belegten Stelle: in einer amtlichen Linkliste ist
    'AGG – Gesetze im Internet' (Stil 'Source Text') eine QUELLENZEILE, sieht
    aber wie eine Ueberschrift aus. Sie setzte den Abschnitt zurueck, und die
    zehn Entscheidungszeilen darunter fielen aus dem Index
    (ANALYSIS.RECRUITING.AGG_RISK, Fund 31.08.2026). Siehe `process()`.
    """
    style = absatz_stil(par).lower()
    t = norm_txt(par.text)
    if not t:
        return None
    if "head" in style or "überschrift" in style or "title" in style:
        return "stil"
    if len(t) > 90 or t.endswith((".", ":", ";", "?")):
        return None
    if alle_az(t) or NORM_MARKER_RE.search(t) or DATUM_RE.search(t):
        return None
    return "heuristik"


def ist_ueberschrift(par):
    """Kurze Absaetze ohne Schlusspunkt gelten als Ueberschrift - AUSSER sie
    tragen selbst eine Fundstelle.

    Grund (Fund 26.08.2026): in den Familien NONCOMPETE und
    RESEARCH.INTERNATIONAL steht jede Fundstelle als eigener kurzer Absatz
    ('BAG, Urteil vom 22.03.2017 – 10 AZR 448/15 – (...)',
    'Entscheidung: EuGH, 15.03.2011 – C-29/10, Koelzsch, ECLI:...').
    Ohne diese Ausnahme wird jede solche Zeile zur Ueberschrift, der
    Abschnittsbezug ('Tragende Rechtsprechung') geht verloren und die Datei
    faellt faelschlich als 'keine Rechtsprechung' aus.
    """
    return ueberschrift_art(par) is not None


def body_items(doc):
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            yield "p", Paragraph(child, doc)
        elif tag == "tbl":
            yield "t", Table(child, doc)


def abschnitt_art(heading):
    h = clean_heading(heading)
    if not h:
        return set(), False
    if any(a in h for a in AUSSCHLUSS):
        return set(), True
    arten = set()
    if any(h.startswith(n) or h == n for n in NORM_ABSCHNITT):
        arten.add("norm")
    if any(w in h for w in RSPR_ABSCHNITT_WORT):
        arten.add("rspr")
    if any(b in h for b in BEIDE_ABSCHNITT):
        arten.update({"norm", "rspr"})
    if ist_quellen_linkliste(heading):
        arten.update({"norm", "rspr"})
    return arten, False


def ist_quellen_linkliste(heading):
    """Amtliche Linkliste? Entscheidet zusaetzlich, dass dort Zeilen OHNE Zitat
    keine unbestimmte Norm werden (siehe `normen_aus(nur_zitate=...)`)."""
    h = clean_heading(heading)
    return bool(h) and any(h.startswith(q) for q in QUELLEN_LINK_ABSCHNITT)


# Woerter, die eine Ueberschrift verdaechtig machen, ein Rechtsgrundlagen-
# Abschnitt zu SEIN, ohne dass eine der Listen oben sie erfasst. Trifft eines
# zu und die Ueberschrift ist weder Abschnitt noch Ausschluss, ist das ein
# Hinweis auf eine neue Bauform - also ein Erkennerausfall, den der Lauf
# meldet, statt die Datei stillschweigend leer zu lassen.
VERDACHT_WORTE = ("norm", "rechtsprechung", "rechtsgrundlage", "rechtsquelle",
                  "fundstelle")

# Eine Warnung, die bei zwei Dritteln aller Dateien anschlaegt, wird nicht
# gelesen. Gemessen am 27.08.2026 ueber 291 Dateien: mit den generischen
# Woertern 'quelle'/'entscheidung'/'urteil'/'vorschrift' und ohne die Filter
# unten schlug diese Warnung 194x an - fast ausschliesslich auf
# Arbeitsschritt-Ueberschriften wie 'Schritt 5 - Entscheidungswirkung und
# menschliche Kontrolle pruefen' oder 'Schritt 12 - Quellen und Pruefstatus
# ausweisen'. Das sind Anweisungen an den Agenten, keine Quellendeklarationen.
# Drei Filter trennen sie: Arbeitsschritt-Praefix, Taetigkeitsverb am Ende,
# Ueberschriftenlaenge.
SCHRITT_PRAEFIX_RE = re.compile(r"^(schritt|step|phase|teil|modul|abschnitt)\b")
TAETIGKEITSVERB = ("prüfen", "pruefen", "ausgeben", "ableiten", "benennen",
                   "bestimmen", "ausweisen", "erstellen", "bilden", "festlegen",
                   "dokumentieren", "einordnen", "klassifizieren", "aufnehmen",
                   "durchführen", "durchfuehren", "vorbereiten", "beantworten")
UEBERSCHRIFT_MAXWORTE = 6


def ueberschrift_verdaechtig(heading):
    """True, wenn die Ueberschrift wie ein Rechtsgrundlagen-Abschnitt aussieht,
    aber von abschnitt_art() weder als Abschnitt noch als Ausschluss erkannt
    wurde - also Verdacht auf eine neue Bauform.

    Bewusst getrennt von abschnitt_art(): die Bewertung aendert NICHTS am
    Extraktionsergebnis, sie erzeugt nur eine Warnung. Deshalb darf sie streng
    sein - eine verpasste neue Bauform faellt ueber die andere Warnung
    'kein_abschnitt_erkannt_trotz_kandidaten' ohnehin auf, und DIE haengt nicht
    an einer Wortliste."""
    arten, ausgeschlossen = abschnitt_art(heading)
    if arten or ausgeschlossen:
        return False
    h = clean_heading(heading)
    if not h or not any(w in h for w in VERDACHT_WORTE):
        return False
    if SCHRITT_PRAEFIX_RE.match(h):
        return False
    if any(h.rstrip(" .:").endswith(v) for v in TAETIGKEITSVERB):
        return False
    return len(h.split()) <= UEBERSCHRIFT_MAXWORTE


# ==================================================================== Tabellen


def table_matrix(t):
    return [[norm_txt(c.text) for c in r.cells] for r in t.rows]


def header_of(rows):
    return [c.lower() for c in rows[0]] if rows else []


def header_has(header, words):
    return any(any(w in h for w in words) for h in header)


def pick_column(rows, predicate, start_row=1):
    if not rows:
        return -1, 0
    ncols = max(len(r) for r in rows)
    best, bestn = -1, 0
    for c in range(ncols):
        n = sum(1 for r in rows[start_row:] if c < len(r) and predicate(r[c]))
        if n > bestn:
            best, bestn = c, n
    return best, bestn


def laengste_uebrige(row, ausgenommen, minlen=15, maxlen=400):
    rest = [c for i, c in enumerate(row) if i not in ausgenommen and c and len(c) > minlen]
    return max(rest, key=len)[:maxlen] if rest else ""


KOPFWORT = ("norm", "normen", "norm / quelle", "norm / regelwerk",
            "norm / rechtsquelle", "vorschrift", "rechtsquelle", "regelwerk",
            "normgruppe", "norm / normgruppe", "primärer anker", "anker")


# ---------------------------------------------------------------------------
# ABGRENZUNGSQUELLEN (Nutzerregel 11.09.2026)
#
# "eine quelle die im skill als 'gehoert NICHT zu' steht, muss nicht im CMS oder
#  im Index liegen und der Agent braucht nicht abfragen, was da drinnen steht."
#
# Erkannt wird das ausschliesslich an der FUNKTIONS- bzw. PFLICHTPRUEFUNGSSPALTE
# der Normtabelle, und dort nur an Formulierungen, die die materielle Pruefung
# ausdruecklich WOANDERS verorten. Die Spalten werden ueber die KOPFZEILE
# bestimmt - nicht ueber `laengste_uebrige()`, sonst entscheidet die Textlaenge,
# welche Spalte gelesen wird (die Rechtsstandsnotiz enthaelt 'nicht verifiziert'
# und haette sonst mitgezaehlt).
#
# 🔴 WARUM EINE PHRASENLISTE UND KEINE HEURISTIK - am Bestand gemessen 11.09.2026:
# ein blosses "Routing oder Negation im Text" trifft 590 Eintraege, davon 552
# PARAGRAPHIERTE Pflichtnormen. Die Zeile
#   "§ 102 BetrVG | Betriebsratsanhoerung vor jeder Kuendigung; bei BR-Bezug
#    Spezialskill aktivieren."
# ist wortgleich gebaut wie eine echte Abgrenzung - die Norm ist aber tragend.
# Der Unterschied liegt nicht im Vorkommen des Routings, sondern darin, ob die
# Zelle sagt, dass hier NUR geroutet und NICHT materiell geprueft wird.
# Deshalb: enumerierte Formulierungen, keine Verallgemeinerung. Wer die Liste
# erweitert, misst vorher den Kollateralschaden an den paragraphierten Normen.
ABGRENZUNG_PHRASEN = (
    # "Nur ... Routing/Router ..." - der Skill leitet weiter, statt zu pruefen
    (r"\bnur\b[^.;]{0,60}\brout(er|ing|en)\b", "nur Routing/Router"),
    (r"\bnur\b[^.;]{0,40}\brouter\b", "nur Router"),
    # ausdruecklicher Ausschluss der materiellen Pruefung
    (r"\bkeine?\s+(materielle[nrs]?|inhaltliche[nrs]?|vertiefte[nrs]?|"
     r"detail\w*|einzel\w*|end\w*|voll\w*)\s*"
     r"(pr(ü|ue)fung|bewertung|aussage|entscheidung|berechnung|freigabe)?", "keine materielle Prüfung"),
    # die Pruefung wird benannt anderswo verortet
    (r"(materielle|detail|end|voll)\w*\s*pr(ü|ue)fung\s+(in|an|im|durch)\s", "Prüfung liegt anderswo"),
    (r"\bberechnung\s+delegieren\b", "Berechnung delegiert"),
    # ausdrueckliche Nichtzustaendigkeit
    (r"\bnicht\s+gegenstand\b", "nicht Gegenstand"),
    # "nicht abschliessend" ALLEIN reicht nicht: "Klauseltypen erkennen, nicht
    # abschliessend bewerten." meint die PRUEFTIEFE, nicht die Unzustaendigkeit -
    # die Norm wird dort sehr wohl angewandt (Selbsttestfall ABGR-NEIN).
    # Es zaehlt nur zusammen mit einer Weitergabe.
    (r"\bnicht\s+abschlie(ß|ss)end\b[^.;]{0,50}"
     r"(\brout|\bdelegier|Spezialskill|Folgeskill|\bin\s+[A-Z][A-Z_]{3,})",
     "nicht abschließend"),
)
ABGRENZUNG_RE = [(re.compile(m, re.I), name) for m, name in ABGRENZUNG_PHRASEN]


def ist_abgrenzung(text):
    """Sagt die Zelle, dass hier NUR geroutet und nicht materiell geprueft wird?
    Rueckgabe: der Name der treffenden Formulierung oder ''.

    `STUFE1_OHNE_ABGRENZUNG=1` schaltet die Regel ab. Das ist kein Notausgang,
    sondern der A/B-Schalter: nur mit ihm laesst sich derselbe Codestand einmal
    mit und einmal ohne die Regel fahren und die zwei Ergebnisse gegeneinander
    halten. Ein Vergleich gegen eine aeltere xlsx taugt dafuer NICHT - dort
    stecken Aenderungen anderer Laeufe mit drin (gemessen 11.09.2026: die
    Rechtsprechungs-Platzhalter hiessen im `_260910_`-Stand noch
    '[negativbefund]', heute '[offen]' - nichts davon gehoert zu dieser Regel)."""
    if os.environ.get("STUFE1_OHNE_ABGRENZUNG"):
        return ""
    t = norm_txt(text or "")
    if not t:
        return ""
    for rx, name in ABGRENZUNG_RE:
        if rx.search(t):
            return name
    return ""


def abgrenzungsspalten(rows, ausser):
    """Indizes der Funktions-/Pflichtpruefungsspalten, aus der KOPFZEILE."""
    header = [norm_txt(h).lower() for h in header_of(rows)]
    return [i for i, h in enumerate(header)
            if ("funktion" in h or "pflichtpr" in h or "pflichtpruef" in h)
            and i not in ausser]


_BRUCH_RE = re.compile(r"^(?:§§?|Art\.?|Artikel)\s*\d+[a-z]?\b"
                       r"(?:\s*(?:Abs\.|Nr\.|Satz|S\.|lit\.|und|bis|ff\.|f\.|,)\s*[\da-z]*\.?)*\s*$")
_VORSPANN_RE = re.compile(r"^(.*?)[\s,:;]*(?:insbesondere|insb\.|namentlich|dort)?[\s,:]*$", re.IGNORECASE)


def _bruchstuecke_ergaenzen(einzelne, zelle, zeile, kopf):
    """Paragraph ohne Gesetz ('§ 14', 'Art. 5') bekommt sein Gesetz, wenn es EINDEUTIG
    in derselben Tabellenzeile steht (seit 25.09.2026) - sonst bleibt er, wie er ist.

    Drei Bauformen, alle aus der Bibliothek:
      1. Gesetz VOR dem ersten Zitat, abgesetzt: 'SGB IV, insbesondere §§ 7, 7a und 28e',
         'DSGVO, insbesondere Art. 5, 6, 9', 'Rom-I-Verordnung, Art. 8 und 9'
         (dann faellt der mitgelieferte nackte Name 'Rom-I-Verordnung' weg).
      2. Gesetz unmittelbar vor DIESEM Zitat: '§ 90 Abs. 3 AO und GAufzV § 3' - dann
         wird das Anhaengsel 'und GAufzV' am vorigen Zitat abgeschnitten.
      3. Gesetz in der Nachbarspalte 'Quelle'/'Gesetz': FIXEDTERM.CHECK 'TzBfG | §§ 14-17'.
    Ein Vorspann, der kein reiner Gesetzesname ist ('Wiener Uebereinkommen ...'),
    zaehlt nicht. Abschaltbar mit STUFE1_OHNE_BRUCHSTUECK=1 (A/B)."""
    if os.environ.get("STUFE1_OHNE_BRUCHSTUECK") == "1":
        return einzelne
    bruch = [e for e in einzelne if _BRUCH_RE.match(e) and not _gesetz_ende(e)]
    if not bruch:
        return einzelne
    z = norm_txt(zelle)
    m0 = NORM_MARKER_RE.search(z)
    vorspann = _VORSPANN_RE.match(z[:m0.start()]).group(1).strip(" ,;:") if m0 else ""
    gesetz_vorn = vorspann if vorspann and NUR_GESETZ_RE.match(vorspann) else ""
    gesetz_spalte = ""
    for i, h in enumerate(kopf):
        if i < len(zeile) and re.search(r"\b(quelle|gesetz)\b", h or "", re.IGNORECASE):
            kand = norm_txt(zeile[i]).strip(" ,;:")
            if kand and kand != z and NUR_GESETZ_RE.match(kand):
                gesetz_spalte = kand
                break
    out = []
    for e in einzelne:
        if e not in bruch:
            out.append(e)
            continue
        m_direkt = re.search(r"(?:^|[\s,;(])([^\s,;()]+)\s+" + re.escape(e.split()[0])
                             + r"\s*" + re.escape(" ".join(e.split()[1:2])) + r"(?![\w])", z)
        direkt = m_direkt.group(1) if m_direkt and NUR_GESETZ_RE.match(m_direkt.group(1)) \
            and _gesetz_ende(m_direkt.group(1)) else ""
        gesetz = direkt or gesetz_vorn or gesetz_spalte
        if not gesetz:
            out.append(e)
            continue
        if direkt and out and re.search(r"\s+(?:und|sowie|bzw\.?|oder)\s+%s$" % re.escape(direkt), out[-1]):
            out[-1] = re.sub(r"\s+(?:und|sowie|bzw\.?|oder)\s+%s$" % re.escape(direkt), "", out[-1])
        out.append("%s %s" % (e, gesetz))
    if gesetz_vorn:
        out = [e for e in out if e != gesetz_vorn]
    return out


def normen_aus(rows, unzerlegbar=None, nur_zitate=False):
    """<unzerlegbar> ist ein optionaler Collector: Zellen, die als Normzelle
    erkannt wurden, aus denen split_normen() aber KEINEN Eintrag gewinnt,
    landen dort. Ohne diesen Collector verschwinden sie lautlos - genau die
    Zellen sind aber der Hinweis auf eine unbekannte Zitierweise.

    <nur_zitate> schaltet die 'unbestimmt'-Rueckfallebene ab. Gesetzt wird das
    fuer die amtlichen LINKLISTEN (31.08.2026): dort ist eine Zeile ohne Zitat
    ein Portal- oder Institutionsverweis ('BBiG - Gesamtausgabe - <URL>',
    'Bundesarbeitsgericht', 'BIBB HA 120'), keine deklarierte Fundstelle - die
    konkreten Paragraphen stehen unter 'Primaere Normen'. Ohne den Schalter
    stand 'Bundesarbeitsgericht' als unbestimmte NORM im Index
    (ANALYSIS.APPRENTICESHIP.EXAM_FAILURE_EXTENSION) und haette als
    Fundstellenlucke gesperrt."""
    n_col, n_n = pick_column(rows, is_norm_cell)
    if n_col < 0 or n_n == 0:
        return []
    abg_cols = abgrenzungsspalten(rows, {n_col})
    out = []
    for r in rows[1:]:
        zelle = r[n_col] if n_col < len(r) else ""
        if not is_norm_cell(zelle):
            # Zelle steht in der Normspalte, nennt aber KEIN Zitat, sondern ein
            # Rechtsgebiet/eine Normgruppe ('EStG und Beitragsrecht',
            # 'Themenabhängige Primärnorm(en)', '[Themenabhängige Hauptnorm(en)]').
            # Das ist eine deklarierte Quelle ohne Fundstelle - also der Befund,
            # nicht Rauschen. Wird als 'unbestimmt' gefuehrt statt verworfen.
            z = norm_txt(zelle)
            if nur_zitate:
                continue
            if z and len(z) <= NORM_MAXLEN and z.lower() not in KOPFWORT \
                    and not alle_az(z):   # Aktenzeichen ist keine Norm
                out.append({"norm": z, "funktion": laengste_uebrige(r, {n_col}),
                            "unbestimmt": True})
            continue
        funktion = laengste_uebrige(r, {n_col})
        # Abgrenzung wird an der KOPFZEILEN-bestimmten Spalte entschieden, nicht
        # an `funktion` (das ist die laengste uebrige Zelle und kann die
        # Rechtsstandsnotiz sein).
        abgrenzung = ""
        for i in abg_cols:
            abgrenzung = abgrenzung or ist_abgrenzung(r[i] if i < len(r) else "")
        einzelne = _bruchstuecke_ergaenzen(split_normen(zelle), zelle, r, header_of(rows))
        if not einzelne and unzerlegbar is not None:
            unzerlegbar.append(norm_txt(zelle)[:160])
        for einzel in einzelne:
            # 'n.F.'/'a.F.' aus dem Normtext in ein eigenes Feld - sonst haengt
            # der Zusatz hinter dem Gesetzesnamen und stoert das Gesetz+Paragraph-
            # Matching der Stufe 2
            mr = RECHTSSTAND_RE.search(einzel)
            eintrag = {"norm": RECHTSSTAND_RE.sub("", einzel), "funktion": funktion}
            if mr:
                eintrag["rechtsstand"] = norm_txt(mr.group(1))
            # ZWEI RIEGEL, beide aus der Messung vom 11.09.2026:
            # 1. '[offen]' & Co sind ein ausdruecklicher Sperrgrund (Nutzerregel
            #    19.08.2026) - eine Abgrenzung darf ihn nie stilllegen.
            # 2. Ein Aktenzeichen in der Normspalte ist keine Norm; es hier
            #    stillzulegen wuerde eine Entscheidung aus der Pruefung nehmen.
            if abgrenzung and not is_platzhalter(eintrag["norm"]) \
                    and not alle_az(eintrag["norm"]):
                eintrag["abgrenzung"] = abgrenzung
            out.append(eintrag)
    return out


def rspr_aus(rows):
    az_col, az_n = pick_column(rows, is_az_cell)
    if az_col < 0 or az_n == 0:
        return []
    # Gerichtsspalte VORRANGIG aus der Kopfzeile, nicht aus dem Inhalt: bei
    # Sammelzeilen ('Weitere Instanzrechtsprechung', 'Arbeitsgerichte') steht dort
    # kein erkennbarer Gerichtsname, die Spalte wuerde inhaltsbasiert gar nicht
    # gefunden und das Kategorie-Label ginge verloren. Genau dieses Label
    # entscheidet in Stufe 2, ob der Platzhalter die Freigabe sperrt
    # (Nutzerregel 17b) - Fund 26.08.2026: 104 von 126 Labels waren weg.
    header = header_of(rows)
    ger_col = next((i for i, h in enumerate(header)
                    if "gericht" in h and "aktenzeichen" not in h), -1)
    if ger_col < 0:
        ger_col, _ = pick_column(rows, lambda s: bool(GERICHT_RE.search(s)) and len(s) < 80)
    dat_col, _ = pick_column(rows, lambda s: bool(DATUM_RE.search(s)) and len(s) < 60)
    out = []
    for r in rows[1:]:
        zelle = r[az_col] if az_col < len(r) else ""
        if not is_az_cell(zelle):
            continue
        gericht = gericht_aus(zelle)
        datum = datum_aus(zelle)
        if ger_col >= 0 and ger_col != az_col and ger_col < len(r):
            gericht = gericht or gericht_aus(r[ger_col])
            datum = datum or datum_aus(r[ger_col])
            # Kein erkannter Gerichtsname, aber die Spalte ist gefuellt: das ist ein
            # Kategorie-Label ('Weitere Instanzrechtsprechung') oder selbst ein
            # Platzhalter. Woertlich uebernehmen statt verwerfen.
            if not gericht and r[ger_col]:
                gericht = r[ger_col][:80]
        if dat_col >= 0 and dat_col < len(r):
            datum = datum or datum_aus(r[dat_col])
        relevanz = laengste_uebrige(r, {az_col, ger_col})
        azs = alle_az(zelle)
        if not azs:
            azs = [norm_txt(zelle)] if is_platzhalter(zelle) else []
        # Mehrere Zitate in EINER Zelle ('4 AZR 521/15, 25.01.2017; 4 AZR 375/16,
        # 25.10.2017'): je Az das zugehoerige Datum statt des ersten der Zelle (24.09.2026).
        z_norm = norm_txt(zelle)
        mehrfach = len(azs) > 1 and len(DATUM_RE.findall(z_norm)) + len(DATUM_ISO_RE.findall(z_norm)) > 1
        vorher = 0
        for az in azs:
            if not gericht and not az:
                continue
            d_az = datum
            i = z_norm.find(az, vorher)
            if mehrfach:
                d_az = datum_zum_az(z_norm, i, i + len(az)) if i >= 0 else datum
            # Gericht je Az aus dem EIGENEN Zitatglied (seit 25.09.2026): 'Entscheidungen:
            # EuGH Rosenbladt, BAG 7 AZR 632/15, BAG 7 AZR 68/14' gab allen drei BAG-Az
            # das erste Gericht der Zelle ('EuGH Rosenbladt', RETIREMENT_CASELAW).
            g_az = gericht_zum_az(z_norm, vorher, i) if i >= 0 else ""
            if i >= 0:
                vorher = i + len(az)
            out.append({"gericht": g_az or gericht, "aktenzeichen": az,
                        "datum": d_az, "relevanz": relevanz})
    return out


def ist_pflichtenliste(header):
    """Ist diese Tabelle eine Fachreview-PFLICHTENLISTE statt einer
    Quellendeklaration?

    Fund 29.08.2026 in JUS.TAX.ANALYSIS.GOBD.DATA_ACCESS / .EINVOICING: im
    Abschnitt 'Tragende Rechtsprechung' steht

        Kategorie | Status | Pflicht fuer Fachreview
        Hoechstrichterliche Rechtsprechung | [offen] | Konkrete BFH-/BVerfG-/
                                                      EuGH-Rechtsprechung nur
                                                      nach Kontextpruefung
                                                      aufnehmen.

    Das ist keine zitierte Entscheidung, sondern eine Auflage fuer das
    Fachreview. Weil `is_az_cell` das '[offen]' der STATUSSPALTE als
    Aktenzeichen gelten laesst, waehlte `pick_column` sie - und im Index stand
    `gericht='Konkrete BFH'` (Spalte 3, mitten im Wort abgeschnitten) mit
    `relevanz='Hoechstrichterliche Rechtsprechung'` (Spalte 1).
    Nutzerentscheidung: "komplett egal, da sperrt nix."

    BEWUSST ENG: geprueft wird nur die Kopfzeile auf 'Pflicht'/'Fachreview'.
    Eine zweispaltige `Status | Feststellung`-Tabelle bleibt unberuehrt - dort
    steht in MEAL_ALLOWANCE_ROUTING und PAYROLL_BRIEFING eine echte
    Rechtsprechungsaussage ('Keine zusaetzliche Instanzentscheidung wird in
    dieser Autorenfassung als tragende allgemeine Regel verwendet'), also ein
    dokumentierter Negativbefund, den der Index behalten muss. Ein breiterer
    Wurf haette diese 3 Zeilen mitgeloescht.
    """
    return any(re.search(r"pflicht|fachreview", h or "", re.IGNORECASE)
               for h in (header or []))


def taugt_als_rspr_rueckfall(rows):
    """Darf diese Tabelle als RUECKFALLEBENE fuer Rechtsprechung dienen?

    Die Ankertabelle ('Norm-/Fundstellen-Anker fuer jede Pruefaussage') greift
    nur, wenn die Datei unter 'Tragende Rechtsprechung' gar nichts fuehrt. Sie
    ist aber keine Quellendeklaration, sondern der PRUEFTEIL: ihre Spalten
    heissen 'Status', 'Status / Schritt', 'Quellenrang'. Weil is_az_cell() jeden
    Platzhalter als Aktenzeichen gelten laesst, waehlt pick_column() dort die
    STATUSSPALTE - und ein eingelesenes Prozessfeld wird zur einzigen
    'Rechtsprechung' der Datei und sperrt die Freigabe mit.

    Fund 29.08.2026, gemessen ueber alle 314 LABOR-Dateien - genau 2 Faelle:
      ANALYSIS.SOCIALSECURITY.PAYROLL_IMPACT
        'Beschluss verifiziert; Wert [Unverifiziert] / Schritt 3'
      ANALYSIS.SOCIALSECURITY.PARTIAL_INCAPACITY_REPORTING
        '[Unverifiziert] bis Veroeffentlichung'

    Kriterium bewusst eng: entweder weist die KOPFZEILE die Tabelle als
    Rechtsprechungstabelle aus, oder es steht irgendwo in ihr ein ECHTES
    Aktenzeichen. Nur-Platzhalter ohne beides ist kein Fundstellenbefund.

    Wirkt NUR auf die Rueckfallebene. Platzhalter in einer echten
    Rechtsprechungstabelle ('1. Hoechstrichterliche Rechtsprechung',
    '2. Einschlaegige Rechtsprechung der Instanzgerichte') bleiben unberuehrt -
    dort ist '[offen]' der Befund 'tragende Quelle nicht benannt' und muss
    weiter sperren (Nutzerregel 17b).
    """
    if header_has(header_of(rows), RSPR_HEADER_WORDS):
        return True
    return any(alle_az(z) for r in rows[1:] for z in r)


def tabellenart(rows, arten):
    """Bei Abschnitten, die beides erlauben ('Rechtsgrundlagen'), entscheidet
    die Kopfzeile bzw. die Trefferdichte, ob Norm- oder Rechtsprechungstabelle."""
    if len(arten) < 2:
        return arten
    header = header_of(rows)
    rspr_kopf = header_has(header, ("aktenzeichen", "entscheidung", "gericht"))
    norm_kopf = header_has(header, NORM_HEADER_WORDS)
    if rspr_kopf and not norm_kopf:
        return {"rspr"}
    if norm_kopf and not rspr_kopf:
        return {"norm"}
    _, n_n = pick_column(rows, is_norm_cell)
    _, a_n = pick_column(rows, is_az_cell)
    return {"rspr"} if a_n > n_n else {"norm"}


# ============================================================== Kopf-Metadaten

KOPF_FELDER = {
    "skill_id": ("skill-id", "skill id", "technische skill-id"),
    "rechtsgebiet": ("rechtsgebiet", "domäne", "domaene", "cluster"),
    "spezialmaterie": ("spezialmaterie", "spezialgebiet"),
    "jurisdiktion": ("jurisdiktion", "rechtsraum", "geltungsraum"),
}

# --------------------------------------------------------------------------
# Der SKILLKOPF maschinenlesbar (Nutzeranforderung 27.08.2026: "die Infos aus
# dem Skillkopf muessen maschinenlesbar sein").
#
# Gemessen ueber alle 314 LABOR-Dateien am 27.08.2026:
#   - JEDE Datei hat eine Kopftabelle; der Anker 'eine Zeile Skill-ID' trifft
#     immer.
#   - ABER: in 67 Dateien enthalten MEHRERE Tabellen eine Skill-ID-Zeile
#     (Subskill-Register, zweite Metadatenbloecke). Der Anker allein reicht
#     also nicht - genommen wird die Tabelle mit den MEISTEN Kernfeldern.
#     Genau diese Verwechslung hat am 27.08. in NONCOMPETE.CHECK dazu gefuehrt,
#     dass eine Aenderungsnotiz die Version aus dem falschen Block zitierte.
#   - 116 verschiedene Feldnamen, davon 93 in weniger als 20 Dateien. Die
#     Schreibvarianten werden ueber KOPF_ALIAS auf einen stabilen Schluessel
#     normalisiert, damit das Ausgabeschema fest ist.
# --------------------------------------------------------------------------

#: In allen 314 Dateien vorhanden - diese Felder machen die Kopftabelle aus.
KOPF_KERNFELDER = ("skill-id", "skill-familie", "version", "status", "autor",
                   "letzte aktualisierung", "freigabegrad", "vorgänger-id")

#: Kopfzeilen-Beschriftungen, die KEIN Feld sind ('Pflichtfeld | Inhalt').
KOPF_NICHTFELD = {"pflichtfeld", "feld", "wert", "inhalt", "ausprägung",
                  "ausprligung", "eintrag", "festlegung", "konkrete befüllung"}

#: Schreibvarianten -> stabiler Schluessel. Nur belegte Varianten.
KOPF_ALIAS = {
    "abweichung / begründung": "abweichungen_begruendung",
    "abweichungen": "abweichungen_begruendung",
    "abweichungen / begründung": "abweichungen_begruendung",
    "abweichungen / begruendung": "abweichungen_begruendung",
    "domäne": "rechtsgebiet",
    "modul / cluster": "cluster",
    "registry-cluster": "cluster",
    "taxonomiestatus": "taxonomiepruefung",
    "taxonomieprüfung": "taxonomiepruefung",
    "vorgänger-id": "vorgaenger_id",
    "vorgaenger-id": "vorgaenger_id",
    "letzte aktualisierung": "letzte_aktualisierung",
    "skill-id": "skill_id",
    "skill-familie": "skill_familie",
    "deutscher titel": "titel_de",
    "englischer titel": "titel_en",
}


def _kopf_schluessel(rohname):
    k = rohname.lower().strip(" :*")
    if k in KOPF_ALIAS:
        return KOPF_ALIAS[k]
    return re.sub(r"[^a-z0-9]+", "_", k.replace("ä", "ae").replace("ö", "oe")
                  .replace("ü", "ue").replace("ß", "ss")).strip("_")


def finde_kopftabelle(doc, erwartete_id=None):
    """Index der Kopftabelle, oder -1.

    **Nutzerregel 27.08.2026: die Skill-ID ist immer die, wie das Dokument
    selbst heisst.** Ueber alle 314 LABOR-Dateien gemessen: 314 von 314, ohne
    Ausnahme. Die Kopftabelle ist damit deterministisch bestimmbar - es ist
    die, deren 'Skill-ID'-Zeile GENAU den Dateinamen traegt. Keine Heuristik
    noetig.

    Das ist wichtig, weil 'Skill-ID' im Dokument in DREI Rollen vorkommt und
    nur eine davon das Kopffeld ist (gemessen: 67 der 314 Dateien fuehren
    mehrere):
      1. Kopffeld            'Skill-ID | JUS.LABOR.ANALYSIS....'   <- das hier
      2. Spaltenueberschrift 'Skill-ID | Abgrenzung / Routing'     (Subskill-Tabelle)
      3. Pruefzeile          'Skill-ID | Ja, entspricht der Registry'
    Der Dateinamen-Abgleich trennt die drei sicher; die Kernfeld-Zaehlung
    bleibt nur als Rueckfallebene, falls kein Wert passt.
    """
    if erwartete_id:
        for ti, t in enumerate(doc.tables):
            for r in table_matrix(t):
                if len(r) >= 2 and r[0].lower().strip(" :*").startswith("skill-id") \
                        and r[1].strip() == erwartete_id:
                    return ti
    bester, beste_zahl = -1, 0
    for ti, t in enumerate(doc.tables):
        rows = table_matrix(t)
        if not rows or max((len(r) for r in rows), default=0) < 2:
            continue
        erste = [r[0].lower().strip(" :*") for r in rows if r and len(r) >= 2]
        if not any(e.startswith("skill-id") for e in erste):
            continue
        treffer = sum(1 for kern in KOPF_KERNFELDER
                      if any(e.startswith(kern) for e in erste))
        if treffer > beste_zahl:
            bester, beste_zahl = ti, treffer
    return bester


def extract_kopf(doc, erwartete_id=None):
    """Vier kanonische Felder (Rueckwaertskompatibilitaet) plus den KOMPLETTEN
    Kopf als Feld/Wert-Abbildung unter '_alle' - maschinenlesbar, mit
    normalisierten Schluesseln und der Rohbeschriftung daneben.

    <erwartete_id> ist der Dateiname ohne Endung. Nach der Nutzerregel vom
    27.08.2026 ist er identisch mit der Skill-ID und bestimmt die Kopftabelle
    eindeutig. Passt keine, wird '_id_stimmt_nicht' gesetzt - das ist ein
    echter Mangel (Datei umbenannt, Kopf nicht nachgezogen oder umgekehrt)."""
    out = {k: "" for k in KOPF_FELDER}
    ti = finde_kopftabelle(doc, erwartete_id)
    out["_id_stimmt_nicht"] = bool(
        erwartete_id and ti >= 0
        and not any(r[0].lower().strip(" :*").startswith("skill-id")
                    and r[1].strip() == erwartete_id
                    for r in table_matrix(doc.tables[ti]) if len(r) >= 2))
    alle, roh = {}, {}
    if ti >= 0:
        for r in table_matrix(doc.tables[ti]):
            if len(r) < 2 or not r[0]:
                continue
            key = r[0].lower().strip(" :*")
            if key in KOPF_NICHTFELD:
                continue
            s = _kopf_schluessel(r[0])
            if s and s not in alle:
                alle[s] = r[1]
                roh[s] = r[0].strip(" :*")
    # Die vier kanonischen Felder bevorzugt AUS DER KOPFTABELLE ziehen; nur
    # wenn sie dort fehlen, dokumentweit suchen (aeltere Bauformen).
    for feld, aliase in KOPF_FELDER.items():
        for s, wert in alle.items():
            if any(s == _kopf_schluessel(a) for a in aliase) and wert:
                out[feld] = wert
                break
    for t in doc.tables:
        rows = table_matrix(t)
        if not rows or max((len(r) for r in rows), default=0) < 2:
            continue
        for r in rows:
            if len(r) < 2:
                continue
            key = r[0].lower().strip(" :*")
            for feld, aliase in KOPF_FELDER.items():
                if out[feld]:
                    continue
                if any(key == a or key.startswith(a) for a in aliase):
                    out[feld] = r[1]
    out["_alle"] = alle
    out["_beschriftungen"] = roh
    out["_kopftabelle"] = ti
    return out


# ==================================================================== Hauptlauf


# ------------------------------------------------- Rueckfallebenen (Bauform F/G)
# Bei einem Teil der Dateien (Familien TRANSFER, RECRUITING, NONCOMPETE) stehen
# unter 'Primäre Normen'/'Tragende Rechtsprechung' GAR KEINE Tabellen - die
# Fundstellen liegen im Fliesstext oder nur in der Ankertabelle
# ('Thema | Primärer Anker | Kontrollanker'). Beide Ebenen greifen deshalb nur,
# wenn der regulaere Abschnitt leer geblieben ist; die verwendete Quelle wird
# je Datei ausgewiesen ('_quelle_normen'/'_quelle_rspr').

# _KEIN_GESETZ: Untergliederungswoerter sind im Deutschen GROSSgeschrieben
# ("Abs.", "Nr.", "Satz") und passen damit genau auf das Muster eines
# Gesetzeskuerzels. Ohne diese Sperre liest ZITAT_RE per Backtracking das Wort
# hinter der Paragraphennummer als Gesetz: bei "Art. 19 Abs. 4, Art. 33 Abs. 2 GG"
# scheitert erst der Versuch "Art. 19 Abs. 4 <Gesetz>" am Komma, dann wird die
# optionale Abs-Gruppe uebersprungen und "Abs" selbst als Gesetz genommen ->
# Geistereintrag "Art. 19 Abs". Am 28.08.2026 an SIEBEN Eintraegen des
# JUS_LABOR-Index belegt ("Art. 19 Abs", "Art. 2 Abs", "Art. 33 Abs", "§ 2 Abs",
# "§ 241 Abs", "§ 613a Abs", "§ 95 Abs"); alle sieben landeten mit einer
# unbrauchbaren URL (/abs/__613a.html) in der Crawl-Liste.
# AUSGESCHRIEBENE Formen mit aufnehmen: '§§ 80 Absatz ...' erzeugte sonst den
# Geist '§§ 80 Absatz' - `Abs\b` greift bei "Absatz" naemlich NICHT (kein
# Wortende nach "Abs"). Am 28.08.2026 im Bestand gefunden, nachdem die
# Kurzformen schon gesperrt waren. Die laengeren Varianten stehen VOR den
# kuerzeren, damit die Alternative nicht vorzeitig abbricht.
# MEHRZAHL- UND BEUGUNGSFORMEN mit aufnehmen (31.08.2026): '§ 613a Abs. 1
# Sätze 1 bis 4, Abs. 2 und Abs. 3 BGB' erzeugte den Geist '§ 613a Abs. 1 Sätze'
# (ANALYSIS.TRANSFER.RIGHTS_OBLIGATIONS). Die Einzahl 'Satz' war gesperrt, die
# Mehrzahl 'Sätze' nicht - derselbe Fehler wie am 28.08. bei 'Absatz'. Die
# laengeren Formen stehen vor den kuerzeren.
_KEIN_GESETZ = (r"(?!(?:Absätzen|Absätze|Absaetze|Absatz|Abs|Nummern|Nummer|Nr|"
                r"Sätzen|Sätze|Saetzen|Saetze|Satz|S|"
                r"Halbsätze|Halbsatz|Halbs|HS|"
                r"Buchstaben|Buchstabe|Buchst|lit|Alternativen|Alternative|Alt|"
                r"Unterabsätze|Unterabsatz|Unterabs|Ziffern|Ziffer|Ziff|"
                r"Anlagen|Anlage|Anhänge|Anhang)\b\.?)")
# SAMMELZITATE IM FLIESSTEXT (seit 2026-08-28)
# ---------------------------------------------
# Vorher erkannte ZITAT_RE nur EIN Zitat bis zum Gesetz. Ein Sammelzitat wie
# "§§ 241 Abs. 2, 305 bis 310, 315, 611a BGB" fiel dadurch GANZ durch: der
# Versuch endete am ersten Komma, und weil kein Alternativzweig griff, entstand
# kein Eintrag. Solche Zitate waren damit weder Befund noch Bedarf - genau das,
# was laut Nutzervorgabe vom 28.08.2026 nicht passieren darf ("weder zu Geistern
# werden noch ignoriert werden noch sperren - in Einzelzeilen aufloesen und
# pruefen"). 36 von 315 Dokumenten haengen an diesem Fliesstext-Pfad.
#
# Der Sammelzweig steht VOR dem Einzelzweig: Python nimmt bei Alternativen die
# ERSTE, die an der linkesten Stelle passt, nicht die laengste. Umgekehrt
# sortiert wuerde der Einzelzweig das Sammelzitat nach dem ersten Glied
# abschneiden.
#
# _GESETZ_KURZ ist hier STRENGER als im Einzelzweig (zwei Grossbuchstaben bzw.
# durchgehend gross): eine Kommaliste kann sonst fremde Zahlen einsammeln -
# "§ 5 BetrVG, 3 Personen" wuerde mit dem lockeren Muster als
# "§ 5, 3 Personen" durchgehen, weil "Personen" grossgeschrieben ist.
# Ein Glied der Aufzaehlung. Die Bereichsformen sind an ZWEI Stellen noetig:
# hinter der Paragraphennummer ('§§ 140-148', '§§ 138d bis 138k') UND hinter dem
# Absatz ('§ 613a Abs. 1 bis Abs. 6 BGB' - dort wird das zweite 'Abs.'
# wiederholt). Ohne die zweite Stelle fiel das Zitat ganz durch.
_BEREICH = r"(?:\s*(?:bis|–|—|-)\s*(?:Abs\.?\s*)?\d+[a-z]?)?"
_TEIL = (r"\d+[a-z]?(?:\s*(?:ff\.|f\.))?" + _BEREICH +
         r"(?:\s+Abs\.?\s*\d+[a-z]?" + _BEREICH + r")?(?:\s+S(?:atz)?\.?\s*\d+)?"
         r"(?:\s+Nr\.?\s*\d+[a-z]?)?(?:\s+(?:Buchst|lit)\.?\s*[a-z]\)?)?")
# `_ROEMISCH_ENDE`: der Zusatz fuer Buchzaehlungen ('SGB V', 'BLV 2026') darf
# NICHT in das naechste Wort hineinlaufen. Ohne den Lookahead verschluckt
# `[IVXL]{1,4}` das fuehrende V/I/X/L des Folgeworts - belegt 29.08.2026 an
# '§ 130 OWiG Verletzung der Aufsichtspflicht …', das als Norm `§ 130 OWiG V`
# im Index landete (TAXCMS.GAPSCAN). Ein erfundenes Gesetz 'OWiG V' ist in
# Stufe 2 nie zu finden und sperrt den Skill dauerhaft.
_ROEMISCH_ENDE = r"(?![A-Za-zÄÖÜäöüß0-9])"
_GESETZ_KURZ = (r"(?:[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*"
                r"|[A-ZÄÖÜ]{2,})(?:\s+(?:[IVXL]{1,4}|\d{1,2})" + _ROEMISCH_ENDE + r")?")
# `*` statt `+`: ein Bereich OHNE Komma ist auch ein Sammelzitat
# ('§§ 140-148 AO' = neun Paragraphen). Mit `+` fiel genau diese Form durch.
_LISTE = _TEIL + r"(?:\s*,\s*" + _TEIL + r")*"

ZITAT_RE = re.compile(
    # 1. Sammelzitat mit §§/§ und Kommaliste
    r"(§§?\s*" + _LISTE + r"\s+" + _KEIN_GESETZ + _GESETZ_KURZ +
    # 2. Sammelzitat mit wiederholtem Art./Artikel
    r"|(?:Art\.|Artikel)\s*" + _TEIL +
    r"(?:\s*,\s*(?:Art\.|Artikel)?\s*" + _TEIL + r")+"
    r"\s+" + _KEIN_GESETZ + _GESETZ_KURZ +
    # 3. Einzelzitat § (unveraendert)
    r"|§§?\s*\d+[a-z]?(?:\s*(?:ff\.|f\.))?"
    r"(?:\s+Abs\.?\s*\d+[a-z]?)?(?:\s+S(?:atz)?\.?\s*\d+)?"
    r"(?:\s+Nr\.?\s*\d+[a-z]?)?(?:\s+(?:Buchst|lit)\.?\s*[a-z]\)?)?"
    r"\s+" + _KEIN_GESETZ + r"[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*"
    r"(?:\s+[IVXL]{1,4}" + _ROEMISCH_ENDE + r")?"
    # 4. Einzelzitat Art. (unveraendert)
    r"|(?:Art\.|Artikel)\s*\d+[a-z]?(?:\s+Abs\.?\s*\d+[a-z]?)?(?:\s+Nr\.?\s*\d+[a-z]?)?"
    r"\s+" + _KEIN_GESETZ + r"[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-/]*"
    r"(?:\s+[IVXL]{1,4}" + _ROEMISCH_ENDE + r")?)"
)


def _eu_nummer_anfuegen(zitat, text, ende):
    """`Art. 11 Durchführungsverordnung` -> `Art. 11 Durchführungsverordnung (EU) Nr. 282/2011`.

    ZITAT_RE endet am Gesetzes-/Rechtsaktnamen; bei EU-Rechtsakten steht die Nummer
    aber DAHINTER und gehoert zum Zitat. Ohne sie ist das Zitat unbestimmt: es sieht
    im Index aus wie ein Dokumentmangel („Artikel ohne konkreten Rechtsakt"), obwohl
    das Dokument voellig korrekt zitiert — Fund 29./30.08.2026 an vier Zeilen
    (`Art. 11`/`Art. 22 Durchführungsverordnung (EU) Nr. 282/2011`,
    `Art. 5 Richtlinie (EU) 2023/970`, `Art. 8 Richtlinie 2001/23/EG`). Sie standen
    als „fehlende Norm" in der Freigabeliste und haben mitgesperrt.

    Angehaengt wird nur, wenn EU_AKT_RE **unmittelbar am Zitatende** ansetzt und der
    Rechtsaktname des Treffers das letzte Wort des Zitats ist — sonst wuerde eine
    beliebige spaetere Nummer im selben Absatz angezogen.
    """
    # Der EU-Treffer muss GENAU dort beginnen, wo der Rechtsaktname im Zitat steht -
    # sonst zieht eine beliebige spaetere Nummer desselben Absatzes herein.
    for m in EU_AKT_RE.finditer(text):
        if m.end() <= ende:                     # Nummer liegt noch im Zitat: nichts zu tun
            continue
        name = m.group(0).split()[0].rstrip(",;")
        if not zitat.rstrip().endswith(name):
            continue
        if m.start() != ende - len(name):       # Name des Treffers = Ende des Zitats
            continue
        return (zitat.rstrip()[: -len(name)] + m.group(0)).strip()
    return zitat


def _prosa_gesetz_taugt(roh):
    """Der Gesetzesteil eines FLIESSTEXT-Zitats muss Abkuerzungsform haben oder
    ein bekannter Gesetzesname sein.

    Grund (31.08.2026): der Einzelzweig von ZITAT_RE laesst hinter dem
    Paragraphen JEDES grossgeschriebene Wort als Gesetz durchgehen. In den
    amtlichen Linklisten steht das Gesetz aber VORN und dahinter die
    Sachbezeichnung: 'BDSG – amtliche Fassung – § 26 Beschäftigtendaten
    einschließlich Bewerber' ergab '§ 26 Beschäftigtendaten'
    (REVIEW.RECRUITING.ASSESSMENT_PROCESS_RISKSCAN). Ein erfundenes Gesetz ist
    doppelt teuer: es sperrt dauerhaft und landet als unbrauchbares Ziel auf der
    Crawl-Liste.

    Kriterium ist NICHT eine Wortliste, sondern die Form: `_gesetz_ende` verlangt
    mindestens zwei Grossbuchstaben bzw. durchgehend gross ('BDSG', 'SGB IX',
    'KI-VO'). Was daran scheitert, muss die Namenskartei kennen
    ('Betriebsverfassungsgesetz') - sonst faellt das Zitat als gesetzlos durch
    und die vorgelagerten Wege (`_vorn_gesetz`, Bindeglied) duerfen es tragen.
    Nie raten: [[project_gesetzesnamen_kartei]].
    """
    if _gesetz_ende(roh):
        return True
    m = re.search(r"([A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-]*)\s*$", roh or "")
    if not m:
        return False
    from gesetzesnamen import abkuerzung as _abk, ist_bekannte_abkuerzung as _ist_abk
    wort = m.group(1)
    return bool(_abk(wort)) or _ist_abk(wort)


# Wörter, die AB DEM ERSTEN ZITAT in einer reinen Zitatliste vorkommen dürfen:
# Marker, Untergliederungen und Bindewörter. Alles andere ist Prosa - und Prosa
# im Zitatbereich heisst, dass die Zelle keine Aufzaehlung ist, sondern ein Satz.
_LISTENWORT = frozenset("""
art artikel anhang anlage
abs absatz absaetze absätze nr nummer satz saetze sätze halbsatz halbs hs
buchst buchstabe lit alt alternative unterabs unterabsatz ziff ziffer var variante
und bis sowie oder bzw ff f i.v.m ivm
""".split())
# Ein nachlaufendes Bindewort unmittelbar vor dem Gesetz gehoert nicht ins Zitat:
# aus '§ 6a Abs. 2 i.V.m. § 3 Abs. 1a UStG' wird sonst '§ 6a Abs. 2 i.V.m. UStG',
# was im CMS nie trifft. Der Paragraph selbst ist richtig, nur der Rest muss weg.
_BINDEWORT_ENDE_RE = re.compile(
    r"\s+(?:und|bis|sowie|oder|bzw\.?|i\.\s*V\.\s*m\.?|iVm)\s*$", re.IGNORECASE)


def _ist_listenwort(w):
    k = w.strip(" .,;").lower()
    # Ein EINZELNER Buchstabe ist der Untergliederungsbuchstabe ('Buchstabe b',
    # 'lit. e') - deutsche Prosa kennt kaum einbuchstabige Woerter.
    if not k or len(k) == 1 or k in _LISTENWORT:
        return True
    # Gesetzeskürzel: mindestens zwei Grossbuchstaben bzw. durchgehend gross -
    # dasselbe Formkriterium wie in `_taugt_als_gesetz`.
    return bool(re.fullmatch(r"[A-ZÄÖÜ][A-Za-zÄÖÜäöüß0-9\-/]*", w.strip(" .,;"))
                and sum(1 for c in w if c.isupper()) >= 2)


def _reine_zitatliste(t):
    """Besteht der Absatz AB DEM ERSTEN ZITAT nur aus Zitaten und Bindegliedern?

    Der Vorspann davor darf Prosa sein ('Verwaltungsauffassung, insbesondere
    UStAE / UStH zu § 1a, § 3d, § 10, § 13 und § 15 UStG') - dort steht die
    Kategorie, und `split_normen` schneidet sie ueber `_konkretisierung` ohnehin
    ab. Prosa NACH dem ersten Zitat ist dagegen das Ausschlusskriterium, und zwar
    aus zwei an echten Faellen gemessenen Gruenden (08.09.2026):

    * `DSGVO: insbesondere Art. 5, 6, 9, 13 bis 15, 22 und 82; ergaenzend § 26
      BDSG.` - hier haengt die ';'-Ebene von `split_normen` den DSGVO-Artikeln
      das **BDSG** an, weil `GESETZ_VORN_RE` an dem Doppelpunkt scheitert. Das
      BDSG hat keine Artikel: **acht erfundene Fundstellen** in einer Zeile.
    * `§§ 2 bis 4 NachwG, besonders § 2 Abs. 2 und 3 sowie Mitteilung vor
      Abreise.` - der Prosaschwanz wandert in das Zitat
      (`§ 2 Abs. 2 und 3 sowie Mitteilung vor Abreise. NachwG`).

    Beide Schwaechen stecken in `split_normen` und sind aelter als dieser Zweig;
    er darf sie nur nicht sichtbar machen. Wer die Bedingung lockert, baut sie
    sich zurueck - dann vorher wieder A/B messen.
    """
    m = re.search(r"§§?\s*\d|\bArt\.?\s*\d|\bArtikel\s*\d", t)
    if not m:
        return False
    # `(?<![0-9])`: der Buchstabe HINTER einer Zahl gehoert zum Paragraphen
    # ('§ 14a', '§ 22 Abs. 4f') und ist kein eigenes Wort. Ohne diese Sperre
    # scheitert die Pruefung an jedem Buchstabenparagraphen - gemessen an sechs
    # von sieben Positivfaellen.
    for w in re.findall(r"(?<![0-9])[A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß0-9\-/.]*", t[m.start():]):
        if not _ist_listenwort(w):
            return False
    return True


def _zitatliste_ganz(text, einzeln):
    """Bauform `§ A, § B und § C GESETZ` - Paragraphenzeichen je Glied, Gesetz nur am ENDE.

    Befund 08.09.2026 an den zwoelf VAT-Skills in `3. finalisiert`: der
    Rechtsgrundlagen-Abschnitt besteht dort aus Aufzaehlungsabsaetzen wie
    `§ 14, § 14a, § 14b, § 14c und § 15 UStG.` **ZITAT_RE trifft nur Zitate, hinter
    denen das Gesetz DIREKT steht** - hier also nur `§ 15 UStG`; die vier
    davorstehenden Paragraphen fielen weg. `split_normen` beherrscht die Bauform
    dagegen vollstaendig (es vererbt das Gesetz rueckwaerts ueber die Segmente),
    bekam bisher aber je Aufruf nur EINEN ZITAT_RE-Treffer und sah die Nachbarn
    darum nie. Gemessen ueber die zwoelf Dokumente: 33 statt 55 Normen, und die
    fehlenden 22 waeren in Stufe 2 nie gegen das CMS geprueft worden - genau die
    Fehlerklasse, in der ein Erkennerausfall wie ein Sachbefund aussieht.

    Der Zweig laeuft NUR im Fliesstext-Rueckfall; Normzellen gehen ohnehin direkt
    durch `split_normen`.

    DREI RIEGEL - der zweite und dritte machen den Zweig erst vertretbar:
    1. **Der Absatz muss selbst eine Zitatliste sein** (`is_norm_cell`) - ein
       Prosasatz wird nicht als Ganzes zerlegt.
    2. **Es muss MEHR herauskommen als bisher.** Sonst bleibt der bisherige Weg
       unberuehrt, und der Zweig kann in der Masse nichts bewegen.
    3. **Kein bisher gefundenes Zitat darf verschwinden.** Das ist die Kennzahl,
       auf die es bei dieser Stufe ankommt - nicht "wie viele Eintraege aendern
       sich", sondern "faellt irgendwo ein Gesetz weg". Statt sie hinterher zu
       messen, ist sie hier erzwungen: geht eines verloren, gilt der ganze
       Absatz als nicht behandelt.

    Abschaltbar mit `STUFE1_OHNE_ZITATLISTE=1`, um die Wirkung isoliert zu
    messen (dieselbe Mechanik wie `STUFE1_OHNE_BINDEGLIED`).
    """
    if os.environ.get("STUFE1_OHNE_ZITATLISTE") == "1" or not einzeln:
        return None
    t = norm_txt(text)
    if not t or not is_norm_cell(t) or not _reine_zitatliste(t):
        return None                                  # Riegel 1
    ganz = []
    for n in split_normen(t):
        n = re.sub(r"[\s.,;]+$", "", n)               # Satzpunkt der letzten Fundstelle
        ges = _gesetz_ende(n)
        if ges:                                      # nachlaufendes Bindewort abtrennen
            kopf = _BINDEWORT_ENDE_RE.sub("", n[:len(n) - len(ges)].rstrip())
            n = ("%s %s" % (kopf.rstrip(" ,;"), ges)).strip()
        if n and _prosa_gesetz_taugt(n) and n not in ganz:
            ganz.append(n)
    if len(ganz) <= len(einzeln):
        return None                                  # Riegel 2
    if any(e not in ganz for e in einzeln):
        return None                                  # Riegel 3
    return ganz


def normen_aus_prosa(absaetze):
    out = []
    for text in absaetze:
        einzeln = []
        for m in ZITAT_RE.finditer(text):
            roh = _eu_nummer_anfuegen(m.group(1), text, m.end(1))
            if not _prosa_gesetz_taugt(roh):
                continue
            for einzel in split_normen(roh):
                if einzel not in einzeln:
                    einzeln.append(einzel)
        ganz = _zitatliste_ganz(text, einzeln)
        for einzel in (ganz if ganz is not None else einzeln):
            out.append({"norm": einzel, "funktion": text[:300]})
    return out


def rspr_aus_prosa(absaetze):
    out = []
    for text in absaetze:
        for m in AZ_RE.finditer(text):
            # az_ok() GENAUSO anwenden wie in alle_az(). Ohne diese Zeile greift
            # der Fliesstextweg an dem Filter vorbei, den AZ_BLOCK fuer die
            # Tabellen setzt - Fund 29.08.2026 in genau der Datei, die der
            # Kommentar bei AZ_BLOCK schon nennt
            # (RESEARCH.INTERNATIONAL.COUNTRY_REQUIREMENTS_UPDATE): 'Art. 1 VO
            # 987/2009' und 'Art. 16 VO 987/2009' standen als '1 VO 987/2009'
            # bzw. '16 VO 987/2009' in der Rechtsprechungsspalte der
            # Freigabeliste und haben als fehlende Rechtsprechung mitgesperrt.
            # Das sind NORMEN, keine Aktenzeichen.
            if not az_ok(m.group(0)):
                continue
            umfeld = text[max(0, m.start() - 90): m.end() + 40]
            out.append({
                "gericht": gericht_aus(umfeld),
                "aktenzeichen": norm_txt(m.group(0)),
                "datum": datum_zum_az(text, m.start(), m.end()),
                "relevanz": text[:300],
            })
    return out


# Nur 'Schritt'/'Step' - bewusst ENGER als SCHRITT_PRAEFIX_RE (das auch
# 'Abschnitt', 'Teil', 'Modul' kennt und in einer Warnung steckt, wo ein
# Fehlgriff nichts kostet). Hier entstehen INDEXZEILEN, und 'Abschnitt B -
# Testfaelle' wuerde die Testfall-Sollwerte einsammeln.
ARBEITSSCHRITT_RE = re.compile(r"^(?:schritt|step)\b")


def belegzitate_aus_arbeitsschritt(text, ueberschrift):
    """Aktenzeichen, die im Fliesstext eines ARBEITSSCHRITTS als Beleg zitiert
    werden - z. B. 'familienstandsbezogene Zuschlaege einkommensmindernd zu
    beruecksichtigen (BSG B 12 KR 16/02 R)' unter 'Schritt 5 - Beitragszuschlag
    fuer familienversicherte Partner pruefen' (ANALYSIS.SOCIALSECURITY.
    PAYROLL_IMPACT). Solche Zitate standen in keiner Spalte des Index: die
    Arbeitsschritt-Ueberschrift ist kein Rechtsgrundlagen-Abschnitt, und die
    Rueckfallebene Fliesstext greift nur, wenn der Skill GAR keine
    Rechtsprechung fuehrt. Gemessen 31.08.2026: 15 echte Belegzitate in
    14 Skills.

    ZWEI WAECHTER, beide notwendig - ohne sie kommt Falsches herein:
    1. **Gericht im Umfeld erforderlich.** 'Art. 6 oder 7 ARB 1/80' nennt kein
       Gericht; die Zeile ist eine Norm. (Zusaetzlich in AZ_BLOCK gesperrt -
       ein Filter auf beiden Wegen, siehe Fund 29.08.2026.)
    2. **Nur Arbeitsschritte, keine Testfall-Umgebung.** Unter
       'Erwartete Hauptpruefungselemente', 'Pass-/Fail-Kriterium',
       'Eingabedaten' und 'Kurzbezeichnung: ...' stehen Testfall-SOLLWERTE, und
       darunter bewusst ERFUNDENE Aktenzeichen: '6 AZR 999/26' ist der
       Pruefgegenstand von RESEARCH.MASSDISMISSAL.CASELAW_UPDATE ('das
       behauptete Aktenzeichen ... nicht in den Rechtsbestand uebernehmen').
       Ein solcher Eintrag waere im CMS nie zu finden, wuerde den Skill
       dauerhaft sperren und als Crawl-Ziel eine nicht existierende
       Entscheidung suchen.
    """
    out = []
    for m in AZ_RE.finditer(text):
        az = norm_txt(m.group(0))
        if not az_ok(az):
            continue
        umfeld = text[max(0, m.start() - 120): m.end() + 60]
        gericht = gericht_aus(umfeld)
        if not gericht:
            continue
        out.append({
            "gericht": gericht,
            "aktenzeichen": az,
            "datum": datum_zum_az(text, m.start(), m.end()),
            "relevanz": ((ueberschrift + " – ") if ueberschrift else "") + text[:260],
        })
    return out


def _dedupe_normen(liste):
    seen, out = set(), []
    for n in liste:
        k = n["norm"].lower().replace(" ", "")
        if k not in seen:
            seen.add(k)
            out.append(n)
    return out


def _dedupe_rspr(liste):
    """Schluessel ist Aktenzeichen + Datum, weil dasselbe Aktenzeichen zwei
    Entscheidungen tragen kann (`10 AZR 330/16`: Anfragebeschluss und Urteil).

    ABER: ein LEERES Datum ist kein zweites Datum, sondern eine fehlende
    Angabe. Steht dasselbe Aktenzeichen einmal mit und einmal ohne Datum, ist
    das eine Dublette - sichtbar geworden am 31.08.2026, als AGG_RISK seine
    EuGH-Faelle sowohl aus 'Tragende Rechtsprechung' (mit Datum) als auch aus
    der amtlichen Linkliste (ohne Datum) bekam: `C-415/10`, `C-423/15` und
    `C-507/18` standen doppelt im Index."""
    mit_datum = {r["aktenzeichen"].lower().replace(" ", "") for r in liste if r["datum"]}
    seen, out = set(), []
    for r in liste:
        az = r["aktenzeichen"].lower().replace(" ", "")
        if not r["datum"] and az in mit_datum:
            continue
        k = (az, r["datum"])
        if k not in seen:
            seen.add(k)
            out.append(r)
    return out


def process(path):
    doc = Document(path)
    # Die Skill-ID ist immer der Dateiname (Nutzerregel 27.08.2026, ueber alle
    # 314 LABOR-Dateien bestaetigt) - damit ist die Kopftabelle eindeutig.
    kopf = extract_kopf(doc, os.path.basename(path).rsplit(".docx", 1)[0])
    normen, rspr, sonstige = [], [], []
    anker_norm, anker_rspr = [], []
    prosa_abs = {"norm": [], "rspr": []}
    verworfen = {"norm": 0, "rspr": 0}
    abgrenzung_raus = []      # Nutzerregel 11.09.2026, siehe ABGRENZUNG_PHRASEN
    abschnitte = []
    warnungen = []
    unbekannte_ueberschriften = []
    unzerlegbar = []
    letzte = ""
    letzte_stil = ""            # Stil der Ueberschrift, die `letzte` gesetzt hat
    beleg_rspr = []             # Belegzitate im Arbeitsschritt-Fliesstext
    for kind, obj in body_items(doc):
        if kind == "p":
            art_u = ueberschrift_art(obj)
            t = norm_txt(obj.text)
            if art_u:
                arten_alt, _ = abschnitt_art(letzte)
                arten_neu, aus_neu = abschnitt_art(t)
                # QUELLENZEILE, DIE WIE EINE UEBERSCHRIFT AUSSIEHT (31.08.2026).
                # In 'Amtliche Quellenlinks' steht 'AGG – Gesetze im Internet'
                # (Stil 'Source Text') als erste Zeile - kurz, ohne Schlusspunkt,
                # ohne Az/Norm/Datum, also heuristisch eine Ueberschrift. Sie
                # setzte den Abschnitt zurueck, und die zehn Entscheidungszeilen
                # darunter fielen aus dem Index (AGG_RISK).
                # Kriterium ist der STIL, nicht der Wortlaut: die Zeile ist
                # anders formatiert als die Ueberschrift, die den Abschnitt
                # eroeffnet hat. Damit greift die Ausnahme NIE gegen eine
                # stilerkannte Ueberschrift (die kommt hier nie als 'heuristik'
                # an) und nie in Dokumenten, deren Ueberschriften durchgehend
                # 'Normal' sind - dort sind Ueberschrift und Zeile gleich
                # formatiert, und die Ausnahme bleibt aus.
                if (art_u == "heuristik" and arten_alt and not arten_neu
                        and not aus_neu and letzte_stil
                        and absatz_stil(obj) != letzte_stil):
                    art_u = None
                else:
                    letzte = t
                    letzte_stil = absatz_stil(obj)
                    if ueberschrift_verdaechtig(letzte) and letzte not in unbekannte_ueberschriften:
                        unbekannte_ueberschriften.append(letzte)
            if art_u is None:
                arten, ausgeschlossen = abschnitt_art(letzte)
                # Schwelle absichtlich niedrig: in NONCOMPETE.* steht jede Norm
                # als eigener Ein-Zeilen-Absatz ('§ 110 GewO', 10 Zeichen)
                if t and len(t) > 3:
                    for a in arten:
                        prosa_abs[a].append(t)
                    # Auch ein Abschnitt OHNE Tabelle ist ein erkannter
                    # Abschnitt. Ohne diese Zeile meldete die Warnung
                    # 'kein_abschnitt_erkannt_trotz_kandidaten' fuer die
                    # 17 Fliesstext-Skills einen Erkennerausfall, den es nicht
                    # gab - der Abschnitt war erkannt, nur sein Inhalt stand in
                    # Absaetzen (Fund 31.08.2026).
                    if arten and letzte not in abschnitte:
                        abschnitte.append(letzte)
                    if not arten and not ausgeschlossen \
                            and ARBEITSSCHRITT_RE.match(clean_heading(letzte)):
                        beleg_rspr.extend(belegzitate_aus_arbeitsschritt(t, letzte))
            continue
        rows = table_matrix(obj)
        if len(rows) < 2:
            continue
        arten, _ = abschnitt_art(letzte)
        arten = tabellenart(rows, arten)
        n_kand = normen_aus(rows, unzerlegbar=unzerlegbar if "norm" in arten else None,
                            nur_zitate=ist_quellen_linkliste(letzte))
        r_kand = rspr_aus(rows)
        ist_anker = "anker" in clean_heading(letzte) or "fundstelle" in clean_heading(letzte)
        if "norm" in arten:
            # Abgrenzungsquellen (Nutzerregel 11.09.2026) wandern NICHT in den
            # Index, aber auch nicht ins Nichts: sie werden gesammelt und im
            # eigenen Blatt der xlsx ausgewiesen. Lautlos verschwinden duerfen
            # sie nicht - sonst sieht eine Regelwirkung wie ein Erkennerausfall
            # aus (und umgekehrt).
            for e in n_kand:
                if e.get("abgrenzung"):
                    abgrenzung_raus.append({"norm": e["norm"],
                                            "funktion": e.get("funktion", ""),
                                            "grund": e["abgrenzung"],
                                            "abschnitt": letzte})
                else:
                    normen.append(e)
            abschnitte.append(letzte)
        elif n_kand:
            verworfen["norm"] += len(n_kand)
            if ist_anker:
                anker_norm.extend(n_kand)
        if "rspr" in arten:
            # DERSELBE WAECHTER WIE AUF DER RUECKFALLEBENE (Fund 29.08.2026).
            # Er stand bisher nur dort - und genau deshalb rutschte im Abschnitt
            # 'Tragende Rechtsprechung' die Fachreview-Pflichtenliste
            # 'Kategorie | Status | Pflicht fuer Fachreview' durch: `is_az_cell`
            # laesst das '[offen]' der STATUSSPALTE als Aktenzeichen gelten,
            # `pick_column` waehlt sie, und im Index stand dann
            #   gericht='Konkrete BFH'  (Spalte 3, mitten im Wort abgeschnitten)
            #   relevanz='Hoechstrichterliche Rechtsprechung'  (Spalte 1)
            # Nutzerentscheidung 29.08.2026 zu diesen Zeilen: "komplett egal,
            # da sperrt nix." Lehre (zum zweiten Mal an einem Tag): einen Filter
            # immer auf ALLEN Wegen anbringen, die dieselbe Groesse erzeugen.
            if not ist_pflichtenliste(header_of(rows)) or taugt_als_rspr_rueckfall(rows):
                rspr.extend(r_kand)
                abschnitte.append(letzte)
            elif r_kand:
                verworfen["rspr"] += len(r_kand)
        elif r_kand:
            verworfen["rspr"] += len(r_kand)
            if ist_anker and taugt_als_rspr_rueckfall(rows):
                anker_rspr.extend(r_kand)
        header = header_of(rows)
        if header_has(header, QUELLEN_HEADER_WORDS) and not arten:
            for r in rows[1:]:
                if r and r[0] and not is_norm_cell(r[0]) and not is_az_cell(r[0]) \
                        and 3 < len(r[0]) <= 160:
                    sonstige.append(r[0])
    seen, uniq = set(), []
    for s in sonstige:
        if s.lower() not in seen:
            seen.add(s.lower())
            uniq.append(s)

    # RUECKFALLEBENEN: Ankertabelle UND Fliesstext, nicht 'entweder/oder'
    # (Korrektur 31.08.2026). Vorher gewann die Ankertabelle per `elif`, und
    # damit verdraengte die Zuordnungstabelle 'Thema | Primaerer Anker |
    # Kontrollanker' die ECHTE Deklaration des Skills: in 17 Skills stehen
    # 'Primaere Normen' und 'Tragende Rechtsprechung' als Fliesstext-Absaetze
    # ('§ 613a Abs. 1 und Abs. 4 BGB – amtliche Quelle. ...'), nicht als
    # Tabelle - in ANALYSIS.TRANSFER.TERMINATION_BAN standen deshalb 12 Normen
    # und 7 Entscheidungen aus der Ankertabelle im Index statt der elf
    # deklarierten Normzeilen und acht Entscheidungen. Die Vereinigung kann
    # nur ADDIEREN; die Reihenfolge (Fliesstext zuerst) entscheidet nur, welche
    # Fundstelle bei einer Dublette den Relevanztext stellt - und das ist die
    # Deklaration des Skills, nicht die Ankerzelle.
    quelle_n, quelle_r = "abschnitt", "abschnitt"
    if not normen:
        kand, herkunft = [], []
        if prosa_abs["norm"]:
            aus_prosa = normen_aus_prosa(prosa_abs["norm"])
            if aus_prosa:
                kand.extend(aus_prosa)
                herkunft.append("fliesstext")
        if anker_norm:
            kand.extend(anker_norm)
            herkunft.append("ankertabelle")
        if kand:
            normen, quelle_n = _dedupe_normen(kand), "+".join(herkunft)
        else:
            quelle_n = "keine"
    if not rspr:
        kand, herkunft = [], []
        if prosa_abs["rspr"]:
            aus_prosa = rspr_aus_prosa(prosa_abs["rspr"])
            if aus_prosa:
                kand.extend(aus_prosa)
                herkunft.append("fliesstext")
        if anker_rspr:
            kand.extend(anker_rspr)
            herkunft.append("ankertabelle")
        if kand:
            rspr, quelle_r = _dedupe_rspr(kand), "+".join(herkunft)
        else:
            quelle_r = "keine"

    # BELEGZITATE IM ARBEITSSCHRITT (31.08.2026): ADDITIV, nicht als
    # Rueckfallebene - sie stehen neben einer vollstaendigen
    # Rechtsprechungstabelle, nicht an ihrer Stelle. Deshalb hier, NACH den
    # Rueckfallebenen, und nur was noch fehlt.
    if beleg_rspr:
        vorhanden = {r["aktenzeichen"].lower().replace(" ", "") for r in rspr}
        neu = [r for r in _dedupe_rspr(beleg_rspr)
               if r["aktenzeichen"].lower().replace(" ", "") not in vorhanden]
        if neu:
            rspr = rspr + neu
            quelle_r = ("arbeitsschritt" if quelle_r == "keine"
                        else quelle_r + "+arbeitsschritt")
            warnungen.append("belegzitat_im_arbeitsschritt: %d" % len(neu))

    # ---- `ff.` aufloesen: eine Zeile JE PARAGRAPH (Nutzerregel 28.08.2026)
    # Hier, NACH der Deduplizierung, damit es fuer alle drei Herkunftswege
    # (Abschnitt, Ankertabelle, Fliesstext) gleich wirkt. Die Grenzen kommen aus
    # der Kartei, die `abschnittsgrenzen.py` schreibt - dieser Extraktor geht
    # nicht ins Netz.
    ff_offen = []
    if normen:
        karte = ff_karte()
        neu = []
        for n in normen:
            erweitert = expandiere_ff([n["norm"]], karte, ff_offen)
            if erweitert == [n["norm"]]:
                neu.append(n)
                continue
            for t in erweitert:
                neu.append(dict(n, norm=t, _ff_aus=n["norm"]))
        normen = _dedupe_normen(neu)
    # DIESELBE AUFLOESUNG FUER DIE ABGEZWEIGTEN EINTRAEGE (Fund im A/B 11.09.2026).
    # Sonst zaehlt das Blatt "Abgrenzungsquellen" die QUELLZEILE, der Index
    # verliert aber die aufgeloesten Einzelparagraphen: bei
    # REVIEW.EMPLOYMENT.SUMMARY stand dort EIN Eintrag `§§ 305 ff. BGB`, aus dem
    # Index verschwanden NEUN (§ 305, 305a, 305b, 305c, 306, 306a, 308, 309, 310).
    # Eine Bilanz, die 1 meldet und 9 bewegt, ist keine Bilanz.
    if abgrenzung_raus:
        karte = ff_karte()
        aufgeloest = []
        for a in abgrenzung_raus:
            erweitert = expandiere_ff([a["norm"]], karte, [])
            if erweitert == [a["norm"]]:
                aufgeloest.append(a)
                continue
            for t in erweitert:
                aufgeloest.append(dict(a, norm=t, _ff_aus=a["norm"]))
        abgrenzung_raus = aufgeloest
    if ff_offen:
        # NICHT still lassen: ein nicht aufgeloestes ff. ist entweder eine
        # fehlende Kartei-Zeile (dann `abschnittsgrenzen.py` nachziehen) oder ein
        # Zitat ohne Gesetz (dann ist das der Befund).
        warnungen.append("ff_ohne_abschnittsgrenze: %d" % len(ff_offen))

    # ---------------------------------------------------------- Warnungen
    # Grundsatz: ein Erkennerausfall darf NIEMALS wie ein Sachbefund aussehen.
    # "keine Rechtsprechung gefunden" kann zwei voellig verschiedene Dinge
    # heissen - der Skill deklariert keine, oder wir haben den Abschnitt nicht
    # erkannt. Nur der erste Fall ist ein Ergebnis; der zweite ist ein Mangel
    # DIESES Skripts und wird hier benannt, damit er nicht in den Index wandert.
    if not kopf["skill_id"]:
        warnungen.append("kopf_ohne_skill_id")
    if kopf.get("_id_stimmt_nicht"):
        warnungen.append("skill_id_weicht_vom_dateinamen_ab")
    fehlende_kopffelder = [f for f in ("rechtsgebiet", "spezialmaterie", "jurisdiktion")
                           if not kopf[f]]
    if fehlende_kopffelder:
        warnungen.append("kopf_unvollstaendig: " + ", ".join(fehlende_kopffelder))
    if not abschnitte:
        # Keine einzige Ueberschrift hat als Rechtsgrundlagen-Abschnitt getragen.
        if verworfen["norm"] or verworfen["rspr"]:
            warnungen.append(
                "kein_abschnitt_erkannt_trotz_kandidaten: %d Normen / %d Urteile lagen "
                "ausserhalb jedes erkannten Abschnitts - Abschnittserkennung pruefen, "
                "NICHT als 'Skill ohne Rechtsgrundlagen' auswerten"
                % (verworfen["norm"], verworfen["rspr"]))
        elif not normen and not rspr:
            warnungen.append(
                "nichts_gefunden_und_keine_kandidaten: weder Abschnitt noch Zitat "
                "irgendwo im Dokument - vor jedem Negativbefund per Hand gegenpruefen")
    if unbekannte_ueberschriften:
        warnungen.append("unbekannte_ueberschrift: "
                         + " | ".join(unbekannte_ueberschriften[:8]))
    if unzerlegbar:
        warnungen.append("norm_nicht_zerlegbar: " + " | ".join(unzerlegbar[:8]))
    if "fliesstext" in quelle_n or "fliesstext" in quelle_r:
        # Rueckfallebene hat getragen - das Ergebnis ist brauchbar, aber schwaecher
        # belegt als eine Tabelle. Sichtbar machen, nicht verstecken.
        warnungen.append("rueckfallebene_fliesstext: Normen=%s / Rechtsprechung=%s"
                         % (quelle_n, quelle_r))

    # ------------------------------------------------- Fundstellenluecken
    # Seit 2026-08-27. Diese vier Befunde stehen NICHT in einer Normen- oder
    # Rechtsprechungstabelle, sondern irgendwo im Dokument - der tabellen-
    # basierte Extraktor oben kann sie strukturell nicht sehen. Alle vier
    # Erkenner sind an diesem Tag gegen alle 314 LABOR-Skills validiert;
    # ihre Fallstricke stehen im Modulkopf von pruefkriterien_zitate.py.
    pruefkriterien = pk.alle_kriterien(
        doc, skill_id=os.path.basename(path).rsplit(".docx", 1)[0])
    for schluessel, wort in (("zitate_ohne_az", "zitat_ohne_aktenzeichen"),
                             ("namenszitate", "namenszitat_ohne_rechtssache"),
                             ("skillkopf_version", "skillkopf_version_uneindeutig"),
                             ("versionsruecksprung", "versionsruecksprung_in_notiz"),
                             ("kopf_unvollstaendig", "kopf_unvollstaendig")):
        if pruefkriterien[schluessel]:
            warnungen.append("%s: %d" % (wort, len(pruefkriterien[schluessel])))

    return {
        "skill_id": kopf["skill_id"] or os.path.basename(path).rsplit(".docx", 1)[0],
        "kopf": kopf.get("_alle", {}),
        "_kopf_beschriftungen": kopf.get("_beschriftungen", {}),
        "_kopftabelle": kopf.get("_kopftabelle", -1),
        "_pruefkriterien": pruefkriterien,
        "_warnungen": warnungen,
        "_unbekannte_ueberschriften": unbekannte_ueberschriften,
        "_nicht_zerlegbar": unzerlegbar,
        "_abgrenzungsquellen": abgrenzung_raus,
        "_quelle_normen": quelle_n,
        "_quelle_rspr": quelle_r,
        "rechtsgebiet": kopf["rechtsgebiet"],
        "spezialmaterie": kopf["spezialmaterie"],
        "jurisdiktion": kopf["jurisdiktion"],
        "normen": normen,
        "rechtsprechung": rspr,
        "sonstige_quellen": uniq,
        "_datei": os.path.basename(path),
        "_pfad": path,
        "_abschnitte": sorted(set(abschnitte)),
        "_verworfen_ausserhalb_abschnitt": verworfen,
    }


def sammle_pfade(quelle):
    """Nimmt entweder eine Textdatei mit einem Pfad je Zeile ODER direkt einen
    Ordner (dann rekursiv alle .docx). Gibt (pfade, uebersprungen, eingangsfehler)
    zurueck - jede Aussortierung wird begruendet, nichts verschwindet still."""
    roh = []
    if os.path.isdir(quelle):
        for wurzel, _dirs, dateien in os.walk(quelle):
            for d in sorted(dateien):
                roh.append(os.path.join(wurzel, d))
    else:
        with open(quelle, encoding="utf-8-sig") as f:
            roh = [z.strip().strip('"').strip("'") for z in f if z.strip()]

    pfade, uebersprungen, eingangsfehler = [], [], []
    gesehen = set()
    for p in roh:
        name = os.path.basename(p)
        schluessel = os.path.normcase(os.path.abspath(p))
        if name.startswith("~$"):
            # Word-Sperrdatei einer geoeffneten Datei - kein Dokument, aber auch
            # kein Fehler. Passiert immer, wenn der Nutzer eine Datei offen hat.
            uebersprungen.append((p, "Word-Sperrdatei (~$...)"))
            continue
        if not p.lower().endswith(".docx"):
            uebersprungen.append((p, "keine .docx-Datei"))
            continue
        if schluessel in gesehen:
            uebersprungen.append((p, "Pfad doppelt in der Eingabe"))
            continue
        if not os.path.isfile(p):
            eingangsfehler.append({"datei": p, "fehler": "Pfad existiert nicht"})
            continue
        if os.path.getsize(p) == 0:
            eingangsfehler.append({"datei": p, "fehler": "Datei ist 0 Byte"})
            continue
        gesehen.add(schluessel)
        pfade.append(p)
    return pfade, uebersprungen, eingangsfehler


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        print("Aufruf: python extract_rechtsgrundlagen_deterministisch.py "
              "<dateiliste.txt | Ordner> <out.json>")
        sys.exit(2)
    liste, outfile = sys.argv[1], sys.argv[2]
    if not os.path.exists(liste):
        print("ABBRUCH: Eingabe '%s' existiert nicht (weder Datei noch Ordner)." % liste)
        sys.exit(2)
    paths, uebersprungen, fehler = sammle_pfade(liste)
    if not paths:
        print("ABBRUCH: keine verarbeitbare .docx in der Eingabe gefunden.")
        for p, grund in uebersprungen[:20]:
            print("  uebersprungen:", os.path.basename(p), "-", grund)
        for fe in fehler[:20]:
            print("  FEHLER", fe["datei"], fe["fehler"])
        sys.exit(2)
    res = []
    for p in paths:
        try:
            res.append(process(p))
        except Exception as e:
            fehler.append({"datei": p, "fehler": repr(e)})
    with open(outfile, "w", encoding="utf-8") as f:
        json.dump({
            "dateien": res,
            "fehler": fehler,
            "uebersprungen": [{"datei": p, "grund": g} for p, g in uebersprungen],
        }, f, ensure_ascii=False, indent=1)
    print("eingelesen: %d | verarbeitet: %d | Fehler: %d | uebersprungen: %d"
          % (len(paths) + len(uebersprungen), len(res), len(fehler), len(uebersprungen)))
    print("Normen gesamt: %d | Rechtsprechung gesamt: %d" % (
        sum(len(r["normen"]) for r in res), sum(len(r["rechtsprechung"]) for r in res)))
    print("ohne Normen: %d | ohne Rechtsprechung: %d | ohne beides: %d" % (
        sum(1 for r in res if not r["normen"]),
        sum(1 for r in res if not r["rechtsprechung"]),
        sum(1 for r in res if not r["normen"] and not r["rechtsprechung"])))
    print("Platzhalter: Normen %d | Rechtsprechung %d" % (
        sum(1 for r in res for n in r["normen"] if is_platzhalter(n["norm"])),
        sum(1 for r in res for x in r["rechtsprechung"] if is_platzhalter(x["aktenzeichen"]))))
    print("ausserhalb der Abschnitte verworfen: Normen %d | Urteile %d" % (
        sum(r["_verworfen_ausserhalb_abschnitt"]["norm"] for r in res),
        sum(r["_verworfen_ausserhalb_abschnitt"]["rspr"] for r in res)))
    # Abgrenzungsquellen laut ausweisen (Nutzerregel 11.09.2026). Eine Regel,
    # die still Eintraege entfernt, ist von einem Erkennerausfall nicht zu
    # unterscheiden - deshalb Zahl UND Aufschluesselung nach Formulierung.
    abg = [a for r in res for a in r["_abgrenzungsquellen"]]
    print("Abgrenzungsquellen (nicht im Index, kein Sperrgrund): %d in %d Skills"
          % (len(abg), sum(1 for r in res if r["_abgrenzungsquellen"])))
    if abg:
        from collections import Counter as _C
        for grund, n in _C(a["grund"] for a in abg).most_common():
            print("    %-28s %d" % (grund, n))

    # ------------------------------------------------------------- Pruefblock
    # Alles, was ein Mensch ansehen muss, bevor der Index als Grundlage fuer
    # Stufe 2 taugt. Bewusst NACH den Summen und deutlich abgesetzt: eine
    # Kennzahl uebersieht man, einen benannten Dateinamen nicht.
    if uebersprungen:
        print("\nUEBERSPRUNGEN (%d):" % len(uebersprungen))
        for p, grund in uebersprungen[:20]:
            print("   %s - %s" % (os.path.basename(p), grund))
        if len(uebersprungen) > 20:
            print("   ... (+%d weitere)" % (len(uebersprungen) - 20))

    if fehler:
        print("\nFEHLER (%d) - diese Dateien fehlen im Index:" % len(fehler))
        for fe in fehler:
            print("   %s: %s" % (os.path.basename(fe["datei"]), fe["fehler"]))

    # Doppelte Skill-IDs brechen die Zuordnung in Stufe 2 (dort ist die Skill-ID
    # der Schluessel) - zwei Dateien mit derselben ID heben sich gegenseitig auf.
    ids = {}
    for r in res:
        ids.setdefault(r["skill_id"], []).append(r["_datei"])
    doppelt = {k: v for k, v in ids.items() if len(v) > 1}
    if doppelt:
        print("\nDOPPELTE SKILL-IDs (%d) - Zuordnung in Stufe 2 ist nicht eindeutig:"
              % len(doppelt))
        for k, v in sorted(doppelt.items()):
            print("   %s: %s" % (k, ", ".join(v)))

    mit_warnung = [r for r in res if r["_warnungen"]]
    if mit_warnung:
        print("\nPRUEFEN (%d von %d Dateien) - Warnung heisst NICHT falsch, sondern "
              "'nicht ohne Blick darauf verwenden':" % (len(mit_warnung), len(res)))
        nach_code = {}
        for r in mit_warnung:
            for w in r["_warnungen"]:
                nach_code.setdefault(w.split(":")[0], []).append(r)
        for code in sorted(nach_code, key=lambda c: -len(nach_code[c])):
            betroffen = nach_code[code]
            print("   [%s] %d Datei(en)" % (code, len(betroffen)))
            for r in betroffen[:5]:
                detail = next((w for w in r["_warnungen"] if w.startswith(code)), code)
                print("       %s" % r["_datei"])
                if ":" in detail:
                    print("           %s" % detail.split(":", 1)[1].strip()[:200])
            if len(betroffen) > 5:
                print("       ... (+%d weitere)" % (len(betroffen) - 5))
    else:
        print("\nPRUEFEN: keine Datei mit Warnung.")

    # Exit-Code, damit ein Aufrufer den Teilausfall nicht uebersehen kann:
    # 0 = alles verarbeitet, 1 = mindestens eine Datei fehlt im Index.
    if fehler:
        print("\nExit 1: %d Datei(en) sind NICHT im Index - der Lauf ist unvollstaendig."
              % len(fehler))
        sys.exit(1)
