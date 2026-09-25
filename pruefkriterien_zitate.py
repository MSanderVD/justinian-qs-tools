# -*- coding: utf-8 -*-
"""Stufe-1-Pruefkriterien fuer UNVOLLSTAENDIGE Zitate (Fundstellenluecken).

Ergaenzt den tabellenbasierten Extraktor um Befunde, die NICHT in einer
Normen- oder Rechtsprechungstabelle stehen, sondern irgendwo im Dokument -
und die deshalb bis zum 27.08.2026 unsichtbar waren. Alle Erkenner und ihre
Grenzen sind an diesem Tag gegen alle 314 LABOR-Skills validiert worden.

Die vier Kriterien
------------------
A  ZITAT OHNE AKTENZEICHEN
   "BAG 01.04.2026" - Gericht und Datum, aber keine Fundstelle. Das ist die
   Rechtsprechungs-Entsprechung der Dauerregel "nur das Gesetz, aber keinen
   Paragraphen - das ist sperrend". Ausbeute am 27.08.2026: 28 Zitate in
   9 Dateien, 26 davon aus dem CMS aufloesbar ueber gericht + entscheidungsdatum.

B  NAMENSZITAT OHNE RECHTSSACHE
   "EuGH Kuecuekdeveci" statt "EuGH 19.01.2010 - C-555/07". ACHTUNG, die
   wichtigste Regel des ganzen Moduls: **je DATEI pruefen, nicht je Zelle.**
   Ein Kurzname im Fliesstext ("Spijkers-Faktoren") ist im deutschen Rechtstext
   voellig korrekt, SOLANGE das Dokument die Rechtssache an mindestens einer
   Stelle vollstaendig fuehrt. Zellweise geprueft meldete der erste Lauf 21
   Spijkers-"Luecken"; je Datei geprueft waren es NULL. Wer das verwechselt,
   zerschiesst 21 richtige Formulierungen.

C  UEBERHOLTE STATUSAUSSAGE
   "Rat hat den Rechtsakt angenommen; Veroeffentlichung im Amtsblatt noch nicht
   dokumentiert". Solche Saetze frieren den Tag ein, an dem sie geschrieben
   wurden. Merkregel: nach "angenommen / Amtsblatt / Inkrafttreten" suchen,
   NICHT nur nach "Entwurf" oder "noch nicht in Kraft" - genau daran ist der
   erste Lauf vorbeigegangen und hat vier Zellen in ONBOARDING.CHECKLIST
   uebersehen.

D  SKILLKOPF-VERSION EINDEUTIG?
   **Nutzerregel 27.08.2026: es zaehlt NUR die Version im Skillkopf - die, die
   wir auch aendern. Alle anderen Versionsangaben im Dokument interessieren
   uns nicht.** Deckblatt, Fliesstext ("Autorenfassung v1.0.0-draft"),
   Aenderungsnotizen: kein Mangel, sondern irrelevant. Der erste Lauf meldete
   davon 112 Dateien - reines Rauschen.
   Gemeldet wird nur, wenn das Kopffeld SELBST nicht eindeutig ist (keines
   oder mehrere mit verschiedenen Werten). Das Feld ist die Zeile
   "Version | <x.y.z>" - die Wertpruefung ist entscheidend, weil auf einem
   Deckblatt "Version" als SPALTENUEBERSCHRIFT ueber "Rechtsstand" steht.

Zaehlregel fuer alle vier: **je ZITAT, nicht je ZELLE.** Enthaelt eine Zelle
daneben ein anderes, vollstaendiges Aktenzeichen, faellt das unvollstaendige
sonst durchs Raster - so sind am 27.08.2026 zwei Fundstellen entwischt.
"""
import re
import unicodedata

S = chr(47)          # Schraegstrich
ND = chr(8211)       # Halbgeviertstrich

# --------------------------------------------------------------- Aktenzeichen
_ROEM = r"(?:I{1,3}|IV|V|VI{1,3}|IX|X{1,2}|XI{1,2})"

#: Alle im Bestand vorkommenden Bauformen. Wer hier eine vergisst, erzeugt
#: Fehlalarme: BVerwG schreibt "5 PA 2.24" mit PUNKT, der BGH "II ZR 202/07"
#: mit roemischem Senat - beide matchen das Arbeitsgerichts-Muster nicht.
AZ_MUSTER = re.compile(
    r"\b\d{1,2}\s?[A-Za-z]{1,6}\s?\d{1,5}\s?" + S + r"\s?\d{1,4}\b"        # 2 AZR 714/08
    r"|\bB\s?\d{1,2}\s?[A-Za-z]{1,3}\s?\d{1,4}\s?" + S + r"\s?\d{1,3}\s?R\b"  # B 3 KR 22/17 R
    r"|\b\d{1,2}\s?[A-Za-z]{1,4}\s?\d{1,4}\.\d{2}\b"                        # 5 PA 2.24
    r"|\b" + _ROEM + r"\s?[A-Za-z]{1,3}\s?\d{1,4}\s?[" + S + r"\.]\s?\d{1,4}\b"  # II ZR 202/07
    r"|\bC" + chr(45) + r"\d{1,4}\s?" + S + r"\s?\d{2}\b"                   # C-555/07
    r"|\bRs\.\s?C" + chr(45) + r"\d+"
    r"|\bECLI:"
    r"|\b\d{1,3}\s?" + S + r"\s?\d{2}\b",                                   # 24/85 (EuGH alt)
    re.I)

#: Wie weit nach dem Datum darf das Aktenzeichen stehen? Nicht anker-, sondern
#: fensterbasiert pruefen: zwischen Datum und Aktenzeichen stehen im Bestand
#: Einschuebe wie "- verb. Rs." oder ", Az.". Anker-Pruefung meldete deshalb
#: vollstaendige BSG- und EuGH-Zitate faelschlich als Luecke.
_AZ_FENSTER = 46

#: Gericht unmittelbar vor einem Datum - die Zitatform, um die es geht.
_ZITAT = re.compile(
    r"\b(BAG|BVerfG|BVerwG|BGH|BSG|BFH|EuGH)\b"
    r"(?:\s*,)?\s*(?:vom\s+|v\.\s*|Urteil\s+vom\s+|Beschlu(?:ss|ß)\s+vom\s+)?"
    r"(\d{1,2})\.\s?(\d{1,2})\.\s?((?:19|20)\d{2})")

#: Rechtsstandsangaben sind KEINE Zitate ("BAG bis 04.12.2024 beruecksichtigt").
_STAND = re.compile(
    r"\bbis\b|\bStand\b|Rechtsstand|gepr(ü|ue)ft|verifiziert|"
    r"ber(ü|ue)cksichtigt|einbezogen|Erstreview|Datenbank|"
    r"Recherche|auffindbar|Aktualit", re.I)

#: Kriterium C - Formulierungen, die den Gesetzgebungsstand festschreiben.
_STATUS = re.compile(
    r"angenommen|Ratsannahme|Amtsblatt|Inkrafttret|in Kraft|"
    r"ver(ö|oe)ffentlich|Ver(ö|oe)ffentlichung|"
    r"noch nicht in Kraft|liegt als Entwurf|ist ein Entwurf", re.I)

#: Versionsnummer, NICHT Datum. Die letzte Gruppe darf hoechstens dreistellig
#: sein - sonst matcht "10.07.2026" und jede Rechtsstandsangabe gilt als
#: Version (Fund 27.08.2026: 337 Fehlalarme in 145 Dateien).
_VERSION = re.compile(r"(?<!\d)\d{1,3}\.\d{1,3}\.\d{1,3}(?:-[A-Za-z]+)?(?!\d)")


def nfc(s):
    """IMMER anwenden, bevor Text mit CMS-Rueckgaben verglichen wird: die
    Textfelder des CMS kommen NFD-zerlegt zurueck ('u' + U+0308 statt 'ue'),
    ein normal getipptes 'Betriebsuebergang' findet sie clientseitig NICHT."""
    return unicodedata.normalize("NFC", str(s or ""))


def _flach(zelle):
    return " ".join(zelle.text.split())


def _stellen(doc):
    """(Ort, Text) fuer jede Zelle und jeden Absatz des Dokuments."""
    for i, p in enumerate(doc.paragraphs):
        t = " ".join(p.text.split())
        if t:
            yield ("A%d" % i, "", t)
    for ti, tab in enumerate(doc.tables):
        kopf = " | ".join(_flach(c) for c in tab.rows[0].cells) if tab.rows else ""
        for ri, row in enumerate(tab.rows):
            for ci, c in enumerate(row.cells):
                t = _flach(c)
                if t:
                    yield ("T%d.%d.%d" % (ti, ri, ci), kopf, t)


def volltext(doc):
    return " ||| ".join(t for _, _, t in _stellen(doc))


# Tabellen des PRUEFTEILS - dort ist ein Gerichtsname keine Quellenangabe,
# sondern ein Hinweis auf eine Entscheidung ("seit BAG 30.01.2025 zusaetzlich
# die Bedeutung des Fehlers pruefen") oder ein Protokolleintrag des Fachreviews.
# Ein Hinweis braucht kein Aktenzeichen; ihn als Fundstellenluecke zu melden,
# erzeugt Arbeit an einer Stelle, die gar keine Quelle deklariert.
#
# Nutzerentscheidung 29.08.2026 ("grenz den erkenner ein"), nachdem die
# Durchsicht der 5 Befunde ergab: nur 2 stehen ueberhaupt in einer
# Rechtsprechungstabelle.
#
# BEWUSST NICHT ausgeschlossen: Tabellen mit 'Quelle' im Kopf
# ('Quelle | Verwendung | Pruefstatus'). Sie DEKLARIEREN eine Quelle - dort ist
# ein Zitat ohne Aktenzeichen ein echter Befund, auch wenn daneben ein
# Pruefstatus steht. Der Filter zielt auf den Pruefteil, nicht auf die
# Quellenmatrix.
_PROZESSTABELLE = re.compile(
    r"\bReviewpunkt\b|\bBeanstandung\b|\bReststatus\b|\bUmsetzung\b"
    r"|\bPr(?:ü|ue)fpunkt\b|\bAutorenhinweis\b|\bPr(?:ü|ue)ffrage\b"
    r"|\bPr(?:ü|ue)ffeld\b|\bArbeitsschritt\b", re.IGNORECASE)


# ============================================================= Kriterium A
def zitate_ohne_aktenzeichen(doc):
    """Gericht + Datum ohne Aktenzeichen. Zaehlt je ZITAT, nicht je Zelle:
    eine Zelle kann ein vollstaendiges und ein unvollstaendiges Zitat tragen.

    Tabellen des Pruefteils bleiben aussen vor (s. _PROZESSTABELLE)."""
    out = []
    for ort, kopf, s in _stellen(doc):
        if _STAND.search(s):
            continue
        if kopf and _PROZESSTABELLE.search(kopf):
            continue
        for m in _ZITAT.finditer(s):
            rest = s[m.end():m.end() + _AZ_FENSTER]
            # Steht IM FENSTER nach dem Datum ein Aktenzeichen, ist das Zitat
            # vollstaendig. Kein Anker: dazwischen stehen Einschuebe.
            if AZ_MUSTER.search(rest):
                continue
            gericht, tt, mm, jjjj = m.groups()
            out.append({
                "ort": ort, "tabellenkopf": kopf[:90], "gericht": gericht,
                "datum": "%s-%02d-%02d" % (jjjj, int(mm), int(tt)),
                "zitat": m.group(0), "zelle": s[:400],
                "befund": "Rechtsprechung ohne Aktenzeichen",
            })
    return out


# ============================================================= Kriterium B
def namenszitate_ohne_rechtssache(doc, namen):
    """namen: {Parteiname: Rechtssachennummer}. Prueft JE DATEI - ein Kurzname
    im Fliesstext ist kein Mangel, solange das Dokument die Nummer einmal
    vollstaendig fuehrt."""
    ganz = volltext(doc)
    out = []
    for partei, rs in namen.items():
        if partei not in ganz:
            continue
        if re.search(re.escape(rs).replace(r"\ ", r"\s*"), ganz):
            continue                      # irgendwo vollstaendig -> in Ordnung
        stellen = [(o, t) for o, _, t in _stellen(doc) if partei in t]
        out.append({
            "partei": partei, "soll": rs, "treffer": len(stellen),
            "ort": stellen[0][0] if stellen else "",
            "zelle": stellen[0][1][:400] if stellen else "",
            "befund": "Entscheidung nur namentlich zitiert, Rechtssache fehlt "
                      "im GESAMTEN Dokument",
        })
    return out


#: Woerter, die nach "EuGH" stehen koennen, ohne ein Parteiname zu sein.
_KEIN_PARTEINAME = {
    "Rechtsprechung", "Entscheidung", "Entscheidungen", "Urteil", "Urteile",
    "Kontext", "Vorlage", "Vorabentscheidung", "Richtlinie", "Verordnung",
    "Auslegung", "Kammer", "Senat", "Recht", "Rechtssache", "Verfahren",
    "Beschluss", "Leitsatz", "Grundsatz", "Linie", "Linien", "Anker",
    "Massenentlassung", "Arbeitszeit", "Urlaub", "Stand", "Rechtsstand",
    "Monitor", "Bezug", "Fragen", "Frage", "Nach", "Seit", "Siehe", "Vgl",
    "Sowie", "Oder", "Und", "Die", "Der", "Das", "Bei", "Zur", "Zum", "Auch",
    "Nur", "Kein", "Keine", "Ohne", "Deutschland", "Union", "Unionsrecht",
    "EuGH", "BAG", "BVerfG", "BGH", "BVerwG", "BSG", "BFH", "Gericht",
    "Gerichte", "Gerichtshof", "Instanzgerichte", "Aktuelle", "Aktuellen",
    "Neue", "Neuere", "Weitere", "Kontrolle", "Praxis", "Fassung", "Erste",
    "Zweite", "Dritte", "Vierte", "Fünfte", "Große", "Kommission", "Rechtsakt",
}
_EUGH_NAME = re.compile(r"\bEuGH\b[^.;:|)]{0,26}?"
                        r"\b([A-ZÄÖÜ][a-zäöüßA-ZÄÖÜ'\-]{3,})")
_RECHTSSACHE = re.compile(r"C" + chr(45) + r"\d{1,4}\s?" + S + r"\s?\d{2}"
                          r"|\b\d{1,3}" + S + r"\d{2}\b"
                          r"|Rs\.\s?C" + chr(45) + r"\d|ECLI:EU:")


def namenszitat_unbekannt(doc):
    """Ergaenzung zu Kriterium B fuer NEUE Skills: faengt Parteinamen, die in
    BEKANNTE_NAMEN noch gar nicht stehen koennen.

    Die feste Liste kann per Definition nur finden, was wir schon kennen -
    fuer einen kuenftigen Skill, der eine bislang unbekannte Entscheidung beim
    Namen nennt, ist sie blind. Dieser Erkenner arbeitet stattdessen ueber
    eine Abwesenheit: nennt das Dokument den EuGH mit einem Eigennamen, fuehrt
    aber im GESAMTEN Text KEINE einzige Rechtssachennummer, fehlt die
    Fundstelle mit hoher Wahrscheinlichkeit.

    Die Null-Bedingung je Datei macht das praezise: ueber alle 314 LABOR-Skills
    am 27.08.2026 nur 3 Treffer, alle drei Fehlalarme (Gerichtsnamen statt
    Parteinamen) - deshalb weiche Einstufung, nicht als harter Mangel.
    """
    ganz = volltext(doc)
    if "EuGH" not in ganz or _RECHTSSACHE.search(ganz):
        return []
    namen = set()
    for m in _EUGH_NAME.finditer(ganz):
        w = m.group(1)
        if w in _KEIN_PARTEINAME or w.isupper() or len(w) < 4:
            continue
        namen.add(w)
    if not namen:
        return []
    return [{"namen": sorted(namen), "befund":
             "Dokument nennt den EuGH mit einem Eigennamen, fuehrt aber "
             "NIRGENDS eine Rechtssachennummer - moegliches Namenszitat ohne "
             "Fundstelle. Weiches Signal, bitte ansehen."}]


# ============================================================= Kriterium C
def ueberholte_statusaussagen(doc, gegenstand):
    """gegenstand: Regex auf den Rechtsakt (z.B. 'Digital[- ]?Omnibus').
    Meldet Zellen, die dessen Gesetzgebungsstand behaupten - solche Saetze
    veralten still."""
    rx = re.compile(gegenstand, re.I)
    out = []
    for ort, kopf, s in _stellen(doc):
        if not rx.search(s) or not _STATUS.search(s):
            continue
        out.append({"ort": ort, "tabellenkopf": kopf[:90], "zelle": s[:400],
                    "befund": "Statusaussage zum Gesetzgebungsverfahren - "
                              "auf Aktualitaet pruefen"})
    return out


# ============================================================= Kriterium D
def skillkopf_version(doc):
    """NUR die Version im SKILLKOPF zaehlt.

    **Nutzerregel vom 27.08.2026, wortwoertlich:** „fuer uns zaehlt nur die
    Version im Skillkopf, den wir ja auch aendern. Alle anderen Versionen im
    Dokument interessieren uns nicht."

    Also: Deckblattangaben, Fliesstext-Fassungen ('Autorenfassung v1.0.0-draft'),
    Versionsnennungen in Aenderungsnotizen und Statussaetzen werden hier
    ausdruecklich NICHT gemeldet. Sie sind kein Mangel - sie sind irrelevant.
    (Der erste Lauf meldete davon 112 Dateien; das war Rauschen.)

    Gemeldet wird ausschliesslich, wenn das Kopffeld selbst nicht eindeutig
    ist: keines vorhanden, oder mehrere mit verschiedenen Werten. Genau das
    ist die Gefahr fuer einen Versions-Bump, der sonst in die falsche Zelle
    schreibt.

    Das Kopffeld ist die Zeile, deren ERSTE Zelle exakt 'Version' lautet UND
    deren zweite Zelle eine reine Versionsnummer ist. Die zweite Bedingung ist
    entscheidend: auf einem Deckblatt steht 'Version' als SPALTENUEBERSCHRIFT
    ueber dem Wert 'Rechtsstand' - ohne die Wertpruefung haelt ein Bumper die
    Kopfzeile fuer das Feld.
    """
    kandidaten = []
    for ti, tab in enumerate(doc.tables):
        for ri, row in enumerate(tab.rows):
            c = [_flach(x) for x in row.cells]
            if len(c) >= 2 and c[0] == "Version" and re.fullmatch(r"\d+\.\d+\.\d+", c[1]):
                kandidaten.append(("T%d.%d" % (ti, ri), c[1]))
    if not kandidaten:
        return [{"ort": "", "wert": "", "befund":
                 "kein eindeutiges Versionsfeld im Skill-Kopf - eine Zeile "
                 "'Version | <x.y.z>' fehlt; ein Versions-Bump haette hier "
                 "keinen Angriffspunkt"}]
    werte = {w for _, w in kandidaten}
    if len(kandidaten) > 1 and len(werte) > 1:
        return [{"ort": ", ".join(o for o, _ in kandidaten),
                 "wert": ", ".join(sorted(werte)),
                 "befund": "MEHRERE Versionsfelder im Skill-Kopf mit "
                           "verschiedenen Werten - ein Bump trifft sonst das "
                           "falsche Feld"}]
    return []


# ============================================================= Kriterium E
_A_MUSTER = re.compile(r"Version\s+((?<!\d)\d{1,3}\.\d{1,3}\.\d{1,3}(?:-[A-Za-z]+)?)"
                       r"\s*;\s*gegen(?:ü|ue)ber\s+v?"
                       r"((?<!\d)\d{1,3}\.\d{1,3}(?:\.\d{1,3})?(?:-[A-Za-z]+)?)", re.I)


def _versionswert(v):
    """Vergleichbarer Wert. Ein Suffix ('-draft') zaehlt als NIEDRIGER als
    dieselbe Nummer ohne - '1.0.1-draft' -> '1.0.1' ist ein Fortschritt, kein
    Rueckschritt. Ohne diese Regel meldet der Erkenner vier Dateien falsch."""
    zahlen = [int(x) for x in re.findall(r"\d+", v.split("-")[0])]
    while len(zahlen) < 3:
        zahlen.append(0)
    return tuple(zahlen[:3]) + (0 if "-" in v else 1,)


def versionsruecksprung(doc, kopf_tabelle=None):
    """Behauptet die Aenderungsnotiz einen Versions-RUECKSCHRITT?

    Die Notiz hat die Form 'Version X; gegenueber Y: ...'. Ist X nicht groesser
    als Y, stimmt etwas nicht - so ist am 27.08.2026 in NONCOMPETE.CHECK
    aufgefallen, dass eine Notiz '2.1.0' nannte, obwohl der Vorgaenger 2.1.1
    war: der Wert stammte aus einem zweiten Metadatenblock.

    Bewusst NUR der Rueckschritt. Der naheliegendere Test 'Notiz-Version gleich
    Feld-Version?' wurde gemessen und verworfen: 52 Treffer, fast alle
    rechtmaessig - die Notiz ist ein Changelog und nennt den Stand IHRES
    Eintrags, nicht den heutigen.
    """
    out = []
    for tab in ([kopf_tabelle] if kopf_tabelle is not None else doc.tables):
        for row in tab.rows:
            c = [_flach(x) for x in row.cells]
            if len(c) < 2 or c[0] != "Letzte Aktualisierung":
                continue
            for m in _A_MUSTER.finditer(c[1]):
                if _versionswert(m.group(1)) <= _versionswert(m.group(2)):
                    out.append({"neu": m.group(1), "alt": m.group(2),
                                "zelle": m.group(0)[:120],
                                "befund": "Aenderungsnotiz behauptet einen "
                                          "Versions-Ruecksprung - meist aus einem "
                                          "zweiten Metadatenblock abgeschrieben"})
    return out


# ============================================================= Kriterium F
#: Im Skillkopf verbindlich. Fehlt eines oder steht auf '[offen]', ist der
#: Kopf unvollstaendig - Stufe 0 haette das schliessen muessen.
KOPF_PFLICHT = ("Skill-ID", "Version", "Status", "Autor", "Fachreview",
                "Rechtsstand", "Freigabegrad", "Quellenstatus", "Vorgänger-ID")


def kopf_unvollstaendig(doc, kopf_tabelle=None):
    """Pflichtfelder im Skillkopf, die fehlen oder noch '[offen]' fuehren.
    Ueber alle 314 LABOR-Skills am 27.08.2026: null Befunde - der Erkenner
    laeuft als billige Absicherung mit, damit ein Rueckfall auffaellt."""
    if kopf_tabelle is None:
        return []
    kopf = {}
    for row in kopf_tabelle.rows:
        c = [_flach(x) for x in row.cells]
        if len(c) >= 2 and c[0]:
            kopf.setdefault(c[0], c[1])
    out = []
    for feld in KOPF_PFLICHT:
        wert = kopf.get(feld)
        if wert is None:
            out.append({"feld": feld, "befund": "Pflichtfeld fehlt im Skillkopf"})
            continue
        # NUR ein Feld melden, das AUSSCHLIESSLICH ein Platzhalter ist. Ein
        # gefuellter Wert, der irgendwo '[offen]' erwaehnt, ist eine bewusste
        # Teilaussage und KEINE Luecke - z.B. "27.08.2026; Verkuendung
        # verifiziert; Verlagsquellen [offen]". Der urspruengliche Test
        # ("[offen]" irgendwo enthalten) hat solche Felder als Luecke gemeldet;
        # ein Ueberschreiben waere Informationsverlust gewesen.
        rein = re.sub(r"[\s\.;,:\-]+", "", wert.lower())
        if rein in ("[offen]", "offen", "[tbd]", "tbd", "[n.n.]", "nn", ""):
            out.append({"feld": feld, "wert": wert[:80],
                        "befund": "Pflichtfeld steht nur auf [offen]"})
    return out


# ============================================================= Sammelaufruf
#: Parteinamen, die im Bestand vorkommen. Erweiterbar; die Nummer ist das Soll.
BEKANNTE_NAMEN = {
    "Kücükdeveci": "C" + chr(45) + "555" + S + "07",
    "Tomann": "C" + chr(45) + "134" + S + "24",
    "Sewel": "C" + chr(45) + "402" + S + "24",
    "Spijkers": "24" + S + "85",
    "Junk": "C" + chr(45) + "188" + S + "03",
    "Akavan": "C" + chr(45) + "44" + S + "08",
    "USDAW": "C" + chr(45) + "80" + S + "14",
    "Süzen": "C" + chr(45) + "13" + S + "95",
    "Klarenberg": "C" + chr(45) + "466" + S + "07",
    "Govaerts": "C" + chr(45) + "344" + S + "18",
    "Abler": "C" + chr(45) + "340" + S + "01",
}


#: Rechtsakte, deren Gesetzgebungsstand ueberwacht wird. **HIER ERWEITERN**,
#: sobald ein neues Vorhaben in die Skills einzieht - Kriterium C kann NICHT
#: generisch arbeiten.
#:
#: Am 27.08.2026 gemessen und verworfen: ein Erkenner ohne benannten Gegenstand
#: liefert 59 Befunde in 42 Dateien, ueberwiegend Fehlalarme - "angenommen"
#: trifft "bei angenommener Berechtigung", "automatisch angenommen". Eine
#: geschaerfte Fassung wurde SCHLECHTER (103 in 65), weil "noch nicht
#: abgeschlossen" und "ist ein Entwurf" Alltagssaetze sind ("die Aufklaerung
#: ist noch nicht abgeschlossen", "Diese Autorenfassung ist ein Entwurf").
#: Der Gegenstand muss benannt werden - das ist keine Bequemlichkeit, sondern
#: die einzige Form, in der das Kriterium praezise ist.
BEOBACHTETE_RECHTSAKTE = [
    r"Digital[- ]?Omnibus|KI-Omnibus|Omnibus on AI",
    r"Entgelttransparenzrichtlinie|Richtlinie \(EU\) 2023" + S + r"970",
    r"GKV-BStabG|GKV-Beitragssatzstabilisierungsgesetz",
]


def finde_kopftabelle(doc, erwartete_id):
    """Die Kopftabelle ueber den Dateinamen. Nutzerregel 27.08.2026: die
    Skill-ID ist immer die, wie das Dokument heisst (314/314 gemessen)."""
    if not erwartete_id:
        return None
    for tab in doc.tables:
        for row in tab.rows:
            c = [_flach(x) for x in row.cells]
            if len(c) >= 2 and c[0].lower().strip(" :*").startswith("skill-id") \
                    and c[1].strip() == erwartete_id:
                return tab
    return None


def alle_kriterien(doc, statusgegenstand=None, skill_id=None):
    """Ein Aufruf, fuenf Befundlisten. Wird von process() mitgefuehrt."""
    gegenstaende = ([statusgegenstand] if statusgegenstand
                    else BEOBACHTETE_RECHTSAKTE)
    status = []
    for g in gegenstaende:
        status.extend(ueberholte_statusaussagen(doc, g))
    kt = finde_kopftabelle(doc, skill_id)
    return {
        "zitate_ohne_az": zitate_ohne_aktenzeichen(doc),
        "namenszitate": namenszitate_ohne_rechtssache(doc, BEKANNTE_NAMEN),
        "namenszitat_unbekannt": namenszitat_unbekannt(doc),
        "statusaussagen": status,
        "skillkopf_version": skillkopf_version(doc),
        "versionsruecksprung": versionsruecksprung(doc, kt),
        "kopf_unvollstaendig": kopf_unvollstaendig(doc, kt),
    }


# ============================================================= Direktaufruf
# Damit die Kriterien nicht nur im grossen Stufe-1-Lauf greifen, sondern auch
# auf eine EINZELNE Datei oder einen Ordner - fuer den Fall "sieh dir diesen
# neuen Skill an", ohne den ganzen Index neu zu bauen.
#
#   python pruefkriterien_zitate.py <datei.docx | ordner> [...]
if __name__ == "__main__":
    import sys
    import os
    from docx import Document

    ziele = []
    for arg in sys.argv[1:]:
        if os.path.isdir(arg):
            ziele += [os.path.join(arg, f) for f in sorted(os.listdir(arg))
                      if f.lower().endswith(".docx") and not f.startswith("~$")]
        elif os.path.isfile(arg):
            ziele.append(arg)
    if not ziele:
        print(__doc__)
        print("Aufruf: python pruefkriterien_zitate.py <datei.docx | ordner> [...]")
        sys.exit(1)

    # Die Integritaetspruefung liegt bei Stufe 0 und wird von hier ueber den
    # absoluten Pfad aufgerufen - eine zweite Kopie wuerde auseinanderlaufen.
    # Sie prueft eine Ebene TIEFER (ZIP/XML-Namespaces) als alles hier und
    # faengt Dateien, die python-docx fuer sauber haelt, Word aber nicht oeffnet.
    import subprocess
    INTEGRITAET = (r"C:\Users\m.sander\.claude\skills\stufe_0_finalisierung"
                   r"\references\check_docx_integrity.py")
    if os.path.isfile(INTEGRITAET):
        for ordner in sorted({os.path.dirname(os.path.abspath(p)) for p in ziele}):
            r = subprocess.run([sys.executable, INTEGRITAET, ordner],
                               capture_output=True, text=True, encoding="utf-8")
            if r.returncode != 0:
                print("INTEGRITAET: Befunde in %s" % ordner)
                for z in (r.stdout or "").strip().splitlines()[-12:]:
                    print("   " + z)
        print()
    else:
        print("Hinweis: check_docx_integrity.py nicht gefunden - "
              "Integritaetspruefung uebersprungen.")
        print()

    gesamt = {}
    for pfad in ziele:
        try:
            doc = Document(pfad)
        except Exception as e:
            print("FEHLER %s: %s" % (os.path.basename(pfad), str(e)[:70]))
            continue
        name = os.path.basename(pfad)
        befunde = alle_kriterien(doc, skill_id=name.rsplit(".docx", 1)[0])
        n = sum(len(v) for v in befunde.values())
        for k, v in befunde.items():
            gesamt[k] = gesamt.get(k, 0) + len(v)
        if n == 0:
            print("OK      %s" % name)
            continue
        print("BEFUND  %s" % name)
        for k, v in befunde.items():
            for e in v:
                zusatz = (e.get("zitat") or ", ".join(e.get("namen", []))
                          or e.get("feld") or e.get("wert")
                          or (e.get("neu") and "%s gegen %s" % (e["neu"], e["alt"]))
                          or "")
                print("   [%-22s] %s" % (k, zusatz))
                print("      %s" % str(e.get("befund", ""))[:150])
                if e.get("zelle"):
                    print("      Text: %s" % str(e["zelle"])[:150])
    print()
    print("Geprueft: %d Datei(en)" % len(ziele))
    for k, v in sorted(gesamt.items()):
        print("   %-24s %d" % (k, v))
