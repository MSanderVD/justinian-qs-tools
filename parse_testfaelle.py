"""Liest `references/testfaelle.md` eines Skill-Packages und normalisiert die
Testfaelle - unabhaengig davon, in welchem der beiden bekannten Formate sie stehen.

WARUM (Befund 2026-08-19)
------------------------
Der v1.1-Builder erzeugt nicht immer dasselbe Format:

  Format A "Tabelle"  (z.B. CONFIDENTIALITY, PROBATION, FREELANCER.CLASSIFY)
      ## TC-CONF-01 - Kurzbezeichnung
      | Feld | Inhalt |
      | Testfall-ID | TC-CONF-01 |
      ... 14 Zeilen inkl. Pass-Kriterien, Fail-Indikatoren, Maschinenanker

  Format B "Prosa"    (z.B. APPRENTICESHIP.DURATION_CHANGE, IDs TF-<n>-01)
      ## TF-27-01 - Kurzbezeichnung
      **Testfall-ID:** TF-27-01
      **Eingabedaten:** ...
      ... nur NEUN Felder: es fehlen Testfalltyp, Nicht-Aktivierung,
      erwartete Rechtsprechung, erwartete Rueckfragen, Pass-Kriterien,
      Fail-Indikatoren und Maschinenanker

  Format A2 "Tabelle, fette Labels"  (z.B. APPRENTICESHIP.MINIMUM_REMUNERATION,
                                      IDs TF-<n>-01 trotz Tabellenform)
      ## TF-24-01 - Kurzbezeichnung
      | **Testfall-ID** | TF-24-01 |
      | **Pass-/Fail-Kriterium** | Pass: ... Fail: ... |
      Zwei Abweichungen gegenueber A, beide gemessen am 2026-08-30: die Labels
      sind fett ausgezeichnet, und Pass und Fail stehen in EINER Zelle. Ohne
      Behandlung wurde kein einziges Feld erkannt, der Fall fiel auf "prosa"
      zurueck und meldete alle Pflichtfelder als fehlend. Das sah nach einer
      lueckenhaften Vorgabe aus, war aber ein Parser-Fehler - siehe _kanon()
      und _teile_pass_fail().

Format B liefert damit genau die Felder NICHT, auf denen die Verdikt-Logik von
Stufe 5 aufsetzt (Fail-Indikator -> Fail, Pass-Kriterien -> Kriterienliste,
erwartete Rechtsprechung -> Quellendeckung). Ersatzweise steht in Format B ein
Abschnitt "Automatisierbare QS-Pruefpunkte" mit fallbezogenen Stichpunkten; die
werden je Testfall zugeordnet und als `qs_pruefpunkte` ausgegeben. Was trotzdem
fehlt, steht je Testfall in `fehlende_felder` und gehoert nach Schritt 4c ins
Protokollfeld `testfall_luecken` - eine Luecke der VORGABE, kein Fail des Agenten.

    python parse_testfaelle.py <testfaelle.md> [--json <ausgabe.json>]
"""
import json
import re
import sys

# Kanonische Feldnamen -> erkannte Schreibweisen (klein, ohne Doppelpunkt)
FELDER = {
    "testfall_id": ["testfall-id", "testfall id"],
    "kurzbezeichnung": ["kurzbezeichnung"],
    "testfalltyp": ["testfalltyp", "testfall-typ"],
    "eingabedaten": ["eingabedaten"],
    "erwartete_aktivierung": ["erwartete aktivierung"],
    "erwartete_nicht_aktivierung": ["erwartete nicht-aktivierung", "erwartete nicht aktivierung"],
    "erwartete_kernnormen": ["erwartete kernnormen"],
    "erwartete_rechtsprechung": ["erwartete rechtsprechung / quellen", "erwartete rechtsprechung/quellen",
                                 "erwartete rechtsprechung", "erwartete quellen"],
    "erwartete_hauptpruefungselemente": ["erwartete hauptprüfungselemente", "erwartete hauptpruefungselemente"],
    "erwartete_rueckfragen": ["erwartete rückfragen", "erwartete rueckfragen"],
    "erwartetes_ergebnis": ["erwartetes ergebnis"],
    "erwartete_risiken": ["erwartete risiken"],
    "erwartete_eskalationsstufe": ["erwartete eskalationsstufe"],
    "pass_kriterien": ["pass-kriterien", "pass kriterien"],
    "fail_indikatoren": ["fail-indikatoren", "fail indikatoren"],
    "maschinenanker": ["maschinenanker"],
    # Format A2 fuehrt Pass und Fail in EINER Zelle ("Pass: ... Fail: ...").
    # Wird nach dem Einlesen von _teile_pass_fail() aufgeteilt.
    # Sechste Etikettvariante, gemessen 2026-09-10 an TERMINATION.MISCONDUCT:
    # "Pass-Kriterien / Fail-Indikatoren" nennt BEIDE Pflichtfelder im Label und
    # sah deshalb aus wie zwei fehlende Felder, obwohl der Wert vollstaendig in
    # der Zelle stand. Dieselbe Fehlerfamilie wie A2: Etikett gesucht, Bauform
    # verfehlt.
    "pass_fail_kombiniert": ["pass-/fail-kriterium", "pass-/fail-kriterien",
                             "pass/fail-kriterium", "pass- und fail-kriterium",
                             "pass-kriterien / fail-indikatoren",
                             "pass-kriterien/fail-indikatoren",
                             "pass-kriterien und fail-indikatoren",
                             "pass-kriterium / fail-indikator"],
    # Format C (Feld-Ueberschriften) fasst Felder zusammen, die A/B trennen.
    # Aufgeloest von _verteile_kombiniert(): der Text geht in BEIDE Zielfelder,
    # weil Normen und Rechtsprechung dort in einem Fliesstext stehen - Kopieren
    # ist ehrlicher als ein geratener Schnitt.
    "kernnormen_rspr_kombiniert": ["erwartete kernnormen und rechtsprechung",
                                   "erwartete kernnormen / rechtsprechung",
                                   "erwartete kernnormen und quellen"],
    "risiken_eskalation_kombiniert": ["erwartete risiken / eskalation",
                                      "erwartete risiken und eskalation",
                                      "erwartete risiken / eskalationsstufe"],
}
# Der 17.09.-Builder fuehrt die Soll-Erwartung als MATRIX `| Prueffeld |
# Soll-Erwartung |` mit eigenen Etiketten (FIXEDTERM.EXTENSION) bzw. als
# beschriftete Absaetze (FIXEDTERM.CHECK). Die Werte sind vollstaendig da, nur
# die Etiketten standen in keiner Aliasliste - also wieder "Etikett gesucht,
# Bauform verfehlt", diesmal fuer die Felder, gegen die Schritt 5 abgleicht.
# "Anker"/"Normen / Anker" nennt Norm UND Entscheidung in einem Fliesstext
# ("§ 14 Abs. 2 Satz 1 TzBfG; BAG 7 AZR 108/20") - deshalb das Kombifeld, das
# den Text in beide Zielfelder legt, statt eines geratenen Schnitts.
FELDER["kernnormen_rspr_kombiniert"] += ["anker", "normen / anker", "normen/anker",
                                         "normen und anker"]
FELDER["erwartetes_ergebnis"] += ["spezialergebnis"]
FELDER["erwartete_aktivierung"] += ["aktivierung"]
# Format C schreibt den Sachverhalt unter eine eigene Ueberschrift statt in ein
# Feld "Eingabedaten"; Format B (RETIREMENT) nummeriert die Faelle im Kopf und
# fuehrt die echte Kennung erst im Block.
FELDER["eingabedaten"] += ["ausführlicher sachverhalt", "ausfuehrlicher sachverhalt",
                           "sachverhalt"]
# Format D fuehrt BEIDES getrennt: "1. Kurzsachverhalt" (die Erzaehlung) und
# "2. Eingabedaten" (Rollen/Dokumente/Stichtag als Liste). `kurzsachverhalt`
# bekommt deshalb ein eigenes Feld statt eines Alias auf `eingabedaten` - wer
# nur eines von beiden an den Agenten gibt, verliert entweder die Tatsachen
# oder die Metadaten. Die Zusammensetzung entscheidet der Aufrufer.
FELDER["kurzsachverhalt"] = ["kurzsachverhalt", "kurz-sachverhalt"]
# Format D laesst in der Aufzaehlung das Wort "Erwartete" weg ("- Ergebnis: …",
# "- Eskalationsstufe: Green, sofern …"). Ohne diese Aliase fehlt genau der
# fachliche Massstab, gegen den Schritt 5 abgleicht - der Testlauf laeuft dann,
# aber das Verdikt haette kein Soll.
FELDER["erwartetes_ergebnis"] += ["ergebnis"]
FELDER["erwartete_eskalationsstufe"] += ["eskalationsstufe"]
FELDER["erwartete_risiken"] += ["risiken"]
FELDER["erwartete_hauptpruefungselemente"] += ["hauptprüfung", "hauptpruefung"]
FELDER["erwartete_hauptpruefungselemente"] += ["erwartete prüfungsschritte",
                                               "erwartete pruefungsschritte"]
# Format E (gemessen 03.10.2026 an NONCOMPETE.DURATION/.COMPENSATION): Ueberschriften
# "### Eingabesachverhalt", "### Richtiges Ergebnis / Kernaussage", "### Richtige
# Risikoeinstufung" ... Kein Alias traf, der Freitext-Rueckfall nahm den LAENGSTEN
# Absatz - in 5 Faellen die MUSTERLOESUNG ("Richtiges Ergebnis: ..."), die dann als
# Nutzereingabe an den Agenten ging. Teuerster Fehlertyp: der Lauf sieht normal aus.
FELDER["eingabedaten"] += ["eingabesachverhalt", "eingabe-sachverhalt", "eingabe"]
FELDER["erwartetes_ergebnis"] += ["richtiges ergebnis / kernaussage", "richtiges ergebnis",
                                  "kernaussage"]
FELDER["erwartete_eskalationsstufe"] += ["richtige risikoeinstufung", "risikoeinstufung"]
FELDER["erwartete_aktivierung"] += ["erwartete skill-aktivierung", "erwartete skill aktivierung"]
FELDER["erwartete_hauptpruefungselemente"] += ["erwarteter prüffokus", "erwarteter prueffokus"]
FELDER["kernnormen_rspr_kombiniert"] += ["erwartete normzitate und fundstellen"]
# Format F (RESEARCH-Skills, gemessen 03.10.2026): Etiketten im SINGULAR ("**Pass-Kriterium:**",
# "**Fail-Kriterium:**") und "Hauptprüfungselemente" ohne "Erwartete" - Pass/Fail galten als fehlend.
FELDER["pass_kriterien"] += ["pass-kriterium", "pass kriterium"]
FELDER["pass_fail_kombiniert"] += ["maschinenanker / pass-fail-kriterien", "maschinenanker/pass-fail-kriterien",
                                   "pass-fail-kriterien", "pass/fail-kriterien"]
FELDER["fail_indikatoren"] += ["fail-kriterium", "fail kriterium", "fail-kriterien", "fail-indikator"]
FELDER["erwartete_hauptpruefungselemente"] += ["hauptprüfungselemente", "hauptpruefungselemente"]
# Ohne diese Felder ist ein Testfall fuer Stufe 5 nur eingeschraenkt bewertbar.
PFLICHT = ["testfall_id", "kurzbezeichnung", "eingabedaten", "erwartete_aktivierung",
           "erwartete_kernnormen", "erwartetes_ergebnis", "erwartete_eskalationsstufe",
           "erwartete_rechtsprechung", "erwartete_rueckfragen", "pass_kriterien",
           "fail_indikatoren"]

_ALIAS = {alias: kanon for kanon, aliase in FELDER.items() for alias in aliase}
# Der Fall-Kopf variiert in DREI Dimensionen (alle am 2026-09-04 gemessen), und
# jede einzelne liess den Parser vorher auf 0-1 Faelle zurueckfallen:
#   Ebene   `##` (die meisten) vs. `###` (COMPANY_CAR_DURING_LITIGATION)
#   Trenner `-` `–` `—` (bekannt) vs. `·` (FIXEDTERM.MAX_DURATION)
#             vs. `:` (RETIREMENT_NO_END_CLAUSE)
#   Kennung `TC-…-01`/`TF-…-01` vs. blosse Nummerierung "Testfall 1"
# Bei blosser Nummerierung wird KEINE Kennung erfunden: der Kopf liefert nur
# einen Platzhalter, die echte ID kommt aus dem Feld `Testfall-ID` im Block
# (setdefault in parse() laesst dem gelesenen Feld den Vortritt).
#   Praefix  keins vs. "Testfall: " VOR der Kennung (TRANSFER.INFORMATION_DUTY,
#            TRANSFER.TERMINATION_BAN, gemessen 2026-09-04) - ohne das optionale
#            Praefix erkannte der Parser dort NULL Faelle.
#   QS-Nummer statt Kennung: `## QS 10.1 - Betriebsstilllegung ...`
#            (RESEARCH.RESTRUCTURING.CASELAW_OVERVIEW und
#            GENERATE.RESTRUCTURING.INTEREST_SOCIAL_PLAN_CHECKLIST, gemessen
#            2026-09-08). Der Kopf traegt hier weder TC/TF noch das Wort
#            "Testfall", sondern eine Gliederungsnummer - der Parser fand NULL
#            Faelle, was wie eine leere Vorgabe aussah. Wie bei der blossen
#            Nummerierung gilt: der Kopf liefert nur einen Platzhalter, die
#            echte Kennung kommt aus dem Feld `Testfall-ID` im Block.
#   Praefix MIT Nummer und beliebigem Trenner vor der Kennung, gemessen
#            2026-09-17 an den Paketen des 17.09.-Builders:
#            `## QS-Testfall TC-EXTENSION-01: Positiv - ...` (FIXEDTERM.EXTENSION)
#            und `## QS-Testfall 1 - TC-FIXEDTERM_CHECK-01 · ...` (FIXEDTERM.CHECK).
#            Das alte Praefix verlangte "Testfall:" MIT Doppelpunkt, traf also
#            weder die eine noch die andere Form. Folge bei EXTENSION: NULL
#            Faelle - eine leer aussehende Vorgabe, die in Wahrheit fuenf
#            vollstaendige Faelle fuehrt. Bei CHECK matchte nur die
#            `nr`-Alternative, wodurch die echte Kennung in der
#            Kurzbezeichnung landete statt im Feld `testfall_id`.
#            Das Praefix ist jetzt optional-greedy; faellt danach keine Kennung,
#            loest das Backtracking auf die `nr`-Alternative zurueck, sodass die
#            blosse Nummerierung ("## Testfall 1 - ...") weiter traegt.
FALL_KOPF = re.compile(
    r"^(#{2,4})\s+(?:\d+[.)]\s*(?=(?:QS-)?(?:TC|TF|Testfall|QS)))?(?:(?:QS-)?Testfall\s*(?:\d+|[A-Z](?=\s))?\s*[:–\-—·]?\s*)?"
    r"(?:(?P<id>(?:TC|TF)[-\w]*\d)|(?:QS-)?Testfall\s+(?P<nr>\d+)"
    r"|QS\s+(?P<qs>\d+(?:\.\d+)*))"
    r"(?:\s*[–\-—·:]\s*(?P<kurz>.+?))?\s*$", re.M)
KOPF_ALLE = re.compile(r"^(#{1,6})\s+\S", re.M)


def _kanon(label):
    # Markdown-Auszeichnung gehoert zur Darstellung, nicht zum Feldnamen: in
    # Format A2 sind die Labels fett ("| **Eingabedaten** |"). Ohne das Strippen
    # wird KEIN Feld erkannt, der Fall faellt auf "prosa" zurueck und meldet
    # saemtliche Pflichtfelder als fehlend - ein Parser-Fehler, der wie ein
    # Mangel der Testfallvorgabe aussieht (Befund 2026-08-30 an
    # APPRENTICESHIP.MINIMUM_REMUNERATION: 9 von 10 Feldern unsichtbar).
    l = re.sub(r"[*`]", "", (label or "").strip())
    # Format D (gemessen 2026-09-10 an den drei AICOMPLIANCE-Skills): die
    # Feldnamen sind NUMMERIERT - "**1. Kurzsachverhalt:**", "**2. Eingabedaten:**",
    # "**4. Maschinenanker / Pass-Fail-Kriterien:**". Ohne das Abschneiden der
    # Gliederungsnummer trifft kein Alias, der Fall faellt auf "prosa" zurueck
    # und meldet ALLE Pflichtfelder als fehlend - inklusive `eingabedaten`, und
    # damit waere ein Testlauf mit LEERER Eingabe gefahren worden.
    l = re.sub(r"^\s*[•\-\*•]?\s*(?:\d{1,2}|[ivxIVX]{1,4})\.\s*", "", l)
    l = re.sub(r"\s+", " ", l.strip().strip(":").lower())
    return _ALIAS.get(l)


# "Pass: ... Fail: ..." in einer Zelle. Getrennt wird am ERSTEN eigenstaendigen
# "Fail:" - der Pass-Teil endet regelmaessig mit "... Eskalationsstufe E0."
# Zweite Trennform, gemessen 2026-09-10: die Zelle formuliert als SATZ
# ("Pass, wenn der Skill ... erkennt. Fail, wenn ..."). Nur mit dem Doppelpunkt
# blieb `fail_indikatoren` leer und tauchte als fehlendes Pflichtfeld auf -
# betroffen waren MISCONDUCT, MINIMUM_REMUNERATION, COMPANY_CAR und
# FIXEDTERM.MAX_DURATION. Der Trenner ist deshalb `[:,]`; getrennt wird
# weiterhin am ERSTEN eigenstaendigen "Fail", und ohne erkennbaren Schnitt
# bleibt `fail_indikatoren` leer statt geraten.
_PASS_FAIL = re.compile(r"^\s*pass\s*[:,]\s*(.*?)(?:\s|^)fail\s*[:,]\s*(.+)$", re.I | re.S)


def _teile_pass_fail(daten):
    """Loest das kombinierte Feld auf. Ohne erkennbaren Schnitt bleibt
    `fail_indikatoren` leer und taucht in `fehlende_felder` auf - lieber eine
    ehrliche Luecke als ein geratener Fail-Indikator."""
    roh = daten.pop("pass_fail_kombiniert", None)
    if not roh:
        return daten
    m = _PASS_FAIL.match(roh)
    if m:
        daten.setdefault("pass_kriterien", m.group(1).strip().rstrip(" .;"))
        daten.setdefault("fail_indikatoren", m.group(2).strip())
        return daten
    # Format F (03.10.2026, RESEARCH): Aufzaehlung "- Pass, wenn …" / "- Fail, wenn …" unter
    # "### 4. Maschinenanker / Pass-Fail-Kriterien" - beliebig viele Punkte je Seite.
    teile = re.findall(r"\b(Pass|Fail)\b\s*(?:,|:)?\s*(.*?)(?=\s-\s*(?:Pass|Fail)\b|\s(?:Pass|Fail),\s|$)", roh, re.S)
    pas = [t.strip(" -;") for k, t in teile if k == "Pass" and t.strip(" -;")]
    fal = [t.strip(" -;") for k, t in teile if k == "Fail" and t.strip(" -;")]
    if pas or fal:
        if pas:
            daten.setdefault("pass_kriterien", " | ".join(pas))
        if fal:
            daten.setdefault("fail_indikatoren", " | ".join(fal))
        daten.setdefault("maschinenanker", roh.strip())
    else:
        daten.setdefault("pass_kriterien", roh.strip())
    return daten


# Format C bildet mehrere kanonische Felder auf EINE Ueberschrift ab.
_KOMBI_ZIELE = {
    "kernnormen_rspr_kombiniert": ("erwartete_kernnormen", "erwartete_rechtsprechung"),
    "risiken_eskalation_kombiniert": ("erwartete_risiken", "erwartete_eskalationsstufe"),
}


# Aktenzeichen deutscher Gerichte und EuGH-Rechtssachen. Bewusst dasselbe enge
# Muster wie in eval_stufe5.py - ein weiteres wuerde hier anders zaehlen als dort.
_AZ_RE = re.compile(r"\b\d{1,2}\s?[A-Za-z]{1,5}\s?\d+/\d{2}\b|\bC-\d+/\d{2}\b")
_NORM_RE = re.compile(r"§§?\s*\d|\bArt(?:ikel|\.)\s*\d")


def _trenne_norm_rspr(roh):
    """(normen, rechtsprechung) aus einem kombinierten "Anker"-Text - oder
    (None, None), wenn die Trennung nicht EINDEUTIG ist.

    Der 17.09.-Builder fuehrt beides in einem Feld:
    "§ 14 Abs. 2 Satz 1 TzBfG; BAG 7 AZR 108/20." Bis hierher landete der ganze
    Text in beiden Zielfeldern (s. _verteile_kombiniert) - richtig, aber stumpf:
    der Soll-Ist-Abgleich prueft dann Normen gegen einen Text, der auch
    Entscheidungen enthaelt, und umgekehrt.

    Getrennt wird an Semikolon/Mittelpunkt und nur nach EINDEUTIGEN Merkmalen:
    ein Segment mit `§`/`Art.` ist eine Norm, eines mit Aktenzeichen oder
    EuGH-Nummer eine Entscheidung. **Traegt ein Segment beides oder keines von
    beiden, wird nicht geraten** - dann faellt die ganze Trennung aus und es
    bleibt beim bisherigen Kopieren in beide Felder. Anlass: Segmente wie
    "ANCH-BASE-01, ANCH-FORM-01/02" (FIXEDTERM.CHECK) sind Ankerkennungen des
    Skills und gehoeren in keines der beiden Felder.
    """
    if not roh:
        return None, None
    normen, rspr, unklar = [], [], 0
    for seg in re.split(r"[;·]|(?<=\.)\s+(?=[A-ZÄÖÜ§])", roh):
        seg = seg.strip().strip(".,; ")
        if not seg:
            continue
        ist_n, ist_r = bool(_NORM_RE.search(seg)), bool(_AZ_RE.search(seg))
        if ist_n and not ist_r:
            normen.append(seg)
        elif ist_r and not ist_n:
            rspr.append(seg)
        elif ist_n and ist_r:
            unklar += 1                      # ein Segment mit beidem: nicht schneiden
    if unklar or not (normen or rspr):
        return None, None
    return "; ".join(normen) or None, "; ".join(rspr) or None


def _verteile_kombiniert(daten):
    """Kombifelder auf ihre Zielfelder legen - getrennt, wo es eindeutig ist.

    Bis 17.09.2026 wurde ausnahmslos KOPIERT, weil ein Trennversuch mitten im
    Fliesstext raten wuerde. Das gilt weiterhin als Rueckfall - aber wo die
    Vorgabe sauber mit Semikolon segmentiert und jedes Segment eindeutig ist
    (`§`/`Art.` gegen Aktenzeichen), ist die Trennung eine Messung, keine
    Schaetzung. `kombinierte_felder` haelt in beiden Faellen fest, was passiert
    ist, damit der Soll-Ist-Abgleich weiss, was er vor sich hat.
    """
    vermerke = []
    for quelle, ziele in _KOMBI_ZIELE.items():
        roh = daten.pop(quelle, None)
        if not roh:
            continue
        if quelle == "kernnormen_rspr_kombiniert":
            n, r = _trenne_norm_rspr(roh)
            if n or r:
                if n:
                    daten.setdefault("erwartete_kernnormen", n)
                if r:
                    daten.setdefault("erwartete_rechtsprechung", r)
                vermerke.append("%s GETRENNT: Normen=%s, Rechtsprechung=%s"
                                % (quelle, "ja" if n else "keine genannt",
                                   "ja" if r else "keine genannt"))
                continue
        for ziel in ziele:
            daten.setdefault(ziel, roh)
        vermerke.append("%s -> %s (nicht eindeutig trennbar, Text in beiden Feldern)"
                        % (quelle, "+".join(ziele)))
    if vermerke:
        daten["kombinierte_felder"] = vermerke
    return daten


TRENNER = re.compile(r"^\s*-{3,}\s*$", re.M)   # eigenstaendige Trennlinie, NICHT |---|---|


def _bloecke(text):
    """(id, kurzbezeichnung, rohtext, ebene) je Testfall - Format-unabhaengig.

    Der Block endet am naechsten Testfall-Kopf, an der ersten eigenstaendigen
    Trennlinie `---` ODER an der naechsten Ueberschrift GLEICHER/HOEHERER Ebene.
    Ohne diesen Schnitt lief der letzte Feldwert in den Folgeabschnitt hinein
    (beobachtet 2026-08-19: die Eskalationsstufe von TF-27-05 endete mit
    "--- ## Automatisierbare QS-Pruefpunkte").

    Die Ebene muss mitlaufen, weil sie beides steuert: bei Format C sind die
    FELDER tiefere Ueberschriften (`###` unter `##`) und duerfen nicht schneiden,
    waehrend bei einem `###`-Fallkopf das nachfolgende `## Empfehlung...` sehr
    wohl schneiden muss.
    """
    treffer = list(FALL_KOPF.finditer(text))
    for i, m in enumerate(treffer):
        ebene = len(m.group(1))
        ende = treffer[i + 1].start() if i + 1 < len(treffer) else len(text)
        for k in KOPF_ALLE.finditer(text, m.end(), ende):
            if len(k.group(1)) <= ebene:
                ende = k.start()
                break
        roh = text[m.end():ende]
        schnitt = TRENNER.search(roh)
        if schnitt:
            roh = roh[:schnitt.start()]
        tid = (m.group("id")
               or ("Testfall %s" % m.group("nr") if m.group("nr") else None)
               or ("QS %s" % m.group("qs")))
        yield tid, (m.group("kurz") or "").strip(), roh, ebene


def _aus_tabelle(roh):
    """Format A: Zeilen '| Feld | Inhalt |'."""
    daten = {}
    for zeile in roh.splitlines():
        if not zeile.strip().startswith("|"):
            continue
        spalten = [s.strip() for s in zeile.strip().strip("|").split("|")]
        if len(spalten) < 2 or set(spalten[0]) <= set("-: "):
            continue
        k = _kanon(spalten[0])
        if k and spalten[1]:
            daten.setdefault(k, spalten[1])
    return daten


def _aus_prosa(roh):
    """Format B: Absaetze '**Feld:** Wert' (Wert darf mehrzeilig sein).

    Zweiter Durchgang seit 2026-09-10 (Format D): dieselbe Bauform fuehrt die
    Erwartungen als AUFZAEHLUNG unter einer Sammelueberschrift -
    "**3. Erwartung an den Skill:**" mit "• Erwartete Kernnormen: …",
    "• Erwartete Eskalationsstufe: Yellow." Der erste Durchgang sieht davon
    nichts, weil kein `**Feld:**` davorsteht; ohne den zweiten fehlen genau die
    Felder, auf denen der Soll-Ist-Abgleich aufsetzt.
    """
    daten = {}
    muster = re.compile(r"\*\*([^*]{3,60}?):\*\*\s*(.*?)(?=\n\s*\*\*[^*]{3,60}?:\*\*|\Z)", re.S)
    for m in muster.finditer(roh):
        k = _kanon(m.group(1))
        if k:
            wert = re.sub(r"\s*\n\s*", " ", m.group(2)).strip()
            if wert:
                daten.setdefault(k, wert)
    # Aufzaehlungsfelder: '• Label: Wert' bzw. '- Label: Wert'. Der Wert endet
    # am naechsten Aufzaehlungszeichen oder an der naechsten fetten Ueberschrift.
    bullet = re.compile(r"(?m)^\s*[••\-\*]\s*\*{0,2}([A-ZÄÖÜ][^:\n]{3,60}?)\*{0,2}\s*:\s*"
                        r"(.*?)(?=\n\s*[••\-\*]\s|\n\s*\*\*|\n\s*#|\Z)", re.S)
    for m in bullet.finditer(roh):
        k = _kanon(m.group(1))
        if k:
            wert = re.sub(r"\s*\n\s*", " ", m.group(2)).strip()
            if wert:
                daten.setdefault(k, wert)
    return daten


def _aus_ueberschriften(roh, ebene):
    """Format C: jedes Feld ist eine tiefere Ueberschrift, der Wert steht darunter.

    Gemessen 2026-09-04 an FIXEDTERM.MAX_DURATION (`## TC-MAX_DURATION-01 · …`
    mit `### Ausfuehrlicher Sachverhalt`, `### Erwartete Kernnormen und
    Rechtsprechung`, …). Vorher erkannte weder die Tabellen- noch die
    Prosa-Lesart ein Feld.
    """
    daten = {}
    kopf = re.compile(r"^#{%d,6}\s+(.+?)\s*$" % (ebene + 1), re.M)
    treffer = list(kopf.finditer(roh))
    for i, m in enumerate(treffer):
        ende = treffer[i + 1].start() if i + 1 < len(treffer) else len(roh)
        k = _kanon(m.group(1))
        if k:
            wert = re.sub(r"\s*\n\s*", " ", roh[m.end():ende]).strip()
            if wert:
                daten.setdefault(k, wert)
    return daten


def _aus_freitext(roh):
    """Letzte Instanz: der Sachverhalt steht als UNBENANNTER Absatz im Block.

    Gemessen 2026-09-17 an FIXEDTERM.CHECK (Paket des 17.09.-Builders): zwischen
    `**Kategorie:** positiv` und `**Erwartete Aktivierung:** …` steht die
    Sachverhaltserzaehlung als reiner Absatz OHNE jedes Label. Alle drei
    Lesarten finden dort ihre Felder, nur `eingabedaten` bleibt leer - also
    genau die teure Variante der bekannten Falle: nicht "Feld fehlt" (das faellt
    auf), sondern ein vollstaendig gefuellter Fall mit leerem Pflichtfeld. Ein
    Lauf damit haette den Agenten mit LEERER Eingabe befragt und dessen Antwort
    bewertet.

    Genommen wird der laengste zusammenhaengende Absatz, der weder Tabellenzeile
    noch Ueberschrift noch beschriftetes Feld ist. Die Mindestlaenge haelt
    Einwuerfe wie "positiv" oder eine Trennzeile draussen. Greift NUR, wenn
    weder `eingabedaten` noch `kurzsachverhalt` etwas hergeben - geraten wird
    also nichts, solange die Vorgabe den Sachverhalt selbst benennt.
    """
    absaetze, aktuell = [], []
    for zeile in roh.splitlines():
        s = zeile.strip()
        ist_feld = (not s or s.startswith(("|", "#", ">"))
                    or re.match(r"^\s*[••\-\*\d]+[\.\)]?\s", zeile)
                    or re.match(r"^\*\*[^*]{3,60}?:\*\*", s))
        if ist_feld:
            if aktuell:
                absaetze.append(" ".join(aktuell))
                aktuell = []
        else:
            aktuell.append(s)
    if aktuell:
        absaetze.append(" ".join(aktuell))
    # Nie eine Soll-Vorgabe als Eingabe (03.10.2026: "Richtiges Ergebnis: ..." ging so an den Agenten).
    soll = re.compile(r"^(richtiges ergebnis|erwartet|pass|fail|soll|der skill (darf|muss|soll))", re.I)
    kandidaten = [a for a in absaetze if len(a) >= 200 and not soll.match(a)]
    if not kandidaten:
        return {}
    return {"eingabedaten": max(kandidaten, key=len)}


_AMPEL_RE = re.compile(r"\b(Gr[üu]n|Green|Gelb|Yellow|Rot|Red|Orange)\b", re.I)
_AMPEL_KANON = {"grün": "Grün", "gruen": "Grün", "green": "Grün",
                "gelb": "Gelb", "yellow": "Gelb", "orange": "Orange",
                "rot": "Rot", "red": "Rot"}


def _ampel_aus_ergebnis(erg):
    """(stufe, alle) - die erwartete Eskalationsstufe aus dem Ergebnisfeld.

    Gemessen 17.09.2026 an 10 Testfaellen zweier Skills: **10 von 10** nennen die
    Stufe woertlich am Anfang des Ergebnisses ("Gruen: E-EX-QUALIFIES", "Rot fuer
    den Verlaengerungsbegriff"). Das ist Extraktion, keine Herleitung - der Wert
    stammt aus der Vorgabe, nur unter einem anderen Etikett.

    🔴 **Mehrwertige Faelle werden NICHT auf einen Wert verkuerzt.** Genau daran
    scheitert die naheliegende Regel "nimm die erste":
      CHECK tc03 "Gelb: … ; Green erst nach Beleg"          -> die ERSTE gilt
      CHECK tc05 "Hauptampel Gruen … Gesamtkommunikation Gelb" -> die LETZTE gilt
    Welche es ist, entscheidet der Satzbau, nicht die Position. Nennt das Feld
    mehr als eine Stufe, bleibt `stufe` deshalb leer und `alle` traegt die Menge;
    der Abgleich prueft dann gegen die Menge statt gegen einen geratenen Wert.
    """
    if not erg:
        return None, []
    roh = [m.group(1) for m in _AMPEL_RE.finditer(erg)]
    alle, gesehen = [], set()
    for r in roh:                                   # Reihenfolge erhalten, Dubletten raus
        k = _AMPEL_KANON.get(r.lower())
        if k and k not in gesehen:
            gesehen.add(k)
            alle.append(k)
    if len(alle) == 1:
        return alle[0], alle
    return None, alle


_PASS_NUR_WENN = re.compile(r"(?is)\bPASS\b\s*(?:nur)?\s*[,:]?\s*wenn\s+(.{10,})$")


def _fail_aus_pass(pk):
    """Abgeleiteter Fail-Indikator aus einem 'PASS nur, wenn …'-Kriterium.

    Ein "PASS nur, wenn X" ist ein NOTWENDIGES Kriterium - seine Verletzung ist
    der Fail. Die Umkehrung ist Logik, keine Erfindung; die FORMULIERUNG kommt
    aber von uns. Deshalb landet der Wert in einem eigenen Feld
    `fail_indikator_abgeleitet` und NIE in `fail_indikatoren`: in einem Protokoll
    darf kein Soll stehen, das der Builder nie geschrieben hat.

    Greift nur bei der ausdruecklichen Form "PASS nur, wenn" - ein blosses
    "Pass: …" beschreibt den Erfolgsfall, ohne ihn zur Bedingung zu machen, und
    daraus folgt kein Fail-Indikator.
    """
    if not pk:
        return None
    m = _PASS_NUR_WENN.search(pk)
    if not m:
        return None
    return ("Abgeleitet aus dem Pass-Kriterium (notwendige Bedingung, Formulierung "
            "von der Auswertung, nicht aus der Vorgabe): Fail, wenn NICHT gilt - "
            + m.group(1).strip().rstrip("."))


def _qs_pruefpunkte(text):
    """Abschnitt 'Automatisierbare QS-Pruefpunkte' -> Stichpunkte je Testfall-ID."""
    m = re.search(r"^##\s+Automatisierbare QS-Pr(?:ü|ue)fpunkte\s*$(.*?)(?=^##\s|\Z)",
                  text, re.M | re.S)
    if not m:
        return {}, []
    punkte = [z.strip(" -*\t") for z in m.group(1).splitlines() if z.strip().startswith(("-", "*"))]
    je_fall, allgemein = {}, []
    for p in punkte:
        # Nur Stichpunkte, die MIT einer Testfall-ID beginnen, sind fallbezogen.
        # Saetze wie "Dokument enthaelt genau fuenf Testfaelle TF-27-01 bis TF-27-05"
        # nennen IDs nur beilaeufig und gehoeren nicht an einen einzelnen Fall.
        kopf = re.match(r"^((?:TC|TF)[-\w]*\d)\b", p.strip())
        if kopf:
            je_fall.setdefault(kopf.group(1), []).append(p)
        else:
            allgemein.append(p)
    return je_fall, allgemein


def parse(text):
    qs_je_fall, qs_allgemein = _qs_pruefpunkte(text)
    faelle = []
    for tid, kurz, roh, ebene in _bloecke(text):
        # Die Lesart mit den MEISTEN erkannten Feldern gewinnt - nicht die ID,
        # nicht die Reihenfolge. Ein fester Vorrang hat die Formate A2 und C
        # jeweils auf die falsche Lesart fallen lassen.
        kandidaten = {"tabelle": _aus_tabelle(roh), "prosa": _aus_prosa(roh),
                      "ueberschriften": _aus_ueberschriften(roh, ebene)}
        format_ = max(kandidaten, key=lambda k: len(kandidaten[k]))
        if not kandidaten[format_]:
            format_ = "prosa"
        # Die Lesarten treten NICHT mehr nur gegeneinander an, sondern werden
        # vereinigt - die Gewinner-Lesart hat Vorrang, fehlende Felder kommen aus
        # den anderen. Anlass (2026-09-10, AICOMPLIANCE.HEALTH_SAFETY und
        # .RESTRUCTURING_TRIGGER): dieselbe Datei fuehrt die grossen Bloecke als
        # UEBERSCHRIFTEN ("### 1. Kurzsachverhalt", "### 2. Eingabedaten") und die
        # Erwartungen als AUFZAEHLUNG ("- Erwartete Kernnormen: …"). Beim reinen
        # Wettbewerb gewann die Prosa-Lesart mit den vielen Aufzaehlungsfeldern -
        # und `eingabedaten` blieb LEER, weil es nur als Ueberschrift existiert.
        # Ein Testlauf waere dann mit leerer Eingabe gefahren worden.
        zusammen = dict(kandidaten[format_])
        for andere in kandidaten.values():
            for k, v in andere.items():
                if v and not zusammen.get(k):
                    zusammen[k] = v
        # Erst wenn keine Lesart den Sachverhalt BENANNT gefunden hat, greift der
        # unbenannte Absatz (s. _aus_freitext). Die Reihenfolge ist wesentlich:
        # ein beschriftetes Feld hat immer Vorrang vor der Absatz-Heuristik.
        if not zusammen.get("eingabedaten") and not zusammen.get("kurzsachverhalt"):
            frei = _aus_freitext(roh)
            if frei:
                zusammen.update(frei)
                zusammen["eingabedaten_aus_freitext"] = True
        daten = _verteile_kombiniert(_teile_pass_fail(zusammen))
        # Abgeleitete Felder, NACH allen Lesarten: ein echtes Feld der Vorgabe
        # hat immer Vorrang, abgeleitet wird nur, wo nichts steht.
        if not daten.get("erwartete_eskalationsstufe"):
            stufe, alle = _ampel_aus_ergebnis(daten.get("erwartetes_ergebnis"))
            if stufe:
                daten["erwartete_eskalationsstufe"] = stufe
                daten["eskalationsstufe_abgeleitet"] = "aus 'Erwartetes Ergebnis'"
            elif alle:
                daten["eskalationsstufe_mehrwertig"] = alle
        if not daten.get("fail_indikatoren"):
            fi = _fail_aus_pass(daten.get("pass_kriterien"))
            if fi:
                daten["fail_indikator_abgeleitet"] = fi
        daten.setdefault("testfall_id", tid)
        daten.setdefault("kurzbezeichnung", kurz)
        daten["format"] = format_
        daten["qs_pruefpunkte"] = qs_je_fall.get(tid, [])
        daten["fehlende_felder"] = [f for f in PFLICHT if not daten.get(f)]
        faelle.append(daten)
    return {
        "anzahl": len(faelle),
        "format": (faelle[0]["format"] if faelle else None),
        "qs_pruefpunkte_allgemein": qs_allgemein,
        "testfaelle": faelle,
    }


if __name__ == "__main__":
    pfad = sys.argv[1]
    ergebnis = parse(open(pfad, encoding="utf-8").read())
    print("Datei: %s" % pfad)
    print("Format: %s | Testfaelle: %d" % (ergebnis["format"], ergebnis["anzahl"]))
    for f in ergebnis["testfaelle"]:
        print("\n%s  %s" % (f["testfall_id"], f.get("kurzbezeichnung", "")[:60]))
        print("   vorhandene Felder: %d | fehlend: %s"
              % (sum(1 for k in FELDER if f.get(k)), ", ".join(f["fehlende_felder"]) or "-"))
        for k in ("erwartete_eskalationsstufe", "erwartete_kernnormen", "erwartete_rechtsprechung",
                  "fail_indikatoren"):
            if f.get(k):
                print("   %-28s %s" % (k + ":", str(f[k])[:88]))
        if f["qs_pruefpunkte"]:
            print("   QS-Pruefpunkte (Ersatz fuer Pass/Fail): %s" % f["qs_pruefpunkte"][0][:80])
    if ergebnis["qs_pruefpunkte_allgemein"]:
        print("\nQS-Pruefpunkte ohne Fallbezug: %d" % len(ergebnis["qs_pruefpunkte_allgemein"]))
    if "--json" in sys.argv:
        ziel = sys.argv[sys.argv.index("--json") + 1]
        json.dump(ergebnis, open(ziel, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("\ngeschrieben:", ziel)
