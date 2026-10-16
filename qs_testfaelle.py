"""QS-Paket: Testfaelle eines JUS-Word-Skills deterministisch aus der .docx - ohne Sprachmodell.

Laeuft in der PhiBox-Sandbox (Agent "JUS QS-Paket Builder"). Quelle ist die .docx des Skill-Pakets unter
/skills/<name_normalized>/<SKILL-ID>.docx (Verzeichnis /skills/manifest.json). Ergebnis:
  testfaelle.md    der Testfallteil der Word-Datei WOERTLICH als Markdown (Ueberschriften, Absaetze, Listen, Tabellen)
  testfaelle.json  dieselben Faelle maschinenlesbar (Parser = parse_testfaelle.py aus Stufe 5) + Herkunft

  python3 qs_testfaelle.py <SKILL-ID> <zielordner> [--skills /skills]
  python3 qs_testfaelle.py --docx <datei.docx> <zielordner>      (Pruefung ausserhalb der Sandbox)

Welcher Teil: "Teil 2" (Autoren-QS). Liefert er keinen Testfall, der Teil, dessen Ueberschrift "Testf" oder "QS"
nennt (NONCOMPETE: "Teil 3: Fuenf vollstaendige QS-Testfaelle").
Exit 0 = geschrieben, 2 = Skill/.docx nicht gefunden, 3 = kein Testfallteil, 4 = 0 Testfaelle oder Testfall ohne
Eingabe (dann wird NICHTS geschrieben - lieber kein QS-Paket als ein falsches)."""
import datetime, hashlib, json, os, re, sys

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import parse_testfaelle as PT

VERSION = "qs_testfaelle 1.0 (16.10.2026)"
TEIL = re.compile(r"^\s*Teil\s+(\d)\b", re.I)
# Kopf eines Testfalls: Kennung TC-/TF-..., "Testfall N", "QS N.N", "QS-Testfall A - TC-...", auch "1. TC-..."
FALLKOPF_TEXT = re.compile(r"^(?:\d+[.)]\s*)?(?:QS-)?(?:Testfall\s*(?:(?:\d+|[A-Z])\b)?\s*[:–\-—·]?\s*)?"
                           r"(?:(?:TC|TF)-[-\w]*\d\b|Testfall\s+\d+\b|QS\s+\d+(?:\.\d+)*\b)")
# Zeilen einer Falltabelle, die NICHT zum Sachverhalt gehoeren (Soll, Kennung, Zweck)
KEIN_SACHVERHALT = re.compile(r"(?i)^(erwart|pass|fail|soll|ziel|qs-|kurzbezeichnung|testfall|maschinen|kategorie|"
                              r"richtig|bewertung|hinweis)")


def _ebene(p):
    s = (p.style.name or "") if p.style is not None else ""
    m = re.match(r"(?:Heading|Überschrift)\s*(\d)", s, re.I)
    return int(m.group(1)) if m else (1 if s.lower() == "title" else 0)


def _absatz(p):
    t = p.text.strip()
    if not t:
        return ""
    # fett gesetztes Etikett am Anfang ("Eingabedaten: ...") als **Etikett:** erhalten - der Parser liest es so
    fett, rest, im_fett = "", [], True
    for r in p.runs:
        if im_fett and r.bold and r.text:
            fett += r.text
        else:
            im_fett = False
            rest.append(r.text)
    if fett.strip().endswith(":") and 2 < len(fett.strip()) < 70:
        t = "**" + fett.strip() + "** " + "".join(rest).strip()
    liste = p._p.pPr is not None and p._p.pPr.numPr is not None
    stil = (p.style.name or "").lower() if p.style is not None else ""
    if liste or stil.startswith(("list", "liste", "aufz")):
        return "- " + t
    return t


def _zeilen(t):
    out = []
    for r in t.rows:
        z = [re.sub(r"\s*\n\s*", " ", c.text.strip()).replace("|", "/") for c in r.cells]
        z = list(dict.fromkeys(z)) if len(set(z)) < len(z) else z     # verbundene Zellen nur einmal
        if any(z):
            out.append(z)
    return out


def _tabelle_md(zeilen):
    # Einspaltige Tabelle mit Etikett im Zellentext ("Kurzsachverhalt Eine Logistik-GmbH ...", FIXEDTERM.EXTENSION):
    # als Absatz "**Kurzsachverhalt:** ..." ausgeben, damit der Parser das Feld erkennt. Text bleibt woertlich.
    if zeilen and all(len(z) == 1 for z in zeilen):
        out = []
        for z in zeilen:
            w = z[0].split(" ")
            for n in range(min(4, len(w) - 1), 0, -1):
                if PT._kanon(" ".join(w[:n])):
                    out.append("**%s:** %s" % (" ".join(w[:n]).rstrip(":"), " ".join(w[n:]))); break
            else:
                out.append(z[0])
        return out
    md = ["| " + " | ".join(z) + " |" for z in zeilen]
    if md:
        md.insert(1, "|" + "---|" * max(len(zeilen[0]), 1))
    return md


def _tabelle_folgt(kinder, i, d):
    """Folgt auf den Absatz (nach Leerabsaetzen) unmittelbar eine Tabelle? Nur dann ist er ein Fallkopf."""
    for el in kinder[i + 1:i + 4]:
        if el.tag.endswith("}tbl"):
            return True
        if el.tag.endswith("}p") and Paragraph(el, d).text.strip():
            return False
    return False


def teile(pfad):
    """{nr: (ueberschrift, markdown)} fuer jeden 'Teil N' der Word-Datei."""
    d = docx.Document(pfad)
    kinder = list(d.element.body.iterchildren())
    erg, nr, aus, basis, fall, zuletzt, letzte_ueberschrift_id = {}, None, [], 1, None, 1, None
    etikett_bauform = False
    offenes_etikett = None   # Etikett als eigener Absatz ("Eingabedaten"), Wert im naechsten (MASSDISMISSAL.ERROR_CONSEQUENCES)
    verbraucht = True    # True = der letzte Fallkopf hat schon seine Falltabelle; eine weitere ID-Tabelle braucht einen eigenen Kopf

    def ende():
        if nr is not None:
            erg[nr] = (aus[0][2:] if aus else "", "\n".join(aus))

    for i, el in enumerate(kinder):
        if el.tag.endswith("}p"):
            p = Paragraph(el, d); e = _ebene(p); t = p.text.strip()
            m = TEIL.match(t)
            if e and m:
                ende()
                nr, aus, basis, fall, zuletzt = m.group(1), ["# " + t], e, None, 1
                continue
            if nr is None:
                continue
            if not e and FALLKOPF_TEXT.match(t) and len(t) < 220 and _tabelle_folgt(kinder, i, d):
                L = fall or zuletzt + 1          # Fallkopf als normaler Absatz (COLLECTIVE.AVE)
                fall = L; aus.append("#" * min(6, max(2, L)) + " " + t); letzte_ueberschrift_id = t; verbraucht = False; etikett_bauform = False
                continue
            if e:
                L = min(6, max(2, e - basis + 1))
                kid = re.search(r"(?:TC|TF)-[-\w]*\d", t)
                if FALLKOPF_TEXT.match(t) and fall and L > fall and kid and letzte_ueberschrift_id \
                        and kid.group(0) in letzte_ueberschrift_id:
                    aus.append(t)                # Unterueberschrift wiederholt nur die Kennung (COLLECTIVE.AUEG_ORDER)
                    continue
                if FALLKOPF_TEXT.match(t):
                    fall = L; letzte_ueberschrift_id = t; verbraucht = False; etikett_bauform = False
                elif fall and L <= fall and PT._kanon(t):
                    L = fall + 1                 # "### Eingabedaten" neben "### TC-..." (OFFBOARDING)
                elif fall and L <= fall:
                    fall = None
                zuletzt = L
                aus.append("#" * L + " " + t)
            else:
                if t and len(t) < 70 and PT._kanon(t) and not t.endswith(":") and \
                        (PT._kanon(t) == "testfall_id" or etikett_bauform):
                    offenes_etikett = t; etikett_bauform = True    # nur in Faellen, die mit "Testfall-ID" als Absatz beginnen
                    continue
                z = _absatz(p)
                if z and offenes_etikett:
                    z = "**%s:** %s" % (offenes_etikett, z); offenes_etikett = None
                if z:
                    aus.append(z)
        elif el.tag.endswith("}tbl") and nr is not None:
            offenes_etikett = None
            zl = _zeilen(Table(el, d))
            # Falltabelle ohne Ueberschrift, erste Zeile "Testfall-ID | TC-..." (RETIREMENT_NO_END_CLAUSE)
            idz = next((z for z in zl[:2] if len(z) >= 2 and PT._kanon(z[0]) == "testfall_id" and FALLKOPF_TEXT.match(z[1])), None)
            if idz and verbraucht:
                L = fall or zuletzt + 1
                fall = L; aus.append("#" * min(6, max(2, L)) + " " + idz[1]); letzte_ueberschrift_id = idz[1]
            if idz:
                verbraucht = True
            aus += [""] + _tabelle_md(zl) + [""]
    ende()
    return erg


def _sachverhalt_aus_zeilen(md, tid):
    """Fall ohne Feld 'Eingabedaten', Sachverhalt auf mehrere Tabellenzeilen verteilt (TRAVEL_EXPENSE.REIMBURSEMENT_CLAIM:
    'Beteiligte und Kontext', 'Reise und Anlass', ...): diese Zeilen WOERTLICH als 'Etikett: Wert' zusammensetzen.
    Nur Zeilen, die kein bekanntes Feld und keine Soll-/Kennungszeile sind."""
    for _tid, _kurz, roh, _eb in PT._bloecke(md):
        if _tid != tid:
            continue
        teile_, soll_tabelle = [], False
        for zeile in roh.splitlines():
            if not zeile.strip().startswith("|"):
                soll_tabelle = False
                continue
            s = [x.strip() for x in zeile.strip().strip("|").split("|")]
            if s[0].lower() in ("prüffeld", "feld", "inhalt"):
                soll_tabelle = bool(re.search(r"(?i)soll|erwart", " ".join(s[1:])))   # Kopf einer Soll-Tabelle
                continue
            if soll_tabelle or len(s) < 2 or set(s[0]) <= set("-: "):
                continue
            if PT._kanon(s[0]) or KEIN_SACHVERHALT.match(s[0]):
                continue
            teile_.append("%s: %s" % (s[0], " ".join(s[1:])))
        return "\n".join(teile_)
    return ""


def testfallteil(pfad):
    """(markdown, ergebnis) - Teil 2, sonst der Teil mit 'Testf'/'QS' in der Ueberschrift."""
    t = teile(pfad)
    reihenfolge = ["2"] + [n for n, (u, _) in sorted(t.items()) if n != "2" and re.search(r"(?i)testf|\bQS\b", u)]
    for n in reihenfolge:
        if n not in t:
            continue
        md = t[n][1]
        erg = PT.parse(md)
        if erg["anzahl"]:
            for f in erg["testfaelle"]:
                if not (f.get("eingabedaten") or f.get("kurzsachverhalt")):
                    sv = _sachverhalt_aus_zeilen(md, f.get("testfall_id"))
                    if sv:
                        f["eingabedaten"] = sv
                        f["eingabedaten_aus_zeilen"] = True
                        f["fehlende_felder"] = [x for x in f.get("fehlende_felder", []) if x != "eingabedaten"]
            erg["teil"] = n
            return md, erg
    return (t["2"][1] if "2" in t else ""), {"anzahl": 0, "testfaelle": []}


def kopf(pfad):
    """Skill-ID, Version, Letzte Aktualisierung aus der Kopftabelle des Skills."""
    out = {}
    for t in docx.Document(pfad).tables[:6]:
        for r in t.rows:
            z = [c.text.strip() for c in r.cells]
            if len(z) >= 2:
                k = z[0].lower()
                if k in ("skill-id", "version", "letzte aktualisierung", "status") and k not in out:
                    out[k] = z[-1]
    return out


def docx_finden(skill_id, skills_root):
    m = json.load(open(os.path.join(skills_root, "manifest.json"), encoding="utf-8"))
    for s in m.get("skills", []):
        if s.get("name") == skill_id:
            p = os.path.join(s["mount_path"], skill_id + ".docx")
            if os.path.exists(p):
                return p, s
            kand = [f for f in os.listdir(s["mount_path"]) if f.lower().endswith(".docx")]
            if len(kand) == 1:
                return os.path.join(s["mount_path"], kand[0]), s
    return None, None


def bauen(pfad, ziel, skill_id=None, paket=None):
    md, erg = testfallteil(pfad)
    if not md:
        print("FEHLER: kein Testfallteil in", pfad); return 3
    leer = [f.get("testfall_id") for f in erg["testfaelle"] if not (f.get("eingabedaten") or f.get("kurzsachverhalt"))]
    if not erg["anzahl"] or leer:
        print("FEHLER: Testfaelle %d, ohne Eingabedaten: %s" % (erg["anzahl"], leer or "-")); return 4
    k = kopf(pfad)
    sha = hashlib.sha256(open(pfad, "rb").read()).hexdigest()
    herkunft = {"skill_id": skill_id or k.get("skill-id"), "version": k.get("version"),
                "letzte_aktualisierung": k.get("letzte aktualisierung"), "quelle": os.path.basename(pfad),
                "teil": erg.get("teil"), "docx_sha256": sha, "skill_paket_id": (paket or {}).get("id"),
                "erzeugt_von": VERSION,
                "erzeugt_am": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")}
    os.makedirs(ziel, exist_ok=True)
    kopfzeilen = ["% QS-Paket: Testfaelle woertlich aus Teil %s der Word-Datei; nicht fuer Q&A-Agenten." % erg.get("teil"),
                  "% Quelle: %s | Version %s | sha256 %s" % (herkunft["quelle"], herkunft["version"], sha[:12]), ""]
    open(os.path.join(ziel, "testfaelle.md"), "w", encoding="utf-8").write("\n".join(kopfzeilen) + md + "\n")
    erg["herkunft"] = herkunft
    json.dump(erg, open(os.path.join(ziel, "testfaelle.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("OK: %d Testfaelle (%s) | Teil %s | Version %s | sha256 %s" % (
        erg["anzahl"], ", ".join(str(f.get("testfall_id")) for f in erg["testfaelle"]), erg.get("teil"),
        herkunft["version"], sha[:12]))
    return 0


if __name__ == "__main__":
    a = sys.argv[1:]
    if a[:1] == ["--docx"]:
        sys.exit(bauen(a[1], a[2]))
    root = a[a.index("--skills") + 1] if "--skills" in a else "/skills"
    pfad, paket = docx_finden(a[0], root)
    if not pfad:
        print("FEHLER: kein Skill-Paket", a[0], "unter", root); sys.exit(2)
    sys.exit(bauen(pfad, a[1], a[0], paket))
