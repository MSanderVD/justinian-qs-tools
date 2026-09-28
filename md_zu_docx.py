"""DOC_TXT (Markdown, wie ihn PhiBox aus der .docx erzeugt) zurueck in eine .docx mit Ueberschriften,
Absaetzen und Tabellen - damit der Stufe-1-Extraktor in der Builder-Sandbox laufen kann, in der beim
Paketbau KEINE .docx liegt (gemessen 25.09.2026: /workspace/skill startet leer).

  python3 md_zu_docx.py <eingabe.md> <ausgabe.docx>
"""
import re
import sys

from docx import Document

TAB_ZEILE = re.compile(r"^\s*\|.*\|\s*$")
TRENN = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")
KOPF = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")


def zellen(zeile):
    z = zeile.strip()
    if z.startswith("|"):
        z = z[1:]
    if z.endswith("|"):
        z = z[:-1]
    teile, akt, esc = [], "", False
    for ch in z:
        if esc:
            akt += ch; esc = False
        elif ch == "\\":
            esc = True
        elif ch == "|":
            teile.append(akt); akt = ""
        else:
            akt += ch
    teile.append(akt)
    return [entmd(t.strip()) for t in teile]


def entmd(t):
    t = re.sub(r"<br\s*/?>", "\n", t, flags=re.I)
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t)
    t = re.sub(r"(?<!\w)\*(.+?)\*(?!\w)", r"\1", t)
    t = re.sub(r"`([^`]*)`", r"\1", t)
    t = re.sub(r"\[([^\]]+)\]\((https?://[^)]+)\)", r"\1 - \2", t)
    return t.replace("&nbsp;", " ")


def zelle_setzen(zelle, text):
    teile = text.split("\n")
    zelle.text = teile[0]
    for t in teile[1:]:
        zelle.add_paragraph(t)


def baue(md_text, ziel):
    d = Document()
    zeilen = md_text.replace("\r\n", "\n").split("\n")
    i = 0
    while i < len(zeilen):
        z = zeilen[i]
        m = KOPF.match(z)
        if m:
            d.add_heading(entmd(m.group(2)), level=min(len(m.group(1)), 9))
            i += 1
            continue
        if TAB_ZEILE.match(z) and i + 1 < len(zeilen) and TRENN.match(zeilen[i + 1]):
            kopf = zellen(z)
            koerper = []
            i += 2
            while i < len(zeilen) and TAB_ZEILE.match(zeilen[i]):
                koerper.append(zellen(zeilen[i]))
                i += 1
            breite = max([len(kopf)] + [len(r) for r in koerper])
            t = d.add_table(rows=1 + len(koerper), cols=breite)
            for r, reihe in enumerate([kopf] + koerper):
                for c in range(breite):
                    zelle_setzen(t.rows[r].cells[c], reihe[c] if c < len(reihe) else "")
            continue
        if z.strip() and not re.match(r"^\s*(-{3,}|\*{3,})\s*$", z):
            text = re.sub(r"^\s*([-*+]|\d+\.)\s+", "", z)
            d.add_paragraph(entmd(text))
        i += 1
    d.save(ziel)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__); sys.exit(2)
    baue(open(sys.argv[1], encoding="utf-8").read(), sys.argv[2])
    print("OK:", sys.argv[2])
