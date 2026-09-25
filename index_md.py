"""Rechtsgrundlagen-Index (Stufe 1 der Justinian-QS-Pipeline) als Markdown-Abschnitt fuer
references/normen-rechtsprechung.md eines PhiBox-Skill-Pakets.

Laeuft in der Builder-Sandbox (PhiBox AVM) und lokal gleich. Der Index ist das Ergebnis von
extract_rechtsgrundlagen_deterministisch.process() - dieselbe Funktion, die den QS-Index baut.

  python3 index_md.py <skill.docx>                                  Abschnitt auf stdout
  python3 index_md.py <skill.docx> --voranstellen <quelle.md> <ziel.md>
        schreibt <ziel.md> = <quelle.md> ohne alten Index-Abschnitt, mit neuem Index direkt
        unter der ersten Ueberschrift (fehlt eine, ganz oben). quelle und ziel duerfen gleich sein.
Exitcode 0 = geschrieben, 2 = Aufruffehler, 3 = keine Norm und keine Entscheidung gefunden
(dann wird NICHTS geschrieben - ein leerer Index waere ein stiller Messfehler).
"""
import os
import re
import sys

HIER = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HIER)
import extract_rechtsgrundlagen_deterministisch as s1  # noqa: E402

VERSION = "v1"
KOPF = "## Rechtsgrundlagen-Index (Pflichtquellen)"
NORM_MARKER = re.compile(r"(§|Art\.?\s*\d|Artikel\s+\d)")


def zelle(t):
    return re.sub(r"\s+", " ", str(t or "")).replace("|", "\\|").strip()


def datum(iso):
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", iso or "")
    return f"{m.group(3)}.{m.group(2)}.{m.group(1)}" if m else (iso or "")


def abschnitt(docx_pfad):
    r = s1.process(docx_pfad)
    normen = r.get("normen") or []
    rspr = r.get("rechtsprechung") or []
    abgr = r.get("_abgrenzungsquellen") or []
    mit = [n for n in normen if NORM_MARKER.search(n.get("norm") or "")]
    ohne = [n for n in normen if not NORM_MARKER.search(n.get("norm") or "")]
    z = [KOPF, "",
         f"> Erzeugt von der Justinian-QS-Pipeline (Stufe 1, justinian-qs-tools {VERSION}) aus "
         f"`{os.path.basename(docx_pfad)}`. Nicht von Hand ändern – maßgeblich ist die Word-Datei.", "",
         "#### Normen", "", "| Norm | Funktion |", "|---|---|"]
    z += [f"| {zelle(n.get('norm'))} | {zelle(n.get('funktion'))} |" for n in mit]
    z += ["", "#### Rechtsprechung", "", "| Gericht | Aktenzeichen | Datum | Relevanz |", "|---|---|---|---|"]
    z += [f"| {zelle(e.get('gericht'))} | {zelle(e.get('aktenzeichen'))} | {zelle(datum(e.get('datum')))} | "
          f"{zelle((e.get('relevanz') or '')[:300])} |" for e in rspr]
    z += ["", "### Ohne Paragraph", ""]
    z += [f"- {zelle(n.get('norm'))}" for n in ohne] or ["- (keine)"]
    z += ["", "### Abgrenzungsquellen", "", "| Norm | Formulierung | Grund | Abschnitt |", "|---|---|---|---|"]
    z += [f"| {zelle(a.get('norm'))} | {zelle(a.get('funktion'))} | {zelle(a.get('grund'))} | {zelle(a.get('abschnitt'))} |"
          for a in abgr]
    return "\n".join(z) + "\n", len(mit) + len(ohne), len(rspr)


def ohne_alten_index(text):
    """Entfernt einen vorhandenen Index-Abschnitt (von KOPF bis zur naechsten '## '-Ueberschrift)."""
    zeilen = text.split("\n")
    out, drin = [], False
    for zl in zeilen:
        if zl.strip().startswith(KOPF):
            drin = True
            continue
        if drin and re.match(r"^#{1,2} ", zl):
            drin = False
        if not drin:
            out.append(zl)
    return "\n".join(out)


def voranstellen(index, text):
    text = ohne_alten_index(text)
    zeilen = text.split("\n")
    for i, zl in enumerate(zeilen):
        if zl.startswith("# "):
            return "\n".join(zeilen[:i + 1] + ["", index.rstrip("\n"), "", "---", ""] + zeilen[i + 1:])
    return index + "\n---\n\n" + text


def main(argv):
    if len(argv) not in (2, 5) or (len(argv) == 5 and argv[2] != "--voranstellen"):
        print(__doc__)
        return 2
    idx, n_norm, n_rspr = abschnitt(argv[1])
    if n_norm == 0 and n_rspr == 0:
        print("ABBRUCH: keine Norm und keine Entscheidung gefunden - nichts geschrieben", file=sys.stderr)
        return 3
    if len(argv) == 2:
        sys.stdout.write(idx)
        return 0
    quelle, ziel = argv[3], argv[4]
    text = open(quelle, encoding="utf-8").read() if os.path.exists(quelle) else "# Rechtsgrundlagen und Rechtsprechung\n"
    neu = voranstellen(idx, text)
    with open(ziel, "w", encoding="utf-8") as f:
        f.write(neu)
    print(f"OK: {ziel} - Index mit {n_norm} Normen, {n_rspr} Entscheidungen ({VERSION})")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
