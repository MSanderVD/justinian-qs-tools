"""Setzt den VORAB berechneten Rechtsgrundlagen-Index (index/<SKILL-ID>.md) vor
references/normen-rechtsprechung.md eines fertig geschriebenen Skill-Pakets (Weg 1, 28.09.2026).

Laeuft in der Builder-Sandbox (PhiBox AVM) als eigener letzter Schritt des Builders v1.2 - ohne .docx:
beim Paketbau liegt keine Word-Datei in /workspace/skill/, deshalb rechnet Stufe 1 den Index lokal
(index_vorberechnen.py) und legt ihn je Skill in dieses Repo.

  python3 index_einsetzen.py --auto [<paketordner>]          (Standard: /workspace/skill)
  python3 index_einsetzen.py <SKILL-ID> [<paketordner>]

--auto nimmt als Skill-ID die ERSTE JUS-ID in <paketordner>/SKILL.md (der Skill-Kopf steht vorn,
Nachbar-Skills danach) - der Builder muss nichts einsetzen. Prueft, bevor es schreibt:
  - die Skill-ID ist die erste JUS-ID in SKILL.md (sonst Exit 5 - falsches Paket / falscher Name)
  - index/<SKILL-ID>.md existiert (sonst Exit 4 - Skill nicht vorberechnet)
Exitcode 0 = geschrieben und nachgeprueft, 2 = Aufruffehler, 4/5 siehe oben, 6 = Nachpruefung fehlgeschlagen.
"""
import hashlib
import json
import os
import re
import sys

HIER = os.path.dirname(os.path.abspath(__file__))
KOPF = "## Rechtsgrundlagen-Index (Pflichtquellen)"
ID_RE = re.compile(r"JUS\.[A-Z0-9_]+(?:\.[A-Z0-9_]+)*")


def ohne_alten_index(text):
    zeilen, out, drin = text.split("\n"), [], False
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
    """Gleiche Regel wie index_md.voranstellen: direkt unter die erste '# '-Ueberschrift, sonst ganz oben."""
    zeilen = ohne_alten_index(text).split("\n")
    for i, zl in enumerate(zeilen):
        if zl.startswith("# "):
            return "\n".join(zeilen[:i + 1] + ["", index.rstrip("\n"), "", "---", ""] + zeilen[i + 1:])
    return index + "\n---\n\n" + "\n".join(zeilen)


def main(argv):
    if len(argv) not in (2, 3) or not (argv[1] == "--auto" or ID_RE.fullmatch(argv[1].strip())):
        print(__doc__)
        return 2
    paket = argv[2] if len(argv) == 3 else "/workspace/skill"
    skill_md = os.path.join(paket, "SKILL.md")
    # Markdown-Maskierung ("JUS\.LABOR\.…", so im Rohexport) vor der Suche aufheben
    # und Fettmarkierung "__JUS.…__" nicht zur ID zaehlen
    ids = [i.rstrip("_") for i in ID_RE.findall(open(skill_md, encoding="utf-8").read().replace("\\.", ".").replace("\\_", "_"))] \
        if os.path.isfile(skill_md) else []
    if not ids:
        print(f"ABBRUCH: keine Skill-ID in {skill_md} - nichts geschrieben", file=sys.stderr)
        return 5
    # eigene ID = erste JUS-ID der SKILL.md (Kopf); Nachbar-Skills stehen erst danach
    sid = ids[0] if argv[1] == "--auto" else argv[1].strip()
    if sid != ids[0]:
        print(f"ABBRUCH: {sid} ist nicht die erste Skill-ID in {skill_md} ({ids[0]}) - nichts geschrieben", file=sys.stderr)
        return 5
    quelle = os.path.join(HIER, "index", sid + ".md")
    if not os.path.isfile(quelle):
        print(f"ABBRUCH: kein vorberechneter Index fuer {sid} (index/{sid}.md fehlt) - nichts geschrieben", file=sys.stderr)
        return 4
    index = open(quelle, encoding="utf-8").read()
    ziel = os.path.join(paket, "references", "normen-rechtsprechung.md")
    text = open(ziel, encoding="utf-8").read() if os.path.exists(ziel) else "# Rechtsgrundlagen und Rechtsprechung\n"
    os.makedirs(os.path.dirname(ziel), exist_ok=True)
    with open(ziel, "w", encoding="utf-8") as f:
        f.write(voranstellen(index, text))
    neu = open(ziel, encoding="utf-8").read()
    if neu.count(KOPF) != 1 or index.rstrip("\n") not in neu:
        print("ABBRUCH: Nachpruefung fehlgeschlagen - Index nicht genau einmal und unveraendert in der Datei", file=sys.stderr)
        return 6
    manifest = os.path.join(HIER, "index", "_MANIFEST.json")
    stand = json.load(open(manifest, encoding="utf-8")).get(sid, {}) if os.path.isfile(manifest) else {}
    print(f"OK: {ziel} - Index {sid} eingesetzt (sha256 {hashlib.sha256(index.encode()).hexdigest()[:12]}, "
          f"Word-Stand {stand.get('docx_sha256', '?')[:12]}, berechnet {stand.get('berechnet', '?')})")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
