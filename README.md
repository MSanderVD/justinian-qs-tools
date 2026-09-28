# justinian-qs-tools

Stufe-1-Extraktor und vorberechnete Rechtsgrundlagen-Indizes der Justinian-QS-Pipeline für den
PhiBox-Builder `JUS Skill Master Builder v1.2`.

**Weg 1 (seit 28.09.2026):** Beim Paketbau liegt keine .docx in der Sandbox. Deshalb wird der Index
lokal je Skill vorberechnet (`index/<SKILL-ID>.md`, Manifest `index/_MANIFEST.json` mit sha256 der
Word-Datei). Der 6. Schritt des Builders führt nur Befehle aus: Er lädt einen festen Tag dieses Repos und ruft

    python3 index_einsetzen.py --auto /workspace/skill

auf. Das Skript nimmt die erste JUS-ID der `SKILL.md` und setzt den Index unverändert direkt unter
die erste Überschrift von `references/normen-rechtsprechung.md`. Ein alter Index-Abschnitt wird dabei ersetzt.

- `index_md.py`: rechnet den Abschnitt aus der .docx (`extract_rechtsgrundlagen_deterministisch.process`),
  also mit derselben Funktion, die den QS-Index baut.
- Quelle der Extraktor-Dateien: `.claude/skills/stufe_1_index_rechtsgrundlagen/references/` (M. Sander).
  Änderungen landen erst mit einem NEUEN Tag im Builder. Das ist Absicht, damit ein Paketbau reproduzierbar bleibt.
- Nur lesend genutzt: kein CMS-Zugriff, keine Schreibwege, keine Geheimnisse.
- Dauerlösung: PhiBox legt die .docx beim Paketbau in die Sandbox (Weg 3).
