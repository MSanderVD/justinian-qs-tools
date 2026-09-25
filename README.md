# justinian-qs-tools

Stufe-1-Extraktor der Justinian-QS-Pipeline für die PhiBox-Builder-Agenten
(`JUS Skill Master Builder v1.2`, `JUS Skill Normen Builder v1.2`).

Der Builder lädt einen festen Tag dieses Repositorys in seine Sandbox und setzt mit

    python3 index_md.py <skill.docx> --voranstellen <quelle.md> <ziel.md>

den Abschnitt „Rechtsgrundlagen-Index (Pflichtquellen)" deterministisch vor
`references/normen-rechtsprechung.md`. Der Index ist das Ergebnis derselben Funktion
(`extract_rechtsgrundlagen_deterministisch.process`), die den QS-Index baut.

- Quelle der Dateien: `.claude/skills/stufe_1_index_rechtsgrundlagen/references/` (M. Sander).
  Änderungen dort landen erst mit einem NEUEN Tag im Builder – bewusst, damit ein Paketbau
  reproduzierbar bleibt.
- Nur lesend genutzt: kein CMS-Zugriff, keine Schreibwege, keine Geheimnisse.
- POC vom 25.09.2026. Dauerlösung: Einbindung durch PhiBox selbst.
