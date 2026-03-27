# Heise Ladebordstein Analyse

Dieses Repository enthaelt eine lokale Auswertung von drei Heise-Forenthreads zum Thema Ladebordsteine.

## Inhalt

- `heise_forum_curbcharger_analyse.ps1`
- Vollberichte und geglaettete Endfassungen fÃ¼r drei Threads
- JSON-Auswertungen pro Thread
- eine Gesamtstatistik Ã¼ber alle drei Threads
- eine Vergleichsdatei `KÃ¶ln_vs_TankE_vs_Serie`

Die grossen HTML-Caches, virtuelle Umgebungen und sonstige Laufzeitdateien sind bewusst nicht versioniert.

## PowerShell-Skript

`heise_forum_curbcharger_analyse.ps1` ist das zentrale Skript fÃ¼r den gesamten Workflow.

Es macht in einem Lauf:

1. Heise-Forenseiten eines Threads abrufen
2. alle Root-Threads finden
3. fÃ¼r jeden Root-Thread die vollstaendige Thread-Ansicht laden
4. alle Posting-URLs deduplizieren
5. alle Posting-Seiten cachen
6. Kommentare lokal parsen und thematisch klassifizieren
7. JSON-Ausgaben und Markdown-Berichte erzeugen

Das Skript arbeitet threadbezogen und schreibt pro Forum getrennte Dateien mit Datum und Threadnamen.

## Parameter

- `-ForumUrl`
  Erwartet eine Heise-Forum-URL wie `https://www.heise.de/forum/heise-online/Kommentare/.../forum-123456/comment/`
- `-SkipFetch`
  Nutzt nur vorhandene lokale Caches und rechnet die Auswertung neu
- `-OnlyAuthors`
  Ã¼berspringt das Nachladen einzelner Posting-Seiten und ist nur fÃ¼r Sonderfaelle gedacht
- `-WorkDir`
  Optionales Arbeitsverzeichnis

## Beispiele

Analyse eines neuen Threads:

```powershell
.\heise_forum_curbcharger_analyse.ps1 -ForumUrl "https://www.heise.de/forum/heise-online/Kommentare/Elektroautos-Koeln-bekommt-Ladebordsteine/forum-519185/comment/"
```

Nur vorhandene Caches neu auswerten:

```powershell
.\heise_forum_curbcharger_analyse.ps1 -ForumUrl "https://www.heise.de/forum/heise-online/Kommentare/Elektroautos-Koeln-bekommt-Ladebordsteine/forum-519185/comment/" -SkipFetch
```

## Erzeugte Dateien

Pro Thread entstehen je nach Lauf unter anderem:

- `YYYY-MM-DD_<thread>_comments_full.json`
- `YYYY-MM-DD_<thread>_summary_full.json`
- `YYYY-MM-DD_<thread>_authors_summary.json`
- `YYYY-MM-DD_<thread>_vollbericht.md`
- `YYYY-MM-DD_<thread>_endfassung.md`

Weitere Gesamtdateien:

- `2026-03-27_gesamtstatistik_drei_threads.md`
- `2026-03-27_gesamtstatistik_drei_threads.json`
- `2026-03-27_KÃ¶ln_vs_TankE_vs_Serie.md`

## Hinweise

- Die Tenor-Einordnung ist heuristisch und keine manuelle Vollkodierung.
- Die Zusammenfuehrung von Autoren erfolgt nur Ã¼ber den sichtbaren Heise-Forennamen.
- HTML-Caches sind fÃ¼r die Reproduzierbarkeit hilfreich, werden hier aber nicht ins Repository eingecheckt.
