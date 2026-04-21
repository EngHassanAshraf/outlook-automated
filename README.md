# Microsoft Outlook Automated

A Python tool that automates saving email attachments and archiving emails from Outlook inbox — with a live terminal UI, config-driven classification, and safe file handling.

> [!NOTE]
> This tool is still under development. Some folder names and sender-normalization rules reflect a specific deployment context and may need adjusting for other environments.

## Features

- **Live Terminal UI** — Rich-powered progress bar pinned to the bottom of the terminal showing percentage, fraction, speed, elapsed time, and a scrolling activity log
- **Config-driven Classification** — Email categories and keywords are defined entirely in `config.yaml`; no code changes needed to add or rename categories
- **Sub-category Support** — Daily reports are further classified into Accommodation, Visitors, Attendance, and site-specific sub-folders
- **Attachment Management** — Saves accepted file types (`docx`, `pdf`, `xlsx`, `pptx`) into an organized folder hierarchy; ignores email signature images automatically
- **Collision-safe Saving** — Duplicate filenames get a numeric suffix (`report(1).pdf`) instead of silently overwriting
- **Smart Date Extraction** — Parses dates from attachment filenames using multiple patterns (`YYYY-MM-DD`, `DD-MM-YYYY`, `YYYYMMDD`, `DDMMYYYY`); falls back to email received date
- **Archive Management** — Moves processed emails to a configurable archive folder
- **Pinned Dependency Versions** — `requirements.txt` pins all runtime deps for reproducible installs

## Installation

### Prerequisites

- Python 3.8+
- Microsoft Outlook installed and logged in
- Windows (required for `pywin32` / MAPI)

### Setup

```bash
pip install -r requirements.txt
```

Runtime dependencies:

- `pywin32==311` — Outlook automation via MAPI
- `psutil==7.2.1` — Disk partition detection
- `python-bidi==0.6.7` — Arabic RTL text display in terminal
- `PyYAML==6.0.3` — Configuration file parsing
- `rich==15.0.0` — Live terminal UI

For development (includes pytest):

```bash
pip install -r requirements-dev.txt
```

## Usage

```bash
# Pass partition letter as argument
python main.py D

# Or run interactively — the tool will prompt
python main.py
```

The tool will:

1. Connect to Outlook via MAPI
2. Scan the inbox and ask whether to include unread emails
3. Show a live progress UI while processing
4. Save attachments to the output directory
5. Move each processed email to the archive folder
6. Display a completion summary

## Project Structure

```text
outlook_auto/
├── main.py              # Orchestration and entry point
├── application.py       # Outlook COM connection and folder access
├── message.py           # Email and attachment domain logic
├── config_manager.py    # YAML config loader with typed getters
├── progress_ui.py       # Rich live terminal UI (ProgressUI class)
├── config.yaml          # All runtime configuration
├── requirements.txt     # Pinned runtime dependencies
├── requirements-dev.txt # Dev dependencies (pytest)
└── README.md
```

## Configuration

All behaviour is controlled by `config.yaml`. The file is the single source of truth — the application has no duplicate hardcoded defaults.

```yaml
output:
  base_folder: "MV"
  year_format: "MV-{year}"

processing:
  process_unread: false
  archive_processed: true
  mark_as_read: true

categories:
  technical:
    keywords: ["الحالة الفنية", "technical"]
    name: "Technical Report - التقرير الفني"
  daily_report:
    keywords: ["اليومي", "اليومى", "الحالة الأمنية", "daily"]
    name: "Daily Report - التقرير اليومي"
    sub_categories:
      accommodation:
        keywords: ["تواجد الملاك", "mv-nc accommodation"]
        name: "Accommodation - تقرير الإقامة"
      visitors:
        keywords: ["تواجد الزائرين"]
        name: "Dayra - تقرير دايرة"
      attendance:
        keywords: ["الحضور", "الإنصراف"]
        name: "الحضور والإنصراف"
  # ... more categories

attachments:
  accepted_types: ["docx", "pdf", "xlsx", "pptx"]
  ignored_files: ["EmailSignature.jpg", "image001.png"]  # signature images

logging:
  level: "INFO"
  format: "%(asctime)s - %(levelname)s - %(message)s"
  file: "outlook_auto.log"
  console: true   # suppressed automatically during live UI

outlook:
  application: "Outlook.Application"
  namespace: "MAPI"
  inbox_folder_number: 6
  archive_root_folder: "Archives"
  archive_folder_name: "Archive"
```

To add a new category, add an entry under `categories` with a `name` and `keywords` list — no code changes required.

## Folder Structure

Attachments are saved to:

```text
{Partition}:\{BaseFolder}\{Year}\{Category}\{Compound}\{SubCategory}\{Month}\{WeekNumber}\{Filename}
```

Example:

```text
D:\MV\MV-2026\Daily Report - التقرير اليومي\MVHYDEPARK\Dayra - تقرير دايرة\4. April\week 2\report_20260410.pdf
```

Technical and weekly-operations categories include a `week N` sub-folder based on the processing date.

## Live Terminal UI

During processing the terminal shows a two-panel display:

```text
╭─ Latest Activity ──────────────────────────────────────────────╮
│  ☑  Daily Report - التقرير اليومي                              │
│  ☑  Weekly Technical Status Report                             │
│  ⚠  Collision: saved as 09-02-2026(1).pdf                      │
│  ☑  Daily Report - التقرير اليومي                              │
│  ☑  Daily Report - التقرير اليومي                              │
╰────────────────────────────────────────────────────────────────╯
╭────────────────────────────────────────────────────────────────╮
│ Processing emails ████████░░░░  57.0%  46 / 80  2.3/s  0:00:32 │
╰────────────────────────────────────────────────────────────────╯
```

- Top panel — last 5 activity lines (archived subjects, warnings, errors)
- Bottom panel — animated bar, percentage, `processed / total`, speed, elapsed time
- On completion — border turns green and shows the summary message
- All raw log output is suppressed from the terminal during processing; full logs go to `outlook_auto.log`

## Logging

Logs are written to `outlook_auto.log` at all times. Console output is suppressed while the live UI is active.

Log levels:

- `INFO` — normal operations (connected, archived, saved)
- `WARNING` — non-fatal issues (filename collision, config key missing)
- `ERROR` — failures (move failed, save failed)

Enable debug logging in `config.yaml`:

```yaml
logging:
  level: "DEBUG"
```

## Troubleshooting

| Problem | Fix |
|---|---|
| Outlook not found | Ensure Outlook is installed and a profile is logged in |
| Permission errors | Run the terminal as Administrator |
| Partition not found | Check the drive letter is mounted and accessible |
| Config errors | Validate `config.yaml` indentation and structure |
| UI not visible | Ensure the terminal supports ANSI escape codes (Windows Terminal recommended) |

## Future Features

- [ ] Parse attachment content and extract structured data
- [ ] Generate weekly summary Excel sheets from extracted data
- [ ] Database integration for tracking processed emails
- [ ] Date-range and sender filtering options
- [ ] Parallel attachment saving for large inboxes
- [ ] GUI configuration interface

## Development

```bash
# Run tests
python -m pytest tests/
```

The project follows a modular layered structure — `main.py` orchestrates, `application.py` handles COM, `message.py` handles domain logic, `config_manager.py` handles config, `progress_ui.py` handles display. Each layer has no knowledge of layers above it.

## License

MIT License — see the LICENSE file for details.
