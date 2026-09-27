# OfficeMate

OfficeMate is a local-first, open-source office assistant for Word, Excel, PowerPoint, CSV, Markdown, and text files. It has no paid API dependency and runs with Python's standard library; native Office support and LibreOffice conversions are enabled when the free optional tools are installed.

## Run

```bash
python app.py
```

Then open http://localhost:8765. By default, OfficeMate manages files inside `./workspace` and keeps snapshots in `./workspace/.backups`, deleted files in the hidden local trash, metadata in a hidden sidecar, and an audit trail at `./workspace/.officemate-audit.jsonl`. Set `OFFICEMATE_WORKSPACE=/path/to/folder` to use another folder. The server binds to `0.0.0.0` for local-network / preview access; it has no authentication, so do not expose it to an untrusted network.

Optional native Office and PDF support:

```bash
pip install -r requirements.txt
```

Install LibreOffice separately to enable document conversion and batch PDF export. Without optional packages, notes, Markdown, CSV processing, file search, and the dashboard work using Python's standard library. OfficeMate does not send documents to an external service.

## 150 local tools

The searchable **150 tools** page includes the original 50 tools plus 100 more, grouped into four categories:

- **Files (33):** file and folder management; safe delete-to-trash and backup restore; tagging and pins; audit/inventory exports; SHA-256 checksums and duplicate detection; largest/recent/oldest/empty-file reports; regex content search; ZIP create and safe extract; batch prefix/suffix/extension rename; filename cleanup; UTF-8 normalization; manual backups and snapshot inventory; folder tree and age reports; clear tags.
- **Text (43):** word, character, line, sentence, paragraph, and keyword counts; reading-time and top-keyword estimates; URL/email/phone/date extraction; email/URL redaction; case conversion; whitespace, punctuation, Unicode, accents, HTML-tag, and newline cleanup; line numbering, sorting, reversing, joining, wrapping, and deduplication; find/replace, compare, merge, HTML conversion, checklist and bullet processing, and a Markdown table of contents.
- **CSV (37):** row/column and missing-value reports; header normalization; trim, dedupe, fill, filter, search, sort, sample, limit, split, merge, transpose, unpivot, and row numbering; column select/drop/rename/add/split; unique values, value counts, numeric/group summaries, email and structure validation; CSV/JSON/Markdown conversion; summary reports and create-from-text.
- **Office (37):** create, inspect, extract, merge, and edit Word documents; export Word tables; Excel worksheet previews, CSV exports, search/replace, cleanup, metadata, and formatting; PowerPoint metadata, slide inventory, outline/speaker-note exports, and text replacement; PDF page count, metadata, and text extraction; Office format reports, batch extraction, backups, and PDF conversion.

All 100 new actions are registered in the API and appear as searchable cards with dynamically generated file/input forms. Actions that rename, clear, delete, or edit existing content require confirmation. Existing files are backed up before in-place Office edits; most transformations instead write a separate output file. ZIP extraction rejects path traversal and limits archive size/entry count. Activity is recorded in the local audit trail.

Some actions depend on optional packages: `python-docx` for Word, `openpyxl` for Excel, `python-pptx` for PowerPoint, `pypdf` for PDF reading, and LibreOffice for format conversion. Tools explain missing dependencies when invoked.

## Included dashboard capabilities

Browse and preview managed files, create/edit text notes, search file names and readable content, use meeting/project/expense templates, convert files with LibreOffice, organize files into type folders, and use the local command bar for file search and extractive summaries. Python packages are optional, and unavailable capabilities explain which free package is required.

OfficeMate is an extensible local-first foundation. Voice input, scheduled reports, and local-LLM integration are not included yet.
