#!/usr/bin/env python3
"""Local-first OfficeMate server. Uses only the Python standard library by default."""
from __future__ import annotations

import csv
import hashlib
import html
import io
import statistics
import difflib
import unicodedata
import textwrap
import zipfile
import datetime as dt
import json
import mimetypes
import os
import re
import shutil
import subprocess
import urllib.parse
from collections import Counter
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WORKSPACE = Path(os.environ.get("OFFICEMATE_WORKSPACE", ROOT / "workspace")).expanduser().resolve()
BACKUPS = WORKSPACE / ".backups"
AUDIT = WORKSPACE / ".officemate-audit.jsonl"
STATIC = ROOT / "static"
SUPPORTED = {".docx", ".xlsx", ".pptx", ".pdf", ".csv", ".txt", ".md", ".json"}

FEATURES = [
("rename","Files","Rename a file","Change a filename without replacing an existing file.",{"path":"File","name":"New filename"}),
("duplicate","Files","Duplicate a file","Create a separate copy next to the original.",{"path":"File","name":"Copy filename (optional)"}),
("delete","Files","Delete safely","Move a file to the local trash after confirmation; a backup is kept.",{"path":"File"}),
("create_folder","Files","Create a folder","Add a named folder inside your workspace.",{"name":"Folder name"}),
("move","Files","Move to folder","Move a file to a workspace folder without overwriting.",{"path":"File","name":"Destination folder"}),
("tag","Files","Tag a file","Add searchable labels to your local file metadata.",{"path":"File","name":"Tags, separated by commas"}),
("pin","Files","Pin a file","Mark a file as a favorite in your workspace.",{"path":"File"}),
("restore_backup","Files","Restore latest backup","Restore the newest available snapshot for a file.",{"path":"File"}),
("export_audit","Files","Export audit trail","Export the local action log as a CSV file.",{}),
("file_hash","Files","Calculate file checksum","Create a SHA-256 checksum to verify file integrity.",{"path":"File"}),
("find_duplicates","Files","Find duplicate files","Find files with identical contents using SHA-256.",{}),
("workspace_stats","Files","Workspace report","View file-type counts and total storage used.",{}),
("word_count","Text","Word count","Count words in a readable document.",{"path":"File"}),
("char_count","Text","Character count","Count characters, with and without spaces.",{"path":"File"}),
("reading_time","Text","Estimate reading time","Estimate reading time at 220 words per minute.",{"path":"File"}),
("uppercase","Text","Uppercase text","Create a new uppercase copy of a text file.",{"path":"File"}),
("lowercase","Text","Lowercase text","Create a new lowercase copy of a text file.",{"path":"File"}),
("title_case","Text","Title case text","Create a title-cased copy of a text file.",{"path":"File"}),
("trim_whitespace","Text","Trim line whitespace","Create a copy with leading/trailing whitespace removed.",{"path":"File"}),
("remove_blank_lines","Text","Remove blank lines","Create a copy with empty lines removed.",{"path":"File"}),
("sort_lines","Text","Sort lines","Create a copy with lines sorted alphabetically.",{"path":"File"}),
("unique_lines","Text","Remove duplicate lines","Create a copy with duplicate lines removed.",{"path":"File"}),
("replace_text","Text","Find and replace","Replace matching text and save a separate copy.",{"path":"File","query":"Find text","replacement":"Replace with"}),
("compare_text","Text","Compare two text files","Show a line-by-line unified diff.",{"path":"First file","second":"Second file"}),
("merge_text","Text","Merge two text files","Combine two files into a new Markdown file.",{"path":"First file","second":"Second file","name":"Output filename (optional)"}),
("text_to_html","Text","Text to HTML","Create a simple escaped HTML document from plain text.",{"path":"File"}),
("markdown_to_html","Text","Markdown to HTML","Convert basic Markdown formatting into a local HTML file.",{"path":"File"}),
("checklist","Text","Turn lines into a checklist","Convert each non-empty line into an unchecked task.",{"path":"File"}),
("extract_text","Text","Extract text to a note","Save readable text from Word, Excel, PowerPoint, or plain text.",{"path":"File"}),
("reverse_lines","Text","Reverse line order","Create a copy with the order of lines reversed.",{"path":"File"}),
("csv_stats","CSV","CSV overview","Report rows, columns, and empty cells.",{"path":"CSV file"}),
("csv_headers","CSV","Normalize CSV headers","Create a copy with clean, unique snake_case headers.",{"path":"CSV file"}),
("csv_trim","CSV","Trim CSV cell spaces","Create a copy with surrounding whitespace removed from cells.",{"path":"CSV file"}),
("csv_dedupe","CSV","Remove duplicate CSV rows","Create a copy with duplicate records removed.",{"path":"CSV file"}),
("csv_sort","CSV","Sort CSV by column","Create a sorted CSV copy using a selected column name.",{"path":"CSV file","name":"Column name"}),
("csv_filter","CSV","Filter CSV rows","Keep rows where a column exactly matches a value.",{"path":"CSV file","query":"Column name","replacement":"Value to match"}),
("csv_markdown","CSV","CSV to Markdown table","Create a Markdown table from a CSV file.",{"path":"CSV file"}),
("csv_split","CSV","Split a large CSV","Split into smaller CSV files while repeating the header.",{"path":"CSV file","name":"Rows per part (default 500)"}),
("csv_merge","CSV","Merge two CSV files","Append matching-column CSV files into one new file.",{"path":"First CSV","second":"Second CSV","name":"Output filename (optional)"}),
("csv_numeric_report","CSV","Numeric column report","Calculate count, sum, mean, minimum, and maximum per numeric column.",{"path":"CSV file"}),
("csv_transpose","CSV","Transpose CSV","Swap rows and columns in a CSV copy.",{"path":"CSV file"}),
("csv_create","CSV","Create CSV from pasted rows","Create a CSV file from comma-separated lines.",{"name":"Filename","text":"CSV content (header on first line)"}),
("docx_create","Office","Create a Word document","Create a simple DOCX from text (requires python-docx).",{"name":"Filename","text":"Document text"}),
("xlsx_from_csv","Office","Create Excel from CSV","Create an XLSX workbook from a CSV (requires openpyxl).",{"path":"CSV file","name":"Output filename (optional)"}),
("xlsx_sheets","Office","Inspect Excel sheets","List sheet names, dimensions, and populated rows.",{"path":"Excel workbook"}),
("pptx_inventory","Office","Inspect PowerPoint slides","List slide count and extracted slide text.",{"path":"Presentation"}),
("pptx_from_outline","Office","Create slides from an outline","Create a simple slide deck from Markdown headings (requires python-pptx).",{"path":"Outline file","name":"Output filename (optional)"}),
("docx_metadata","Office","Inspect Word metadata","Show title, author, subject, and paragraph count.",{"path":"Word document"}),
("batch_convert_pdf","Office","Batch convert to PDF","Convert all supported Office files using LibreOffice.",{}),
("batch_rename_prefix","Files","Batch add filename prefix","Prefix all top-level workspace files without overwriting.",{"name":"Filename prefix"}),
]

EXTRA_FEATURES = [
# 20 additional file/workspace actions
("create_workspace_readme","Files","Create workspace README","Add a starter guide to your managed folder.",{"name":"Filename (optional)"}),
("clean_empty_folders","Files","Remove empty folders","Remove empty folders after confirmation; files are never deleted.",{}),
("largest_files","Files","Largest files","List the largest files in your workspace.",{"name":"Number of results (default 10)"}),
("recent_files","Files","Recently changed files","List files changed most recently.",{"name":"Number of results (default 10)"}),
("oldest_files","Files","Oldest files","Find files with the oldest modification dates.",{"name":"Number of results (default 10)"}),
("file_inventory","Files","Export file inventory","Create a CSV with names, types, sizes, and modification dates.",{"name":"Output filename (optional)"}),
("regex_search","Files","Search with a regular expression","Search readable file contents using a Python regular expression.",{"query":"Regular expression"}),
("export_file_hashes","Files","Export file checksums","Create a CSV inventory of SHA-256 checksums.",{"name":"Output filename (optional)"}),
("create_zip","Files","Create a ZIP archive","Archive selected workspace files (comma-separated paths).",{"query":"File paths, separated by commas","name":"Archive filename"}),
("extract_zip","Files","Extract a ZIP archive","Extract files safely inside a new workspace folder.",{"path":"ZIP archive","name":"Destination folder"}),
("batch_suffix","Files","Batch add filename suffix","Add a suffix to top-level filenames without overwriting.",{"name":"Suffix (e.g. _final)"}),
("batch_extension","Files","Batch change extensions","Rename selected top-level files to a new extension (confirmation required).",{"name":"Current extension (e.g. .txt)","query":"New extension (e.g. .md)"}),
("sanitize_names","Files","Clean filenames","Replace awkward filename characters with safe underscores.",{}),
("normalize_encoding","Files","Normalize text encoding","Read with replacement fallback and save a UTF-8 copy.",{"path":"Text file","name":"Output filename (optional)"}),
("backup_now","Files","Snapshot a file now","Create a timestamped local backup without editing the source.",{"path":"File"}),
("backup_inventory","Files","List file backups","Show saved snapshot names and dates for a file.",{"path":"File"}),
("folder_tree","Files","Workspace folder tree","Show folders and file counts in a compact tree.",{}),
("file_age_report","Files","File age report","Summarize workspace files by age in days.",{}),
("empty_files","Files","Find empty files","List zero-byte files that may need attention.",{}),
("clear_tags","Files","Clear file tags","Remove tags for a file after confirmation.",{"path":"File"}),
# 25 additional text actions
("line_count","Text","Line count","Count lines in a readable document.",{"path":"File"}),
("sentence_count","Text","Sentence count","Estimate sentences using punctuation boundaries.",{"path":"File"}),
("paragraph_count","Text","Paragraph count","Count non-empty text paragraphs.",{"path":"File"}),
("find_urls","Text","Extract web links","List URLs found in a document.",{"path":"File"}),
("extract_emails","Text","Extract email addresses","Find email addresses in readable text.",{"path":"File"}),
("extract_phone_numbers","Text","Extract phone-like numbers","Find phone number patterns in readable text.",{"path":"File"}),
("extract_dates","Text","Extract dates","Find common numeric and written date patterns.",{"path":"File"}),
("redact_emails","Text","Redact email addresses","Create a copy with email addresses masked.",{"path":"File"}),
("redact_urls","Text","Redact URLs","Create a copy with web links masked.",{"path":"File"}),
("clean_punctuation","Text","Clean repeated punctuation","Create a copy with repeated punctuation normalized.",{"path":"File"}),
("normalize_unicode","Text","Normalize Unicode text","Create a Unicode NFKC-normalized copy.",{"path":"File"}),
("remove_accents","Text","Remove accents","Create a copy with Latin diacritics removed.",{"path":"File"}),
("remove_html_tags","Text","Strip HTML tags","Create a plain text copy of HTML-like content.",{"path":"File"}),
("normalize_newlines","Text","Normalize line endings","Create a UTF-8 copy with consistent LF newlines.",{"path":"File"}),
("indent_text","Text","Indent paragraphs","Create a copy with each non-empty line indented.",{"path":"File","name":"Indent spaces (default 2)"}),
("wrap_text","Text","Wrap long lines","Create a copy wrapped to a target width.",{"path":"File","name":"Line width (default 80)"}),
("number_lines","Text","Add line numbers","Create a copy with numbered lines.",{"path":"File"}),
("remove_line_numbers","Text","Remove line numbers","Create a copy with leading line numbers removed.",{"path":"File"}),
("extract_bullets","Text","Extract bullet points","Save bullet and numbered list items to a new note.",{"path":"File"}),
("remove_bullets","Text","Remove bullet markers","Create a copy with list markers removed.",{"path":"File"}),
("keyword_count","Text","Count a keyword","Count exact case-insensitive matches for a phrase.",{"path":"File","query":"Keyword or phrase"}),
("top_keywords","Text","Top keywords","List frequent meaningful words in a document.",{"path":"File","name":"Number of keywords (default 10)"}),
("generate_toc","Text","Generate a Markdown contents list","Build a linked table of contents from Markdown headings.",{"path":"File"}),
("reverse_words","Text","Reverse word order","Create a copy with words reversed per line.",{"path":"File"}),
("join_lines","Text","Join lines","Combine lines using a chosen separator.",{"path":"File","name":"Separator (default space)"}),
# 25 additional CSV/data actions
("csv_columns","CSV","List CSV columns","Show column names and positions.",{"path":"CSV file"}),
("csv_null_report","CSV","CSV missing-value report","Count blank cells by column.",{"path":"CSV file"}),
("csv_fill_blanks","CSV","Fill blank CSV cells","Create a copy with empty cells replaced by a chosen value.",{"path":"CSV file","replacement":"Replacement value"}),
("csv_drop_columns","CSV","Drop CSV columns","Create a copy excluding named columns (comma-separated).",{"path":"CSV file","name":"Columns to drop, comma-separated"}),
("csv_select_columns","CSV","Select CSV columns","Create a copy containing only the requested columns.",{"path":"CSV file","name":"Columns to keep, comma-separated"}),
("csv_rename_column","CSV","Rename a CSV column","Rename one header and create a separate copy.",{"path":"CSV file","query":"Current column name","replacement":"New column name"}),
("csv_add_column","CSV","Add a CSV column","Add a constant-valued column to every data row.",{"path":"CSV file","name":"New column name","replacement":"Value for each row"}),
("csv_drop_empty_rows","CSV","Remove empty CSV rows","Create a copy without entirely blank rows.",{"path":"CSV file"}),
("csv_limit_rows","CSV","Limit CSV rows","Keep the header and first N data rows.",{"path":"CSV file","name":"Maximum data rows"}),
("csv_sample_rows","CSV","Sample CSV rows","Create a deterministic sample of rows for testing.",{"path":"CSV file","name":"Sample size"}),
("csv_search_rows","CSV","Search CSV rows","Keep rows containing a phrase in any cell.",{"path":"CSV file","query":"Text to find"}),
("csv_replace_values","CSV","Replace CSV values","Replace a value across every cell.",{"path":"CSV file","query":"Find value","replacement":"Replace with"}),
("csv_unique_values","CSV","List unique column values","Report distinct values in a selected column.",{"path":"CSV file","name":"Column name"}),
("csv_value_counts","CSV","Count column values","Count occurrences of each value in a selected column.",{"path":"CSV file","name":"Column name"}),
("csv_group_sum","CSV","Group and sum CSV values","Sum a numeric column for each group in another column.",{"path":"CSV file","name":"Numeric column","query":"Group-by column"}),
("csv_to_json","CSV","Convert CSV to JSON","Create a JSON array of objects using CSV headers.",{"path":"CSV file"}),
("json_to_csv","CSV","Convert JSON to CSV","Convert a JSON array of objects into a spreadsheet-ready CSV.",{"path":"JSON file","name":"Output filename (optional)"}),
("csv_validate_email","CSV","Validate email column","Report rows with malformed email values.",{"path":"CSV file","name":"Email column"}),
("csv_validate","CSV","Validate CSV structure","Check for inconsistent row widths and duplicate headers.",{"path":"CSV file"}),
("csv_summary_report","CSV","Create a CSV summary report","Create a Markdown report with row count, columns, and missing values.",{"path":"CSV file"}),
("csv_sort_numeric","CSV","Sort CSV numerically","Sort rows by a numeric column.",{"path":"CSV file","name":"Numeric column"}),
("csv_split_column","CSV","Split a CSV column","Split one column into multiple columns using a delimiter.",{"path":"CSV file","name":"Column name","query":"Delimiter (default comma)"}),
("csv_unpivot","CSV","Unpivot CSV columns","Convert wide data columns into attribute/value rows.",{"path":"CSV file","name":"Identifier columns, comma-separated"}),
("csv_add_row_number","CSV","Add row numbers","Create a copy with a sequential row_number column.",{"path":"CSV file"}),
("csv_column_stats","CSV","Column statistics","Show non-empty count, distinct count, and min/max lengths.",{"path":"CSV file","name":"Column name"}),
# 30 additional Office/PDF actions
("docx_outline","Office","Word heading outline","Extract heading structure from a DOCX document.",{"path":"Word document"}),
("docx_count","Office","Word document statistics","Count words, paragraphs, tables, and sections.",{"path":"Word document"}),
("docx_tables_csv","Office","Export Word tables to CSV","Extract Word tables into a single CSV file.",{"path":"Word document","name":"Output filename (optional)"}),
("docx_replace","Office","Find and replace in Word","Replace text in Word paragraphs and tables (backup required).",{"path":"Word document","query":"Find text","replacement":"Replace with"}),
("docx_append","Office","Append text to Word","Add a paragraph to the end of a DOCX (backup required).",{"path":"Word document","text":"Text to append"}),
("docx_merge","Office","Merge Word documents","Append paragraphs and tables from a second DOCX.",{"path":"First Word document","second":"Second Word document","name":"Output filename (optional)"}),
("docx_set_metadata","Office","Set Word document properties","Update title and subject (backup required).",{"path":"Word document","name":"Document title","query":"Subject"}),
("xlsx_preview","Office","Preview a worksheet","Show the first rows from a selected sheet.",{"path":"Excel workbook","name":"Sheet name (optional)"}),
("xlsx_export_sheet_csv","Office","Export one Excel sheet to CSV","Save a selected worksheet as CSV.",{"path":"Excel workbook","name":"Sheet name (optional)"}),
("xlsx_export_all_csv","Office","Export all Excel sheets","Create one CSV per worksheet.",{"path":"Excel workbook"}),
("xlsx_find_text","Office","Search an Excel workbook","Find cells containing a phrase across sheets.",{"path":"Excel workbook","query":"Text to find"}),
("xlsx_replace_value","Office","Replace Excel cell values","Replace matching cell values (backup required).",{"path":"Excel workbook","query":"Find value","replacement":"Replace with"}),
("xlsx_remove_empty_rows","Office","Remove empty Excel rows","Create a cleaned workbook copy without empty rows.",{"path":"Excel workbook","name":"Output filename (optional)"}),
("xlsx_add_sheet","Office","Add an Excel sheet","Add a named worksheet (backup required).",{"path":"Excel workbook","name":"New sheet name"}),
("xlsx_rename_sheet","Office","Rename an Excel sheet","Rename a worksheet (backup required).",{"path":"Excel workbook","name":"Current sheet name","query":"New sheet name"}),
("xlsx_delete_empty_sheets","Office","Remove empty Excel sheets","Delete empty worksheets (backup required).",{"path":"Excel workbook"}),
("xlsx_style_header","Office","Format Excel headers","Bold and shade the first row of each sheet (backup required).",{"path":"Excel workbook"}),
("xlsx_freeze_header","Office","Freeze Excel header rows","Freeze panes below row one in each sheet (backup required).",{"path":"Excel workbook"}),
("xlsx_metadata","Office","Excel workbook summary","Show sheet names, dimensions, and calculation properties.",{"path":"Excel workbook"}),
("pptx_export_outline","Office","Export slide outline","Create a Markdown outline from slide titles and text.",{"path":"PowerPoint"}),
("pptx_export_notes","Office","Export speaker notes","Save slide speaker notes as a text file.",{"path":"PowerPoint"}),
("pptx_replace_text","Office","Find and replace in slides","Replace text in slide shapes (backup required).",{"path":"PowerPoint","query":"Find text","replacement":"Replace with"}),
("pptx_set_metadata","Office","Set presentation properties","Update title and subject (backup required).",{"path":"PowerPoint","name":"Presentation title","query":"Subject"}),
("pptx_slide_count","Office","PowerPoint deck statistics","Report slide count, shape count, and notes availability.",{"path":"PowerPoint"}),
("pdf_page_count","Office","Count PDF pages","Report PDF page count using the optional pypdf package.",{"path":"PDF file"}),
("pdf_extract_text","Office","Extract PDF text","Save selectable PDF text into a local note (requires pypdf).",{"path":"PDF file","name":"Output filename (optional)"}),
("pdf_metadata","Office","Inspect PDF metadata","Show title, author, creator, and page count (requires pypdf).",{"path":"PDF file"}),
("office_format_report","Office","Office file inventory","Summarize DOCX, XLSX, PPTX, and PDF files by format.",{}),
("batch_office_to_text","Office","Batch extract Office text","Extract readable text from all Word, Excel, and PowerPoint files.",{}),
("batch_office_backup","Office","Back up all Office files","Create timestamped backups of all Office documents.",{}),
]
FEATURES.extend(EXTRA_FEATURES)



def now():
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


def safe_path(relative: str) -> Path:
    candidate = (WORKSPACE / relative).resolve()
    if candidate == WORKSPACE or WORKSPACE not in candidate.parents:
        raise ValueError("That path is outside the OfficeMate workspace.")
    if any(part.startswith(".") for part in candidate.relative_to(WORKSPACE).parts):
        raise ValueError("Hidden system files are not available here.")
    return candidate


def audit(operation: str, relative: str, detail: str = ""):
    WORKSPACE.mkdir(parents=True, exist_ok=True)
    with AUDIT.open("a", encoding="utf-8") as f:
        f.write(json.dumps({"timestamp": now(), "operation": operation, "file": relative, "detail": detail}, ensure_ascii=False) + "\n")


def backup(path: Path):
    if not path.exists() or not path.is_file():
        return None
    BACKUPS.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    destination = BACKUPS / f"{path.name}.{stamp}.bak"
    shutil.copy2(path, destination)
    return destination


def file_text(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in {".txt", ".md", ".csv", ".json", ".html", ".htm"}:
        return path.read_text(encoding="utf-8", errors="replace")[:100000]
    if ext == ".docx":
        try:
            from docx import Document
            doc = Document(path)
            paragraphs = [p.text for p in doc.paragraphs if p.text]
            for table in doc.tables:
                paragraphs.extend(" | ".join(cell.text for cell in row.cells) for row in table.rows)
            return "\n".join(paragraphs)[:100000]
        except ImportError:
            return "Install python-docx to preview Word documents."
    if ext == ".pptx":
        try:
            from pptx import Presentation
            prs = Presentation(path)
            return "\n\n".join("\n".join(sh.text for sh in slide.shapes if getattr(sh, "has_text_frame", False)) for slide in prs.slides)[:100000]
        except ImportError:
            return "Install python-pptx to preview PowerPoint presentations."
    if ext == ".xlsx":
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            result = []
            for ws in wb.worksheets:
                result.append(f"[{ws.title}]")
                result.extend(" | ".join("" if v is None else str(v) for v in row) for row in ws.iter_rows(values_only=True))
            return "\n".join(result)[:100000]
        except ImportError:
            return "Install openpyxl to preview Excel workbooks."
    return f"Preview is not available for {ext or 'this file type'}."


def summarize(text: str, limit=5) -> str:
    sentences = re.split(r"(?<=[.!?])\s+|\n+", text.strip())
    sentences = [re.sub(r"\s+", " ", s).strip() for s in sentences if len(s.strip()) > 25]
    if not sentences:
        return "There is not enough readable text to summarize."
    # Rank by useful signal while retaining the source order.
    words = re.findall(r"[a-zA-Z]{4,}", text.lower())
    common = Counter(w for w in words if w not in {"that", "with", "from", "this", "have", "were", "will", "your", "about", "there", "their", "which", "been", "would", "could", "should"})
    ranked = sorted(enumerate(sentences), key=lambda x: (sum(common[w.lower()] for w in re.findall(r"[a-zA-Z]{4,}", x[1])) / max(len(x[1].split()), 1)), reverse=True)
    chosen = sorted(i for i, _ in ranked[:limit])
    return "\n".join(f"• {sentences[i]}" for i in chosen)


def list_files():
    WORKSPACE.mkdir(parents=True, exist_ok=True)
    items = []
    try: metadata = json.loads((WORKSPACE / ".officemate-meta.json").read_text(encoding="utf-8"))
    except Exception: metadata = {}
    for path in WORKSPACE.rglob("*"):
        if not path.is_file() or any(part.startswith(".") for part in path.relative_to(WORKSPACE).parts):
            continue
        rel = path.relative_to(WORKSPACE).as_posix()
        try:
            stat = path.stat()
            items.append({"name": path.name, "path": rel, "ext": path.suffix.lower(), "size": stat.st_size,
                          "modified": dt.datetime.fromtimestamp(stat.st_mtime).astimezone().isoformat(timespec="minutes"),
                          "kind": {".docx": "Word", ".xlsx": "Excel", ".pptx": "PowerPoint", ".pdf": "PDF", ".csv": "CSV", ".md": "Markdown", ".txt": "Text", ".html": "HTML"}.get(path.suffix.lower(), "File"),
                          "tags": metadata.get(rel, {}).get("tags", []), "pinned": metadata.get(rel, {}).get("pinned", False)})
        except OSError:
            continue
    items.sort(key=lambda x: x["modified"], reverse=True)
    return items


def tool_output(name: str, content: str | bytes, operation: str) -> str:
    name = Path(name).name.strip()
    if not name or name in {".", ".."}:
        raise ValueError("Enter a valid output filename.")
    path = safe_path(name)
    if path.exists():
        raise FileExistsError(f"{name} already exists. Choose a different output filename.")
    path.parent.mkdir(parents=True, exist_ok=True)
    if isinstance(content, bytes): path.write_bytes(content)
    else: path.write_text(content, encoding="utf-8")
    audit(operation, name)
    return name


def require_file(rel: str, extensions=None) -> Path:
    if not rel: raise ValueError("Choose a file first.")
    path = safe_path(rel)
    if not path.is_file(): raise FileNotFoundError(f"File not found: {rel}")
    if extensions and path.suffix.lower() not in extensions: raise ValueError("This tool does not support that file type.")
    return path


def csv_data(path: Path):
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows: raise ValueError("This CSV file is empty.")
    return rows


def csv_string(rows):
    out = io.StringIO(newline="")
    writer = csv.writer(out, lineterminator="\n")
    writer.writerows(rows)
    return out.getvalue()



def run_extra_tool(action, data, values, rel, name):
    """Second wave of focused tools. Returns None for actions handled by run_tool."""
    # File and workspace tools.
    if action in {"create_workspace_readme","clean_empty_folders","largest_files","recent_files","oldest_files","file_inventory","regex_search","export_file_hashes","create_zip","extract_zip","batch_suffix","batch_extension","sanitize_names","normalize_encoding","backup_now","backup_inventory","folder_tree","file_age_report","empty_files","clear_tags"}:
        if action == "create_workspace_readme":
            outname=name or "Workspace README.md"
            body="# Workspace guide\n\nThis folder is managed locally by OfficeMate.\n\n- Files stay on this device.\n- Backups are stored in the hidden `.backups` folder.\n- Actions are recorded in the local audit trail.\n"
            out=tool_output(outname,body,"create workspace guide");return {"message":f"Created {out}.","path":out}
        if action == "clean_empty_folders":
            if not data.get("confirm"):raise PermissionError("Confirm removing empty folders first.")
            dirs=sorted((p for p in WORKSPACE.rglob("*") if p.is_dir() and not any(x.startswith(".") for x in p.relative_to(WORKSPACE).parts)),key=lambda p:len(p.parts),reverse=True);removed=[]
            for p in dirs:
                try:p.rmdir();removed.append(p.relative_to(WORKSPACE).as_posix());audit("remove empty folder",p.relative_to(WORKSPACE).as_posix())
                except OSError:pass
            return {"message":f"Removed {len(removed)} empty folder(s).\n"+"\n".join(removed[:50])}
        if action in {"largest_files","recent_files","oldest_files"}:
            rows=list_files();n=max(1,min(100,int(name or "10")))
            key={"largest_files":lambda x:x["size"],"recent_files":lambda x:x["modified"],"oldest_files":lambda x:x["modified"]}[action]
            rows=sorted(rows,key=key,reverse=action!="oldest_files")[:n];audit(action,"workspace",f"{len(rows)} results")
            return {"message":"\n".join(f"{r['path']} · {pretty_bytes(r['size'])} · {r['modified']}" for r in rows) or "No files in the workspace."}
        if action in {"file_inventory","export_file_hashes"}:
            rows=[["path","type","size_bytes","modified"]]
            for f in list_files():
                row=[f["path"],f["kind"],f["size"],f["modified"]]
                if action=="export_file_hashes":row.append(hashlib.sha256(safe_path(f["path"]).read_bytes()).hexdigest())
                rows[0]=["path","type","size_bytes","modified","sha256"] if action=="export_file_hashes" else rows[0]
                rows.append(row)
            out=tool_output(name or ("File checksums.csv" if action=="export_file_hashes" else "File inventory.csv"),csv_string(rows),action.replace("_"," "))
            return {"message":f"Created {out} with {len(rows)-1} file records.","path":out}
        if action == "regex_search":
            try:pattern=re.compile(values.get("query", ""),re.I)
            except re.error as e:raise ValueError(f"Invalid regular expression: {e}")
            if not values.get("query"):raise ValueError("Enter a regular expression.")
            hits=[]
            for f in list_files():
                if f["ext"] not in {".txt",".md",".csv",".json",".docx",".pptx",".xlsx"}:continue
                try:
                    for i,line in enumerate(file_text(safe_path(f["path"])).splitlines(),1):
                        if pattern.search(line):hits.append(f"{f['path']}:{i}: {line.strip()[:180]}")
                except Exception:pass
            audit("regex search","workspace",values["query"])
            return {"message":"\n".join(hits[:100]) if hits else "No matches found."}
        if action == "create_zip":
            paths=[x.strip() for x in values.get("query","").split(",") if x.strip()]
            if not paths:raise ValueError("Enter one or more workspace file paths.")
            buf=io.BytesIO()
            with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as z:
                for item in paths:
                    p=require_file(item);z.write(p,p.relative_to(WORKSPACE).as_posix())
            outname=name or "OfficeMate archive.zip"
            if not outname.lower().endswith(".zip"):outname+=".zip"
            out=tool_output(outname,buf.getvalue(),"create ZIP archive");return {"message":f"Archived {len(paths)} file(s) into {out}.","path":out}
        if action == "extract_zip":
            src=require_file(rel,{".zip"});folder=name or f"{src.stem}_extracted"
            if Path(folder).name!=folder:raise ValueError("Enter a simple destination folder name.")
            dest=safe_path(folder)
            if dest.exists():raise FileExistsError(f"Folder {folder} already exists.")
            total=0
            with zipfile.ZipFile(src) as z:
                infos=z.infolist()
                if len(infos)>1000:raise ValueError("Archive contains too many entries (limit 1,000).")
                for info in infos:
                    clean=info.filename.replace("\\","/");pp=Path(clean)
                    if pp.is_absolute() or ".." in pp.parts or clean.startswith("/") or any(part.startswith(".") for part in pp.parts):raise ValueError("Archive contains an unsafe or hidden path; nothing was extracted.")
                    mode=info.external_attr>>16
                    if mode & 0o170000 == 0o120000:raise ValueError("Archive contains a symbolic link; nothing was extracted.")
                    total+=info.file_size
                    if total>100*1024*1024:raise ValueError("Archive exceeds the 100 MB extraction limit.")
                dest.mkdir(parents=True)
                for info in infos:
                    target=(dest/info.filename.replace("\\","/")).resolve()
                    if dest.resolve() not in target.parents and target!=dest.resolve():raise ValueError("Unsafe archive path.")
                    if info.is_dir():target.mkdir(parents=True,exist_ok=True)
                    else:
                        target.parent.mkdir(parents=True,exist_ok=True)
                        with z.open(info) as source, target.open("wb") as output:shutil.copyfileobj(source,output)
            audit("extract ZIP",rel,f"to {folder}");return {"message":f"Extracted {len(infos)} archive entries into {folder}/."}
        if action in {"batch_suffix","batch_extension","sanitize_names"}:
            if not data.get("confirm"):raise PermissionError("Confirm this batch rename before continuing.")
            changed=[]
            for f in list_files():
                src=safe_path(f["path"])
                if src.parent!=WORKSPACE:continue
                if action=="batch_suffix":
                    suffix=name
                    if not suffix or any(c in suffix for c in "/\\\\"):raise ValueError("Enter a suffix without slashes.")
                    dest=src.with_name(src.stem+suffix+src.suffix)
                elif action=="batch_extension":
                    old=values.get("name","");new=values.get("query","")
                    if not old or not new:raise ValueError("Enter both the current and new extension.")
                    old=old if old.startswith(".") else "."+old;new=new if new.startswith(".") else "."+new
                    if not re.fullmatch(r"[.][A-Za-z0-9]{1,12}",old) or not re.fullmatch(r"[.][A-Za-z0-9]{1,12}",new):raise ValueError("Extensions must contain only letters and numbers.")
                    if src.suffix.lower()!=old.lower():continue
                    dest=src.with_suffix(new)
                else:
                    safe=re.sub(r"[^A-Za-z0-9._ -]+","_",src.name).strip(" .") or "file"
                    dest=src.with_name(safe)
                if dest==src:continue
                if dest.exists():continue
                backup(src);src.rename(dest);changed.append(f"{src.name} → {dest.name}");audit(action,f["path"],f"to {dest.name}")
            return {"message":f"Renamed {len(changed)} file(s).\n"+"\n".join(changed[:50])}
        if action == "normalize_encoding":
            src=require_file(rel,{".txt",".md",".csv",".json",".html",".htm"});text=src.read_bytes().decode("utf-8",errors="replace");outname=name or f"{src.stem}_utf8{src.suffix}"
            out=tool_output(outname,text,"normalize encoding");return {"message":f"Created UTF-8 copy {out}.","path":out}
        if action == "backup_now":
            src=require_file(rel);dest=backup(src);audit("manual backup",rel,dest.name)
            return {"message":f"Snapshot saved: {dest.name}"}
        if action == "backup_inventory":
            src=require_file(rel);choices=sorted(BACKUPS.glob(src.name+".*.bak"),key=lambda p:p.stat().st_mtime,reverse=True);audit("list backups",rel)
            return {"message":"\n".join(f"{p.name} · {dt.datetime.fromtimestamp(p.stat().st_mtime).astimezone().isoformat(timespec='minutes')} · {pretty_bytes(p.stat().st_size)}" for p in choices) or "No snapshots found."}
        if action == "folder_tree":
            lines=[WORKSPACE.name+"/"]
            for p in sorted(WORKSPACE.rglob("*")):
                if any(x.startswith(".") for x in p.relative_to(WORKSPACE).parts):continue
                depth=len(p.relative_to(WORKSPACE).parts);lines.append("  "*(depth-1)+("📁 " if p.is_dir() else "• ")+p.name)
            audit("folder tree","workspace");return {"message":"\n".join(lines[:300])}
        if action == "file_age_report":
            today=dt.datetime.now().timestamp();bins=Counter()
            for f in list_files():
                days=max(0,int((today-safe_path(f["path"]).stat().st_mtime)/86400));label="< 1 day" if days<1 else "1–7 days" if days<7 else "8–30 days" if days<30 else "31–90 days" if days<90 else "> 90 days";bins[label]+=1
            audit("file age report","workspace");return {"message":"Files by age\n"+"\n".join(f"{k}: {v}" for k,v in bins.items())}
        if action == "empty_files":
            rows=[f["path"] for f in list_files() if f["size"]==0];audit("find empty files","workspace",str(len(rows)))
            return {"message":"\n".join(rows) if rows else "No empty files found."}
        if action == "clear_tags":
            if not data.get("confirm"):raise PermissionError("Confirm clearing these tags first.")
            require_file(rel);meta={}
            try:meta=json.loads((WORKSPACE/".officemate-meta.json").read_text())
            except Exception:pass
            if rel in meta:meta[rel]["tags"]=[]
            (WORKSPACE/".officemate-meta.json").write_text(json.dumps(meta,indent=2));audit("clear tags",rel)
            return {"message":f"Cleared tags for {rel}."}
    # Text analytics and transforms.
    text_actions={"line_count","sentence_count","paragraph_count","find_urls","extract_emails","extract_phone_numbers","extract_dates","redact_emails","redact_urls","clean_punctuation","normalize_unicode","remove_accents","remove_html_tags","normalize_newlines","indent_text","wrap_text","number_lines","remove_line_numbers","extract_bullets","remove_bullets","keyword_count","top_keywords","generate_toc","reverse_words","join_lines"}
    if action in text_actions:
        src=require_file(rel);text=file_text(src);stem=src.stem
        if action=="line_count":audit(action,rel);return {"message":f"{len(text.splitlines()):,} lines"}
        if action=="sentence_count":
            n=len([x for x in re.split(r"(?<=[.!?])\s+",text.strip()) if x]);audit(action,rel);return {"message":f"{n:,} sentences (estimate)"}
        if action=="paragraph_count":
            n=len([x for x in re.split(r"\n\s*\n",text.strip()) if x.strip()]);audit(action,rel);return {"message":f"{n:,} paragraphs"}
        if action=="find_urls":res=re.findall(r"https?://[^\s<>()]+",text);audit(action,rel);return {"message":"\n".join(dict.fromkeys(res)) or "No URLs found."}
        if action=="extract_emails":res=re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",text);audit(action,rel);return {"message":"\n".join(dict.fromkeys(res)) or "No email addresses found."}
        if action=="extract_phone_numbers":res=re.findall(r"(?<!\w)(?:\+?\d[\d .()-]{7,}\d)(?!\w)",text);audit(action,rel);return {"message":"\n".join(dict.fromkeys(x.strip() for x in res)) or "No phone-like values found."}
        if action=="extract_dates":res=re.findall(r"\b(?:\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}|(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2},?\s+\d{4})\b",text,re.I);audit(action,rel);return {"message":"\n".join(dict.fromkeys(res)) or "No common date patterns found."}
        if action=="keyword_count":
            q=values.get("query","")
            if not q:raise ValueError("Enter a keyword or phrase.")
            n=len(re.findall(re.escape(q),text,re.I));audit(action,rel,q);return {"message":f"{q}: {n} occurrence(s)"}
        if action=="top_keywords":
            stop={"this","that","with","from","have","were","will","your","about","there","their","which","been","would","could","should","into","when","what","then","than","them","they","also","some","such","more","most","only","over","under","each","very","because","while","where","after","before"}
            counts=Counter(w.lower() for w in re.findall(r"[A-Za-z][A-Za-z'-]{2,}",text) if w.lower() not in stop);n=max(1,min(50,int(name or "10")));audit(action,rel)
            return {"message":"\n".join(f"{w}: {c}" for w,c in counts.most_common(n)) or "No keywords found."}
        if action=="generate_toc":
            heads=re.findall(r"^(#{1,6})\s+(.+)$",text,re.M);lines=["## Contents",""]+["  "*(len(h)-1)+f"- [{title}](#{re.sub(r'[^a-z0-9 -]','',title.lower()).replace(' ','-')})" for h,title in heads];out=tool_output(f"{stem}_toc.md","\n".join(lines)+"\n","generate table of contents");return {"message":f"Created {out}.","path":out}
        if action=="find_urls":pass
        if action=="redact_emails":new=re.sub(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}","[REDACTED EMAIL]",text);suffix="redacted_emails"
        elif action=="redact_urls":new=re.sub(r"https?://[^\s<>()]+","[REDACTED URL]",text);suffix="redacted_urls"
        elif action=="clean_punctuation":new=re.sub(r"([!?.,])\1{1,}",r"\1",text);suffix="clean_punctuation"
        elif action=="normalize_unicode":new=unicodedata.normalize("NFKC",text);suffix="normalized"
        elif action=="remove_accents":new="".join(c for c in unicodedata.normalize("NFKD",text) if not unicodedata.combining(c));suffix="noaccents"
        elif action=="remove_html_tags":new=re.sub(r"<[^>]*>","",text);new=html.unescape(new);suffix="stripped_html"
        elif action=="normalize_newlines":new=text.replace("\r\n","\n").replace("\r","\n");suffix="lf"
        elif action=="indent_text":
            try:n=max(0,min(32,int(name or "2")))
            except ValueError:raise ValueError("Indent must be a number.")
            new="\n".join((" "*n+x if x.strip() else "") for x in text.splitlines())+"\n";suffix="indented"
        elif action=="wrap_text":
            try:n=max(20,min(240,int(name or "80")))
            except ValueError:raise ValueError("Line width must be a number.")
            new="\n".join(textwrap.fill(line,width=n) if line.strip() else "" for line in text.splitlines())+"\n";suffix="wrapped"
        elif action=="number_lines":new="\n".join(f"{i:04d}: {line}" for i,line in enumerate(text.splitlines(),1))+"\n";suffix="numbered"
        elif action=="remove_line_numbers":new=re.sub(r"(?m)^\s*\d+[.:)]\s?","",text);suffix="unnumbered"
        elif action=="extract_bullets":
            items=[re.sub(r"^\s*(?:[-*+]\s+|\d+[.)]\s+)","",x).strip() for x in text.splitlines() if re.match(r"^\s*(?:[-*+]\s+|\d+[.)]\s+)",x)];new="# Extracted list items\n\n"+"\n".join("- "+x for x in items)+"\n";suffix="bullets"
        elif action=="remove_bullets":new=re.sub(r"(?m)^\s*(?:[-*+]\s+|\d+[.)]\s+)","",text);suffix="nobullets"
        elif action=="reverse_words":new="\n".join(" ".join(reversed(line.split())) for line in text.splitlines())+"\n";suffix="reversed_words"
        elif action=="join_lines":sep=name if name else " ";new=sep.join(x.strip() for x in text.splitlines() if x.strip())+"\n";suffix="joined"
        else:raise ValueError("Unsupported text tool.")
        out=tool_output(f"{stem}_{suffix}.txt",new,action.replace("_"," "));return {"message":f"Created {out}.","path":out}
    # CSV transformations and reports.
    csv_actions={"csv_columns","csv_null_report","csv_fill_blanks","csv_drop_columns","csv_select_columns","csv_rename_column","csv_add_column","csv_drop_empty_rows","csv_limit_rows","csv_sample_rows","csv_search_rows","csv_replace_values","csv_unique_values","csv_value_counts","csv_group_sum","csv_to_json","json_to_csv","csv_validate_email","csv_validate","csv_summary_report","csv_sort_numeric","csv_split_column","csv_unpivot","csv_add_row_number","csv_column_stats"}
    if action in csv_actions:
        src=require_file(rel,{".csv",".json"} if action=="json_to_csv" else {".csv"})
        if action=="json_to_csv":
            try:obj=json.loads(src.read_text(encoding="utf-8-sig"))
            except json.JSONDecodeError as e:raise ValueError(f"Invalid JSON: {e}")
            if not isinstance(obj,list) or any(not isinstance(x,dict) for x in obj):raise ValueError("Expected a JSON array of objects.")
            headers=list(dict.fromkeys(k for row in obj for k in row));rows=[headers]+[[row.get(h,"") for h in headers] for row in obj];out=tool_output(name or f"{src.stem}.csv",csv_string(rows),"JSON to CSV");return {"message":f"Created {out}.","path":out}
        rows=csv_data(src);header=rows[0];body=rows[1:];width=max(map(len,rows))
        if action=="csv_columns":audit(action,rel);return {"message":"\n".join(f"{i+1}. {x}" for i,x in enumerate(header))}
        if action=="csv_null_report":
            counts=[sum(1 for r in body if i>=len(r) or not r[i].strip()) for i in range(len(header))];audit(action,rel);return {"message":"\n".join(f"{h}: {n} blank(s)" for h,n in zip(header,counts))}
        if action=="csv_validate":
            bad=[i+2 for i,r in enumerate(body) if len(r)!=len(header)];dups=[x for x,c in Counter(header).items() if c>1];audit(action,rel)
            return {"message":f"Header columns: {len(header)}\nData rows: {len(body)}\nRows with inconsistent column counts: {len(bad)}"+(f" (lines {bad[:20]})" if bad else "")+"\nDuplicate headers: "+(", ".join(dups) if dups else "none")}
        if action=="csv_validate_email":
            col=name
            if col not in header:raise ValueError(f"Column not found. Available: {', '.join(header)}")
            i=header.index(col);bad=[(n+2,r[i] if i<len(r) else "") for n,r in enumerate(body) if i>=len(r) or not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+",r[i])];audit(action,rel,col)
            return {"message":("Invalid email rows:\n"+"\n".join(f"Row {n}: {v}" for n,v in bad[:100])) if bad else "All non-empty values look like valid email addresses."}
        if action=="csv_column_stats":
            col=name
            if col not in header:raise ValueError(f"Column not found. Available: {', '.join(header)}")
            vals=[r[header.index(col)] for r in body if header.index(col)<len(r) and r[header.index(col)]!=""];audit(action,rel,col)
            return {"message":f"Column: {col}\nNon-empty: {len(vals)}\nDistinct: {len(set(vals))}\nShortest: {min((len(x) for x in vals),default=0)} chars\nLongest: {max((len(x) for x in vals),default=0)} chars"}
        if action=="csv_unique_values" or action=="csv_value_counts":
            col=name
            if col not in header:raise ValueError(f"Column not found. Available: {', '.join(header)}")
            i=header.index(col);counts=Counter(r[i] if i<len(r) else "" for r in body);audit(action,rel,col)
            result=counts.most_common(100) if action=="csv_value_counts" else [(v,1) for v in dict.fromkeys(r[i] if i<len(r) else "" for r in body)]
            return {"message":"\n".join(f"{v}: {n}" if action=="csv_value_counts" else v for v,n in result) or "No values found."}
        if action=="csv_summary_report":
            blanks=sum(1 for r in body for i in range(len(header)) if i>=len(r) or not r[i].strip());md=f"# CSV Summary: {src.name}\n\n- Data rows: {len(body)}\n- Columns: {len(header)}\n- Blank cells: {blanks}\n\n## Columns\n"+"\n".join(f"- **{h}**" for h in header);out=tool_output(f"{src.stem}_summary.md",md,"CSV summary report");return {"message":f"Created {out}.","path":out}
        if action=="csv_to_json":
            objs=[dict(zip(header,(r+[""]*len(header))[:len(header)])) for r in body];out=tool_output(f"{src.stem}.json",json.dumps(objs,ensure_ascii=False,indent=2),"CSV to JSON");return {"message":f"Created {out}.","path":out}
        if action=="csv_group_sum":
            numeric=name;group=values.get("query","")
            if group not in header or numeric not in header:raise ValueError(f"Select valid columns. Available: {', '.join(header)}")
            gi,ni=header.index(group),header.index(numeric);sums=Counter()
            for row in body:
                if max(gi,ni)>=len(row):continue
                try:sums[row[gi]]+=float(row[ni].replace(",",""))
                except ValueError:continue
            return {"message":f"{group} → sum({numeric})\n"+"\n".join(f"{k}: {v:g}" for k,v in sums.items())}
        if action=="csv_add_row_number":
            rows=[header+["row_number"]]+[r+[str(i)] for i,r in enumerate(body,1)];suffix="numbered"
        elif action=="csv_fill_blanks":
            val=values.get("replacement","");rows=[header]+[[c if c.strip() else val for c in r] for r in body];suffix="filled"
        elif action=="csv_drop_columns" or action=="csv_select_columns":
            cols=[x.strip() for x in name.split(",") if x.strip()]
            if not cols:raise ValueError("Enter one or more column names.")
            missing=[x for x in cols if x not in header]
            if missing:raise ValueError("Unknown columns: "+", ".join(missing))
            keep=cols if action=="csv_select_columns" else [x for x in header if x not in cols];idx=[header.index(x) for x in keep];rows=[[header[i] for i in idx]]+[[r[i] if i<len(r) else "" for i in idx] for r in body];suffix="selected" if action=="csv_select_columns" else "dropped"
        elif action=="csv_rename_column":
            old=values.get("query","");new=values.get("replacement","")
            if old not in header or not new:raise ValueError("Enter an existing header and a new name.")
            header[header.index(old)]=new;rows=[header]+body;suffix="renamed"
        elif action=="csv_add_column":
            if not name or name in header:raise ValueError("Enter a new, unique column name.")
            rows=[header+[name]]+[r+[values.get("replacement","")] for r in body];suffix="added"
        elif action=="csv_drop_empty_rows":
            rows=[header]+[r for r in body if any(c.strip() for c in r)];suffix="nonempty"
        elif action=="csv_limit_rows" or action=="csv_sample_rows":
            try:n=max(0,min(len(body),int(name)))
            except ValueError:raise ValueError("Enter a whole number of rows.")
            if action=="csv_limit_rows":subset=body[:n];suffix="limited"
            else:
                import random
                subset=random.Random(0).sample(body,n);suffix="sample"
            rows=[header]+subset
        elif action=="csv_search_rows":
            q=values.get("query","");
            if not q:raise ValueError("Enter a search phrase.")
            rows=[header]+[r for r in body if any(q.lower() in c.lower() for c in r)];suffix="search"
        elif action=="csv_replace_values":
            q=values.get("query","");replacement=values.get("replacement","")
            if not q:raise ValueError("Enter the value to find.")
            rows=[[c.replace(q,replacement) for c in r] for r in rows];suffix="replaced"
        elif action=="csv_sort_numeric":
            if name not in header:raise ValueError("Column not found.")
            i=header.index(name)
            def numeric_key(r):
                try:return (0,float(r[i].replace(",","")))
                except (ValueError,IndexError):return (1,0)
            rows=[header]+sorted(body,key=numeric_key);suffix="sorted_numeric"
        elif action=="csv_split_column":
            if name not in header:raise ValueError("Column not found.")
            i=header.index(name);sep=values.get("query",",") or ",";max_parts=max((len(r[i].split(sep)) for r in body if i<len(r)),default=1);newheaders=header[:i]+[f"{name}_{j+1}" for j in range(max_parts)]+header[i+1:];newbody=[]
            for r in body:
                r=r+[""]*max(0,len(header)-len(r));parts=r[i].split(sep);newbody.append(r[:i]+parts+[""]*(max_parts-len(parts))+r[i+1:])
            rows=[newheaders]+newbody;suffix="split"
        elif action=="csv_unpivot":
            ids=[x.strip() for x in name.split(",") if x.strip()]
            if not ids or any(x not in header for x in ids):raise ValueError(f"Enter identifier column names from: {', '.join(header)}")
            idx=[header.index(x) for x in ids];other=[i for i in range(len(header)) if i not in idx];rows=[ids+["attribute","value"]]
            for r in body:
                r=r+[""]*max(0,len(header)-len(r));rows.extend([[r[i] for i in idx]+[header[j],r[j]] for j in other])
            suffix="unpivoted"
        else:raise ValueError("Unsupported CSV action.")
        out=tool_output(f"{src.stem}_{suffix}.csv",csv_string(rows),action.replace("_"," "));return {"message":f"Created {out} with {max(0,len(rows)-1)} data rows.","path":out}
    # Office file operations, backed up before any in-place change.
    office_actions={x[0] for x in EXTRA_FEATURES if x[1]=="Office" and x[0] not in {"office_format_report","batch_office_to_text","batch_office_backup"}}
    if action in office_actions:
        src=require_file(rel)
        if action.startswith("docx_"):
            if src.suffix.lower()!=".docx":raise ValueError("Choose a DOCX file.")
            try:from docx import Document
            except ImportError:raise ValueError("Install python-docx for Word tools: pip install python-docx")
            doc=Document(src)
            if action=="docx_outline":
                heads=[(p.style.name,p.text) for p in doc.paragraphs if p.style and p.style.name.startswith("Heading") and p.text];return {"message":"\n".join(f"{s}: {t}" for s,t in heads) or "No heading styles found."}
            if action=="docx_count":
                txt=" ".join(p.text for p in doc.paragraphs);words=len(re.findall(r"\b\w+\b",txt));return {"message":f"Words: {words}\nParagraphs: {len(doc.paragraphs)}\nTables: {len(doc.tables)}\nSections: {len(doc.sections)}"}
            if action=="docx_tables_csv":
                rows=[]
                for ti,table in enumerate(doc.tables,1):
                    if ti>1:rows.append([])
                    rows.append([f"Table {ti}"])
                    rows.extend([[c.text.replace("\n"," ") for c in row.cells] for row in table.rows])
                if not rows:raise ValueError("No tables found in this Word document.")
                out=tool_output(name or f"{src.stem}_tables.csv",csv_string(rows),"export Word tables");return {"message":f"Created {out}.","path":out}
            if action=="docx_merge":
                other=require_file(values.get("second", ""),{".docx"});doc2=Document(other);merged=Document()
                for source in (doc,doc2):
                    for p in source.paragraphs:merged.add_paragraph(p.text,style=p.style.name if p.style else None)
                    for table in source.tables:
                        t=merged.add_table(rows=len(table.rows),cols=len(table.columns))
                        for ri,row in enumerate(table.rows):
                            for ci,c in enumerate(row.cells):t.cell(ri,ci).text=c.text
                outname=name or f"{src.stem}_merged.docx";buf=io.BytesIO();merged.save(buf);out=tool_output(outname,buf.getvalue(),"merge Word documents");return {"message":f"Created {out}.","path":out}
            if action in {"docx_replace","docx_append","docx_set_metadata"}:
                if not data.get("confirm"):raise PermissionError("Confirm this Word document edit first.")
                backup(src)
                if action=="docx_replace":
                    q=values.get("query","");repl=values.get("replacement","")
                    if not q:raise ValueError("Enter text to find.")
                    count=0
                    for p in doc.paragraphs:
                        if q in p.text:p.text=p.text.replace(q,repl);count+=1
                    for t in doc.tables:
                        for row in t.rows:
                            for cell in row.cells:
                                if q in cell.text:cell.text=cell.text.replace(q,repl);count+=1
                    detail=f"Replaced text in {count} paragraph/cell(s)."
                elif action=="docx_append":doc.add_paragraph(str(data.get("text", "")));detail="Appended paragraph."
                else:doc.core_properties.title=name;doc.core_properties.subject=values.get("query","");detail="Updated properties."
                doc.save(src);audit(action,rel,detail);return {"message":f"{detail} Backup saved."}
        if action.startswith("xlsx_"):
            if src.suffix.lower()!=".xlsx":raise ValueError("Choose an XLSX workbook.")
            try:from openpyxl import load_workbook
            except ImportError:raise ValueError("Install openpyxl for Excel tools: pip install openpyxl")
            wb=load_workbook(src)
            if action=="xlsx_preview":
                ws=wb[name] if name and name in wb.sheetnames else wb.active;lines=[f"{ws.title} · {ws.max_row} rows × {ws.max_column} columns"]
                lines.extend(" | ".join("" if v is None else str(v) for v in row) for row in ws.iter_rows(min_row=1,max_row=min(ws.max_row,20),values_only=True));return {"message":"\n".join(lines)}
            if action in {"xlsx_export_sheet_csv","xlsx_export_all_csv"}:
                sheets=[wb[name] if name else wb.active] if action=="xlsx_export_sheet_csv" and (not name or name in wb.sheetnames) else wb.worksheets if action=="xlsx_export_all_csv" else []
                if not sheets:raise ValueError("Select an existing sheet name." if action=="xlsx_export_sheet_csv" else "No sheets found.")
                made=[]
                for ws in sheets:
                    rows=[["" if v is None else v for v in row] for row in ws.iter_rows(values_only=True)];out=tool_output(f"{src.stem}_{re.sub(r'[^A-Za-z0-9_-]','_',ws.title)}.csv",csv_string(rows),"export worksheet CSV");made.append(out)
                return {"message":"Created CSV exports:\n"+"\n".join(made),"path":made[0] if len(made)==1 else ""}
            if action=="xlsx_find_text":
                q=values.get("query","");hits=[]
                for ws in wb.worksheets:
                    for row in ws.iter_rows():
                        for cell in row:
                            if q.lower() in str(cell.value or "").lower():hits.append(f"{ws.title}!{cell.coordinate}: {cell.value}")
                return {"message":"\n".join(hits[:100]) if hits else "No matching cells found."}
            if action=="xlsx_metadata":return {"message":f"Sheets: {len(wb.sheetnames)}\nNames: {', '.join(wb.sheetnames)}\nCreator: {wb.properties.creator or '—'}\nTitle: {wb.properties.title or '—'}"}
            if action=="xlsx_remove_empty_rows":
                for ws in wb.worksheets:
                    for i in range(ws.max_row,0,-1):
                        if all(c.value is None for c in ws[i]):ws.delete_rows(i)
                outname=name or f"{src.stem}_cleaned.xlsx";buf=io.BytesIO();wb.save(buf);out=tool_output(outname,buf.getvalue(),"remove empty Excel rows");return {"message":f"Created {out}.","path":out}
            mutating={"xlsx_replace_value","xlsx_add_sheet","xlsx_rename_sheet","xlsx_delete_empty_sheets","xlsx_style_header","xlsx_freeze_header"}
            if action in mutating:
                if not data.get("confirm"):raise PermissionError("Confirm this workbook edit first.")
                backup(src)
                if action=="xlsx_replace_value":
                    q=values.get("query","");count=0
                    for ws in wb.worksheets:
                        for row in ws.iter_rows():
                            for c in row:
                                if c.value is not None and str(c.value)==q:c.value=values.get("replacement","");count+=1
                    detail=f"Replaced {count} matching cell(s)."
                elif action=="xlsx_add_sheet":wb.create_sheet(name);detail=f"Added worksheet {name}."
                elif action=="xlsx_rename_sheet":
                    if name not in wb.sheetnames:raise ValueError("Current sheet name not found.")
                    wb[name].title=values.get("query","");detail=f"Renamed sheet {name}."
                elif action=="xlsx_delete_empty_sheets":
                    removed=[]
                    for ws in list(wb.worksheets):
                        if ws.max_row==1 and ws.max_column==1 and ws["A1"].value is None and len(wb.worksheets)>1:removed.append(ws.title);wb.remove(ws)
                    detail=f"Removed {len(removed)} empty sheet(s)."
                elif action=="xlsx_style_header":
                    from openpyxl.styles import Font, PatternFill
                    for ws in wb.worksheets:
                        for c in ws[1]:c.font=Font(bold=True,color="FFFFFF");c.fill=PatternFill("solid",fgColor="477354")
                    detail="Formatted first-row headers."
                else:
                    for ws in wb.worksheets:ws.freeze_panes="A2"
                    detail="Froze header rows."
                wb.save(src);audit(action,rel,detail);return {"message":f"{detail} Backup saved."}
        if action.startswith("pptx_"):
            if src.suffix.lower()!=".pptx":raise ValueError("Choose a PPTX presentation.")
            try:from pptx import Presentation
            except ImportError:raise ValueError("Install python-pptx for PowerPoint tools: pip install python-pptx")
            prs=Presentation(src)
            if action=="pptx_slide_count":return {"message":f"Slides: {len(prs.slides)}\nShapes: {sum(len(s.shapes) for s in prs.slides)}\nSpeaker notes: supported by installed python-pptx"}
            if action=="pptx_export_outline":
                lines=[]
                for i,slide in enumerate(prs.slides,1):
                    title=slide.shapes.title.text if slide.shapes.title else f"Slide {i}";lines.append(f"## {title}")
                    lines.extend("- "+sh.text.replace("\n"," ") for sh in slide.shapes if getattr(sh,"has_text_frame",False) and sh!=slide.shapes.title and sh.text.strip())
                    lines.append("")
                out=tool_output(f"{src.stem}_outline.md","\n".join(lines),"export presentation outline");return {"message":f"Created {out}.","path":out}
            if action=="pptx_export_notes":
                lines=[]
                for i,slide in enumerate(prs.slides,1):
                    notes=""
                    try:notes=slide.notes_slide.notes_text_frame.text
                    except Exception:pass
                    lines.extend([f"Slide {i}",notes.strip(),""])
                out=tool_output(f"{src.stem}_notes.txt","\n".join(lines),"export speaker notes");return {"message":f"Created {out}.","path":out}
            if action in {"pptx_replace_text","pptx_set_metadata"}:
                if not data.get("confirm"):raise PermissionError("Confirm this presentation edit first.")
                backup(src)
                if action=="pptx_replace_text":
                    q=values.get("query","")
                    if not q:raise ValueError("Enter text to find.")
                    count=0
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if getattr(shape,"has_text_frame",False):
                                for para in shape.text_frame.paragraphs:
                                    for run in para.runs:
                                        if q in run.text:run.text=run.text.replace(q,values.get("replacement", ""));count+=1
                    detail=f"Replaced text in {count} run(s)."
                else:prs.core_properties.title=name;prs.core_properties.subject=values.get("query","");detail="Updated presentation properties."
                prs.save(src);audit(action,rel,detail);return {"message":f"{detail} Backup saved."}
        if action.startswith("pdf_"):
            if src.suffix.lower()!=".pdf":raise ValueError("Choose a PDF file.")
            try:from pypdf import PdfReader
            except ImportError:raise ValueError("Install pypdf for PDF tools: pip install pypdf")
            reader=PdfReader(str(src));meta=reader.metadata or {}
            if action=="pdf_page_count":return {"message":f"{len(reader.pages)} pages"}
            if action=="pdf_metadata":return {"message":f"Pages: {len(reader.pages)}\nTitle: {meta.get('/Title') or '—'}\nAuthor: {meta.get('/Author') or '—'}\nCreator: {meta.get('/Creator') or '—'}"}
            text="\n\n".join(p.extract_text() or "" for p in reader.pages);out=tool_output(name or f"{src.stem}_extracted.txt",text,"extract PDF text");return {"message":f"Created {out} with selectable PDF text.","path":out}
    # Reports and batch utilities.
    if action=="office_format_report":
        rows=list_files();counts=Counter(f["ext"] for f in rows if f["ext"] in {".docx",".xlsx",".pptx",".pdf"});audit(action,"workspace");return {"message":"Office files by format\n"+"\n".join(f"{k}: {counts.get(k,0)}" for k in [".docx",".xlsx",".pptx",".pdf"])}
    if action=="batch_office_to_text":
        made=[]
        for f in list_files():
            src=safe_path(f["path"])
            if src.suffix.lower() not in {".docx",".xlsx",".pptx"}:continue
            outname=f"{src.stem}_extracted.txt"
            extracted=file_text(src)
            if extracted.startswith("Install python-"):continue
            try:out=tool_output(outname,extracted,"batch extract text");made.append(out)
            except FileExistsError:continue
        audit(action,"workspace",f"{len(made)} extracted");return {"message":f"Extracted text from {len(made)} Office file(s).\n"+"\n".join(made)}
    if action=="batch_office_backup":
        backed=[]
        for f in list_files():
            src=safe_path(f["path"])
            if src.suffix.lower() in {".docx",".xlsx",".pptx",".pdf"}:
                b=backup(src);backed.append(src.name);audit("batch backup",f["path"],b.name)
        return {"message":f"Created {len(backed)} backup snapshot(s).\n"+"\n".join(backed)}
    return None

def run_tool(data):
    action = str(data.get("action", ""))
    if action not in {item[0] for item in FEATURES}: raise ValueError("Unknown tool.")
    values = {k: str(v or "").strip() for k, v in data.items()}
    rel = values.get("path", "")
    name = values.get("name", "")
    extra = run_extra_tool(action, data, values, rel, name)
    if extra is not None: return extra
    # Workspace and filesystem tools.
    if action == "create_folder":
        if not name or Path(name).name != name: raise ValueError("Enter a simple folder name.")
        dest=safe_path(name)
        if dest.exists(): raise FileExistsError("That folder already exists.")
        dest.mkdir(); audit("create folder", name); return {"message": f"Created folder {name}/."}
    if action in {"rename", "duplicate", "delete", "move", "tag", "pin", "restore_backup", "file_hash"}:
        src=safe_path(rel) if action == "restore_backup" else require_file(rel)
        if action == "rename":
            if not data.get("confirm"): raise PermissionError("Confirm the rename first.")
            if not name or Path(name).name != name or name.startswith("."): raise ValueError("Enter a simple visible new filename.")
            dest=src.with_name(name)
            if dest.exists(): raise FileExistsError(f"{name} already exists.")
            backup(src); src.rename(dest); audit("rename", rel, f"to {dest.relative_to(WORKSPACE).as_posix()}"); return {"message": f"Renamed to {dest.name}.","path":dest.relative_to(WORKSPACE).as_posix()}
        if action == "duplicate":
            target=name or f"{src.stem} copy{src.suffix}"
            if Path(target).name != target: raise ValueError("Enter a simple copy filename.")
            dest=safe_path(target)
            if dest.exists(): raise FileExistsError(f"{target} already exists.")
            shutil.copy2(src,dest); audit("duplicate",rel,f"to {target}"); return {"message":f"Created {target}.","path":target}
        if action == "delete":
            if not data.get("confirm"): raise PermissionError("Confirm the delete first.")
            b=backup(src); trash=WORKSPACE/".trash";trash.mkdir(exist_ok=True)
            stamp=dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f");dest=trash/f"{stamp}-{src.name}"
            shutil.move(str(src),str(dest));audit("delete to trash",rel,f"backup: {b.name if b else 'none'}")
            return {"message":f"Moved {src.name} to local trash. A backup was saved."}
        if action == "move":
            folder=name
            if not folder or Path(folder).name!=folder: raise ValueError("Enter a simple folder name.")
            dest_dir=safe_path(folder)
            if not dest_dir.is_dir(): raise FileNotFoundError(f"Folder not found: {folder}")
            dest=dest_dir/src.name
            if dest.exists(): raise FileExistsError(f"{folder}/{src.name} already exists.")
            backup(src);shutil.move(str(src),str(dest));audit("move",rel,f"to {folder}/")
            return {"message":f"Moved to {folder}/{src.name}.","path":dest.relative_to(WORKSPACE).as_posix()}
        if action in {"tag","pin"}:
            meta={}
            if (WORKSPACE/".officemate-meta.json").exists():
                try: meta=json.loads((WORKSPACE/".officemate-meta.json").read_text())
                except Exception: pass
            entry=meta.setdefault(rel,{"tags":[],"pinned":False})
            if action=="tag": entry["tags"]=[t.strip() for t in name.split(",") if t.strip()]
            else: entry["pinned"]=not entry.get("pinned",False)
            (WORKSPACE/".officemate-meta.json").write_text(json.dumps(meta,indent=2),encoding="utf-8")
            audit(action,rel,", ".join(entry["tags"]) if action=="tag" else ("pinned" if entry["pinned"] else "unpinned"))
            return {"message":f"Updated {action} metadata for {src.name}."}
        if action == "restore_backup":
            if not data.get("confirm"): raise PermissionError("Confirm restoring the backup first.")
            choices=sorted(BACKUPS.glob(src.name+".*.bak"),key=lambda p:p.stat().st_mtime,reverse=True)
            if not choices: raise FileNotFoundError("No backup exists for this file.")
            backup(src);shutil.copy2(choices[0],src);audit("restore backup",rel,choices[0].name)
            return {"message":f"Restored the latest backup for {src.name}. The previous version was backed up."}
        if action == "file_hash":
            h=hashlib.sha256()
            with src.open("rb") as f:
                for chunk in iter(lambda:f.read(1024*1024),b""): h.update(chunk)
            audit("calculate checksum",rel,"SHA-256")
            return {"message":f"SHA-256\n{h.hexdigest()}"}
    if action == "export_audit":
        rows=[["timestamp","operation","file","detail"]]
        if AUDIT.exists():
            for line in AUDIT.read_text(encoding="utf-8").splitlines():
                try:
                    x=json.loads(line);rows.append([x.get(k,"") for k in rows[0]])
                except json.JSONDecodeError: pass
        out=tool_output("OfficeMate activity.csv",csv_string(rows),"export audit")
        return {"message":f"Exported the audit trail to {out}.","path":out}
    if action == "find_duplicates":
        buckets={}
        for x in list_files():
            p=safe_path(x["path"]);h=hashlib.sha256(p.read_bytes()).hexdigest();buckets.setdefault(h,[]).append(x["path"])
        groups=[v for v in buckets.values() if len(v)>1];audit("find duplicates","workspace",f"{len(groups)} groups")
        return {"message": "No identical files found." if not groups else "Identical content groups:\n"+"\n".join(" • "+" = ".join(g) for g in groups)}
    if action == "workspace_stats":
        rows=list_files();counts=Counter(x["kind"] for x in rows);sizes=Counter()
        for x in rows:sizes[x["kind"]]+=x["size"]
        audit("workspace report","workspace")
        lines=[f"Total files: {len(rows)}",f"Total size: {pretty_bytes(sum(x['size'] for x in rows))}"]+[f"{k}: {v} file(s), {pretty_bytes(sizes[k])}" for k,v in sorted(counts.items())]
        return {"message":"Workspace summary\n"+"\n".join(lines)}
    # Batch actions.
    if action == "batch_rename_prefix":
        if not name or any(c in name for c in "/\\"): raise ValueError("Enter a short prefix without slashes.")
        changed=[]
        for x in list_files():
            p=safe_path(x["path"])
            if p.parent!=WORKSPACE or p.name.startswith(name):continue
            dest=p.with_name(name+p.name)
            if dest.exists():continue
            backup(p);p.rename(dest);changed.append(dest.name);audit("batch rename",x["path"],f"to {dest.name}")
        return {"message":f"Prefixed {len(changed)} file(s).\n"+"\n".join(changed[:30])}
    if action == "batch_convert_pdf":
        soffice=shutil.which("soffice")
        if not soffice: raise ValueError("Install LibreOffice to batch convert files to PDF.")
        converted=[];failures=[]
        for x in list_files():
            src=safe_path(x["path"])
            if src.suffix.lower() not in {".docx",".xlsx",".pptx",".odt",".ods",".odp"}:continue
            dest=src.with_suffix(".pdf")
            if dest.exists():continue
            p=subprocess.run([soffice,"--headless","--convert-to","pdf","--outdir",str(src.parent),str(src)],capture_output=True,text=True,timeout=90)
            if p.returncode==0 and dest.exists(): converted.append(dest.name);audit("batch convert",x["path"],"to pdf")
            else:failures.append(src.name)
        return {"message":f"Converted {len(converted)} file(s) to PDF.\n"+("Could not convert: "+", ".join(failures) if failures else "")}
    # Text analysis, transforms, and Office extraction.
    if action in {"csv_create"}:
        content=data.get("text","");outname=name or "New data.csv"
        if not outname.lower().endswith(".csv"):outname+=".csv"
        out=tool_output(outname,str(content),"create CSV");return {"message":f"Created {out}.","path":out}
    if action == "docx_create":
        try: from docx import Document
        except ImportError: raise ValueError("Install python-docx to create Word files: pip install python-docx")
        outname=name or "New document.docx"
        if not outname.lower().endswith(".docx"):outname+=".docx"
        doc=Document()
        for paragraph in str(data.get("text","")).split("\n"):
            if paragraph.startswith("# "):doc.add_heading(paragraph[2:],0)
            elif paragraph.startswith("## "):doc.add_heading(paragraph[3:],1)
            else:doc.add_paragraph(paragraph)
        buf=io.BytesIO();doc.save(buf);out=tool_output(outname,buf.getvalue(),"create Word document")
        return {"message":f"Created {out}.","path":out}
    if action == "csv_merge":
        first=require_file(rel,{".csv"});second=require_file(values.get("second", ""),{".csv"});a=csv_data(first);b=csv_data(second)
        if a[0]!=b[0]:raise ValueError("The CSV headers must match to merge these files.")
        rows=a+b[1:];outname=name or f"{first.stem}_merged.csv";out=tool_output(outname,csv_string(rows),"merge CSV")
        return {"message":f"Merged {len(rows)-1} data rows into {out}.","path":out}
    if action == "compare_text":
        first=require_file(rel);second=require_file(values.get("second", ""));a=file_text(first).splitlines();b=file_text(second).splitlines()
        diff="\n".join(difflib.unified_diff(a,b,fromfile=rel,tofile=values["second"],lineterm="")) or "The files are identical."
        audit("compare",rel,values["second"]);return {"message":diff[:15000]}
    if action == "merge_text":
        first=require_file(rel);second=require_file(values.get("second", ""));outname=name or f"{first.stem}_merged.md"
        if not Path(outname).suffix:outname+=".md"
        out=tool_output(outname,file_text(first)+"\n\n---\n\n"+file_text(second),"merge text")
        return {"message":f"Merged files into {out}.","path":out}
    if action.startswith("csv_") or action in {"xlsx_from_csv"}:
        src=require_file(rel,{".csv"});rows=csv_data(src);header=rows[0];body=rows[1:]
        if action=="csv_stats":
            widths=max((len(r) for r in rows),default=0);empty=sum(1 for r in body for c in r if not c.strip());audit("CSV overview",rel)
            return {"message":f"{len(body)} data rows\n{widths} columns\n{empty} empty cells\nHeaders: {', '.join(header)}"}
        if action=="csv_numeric_report":
            lines=[]
            for col in range(max((len(r) for r in rows),default=0)):
                nums=[]
                for r in body:
                    try:nums.append(float(r[col].replace(",","")))
                    except (ValueError,IndexError):pass
                if nums:lines.append(f"{header[col] if col<len(header) else 'Column '+str(col+1)}: n={len(nums)}, sum={sum(nums):g}, mean={statistics.mean(nums):.2f}, min={min(nums):g}, max={max(nums):g}")
            audit("CSV numeric report",rel)
            return {"message":"\n".join(lines) if lines else "No numeric columns found."}
        if action=="csv_headers":
            clean=[];used=set()
            for cell in header:
                x=re.sub(r"[^a-z0-9]+","_",cell.strip().lower()).strip("_") or "column"
                base=x;n=2
                while x in used:x=f"{base}_{n}";n+=1
                clean.append(x);used.add(x)
            rows=[clean]+body;outname=name or f"{src.stem}_headers.csv"
        elif action=="csv_trim":
            rows=[[c.strip() for c in r] for r in rows];outname=name or f"{src.stem}_trimmed.csv"
        elif action=="csv_dedupe":
            seen=set();outrows=[header]
            for row in body:
                key=tuple(row)
                if key not in seen:seen.add(key);outrows.append(row)
            rows=outrows;outname=name or f"{src.stem}_unique.csv"
        elif action=="csv_sort":
            col=name
            if not col:raise ValueError("Enter a column name to sort by.")
            try:idx=header.index(col)
            except ValueError:raise ValueError(f"Column not found. Available columns: {', '.join(header)}")
            rows=[header]+sorted(body,key=lambda r:(idx>=len(r),r[idx].lower() if idx<len(r) else ""));outname=f"{src.stem}_sorted.csv"
        elif action=="csv_filter":
            col=values.get("query","");val=values.get("replacement","")
            if col not in header:raise ValueError(f"Column not found. Available columns: {', '.join(header)}")
            idx=header.index(col);rows=[header]+[r for r in body if idx<len(r) and r[idx]==val];outname=f"{src.stem}_filtered.csv"
        elif action=="csv_transpose":
            width=max(map(len,rows));rows=[r+([""]*(width-len(r))) for r in rows];rows=[list(r) for r in zip(*rows)];outname=f"{src.stem}_transposed.csv"
        elif action=="csv_split":
            try:limit=max(1,min(100000,int(name or "500")))
            except ValueError:raise ValueError("Rows per part must be a whole number.")
            outputs=[]
            for i,start in enumerate(range(1,len(rows),limit),1):
                outname=f"{src.stem}_part{i:02}.csv";outputs.append(tool_output(outname,csv_string([header]+rows[start:start+limit]),"split CSV"))
            return {"message":f"Created {len(outputs)} part file(s):\n"+"\n".join(outputs)}
        elif action=="csv_markdown":
            widths=[max(len(str(r[i])) if i<len(r) else 0 for r in rows) for i in range(len(header))]
            line=lambda r:"| "+" | ".join((str(r[i]) if i<len(r) else "").ljust(widths[i]) for i in range(len(header)))+" |"
            md="\n".join([line(header),"| "+" | ".join("-"*max(3,w) for w in widths)+" |"]+[line(r) for r in body]);outname=f"{src.stem}.md"
            out=tool_output(outname,md,"CSV to Markdown");return {"message":f"Created {out}.","path":out}
        elif action=="xlsx_from_csv":
            try:from openpyxl import Workbook
            except ImportError:raise ValueError("Install openpyxl to create Excel files: pip install openpyxl")
            outname=name or f"{src.stem}.xlsx"
            if not outname.lower().endswith(".xlsx"):outname+=".xlsx"
            wb=Workbook();ws=wb.active;ws.title="Data"
            for row in rows:ws.append(row)
            ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
            for cells in ws.columns:
                width=min(48,max(10,max((len(str(c.value or "")) for c in cells),default=10)+2));ws.column_dimensions[cells[0].column_letter].width=width
            buf=io.BytesIO();wb.save(buf);out=tool_output(outname,buf.getvalue(),"create Excel workbook")
            return {"message":f"Created {out} from {len(body)} CSV rows.","path":out}
        else:raise ValueError("Unsupported CSV tool.")
        out=tool_output(outname,csv_string(rows),action.replace("_"," "));return {"message":f"Created {out}.","path":out}
    if action == "xlsx_sheets":
        src=require_file(rel,{".xlsx"})
        try:from openpyxl import load_workbook
        except ImportError:raise ValueError("Install openpyxl to inspect Excel files.")
        wb=load_workbook(src,read_only=True,data_only=False)
        lines=[f"{ws.title}: {ws.max_row} rows × {ws.max_column} columns" for ws in wb.worksheets]
        wb.close();audit("inspect Excel",rel);return {"message":"Workbook sheets\n"+"\n".join(lines)}
    if action == "pptx_inventory":
        src=require_file(rel,{".pptx"})
        try:from pptx import Presentation
        except ImportError:raise ValueError("Install python-pptx to inspect PowerPoint files.")
        prs=Presentation(src);lines=[f"{len(prs.slides)} slides"]
        for i,slide in enumerate(prs.slides,1):
            text=" · ".join(sh.text.replace("\n"," ") for sh in slide.shapes if getattr(sh,"has_text_frame",False))
            lines.append(f"Slide {i}: {text[:240] or '(no text)'}")
        audit("inspect PowerPoint",rel);return {"message":"\n".join(lines)}
    if action == "docx_metadata":
        src=require_file(rel,{".docx"})
        try:from docx import Document
        except ImportError:raise ValueError("Install python-docx to inspect Word metadata.")
        doc=Document(src);p=doc.core_properties
        audit("inspect Word metadata",rel)
        return {"message":f"Title: {p.title or '—'}\nAuthor: {p.author or '—'}\nSubject: {p.subject or '—'}\nCreated: {p.created or '—'}\nParagraphs: {len(doc.paragraphs)}\nTables: {len(doc.tables)}"}
    if action == "pptx_from_outline":
        src=require_file(rel,{".txt",".md"});
        try:from pptx import Presentation
        except ImportError:raise ValueError("Install python-pptx to create presentations.")
        prs=Presentation();title_layout=prs.slide_layouts[0];body_layout=prs.slide_layouts[1];lines=file_text(src).splitlines();title=src.stem
        for line in lines:
            if line.startswith("# "):title=line[2:].strip();break
        slide=prs.slides.add_slide(title_layout);slide.shapes.title.text=title
        sub=slide.placeholders[1] if len(slide.placeholders)>1 else None
        if sub:sub.text="Created locally by OfficeMate"
        current=None
        for line in lines:
            if line.startswith("## "):
                current=prs.slides.add_slide(body_layout);current.shapes.title.text=line[3:].strip()
            elif line.startswith(("- ","* ")) and current:
                tf=current.placeholders[1].text_frame
                para=tf.paragraphs[0] if not tf.text else tf.add_paragraph();para.text=line[2:].strip();para.level=0
            elif line.strip() and current:
                tf=current.placeholders[1].text_frame;para=tf.add_paragraph() if tf.text else tf.paragraphs[0];para.text=line.strip()
        outname=name or f"{src.stem}.pptx"
        if not outname.lower().endswith(".pptx"):outname+=".pptx"
        buf=io.BytesIO();prs.save(buf);out=tool_output(outname,buf.getvalue(),"create presentation");return {"message":f"Created {out} with {len(prs.slides)} slides.","path":out}
    if action in {"word_count","char_count","reading_time"}:
        src=require_file(rel);text=file_text(src);words=re.findall(r"\b[\w'-]+\b",text)
        audit(action,rel)
        if action=="word_count":msg=f"{len(words):,} words"
        elif action=="char_count":
            compact=len(re.sub(r"\\s","",text));msg=f"{len(text):,} characters including spaces\n{compact:,} excluding whitespace"
        else:msg=f"About {max(1,round(len(words)/220))} minute(s) to read ({len(words):,} words at 220 wpm)."
        return {"message":msg}
    # Remaining text-to-text operations create a new sibling file.
    src=require_file(rel);text=file_text(src);stem=src.stem
    if action=="uppercase":new=text.upper();outname=f"{stem}_uppercase.txt"
    elif action=="lowercase":new=text.lower();outname=f"{stem}_lowercase.txt"
    elif action=="title_case":new=text.title();outname=f"{stem}_titlecase.txt"
    elif action=="trim_whitespace":new="\n".join(x.strip() for x in text.splitlines())+"\n";outname=f"{stem}_trimmed.txt"
    elif action=="remove_blank_lines":new="\n".join(x for x in text.splitlines() if x.strip())+"\n";outname=f"{stem}_noblanks.txt"
    elif action=="sort_lines":new="\n".join(sorted(text.splitlines(),key=str.lower))+"\n";outname=f"{stem}_sorted.txt"
    elif action=="unique_lines":new="\n".join(dict.fromkeys(text.splitlines()))+"\n";outname=f"{stem}_unique.txt"
    elif action=="replace_text":
        q=values.get("query","");
        if not q:raise ValueError("Enter the text to find.")
        new=text.replace(q,values.get("replacement",""));outname=f"{stem}_replaced.txt"
    elif action=="reverse_lines":new="\n".join(reversed(text.splitlines()))+"\n";outname=f"{stem}_reversed.txt"
    elif action=="text_to_html":new="<!doctype html><meta charset='utf-8'><title>"+html.escape(src.stem)+"</title><pre>"+html.escape(text)+"</pre>";outname=f"{stem}.html"
    elif action=="markdown_to_html":
        escaped=html.escape(text);lines=[]
        for line in escaped.splitlines():
            m=re.match(r"^(#{1,6})\s+(.*)$",line)
            if m:level=len(m.group(1));lines.append(f"<h{level}>{m.group(2)}</h{level}>")
            elif line.startswith("- "):lines.append("<p>• "+line[2:]+"</p>")
            elif line.strip():lines.append("<p>"+line+"</p>")
        new="<!doctype html><meta charset='utf-8'><title>"+html.escape(src.stem)+"</title><article>"+"\n".join(lines)+"</article>";outname=f"{stem}.html"
    elif action=="checklist":new="# Tasks\n\n"+"\n".join("- [ ] "+x.strip() for x in text.splitlines() if x.strip())+"\n";outname=f"{stem}_checklist.md"
    elif action=="extract_text":new=text;outname=f"{stem}_extracted.txt"
    else:raise ValueError("This tool is not implemented yet.")
    out=tool_output(outname,new,action.replace("_"," "));return {"message":f"Created {out}.","path":out}


def pretty_bytes(n):
    if n<1024:return f"{n} B"
    if n<1024*1024:return f"{n/1024:.1f} KB"
    return f"{n/1048576:.1f} MB"


class Handler(BaseHTTPRequestHandler):
    server_version = "OfficeMate/1.0"

    def log_message(self, *_):
        pass

    def send_json(self, data, status=200):
        raw = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(raw)

    def body(self):
        length = min(int(self.headers.get("Content-Length", 0)), 5_000_000)
        return json.loads(self.rfile.read(length) or b"{}")

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        route = parsed.path
        if route == "/" or route == "/index.html":
            return self.serve_static("index.html")
        if route.startswith("/static/"):
            return self.serve_static(route.removeprefix("/static/"))
        if route == "/api/tools":
            self.send_json({"count": len(FEATURES), "tools": [{"id": i, "category": c, "title": t, "description": d, "fields": f} for i,c,t,d,f in FEATURES]})
        elif route == "/api/status":
            files = list_files()
            deps = {}
            for name, module in [("Word", "docx"), ("Excel", "openpyxl"), ("PowerPoint", "pptx"), ("LibreOffice", "soffice")]:
                if module == "soffice": deps[name] = shutil.which("soffice") is not None
                else:
                    try: __import__(module); deps[name] = True
                    except ImportError: deps[name] = False
            self.send_json({"workspace": str(WORKSPACE), "files": len(files), "storage": sum(f["size"] for f in files), "files_list": files[:8], "tools": deps, "templates": 3, "tool_count": len(FEATURES)})
        elif route == "/api/files":
            self.send_json({"files": list_files()})
        elif route.startswith("/api/file/"):
            rel = urllib.parse.unquote(route[len("/api/file/"):])
            try:
                path = safe_path(rel)
                if not path.is_file(): raise FileNotFoundError
                audit("read", rel)
                self.send_json({"path": rel, "name": path.name, "text": file_text(path), "editable": path.suffix.lower() in {".txt", ".md", ".csv"}})
            except FileNotFoundError: self.send_json({"error": "File not found."}, 404)
            except ValueError as e: self.send_json({"error": str(e)}, 400)
        elif route == "/api/audit":
            rows = []
            if AUDIT.exists():
                for line in AUDIT.read_text(encoding="utf-8").splitlines()[-100:]:
                    try: rows.append(json.loads(line))
                    except json.JSONDecodeError: pass
            self.send_json({"events": rows[::-1]})
        elif route == "/api/search":
            q = urllib.parse.parse_qs(parsed.query).get("q", [""])[0].strip().lower()
            hits = []
            if q:
                audit("search", "workspace", q)
                for f in list_files():
                    path = safe_path(f["path"])
                    text = file_text(path)
                    if q in f["name"].lower() or q in text.lower():
                        idx = text.lower().find(q)
                        hits.append({"path": f["path"], "name": f["name"], "snippet": (text[max(0,idx-70):idx+180].replace("\n", " ") if idx >= 0 else "Matching filename")})
            self.send_json({"results": hits[:50]})
        else:
            self.send_json({"error": "Not found"}, 404)

    def serve_static(self, name):
        target = (STATIC / name).resolve()
        if STATIC.resolve() not in target.parents and target != STATIC.resolve(): return self.send_json({"error": "Not found"}, 404)
        try: raw = target.read_bytes()
        except OSError: return self.send_json({"error": "Not found"}, 404)
        self.send_response(200)
        self.send_header("Content-Type", mimetypes.guess_type(target.name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers(); self.wfile.write(raw)

    def do_POST(self):
        try:
            data = self.body()
            route = urllib.parse.urlparse(self.path).path
            if route == "/api/tool":
                before = AUDIT.stat().st_size if AUDIT.exists() else 0
                result = run_tool(data)
                after = AUDIT.stat().st_size if AUDIT.exists() else 0
                if after == before: audit("tool action", str(data.get("path") or "workspace"), str(data.get("action") or "unknown"))
                return self.send_json(result)
            if route == "/api/files":
                name = str(data.get("name", "")).strip()
                if not name or Path(name).name != name or Path(name).suffix.lower() not in {".txt", ".md", ".csv"}:
                    return self.send_json({"error": "Choose a simple .txt, .md, or .csv filename."}, 400)
                path = safe_path(name)
                if path.exists() and not data.get("confirm_overwrite"):
                    return self.send_json({"error": "That file already exists. Confirm overwrite to replace it."}, 409)
                if path.exists(): backup(path)
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(str(data.get("text", "")), encoding="utf-8")
                audit("create" if not data.get("confirm_overwrite") else "overwrite", name)
                return self.send_json({"ok": True, "path": name})
            if route == "/api/save":
                rel = str(data.get("path", "")); path = safe_path(rel)
                if path.suffix.lower() not in {".txt", ".md", ".csv"}: return self.send_json({"error": "Only text and CSV files can be edited in this workspace."}, 400)
                if not path.exists(): return self.send_json({"error": "File not found."}, 404)
                if not data.get("confirm"): return self.send_json({"error": "Confirm this edit before saving."}, 409)
                backup(path); path.write_text(str(data.get("text", "")), encoding="utf-8"); audit("edit", rel)
                return self.send_json({"ok": True})
            if route == "/api/command":
                command = str(data.get("command", "")).strip()
                if not command: return self.send_json({"error": "Enter a command."}, 400)
                match = re.search(r"(?:summari[sz]e|recap).*?(?:this\s+|file\s+)?([A-Za-z0-9_./ -]+?\.(?:txt|md|docx|pptx|csv))(?=$|\s)", command, re.I)
                if match:
                    rel = re.sub(r"^(?:this|file)\s+", "", match.group(1).strip(), flags=re.I); path = safe_path(rel)
                    if not path.exists(): return self.send_json({"error": f"I couldn't find {rel}."}, 404)
                    result = summarize(file_text(path)); audit("summarize", rel, "Local extractive summary")
                    return self.send_json({"message": result, "action": "summary", "file": rel})
                match = re.search(r"(?:find|search for)\s+(.+?)(?:\s+in\s+(?:my\s+)?files?)?$", command, re.I)
                if match:
                    q = match.group(1).strip().strip('"\'')
                    audit("search", "workspace", q)
                    hits = [f for f in list_files() if q.lower() in f["name"].lower() or q.lower() in file_text(safe_path(f["path"])).lower()]
                    return self.send_json({"message": f"Found {len(hits)} matching file{'s' if len(hits)!=1 else ''}.", "action": "search", "results": hits[:10]})
                return self.send_json({"message": "I can search your files or summarize a text, Word, PowerPoint, or CSV file. Try “summarize report.md” or “find budget”.", "action": "help"})
            if route == "/api/organize":
                # Safe, reversible-by-backup folder organization; never overwrite destinations.
                moves = []
                buckets = {".docx":"Word", ".xlsx":"Excel", ".pptx":"PowerPoint", ".pdf":"PDF", ".csv":"Spreadsheets", ".txt":"Notes", ".md":"Notes"}
                for item in list_files():
                    src = safe_path(item["path"]); folder = buckets.get(src.suffix.lower())
                    if not folder or src.parent != WORKSPACE: continue
                    dest_dir = WORKSPACE / folder; dest_dir.mkdir(exist_ok=True); dest = dest_dir / src.name
                    if dest.exists(): continue
                    backup(src); shutil.move(str(src), str(dest)); moves.append({"from": item["path"], "to": dest.relative_to(WORKSPACE).as_posix()}); audit("organize", item["path"], f"Moved to {folder}/")
                return self.send_json({"ok": True, "moves": moves, "message": f"Sorted {len(moves)} file{'s' if len(moves)!=1 else ''} into type folders."})
            if route == "/api/convert":
                rel = str(data.get("path", "")); target_ext = str(data.get("to", "")).lower().lstrip(".")
                src = safe_path(rel)
                if not src.is_file(): return self.send_json({"error": "Source file not found."}, 404)
                dest = src.with_suffix("." + target_ext)
                if dest.exists() and not data.get("confirm_overwrite"): return self.send_json({"error": "Destination exists. Confirm overwrite first."}, 409)
                if dest.exists(): backup(dest)
                if src.suffix.lower() in {".txt", ".md"} and target_ext in {"txt", "md"}:
                    dest.write_text(src.read_text(encoding="utf-8", errors="replace"), encoding="utf-8")
                elif shutil.which("soffice") and target_ext in {"pdf", "docx", "xlsx", "pptx", "csv"}:
                    outdir = WORKSPACE / ".convert-tmp"; outdir.mkdir(exist_ok=True)
                    proc = subprocess.run(["soffice", "--headless", "--convert-to", target_ext, "--outdir", str(outdir), str(src)], capture_output=True, text=True, timeout=90)
                    produced = outdir / (src.stem + "." + target_ext)
                    if proc.returncode or not produced.exists(): return self.send_json({"error": "Conversion failed. " + (proc.stderr[-300:] or proc.stdout[-300:])}, 422)
                    shutil.move(str(produced), str(dest)); shutil.rmtree(outdir, ignore_errors=True)
                else:
                    return self.send_json({"error": "This conversion requires LibreOffice. Install it and retry."}, 422)
                audit("convert", rel, f"to {target_ext}")
                return self.send_json({"ok": True, "path": dest.relative_to(WORKSPACE).as_posix()})
            if route == "/api/template":
                template = data.get("template")
                templates = {
                    "meeting": ("Meeting Notes.md", "# Meeting notes\n\n**Date:** \n**Attendees:** \n**Facilitator:** \n\n## Agenda\n- \n\n## Discussion\n\n## Decisions\n- \n\n## Action items\n| Task | Owner | Due date |\n|---|---|---|\n| | | |\n"),
                    "project": ("Project Plan.md", "# Project plan\n\n**Project:** \n**Owner:** \n**Last updated:** \n\n## Objective\n\n## Milestones\n| Milestone | Owner | Target date | Status |\n|---|---|---|---|\n| | | | Not started |\n\n## Risks and dependencies\n\n## Next steps\n- \n"),
                    "expenses": ("Expense Tracker.csv", "Date,Description,Category,Amount,Payment method,Notes\n")}
                if template not in templates: return self.send_json({"error": "Unknown template."}, 400)
                name, content = templates[template]; path = safe_path(name)
                if path.exists() and not data.get("confirm_overwrite"): return self.send_json({"error": "That template file already exists. Confirm overwrite first."}, 409)
                if path.exists(): backup(path)
                path.write_text(content, encoding="utf-8"); audit("create from template", name, template)
                return self.send_json({"ok": True, "path": name})
            self.send_json({"error": "Not found"}, 404)
        except FileExistsError as e: self.send_json({"error": str(e)}, 409)
        except PermissionError as e: self.send_json({"error": str(e)}, 409)
        except FileNotFoundError as e: self.send_json({"error": str(e) or "File not found."}, 404)
        except (ValueError, json.JSONDecodeError) as e: self.send_json({"error": str(e)}, 400)
        except subprocess.TimeoutExpired: self.send_json({"error": "Conversion timed out."}, 504)
        except Exception as e: self.send_json({"error": f"Operation failed: {e}"}, 500)


def main():
    WORKSPACE.mkdir(parents=True, exist_ok=True)
    host = os.environ.get("OFFICEMATE_HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", os.environ.get("OFFICEMATE_PORT", "8765")))
    print(f"OfficeMate workspace: {WORKSPACE}\nOpen http://localhost:{port}")
    ThreadingHTTPServer((host, port), Handler).serve_forever()

if __name__ == "__main__":
    main()
