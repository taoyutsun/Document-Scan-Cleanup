# Document Scan & Cleanup

**Local document content and template cleanup for Windows** · MIT License

[繁體中文](README.md) | **English**

Before sharing documents with customers, suppliers or partners, scan for text you specify, review the matches, remove selected content and save a separate copy. Supports PowerPoint, Word, Excel, PDF, email and common text formats while preserving existing document structure where supported.

## Why this tool?

Individuals and organizations often share a single document template rather than maintaining separate internal and external versions. A finished document may need adjustments before delivery: internal notes, draft markings, names, contact information, project codes or other content that should not appear in the shared copy.

Manual cleanup takes time, particularly when text appears in headers, footers, slide masters, unused layouts or internal names and properties. Copying PowerPoint content into a different template can also disturb fonts, images and text boxes, requiring slide-by-slide repairs.

This utility combines **scan → review and select → clean a copy → verify the copy**, reducing repeated searching and content transfer.

**You define what to remove.** It matches literal text supplied by the user; it does not automatically identify every piece of personal or confidential information or determine whether a document is public.

## Download and start

1. Download the Windows x64 portable ZIP from [Releases](https://github.com/taoyutsun/Document-Scan-Cleanup/releases/latest).
2. Extract the entire archive and run `DocumentScanCleanup.exe`. Modern formats do not require Python or Office.
3. Add documents with the file selection button.
4. Enter semicolon-separated text, for example `Draft; Internal Use Only; Project code; client@example.invalid`. Replace these fictional examples with your own terms.
5. Scan, double-click a result to inspect its content, then adjust the checkboxes.
6. Clean selected items and save. **Original files are kept and never overwritten.**
7. Open the output and check content, line wrapping and layout before sharing.

The GUI currently uses Traditional Chinese. The executable is unsigned; Windows may show a SmartScreen prompt. Execution is subject to your device's software settings.

## Personal keyword profiles (since 3.2.0)

Keep general built-in examples or save your own working lists. **Personal profiles are stored in your user settings directory, outside the EXE, GitHub source and release ZIP.** Replacing the executable does not remove them.

1. Enter semicolon-separated terms, click **儲存清單** (Save list), and name the profile. Saving updates the currently loaded local profile; starting from generic defaults creates a new profile.
2. Click **設為啟動預設** (Set startup default) to load it on future GUI launches. Unsaved edits must be named and saved first.
3. Switch saved lists in the dropdown. **載入清單** (Load list) imports a UTF-8 JSON file into local settings without modifying the original.
4. **恢復通用預設** (Restore generic defaults) resets both the current terms and future GUI startup. Saved personal profiles remain available.

On Windows, lists live in `%APPDATA%\DocumentScanCleanup\profiles\`; `keyword-settings.json` in the parent folder records the startup selection. Separate identifiers prevent name collisions. Edits are not saved automatically. Invalid startup settings/profile files trigger a visible fallback to generic defaults without overwriting the damaged files.

**開啟清單資料夾** opens the actual profile folder; the interface shows the saved profile count. **重新載入啟動清單** rereads the saved startup selection after external changes, asking before discarding unsaved edits. Profiles are in the `profiles` subfolder; `keyword-settings.json` stores only the selection.

Results and logs resize with the window while space remains reserved for author links at the bottom, without requiring maximization.

Import example, using fictional terms only:

```json
{
  "schema_version": 1,
  "name": "Example working list",
  "keywords": ["Draft", "Internal Use Only", "Project Cedar", "Example term"]
}
```

Lists accept 1–100 nonempty terms, each up to 150 characters; individual terms cannot contain semicolons or line breaks. Keep personal lists and generated reports private; do not include them in public documentation, screenshots or issue reports.

**CLI selection is explicit:** without keyword options, it uses generic defaults and ignores the GUI startup profile. Use `--keywords-profile .\my-list.json` or `--keywords`, exclusively. With `--gui`, an explicit list applies to that launch only and does not change the saved startup default.

## Supported formats

| Format | Scan scope | Cleanup scope / requirements |
| --- | --- | --- |
| PPTX | Slides, masters, layouts, notes, supported names and properties | Selected text/names; optional unused master/layout/theme pruning, preserving active formatting dependencies |
| DOCX | Body, headers, footers, tables, text boxes, comments, revisions, styles and properties | Selected text and supported names, with synchronized style references |
| XLSX | Visible/hidden text cells, headers/footers, traditional comments, some drawings and properties | Selected text; formulas, numbers and unselected cells are retained, including other shared-string references |
| PDF | Parsable text-show blocks and ordinary metadata, with page numbers | **Removes an entire selected text-show block**, preserving other text positions, images and page geometry |
| EML | Subject, plain/HTML body and supported attachments | Selected subject/body content; attachments and inline images remain unchanged and may retain matches |
| PPT / PPS | Converted to PPTX first | Requires desktop PowerPoint; output is a modern-format copy |
| DOC / RTF | Converted to DOCX first | Requires desktop Word; output is a modern-format copy |
| XLS | Converted to XLSX first | Requires desktop Excel; output is a modern-format copy |
| TXT / MD / LOG | Common encodings and text lines | Selected keywords or lines, preserving encoding, BOM and line endings |
| CSV / TSV | Text content | Keyword removal only, retaining quotes, delimiters and rows |
| JSON / XML | Keyword matches in text and attributes | **Scan only**, to avoid breaking structure or field names |

MSG, macro-enabled formats, OpenDocument, image OCR and automatic cleanup of embedded Office/OLE documents are not supported.

## Matching and selection

- Case-insensitive **literal substring matching**, not regular expressions. Review context to avoid removing unintended portions of longer strings.
- General default examples include Confidential, Internal Use Only, Draft and Chinese markings. Replace them with full names, contact strings, project codes or other specific text as appropriate.
- Keyword mode removes matched text only in selected supported items. PowerPoint/Word text split across formatting runs and email HTML text split across tags are supported.
- Block mode clears a selected supported paragraph or line while retaining its frame and formatting; it does not remove an entire text box or table.
- **PDF always removes a complete selected text-show block.** A block may include other words in the sentence, or cover only part of a longer paragraph. Inspect the full text; no sentence repair or reflow is performed.
- PDF, Excel and email cleanup items start **unchecked** in the GUI. CLI `--clean` selects all cleanable matches by default.
- Unused PowerPoint template pruning is separate from text selection; it does not independently rename active, unselected themes.

Names/properties, Excel headers/footers and email subjects use keyword removal only. CSV/TSV do not support clearing entire lines. Text removal can change Word wrapping or pagination.

## Outputs and originals

The GUI creates a new `Document_Cleanup_timestamp` folder, with numbered per-file subfolders containing the cleaned copy, TXT/JSON change reports and remaining matches. A batch report is also written. Scan-only reports can be saved without cleaning.

If the source changes after scanning, scan again. Existing outputs are never overwritten. A match can appear in multiple storage locations or references, so match counts do not represent unique pieces of sensitive information.

**Reports contain file paths and text excerpts.** Protect them like the source document. Submit synthetic reproductions when reporting issues, not real client documents or unedited reports.

## Local processing and limits

- No automatic document upload, API keys, telemetry or update checks. Author/source links open the system browser only when clicked.
- Formal sensitivity labels, unknown `customXml`, Office field instructions and unsupported identifiers are retained. The utility does not remove access controls or reclassify documents.
- Signed, encrypted, protected or unsupported structures are not cleaned. EML with DKIM, ARC, S/MIME, PGP or encapsulated message attachments is scan-only.
- No OCR. Images, special nested PDF content/forms/attachments and embedded Office documents are not fully inspected. The UI distinguishes completed, partial and failed scans. **Zero matches do not establish that all sensitive information has been removed.**
- Legacy conversion disables macros/events and link updates where supported, and uses only a separately identified Office process. Office settings remain managed by the device environment. Conversion may affect fonts or compatibility objects. Close PowerPoint or save a PPTX manually if it is already running.
- XLSX formulas and cached values are retained, not recalculated. Open the copy in Excel and recalculate when needed.
- General file / uncompressed Office limit: 600 MB and 25,000 ZIP entries. Legacy conversion: 100 MB / 90 seconds. Email attachment recursion: depth 2, 100 MB per attachment, 200 MB total decoded payload.

## Source, CLI and building

Recommended environment: **Windows x64 / Python 3.10**, used for this release's validation.

```powershell
git clone https://github.com/taoyutsun/Document-Scan-Cleanup.git
cd Document-Scan-Cleanup
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -X utf8 document_scan_cleanup.py --gui
```

Replace these example files and terms with your own:

```powershell
python -X utf8 document_scan_cleanup.py --scan .\example.docx .\example.pdf --keywords 'Draft;Project code' --output-dir .\scan-results
python -X utf8 document_scan_cleanup.py --clean .\example.pptx --keywords 'Internal Use Only' --prune --output-dir .\clean-results
```

CLI also supports `--mode keyword|block`, `--prune` and `--selection-json` (absolute source paths mapped to finding ID arrays). The windowed EXE has no terminal output but accepts CLI arguments and writes files/reports. Failures are written to `errors.report.json`.

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -p 'test_*.py' -v
.\scripts\build.ps1 -Python .\.venv\Scripts\python.exe
.\scripts\package.ps1 -Python .\.venv\Scripts\python.exe
```

Build outputs are in `dist/` and are excluded from source control. Tests generate synthetic documents and contain no private or real organization files. [Testing](docs/TESTING.md) · [Changelog](CHANGELOG.md) · [Third-party notices](THIRD_PARTY_NOTICES.md)

## Author and license

Designed and maintained by **Arthur Tao**, under the [MIT License](LICENSE). You are welcome to use and share it while retaining the author and source attribution.

- [Arthur ASK blog](https://taoyutsun.blogspot.com/)
- [Facebook](https://facebook.com/arthurtaoyutsun)
- [GitHub source](https://github.com/taoyutsun/Document-Scan-Cleanup)

Third-party components keep their own licenses. See `third-party-licenses/` and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
