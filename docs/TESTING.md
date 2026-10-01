# 測試與發布驗證 / Testing and release validation

[繁體中文 README](../README.md) · [English README](../README.en.md)

Run on Windows x64 with Python 3.10 and the development requirements:

```powershell
python -m pip install -r requirements-dev.txt
python -m unittest discover -s tests -p 'test_*.py' -v
python scripts/audit_public.py
.\scripts\build.ps1
.\scripts\package.ps1
```

## Synthetic test coverage

Tests create their own synthetic Office, PDF, email and text documents in temporary directories. No actual organization templates, customer documents, emails or locally generated business reports are included.

- Selected and unselected content, invalid selections, changed source hashes and output collision handling.
- Word/PPT split formatting runs, style references, images, field instructions, sensitivity-label preservation, protection and signatures.
- PowerPoint unused structure pruning, including active, unselected theme preservation.
- Excel shared strings, hidden cells, unselected cells and formula/protection preservation.
- Email plain/HTML content, nested formatting, attachment payloads, address fields and signature rejection.
- PDF whole text-show block removal, retained text positioning and page geometry, metadata, split strings, signatures and encryption.
- Text BOM, encoding, byte order, line endings and CSV row structure.
- Generic custom keywords, centralized author metadata and hidden GUI startup.

GitHub Actions runs the synthetic suite, the public-file audit and a Windows build. It does not require or automate desktop Office.

## Portable executable checks

Version 3.1.0 passed 31 synthetic automated tests and 11 additional cleanup cases through the frozen Windows executable. The hidden GUI startup check confirmed the release version, author metadata and public source link. These checks use generated samples and do not establish full coverage of every document layout.

Before publication, the packaged executable is checked separately for hidden GUI startup and CLI cleanup of synthetic samples. The archive and embedded Python code are audited for private workspace markers; the release ZIP is created from a source allowlist, not a recursive copy of the development workspace.

For legacy Office conversion, desktop Office must be installed. Treat conversion as a separate integration check and inspect the resulting modern-format copy. Automated structure checks do not guarantee every document's rendered layout or detect all sensitive content.

## Manual checks

Open output copies and review formatting, page breaks and the removal range, especially for PDF text-show blocks. For email, confirm retained attachments are appropriate for sharing. For Excel, recalculate formulas in the application when needed.

Report issues with synthetic reproductions only. Generated reports contain source paths and content excerpts and must not be committed or published with real files.
