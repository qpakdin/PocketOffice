# Pocket Office

Two parts that work together:

1. **Android app** (`app/`, Java, built by GitHub Actions - no Android Studio needed): viewer for every Office format with the real page layout, and a plain-text editor. Text files work without any server.
2. **Converter server** (`converter/`, Python): LibreOffice does the real page layout, and the **web panel** (the Word / Excel / PowerPoint editor) is served by it at `http://127.0.0.1:8585/`. It runs in Termux (Debian proot) on the phone or on a PC.

## What the app does
- **Office files** (Word / Excel / PowerPoint / OpenDocument, all common formats): converted to PDF by the server and shown page by page (zoom up to 2000 %, swipe, go-to-page). Export PDF, PNG pages, TXT to `Download/PocketOffice/`.
  Word: docx docm dotx dotm doc dot odt ott fodt rtf sxw stw - Excel: xlsx xlsm xltx xltm xlsb xls xlt ods ots fods sxc - PowerPoint: pptx pptm ppsx ppsm potx potm ppt pps pot odp otp fodp sxi.
- **Text files** (txt, py, sh, json, md, csv, configs ...): opened and edited inside the app (undo/redo, find/replace, encodings, line endings). Saving always writes a new copy.
- **Edit in browser**: Export / Edit > "Edit a copy in the browser..." uploads a *copy* to the server and opens the web panel on it (docx/docm/dotx/dotm, xlsx/xlsm/xltx/xltm, pptx/pptm/ppsx/ppsm/potx/potm only).
- **New file**: Home > "New Word / Excel / PowerPoint file" opens the panel's create screen (14 OOXML types; odt/ods/odp/rtf/doc/xls/ppt families are created through their OOXML twin and converted by LibreOffice on save; `.xlsb` and StarOffice 1.x `sxw stw sxc sxi` cannot be created).
- **Multi-window**: every opened file gets its own window (`documentLaunchMode=always`); "Open in a new window" on Home and in the Export menu, split-screen/freeform supported. In the browser every tab is an independent editing session (copy/paste between tabs works).

## Safety rules (kept by design)
- The original file is only read. Edits happen in a server-side copy; saving = downloading a new file.
- Editing is opt-in; it never starts when a file is opened.
- Before editing a preflight report (signatures, macros, protection, tracked changes, fields, shapes, protected sheets, charts ...) is shown and logged by the server (stdout and `~/.pocket_office/edit_log.txt`). Protected documents are refused.
- Before saving the result is verified (package diff, well-formed XML, text invariants, locked paragraphs untouched, LibreOffice can open it).
- Word edits are surgical: only edited paragraphs change; paragraphs with fields, pictures, shapes or tracked changes are locked.

## Server
Needs `python3`, LibreOffice, `poppler-utils` (pdftotext, pdfinfo, pdftoppm) and `lxml` (`apt install python3-lxml`).

```
python3 pocket_office_server.py            # 127.0.0.1:8585
python3 pocket_office_server.py --lan      # reachable from the phone (PC)
python3 pocket_office_server.py --fonts DIR  # font library
```
The banner must say `v4.0` and `Editors : Word + Excel + PowerPoint ready`. The app's default server address is `http://127.0.0.1:8585` (an old saved `:8765` is migrated automatically). See `converter/README_FA.md` (Persian) for the full setup.

## Files (converter/)
| file | role |
|---|---|
| `pocket_office_server.py` | HTTP server, conversion to PDF, hosts the panel, edit endpoints, host/origin guard, font library |
| `po_docx_edit.py` | surgical Word engine |
| `po_office_edit.py` | Word / Excel / PowerPoint sessions (undo, settings preview/apply, find, verify) |
| `po_new_file.py` | creates valid empty files |
| `wordtk/` | Word settings engine (13 feature groups; needs `lxml`) |
| `panel/` | the web panel (vanilla JS, offline) |

## Tests
```
mkdir -p /tmp/ed && cp tests/fixtures/* tests/*.py tests/*.js /tmp/ed/ && cd /tmp/ed
python3 t1.py; ... python3 t8.py; node ui_test.js
```
Not tested anywhere yet: real Microsoft Office opening the edited/new files, the panel on a real phone browser, and an actual APK build/run of the latest Java (only a flow compile check without android.jar was possible).

## Build
Push to GitHub; `.github/workflows/build-apk.yml` builds the APK.
