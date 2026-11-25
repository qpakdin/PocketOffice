# Pocket Office — HANDOVER (continue in a new chat)

Owner language: Persian (answer in Persian). Owner priorities: **never damage the original file**, report every limitation BEFORE editing
and log it in Termux, light tools (not a "studio"), supporting every format matters most.

## 1. What the project is now
Two parts:
1. **Android APK** (`app/`, Java, built by GitHub Actions, no Android Studio): viewer + plain-text editor. Works alone for text files.
2. **Converter server** (`converter/`, Python, runs in Termux/Debian proot or on a Windows PC): LibreOffice does the real page layout, and the
   editors live here. NEW in this session: a **local web panel** (`converter/panel/`, served at `http://127.0.0.1:8585/`) = the editor UI.
   Owner's idea (approved by me): web panel = editor (Word/Excel/PowerPoint, new files, settings, find/replace); APK = viewer + text files.
   Each browser tab is an independent editing session ("multi-window"), copy/paste between tabs works natively.

## 2. Files (converter/)
| file | role |
|---|---|
| `pocket_office_server.py` v4.0 | HTTP server. Default port **8585**. Converts any supported format to PDF, hosts the panel, edit endpoints, host/origin guard (DNS-rebinding), font library index (only needed fonts loaded per file). |
| `po_docx_edit.py` | surgical Word engine (byte-range XML edits; paragraphs with fields/pictures/shapes/tracked changes are LOCKED). |
| `po_office_edit.py` | sessions: `WordSession` (surgical ops + wordtk settings), `XlsxSession`, `PptxSession`; undo = snapshot stack; settings preview (dry run) / apply; find; verify. |
| `po_new_file.py` | creates valid empty OOXML (docx/dotx/docm/dotm, xlsx/xltx/xlsm/xltm, pptx/ppsx/potx/pptm/ppsm/potm). Other formats = edit the OOXML twin, converted by LibreOffice on export (`EDIT_AS`). |
| `wordtk/` | owner's own Word-settings engine (copied from his `wordtk_project.zip`, Tk GUI removed; needs `lxml`). 13 feature groups, schema-driven forms. |
| `panel/index.html, panel.js, panel.css` | the web panel (vanilla JS, offline). |
| `README_FA.md` | setup guide (v4, port 8585). |

Server needs: `python3`, LibreOffice, `poppler-utils` (pdftotext, pdfinfo, **pdftoppm** for preview images), `lxml` (`apt install python3-lxml`).
Run: `python3 pocket_office_server.py` (add `--lan` on a PC; `--fonts DIR` for a font library). Banner must say `v4.0` and `Editors : Word + Excel + PowerPoint ready`.

## 3. Verified in the sandbox (real LibreOffice 24.2 + headless Chromium)
* Word surgical engine: ~700 random edits, all outputs open in LibreOffice, locked paragraphs byte-identical (tests t1–t3).
* wordtk settings through `WordSession`, dry-run preview, undo, verify (t4).
* Excel engine: text/format/layout/page settings, `set_cell`, calcChain removal + `fullCalcOnLoad`, undo, verify (t5).
* PowerPoint engine: set_text/format/para/insert/delete, settings, hide slides, find, verify (t6).
* New files: all 14 OOXML variants valid + LibreOffice opens; conversion to odt/ott/rtf/doc/dot/fodt, ods/ots/xls/xlt/fods, odp/otp/ppt/pps/pot/fodp OK (t7).
* HTTP API incl. host guard, new files, export formats, page PNGs (t8: 44 checks passed).
* Panel in Chromium (phone viewport): home, new docx, typing, find/replace preview+apply, settings forms (13 groups), preview image, save/verify, new xlsx grid + cell typing (ui_test.js).

## 4. STATUS / OPEN POINTS (updated in session 6)
DONE this session:
* `ui_test.js`: change-counter expectation fixed (one Apply = one history step), grid re-query after re-render, formula `=B1*2` via Set, home → sessions list: **15/15 PASS, no console errors**. Panel now has `<link rel="icon" href="data:,">` and the server answers `/favicon.ico` with 204. Formula cells with no cached value show the formula text in the grid.
* t4/t5 benign slips fixed in the tests. All of t1–t8 pass.
* APK: `DEFAULT_SERVER` = `http://127.0.0.1:8585` (a saved `:8765` is migrated automatically). Reader menu (Export ▾, exact mode): "Edit a copy in the browser…" (uploads a COPY via `EditSession.open`, then `ACTION_VIEW <server>/panel/#/s/<sid>`; OOXML families only, otherwise an explanation dialog) and "Open this file in a new window". Home: "New Word / Excel / PowerPoint file" (opens `<server>/panel/#/`) and "Open a new window".
* Multi-window: manifest `documentLaunchMode="always"`, `resizeableActivity="true"`, extended `configChanges` (resize must not recreate the activity); new window = `FLAG_ACTIVITY_NEW_DOCUMENT|MULTIPLE_TASK|LAUNCH_ADJACENT` (+ `FLAG_GRANT_READ_URI_PERMISSION` when a file is passed).
* Decision: the old Java Word editor (`EditorScreen.java`, untested on a device) was DELETED; `EditSession.java` shrank to the upload client. The panel is the only editor.
* New test `tests/edit_in_browser_test.js` simulates exactly what the APK sends (raw POST to `/edit/open`) and opens `/panel/#/s/<sid>` in Chromium for docx/xlsx/pptx; also checks refusal of `.doc`.
* `README.md` and `converter/README_FA.md` rewritten for v4 (panel, port 8585, file list, lxml, pdftoppm, new-file formats, `xlsb`/`sxw stw sxc sxi` cannot be created).

* Session 7 (field report with phone screenshots): on the phone the Termux/Debian server ran WITHOUT `lxml` -> editors off -> panel "Not done" and app "Cannot start editing". Fixes: actionable error text + banner (`install_hint()`), `po_doctor.py` (checks files/lxml/soffice/poppler/LAN address), `install_deps.sh`, README_FA troubleshooting, app dialogs explain that `127.0.0.1` = the phone itself and how to reach a PC (`--lan`, IP, same Wi-Fi, firewall). Verified server-side: simulated missing lxml gives the clear message; access via the machine's LAN IP works, a foreign Host header is rejected (403). NOT verified on the user's real PC/phone network.

STILL OPEN:
1. Not tested anywhere: real Microsoft Office opening our edited/new files; the panel on a real phone browser; the APK built and run on a device (only a flow compile check with `java -m jdk.compiler/com.sun.tools.javac.Main` was possible — all errors are just missing Android symbols, none from project symbols). First thing to do: push to GitHub, let Actions build, and check the multi-window behaviour (`documentLaunchMode`, `LAUNCH_ADJACENT` needs Android 7+; on 6 the flag is ignored).
2. Ideas not done (deliberately, each touches the "never damage the file" guarantee and needs its own tests): Excel insert/delete rows, pptx shapes/moving, Word header/footer editing UI beyond wordtk settings, panel PDF-side text copy, auth token for `--lan` (the app and the panel would both need to send it).

## 5. Design rules to keep (owner's requirements)
* Original file is only READ. Edits happen in a server-side copy; saving = download of a new file (`/edit/export?as=ext`).
* Editing is opt-in, never automatic when a file is opened.
* Before editing: preflight report (signatures, macros, protection → BLOCK, tracked changes, fields, shapes, sheet protection, charts …) shown AND logged by the server (stdout + `~/.pocket_office/edit_log.txt`). Before saving: verification (package diff, XML well-formed, text invariants, locked paragraphs untouched, LibreOffice opens it).
* Document protection (`documentProtection enforcement`, protected sheets, pptx `modifyVerifier`) is respected = editing refused.
* Honesty: say what is untested.

## 6. How to run the tests (paths inside the scripts assume `/tmp/ed` + `/home/claude/po/PocketOffice`; fixtures are in `tests/fixtures/`)
```
mkdir -p /tmp/ed && cp tests/fixtures/* tests/*.py tests/*.js /tmp/ed/ && cd /tmp/ed
python3 t1.py; python3 t2.py; python3 t3.py; python3 t4.py; python3 t5.py; python3 t6.py; python3 t7.py; python3 t8.py
node edit_in_browser_test.js
node ui_test.js     # needs puppeteer-core + Chromium (sandbox: /opt/npm-tools/node_modules/puppeteer-core, /opt/pw-browsers/chromium-1243/...)
```
Never use `pkill -f` with a pattern that appears in your own command line (it kills the shell).

## 7. Server API summary (all JSON unless noted; `sid` = session)
`POST /edit/open?ext=&name=` (body=file) · `POST /edit/new?ext=&name=` · `GET /edit/sessions` · `GET /edit/catalog` · `GET /edit/open_payload|blocks|schema|find|cells|grid|render|page|preview|verify|export?sid=…` ·
`POST /edit/apply` {ops} · `POST /edit/settings` {group,values,scope,dry} · `POST /edit/undo` · `POST /edit/close`.
Ops: Word/PPT `set_text, format, para, insert_paragraph_after, delete_paragraph` (+Word `insert_row, delete_row, page`); Excel `set_cell`.
