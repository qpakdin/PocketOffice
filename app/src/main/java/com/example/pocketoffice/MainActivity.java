package com.example.pocketoffice;

import android.Manifest;
import android.annotation.SuppressLint;
import android.app.Activity;
import android.app.AlertDialog;
import android.app.ProgressDialog;
import android.content.DialogInterface;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageManager;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Paint;
import android.graphics.RectF;
import android.graphics.pdf.PdfDocument;
import android.graphics.pdf.PdfRenderer;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.ParcelFileDescriptor;
import android.print.PrintAttributes;
import android.print.PrintDocumentAdapter;
import android.print.PrintManager;
import android.text.InputType;
import android.util.Base64;
import android.view.Gravity;
import android.view.View;
import android.webkit.JavascriptInterface;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Button;
import android.widget.EditText;
import android.widget.FrameLayout;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.TextView;
import android.widget.Toast;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

public class MainActivity extends Activity {
    private static final int OPEN_FILE = 10, REQ_STORAGE = 20;
    private static final String HOST = "app.local", DEFAULT_SERVER = "http://127.0.0.1:8585", OLD_DEFAULT_SERVER = "http://127.0.0.1:8765";

    private static final class PageFile {
        final File file; final int w, h;
        PageFile(File f, int w, int h) { this.file = f; this.w = w; this.h = h; }
    }
    private interface Job { Storage.Saved run() throws Exception; }

    // shared state
    private volatile Uri currentUri;
    private String currentName = "document";
    private String currentExt = "";
    private TextEditorScreen editor;
    private static final int SAVE_TEXT = TextEditorScreen.REQ_SAVE;
    private boolean exact = false;
    private Button pageBtn, zoomBtn;
    private Runnable pendingAfterPermission;

    // exact mode (real layout via LibreOffice converter)
    private PageView pageView;
    private File exactPdf;

    // quick mode (approximate, offline WebView preview)
    private WebView webView;
    private int pageCount = 0, currentPage = 1;
    private volatile boolean ready = false;
    private final List<PageFile> exportPages = new ArrayList<>();
    private File exportDir;
    private volatile boolean exportCancelled = false;
    private volatile String exportKind = "pdf";
    private ProgressDialog progress;

    @Override public void onCreate(Bundle saved) {
        super.onCreate(saved);
        buildHome();
        handleIncoming(getIntent());
    }

    @Override protected void onNewIntent(Intent intent) { super.onNewIntent(intent); setIntent(intent); handleIncoming(intent); }

    private void handleIncoming(Intent intent) {
        if (intent == null) return;
        Uri uri = intent.getData();
        if (uri == null && Intent.ACTION_SEND.equals(intent.getAction())) uri = intent.getParcelableExtra(Intent.EXTRA_STREAM);
        if (uri != null) openDoc(uri);
    }

    // ------------------------------------------------------------ Home / settings

    private String serverUrl() {
        String s = getSharedPreferences("po", MODE_PRIVATE).getString("server", DEFAULT_SERVER).trim();
        if (s.equals(OLD_DEFAULT_SERVER) || s.isEmpty()) s = DEFAULT_SERVER;      // the v4 server listens on 8585
        if (!s.startsWith("http://") && !s.startsWith("https://")) s = "http://" + s;
        while (s.endsWith("/")) s = s.substring(0, s.length() - 1);
        return s;
    }

    private void buildHome() {
        destroyViews();
        LinearLayout root = new LinearLayout(this); root.setOrientation(LinearLayout.VERTICAL); root.setBackgroundColor(Color.rgb(248, 249, 252));
        TextView title = new TextView(this); title.setText("Pocket Office"); title.setTextSize(25); title.setTextColor(Color.rgb(25, 35, 55)); title.setPadding(20, 24, 20, 8); root.addView(title);
        TextView sub = new TextView(this);
        sub.setText("Office files are shown exactly as the real layout engine prints them (page size, margins, text boxes, tables). Save as PDF, PNG pages or TXT.\n\nOffice: Word, Excel and PowerPoint in all common formats (docx docm dotx doc rtf odt, xlsx xlsm xls ods, pptx ppsx ppt odp ...).\nText: txt, py, sh, json, md, csv ... any text file opens in a simple editor. Edits are always saved as a COPY; the original is never changed.\n\nSaved files go to Download/PocketOffice");
        sub.setTextSize(14); sub.setTextColor(Color.GRAY); sub.setPadding(20, 0, 20, 18); root.addView(sub);
        Button open = new Button(this); open.setText("Open Office file"); open.setAllCaps(false); open.setTextSize(16); open.setOnClickListener(v -> chooseFile());
        LinearLayout.LayoutParams lp = new LinearLayout.LayoutParams(-1, -2); lp.setMargins(16, 6, 16, 6); root.addView(open, lp);
        Button set = new Button(this); set.setText("Converter settings / test"); set.setAllCaps(false); set.setOnClickListener(v -> askServer());
        LinearLayout.LayoutParams lp2 = new LinearLayout.LayoutParams(-1, -2); lp2.setMargins(16, 6, 16, 6); root.addView(set, lp2);
        Button fresh = new Button(this); fresh.setText("New Word / Excel / PowerPoint file (browser editor)"); fresh.setAllCaps(false); fresh.setOnClickListener(v -> openPanel("/panel/#/"));
        LinearLayout.LayoutParams lp3 = new LinearLayout.LayoutParams(-1, -2); lp3.setMargins(16, 6, 16, 6); root.addView(fresh, lp3);
        Button win = new Button(this); win.setText("Open a new window"); win.setAllCaps(false); win.setOnClickListener(v -> openNewWindow(null));
        LinearLayout.LayoutParams lp4 = new LinearLayout.LayoutParams(-1, -2); lp4.setMargins(16, 6, 16, 6); root.addView(win, lp4);
        ScrollView scroll = new ScrollView(this); scroll.setFitsSystemWindows(true); scroll.addView(root); setContentView(scroll);
    }

    // ------------------------------------------------------------ Browser editor + extra windows

    /** Opens the converter's web panel (the editor) in the phone's browser. 'path' is relative to the converter address. */
    private void openPanel(String path) {
        try { startActivity(new Intent(Intent.ACTION_VIEW, Uri.parse(serverUrl() + path))); }
        catch (Exception e) { toast("No browser found to open " + serverUrl() + path); }
    }

    /** Another independent window of this app (split-screen / freeform), optionally already showing 'uri'. */
    private void openNewWindow(Uri uri) {
        Intent i = new Intent(this, MainActivity.class);
        i.addFlags(Intent.FLAG_ACTIVITY_NEW_DOCUMENT | Intent.FLAG_ACTIVITY_MULTIPLE_TASK | Intent.FLAG_ACTIVITY_LAUNCH_ADJACENT);
        if (uri != null) { i.setData(uri); i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION); }
        try { startActivity(i); } catch (Exception e) { toast("Could not open a new window: " + e.getMessage()); }
    }

    private static final String NET_HELP =
            "127.0.0.1 means THIS phone. To use a converter on a PC:\n"
            + "1) start it on the PC with:  python pocket_office_server.py --lan\n"
            + "2) type the address it prints (In the app: http://192.168.x.x:8585) here\n"
            + "3) phone and PC on the same Wi-Fi; allow Python in the Windows firewall (Private network)\n"
            + "To use the phone itself: start the server in Termux/Debian first.";

    private void askServer() {
        final EditText input = new EditText(this); input.setSingleLine(true); input.setText(serverUrl());
        new AlertDialog.Builder(this).setTitle("Converter address")
                .setMessage("The converter (LibreOffice) runs in Termux on this phone, or on a PC in your network.\nDefault: " + DEFAULT_SERVER + "  (= this phone)\nFor a PC: http://PC-address:8585  (start the PC server with --lan)")
                .setView(wrap(input)).setNegativeButton("Cancel", null)
                .setNeutralButton("Test", (d, w) -> {
                    getSharedPreferences("po", MODE_PRIVATE).edit().putString("server", input.getText().toString().trim()).apply();
                    final String url = serverUrl();
                    new Thread(() -> {
                        final String info = ExactClient.ping(url);
                        runOnUiThread(() -> new AlertDialog.Builder(this).setTitle(info != null ? "Connected" : "Not reachable").setMessage(info != null ? info + (info.contains("editors=MISSING") ? "\n\nEDITING IS OFF on that converter: install lxml there (apt install -y python3-lxml) and restart it, or run: python3 po_doctor.py" : "") : "No converter answered at\n" + url + "\n\n" + NET_HELP).setPositiveButton("OK", null).show());
                    }).start();
                })
                .setPositiveButton("Save", (d, w) -> { getSharedPreferences("po", MODE_PRIVATE).edit().putString("server", input.getText().toString().trim()).apply(); toast("Saved"); }).show();
    }

    private void chooseFile() {
        Intent i = new Intent(Intent.ACTION_OPEN_DOCUMENT); i.addCategory(Intent.CATEGORY_OPENABLE); i.setType("*/*");
        startActivityForResult(i, OPEN_FILE);
    }

    @Override protected void onActivityResult(int request, int result, Intent data) {
        super.onActivityResult(request, result, data);
        if (request == OPEN_FILE && result == RESULT_OK && data != null && data.getData() != null) openDoc(data.getData());
        else if (request == SAVE_TEXT && result == RESULT_OK && data != null && editor != null) editor.saveTo(data.getData());
    }

    // ------------------------------------------------------------ Open

    private void openDoc(final Uri uri) {
        String name = Storage.displayName(this, uri);
        if (name == null) name = uri.getLastPathSegment() != null ? uri.getLastPathSegment() : "document";
        final String fname = name, ext = Formats.extOf(name);
        if (Formats.isOffice(ext)) {
            currentUri = uri; currentName = fname; currentExt = ext;
            startExact();
            return;
        }
        // not an Office extension: is it readable text (txt, py, sh, json, md, csv, config files ...)?
        new Thread(() -> {
            boolean text = false;
            try (InputStream in = getContentResolver().openInputStream(uri)) {
                byte[] head = new byte[65536]; int n = 0, r;
                while (in != null && n < head.length && (r = in.read(head, n, head.length - n)) > 0) n += r;
                text = Formats.looksLikeText(head, n);
            } catch (Exception ignored) {}
            final boolean isText = text;
            runOnUiThread(() -> {
                if (isFinishing()) return;
                if (isText) openTextEditor(uri, fname);
                else new AlertDialog.Builder(this).setTitle("Unsupported file")
                        .setMessage("\"" + fname + "\" is not an Office document or a text file.\n\nPDF files: use Pocket PDF.\nOffice: docx docm dotx doc rtf odt, xlsx xlsm xls ods, pptx ppsx ppt odp and similar.")
                        .setPositiveButton("OK", null).show();
            });
        }).start();
    }

    private void openTextEditor(Uri uri, String name) {
        destroyViews();
        exact = false;
        editor = new TextEditorScreen(this, new TextEditorScreen.Host() {
            @Override public void goHome() { buildHome(); }
            @Override public void toast(String s) { MainActivity.this.toast(s); }
        });
        setContentView(editor.build(uri, name));
    }

    private void startExact() {
        final String server = serverUrl(), type = currentExt;
        final Uri uri = currentUri;
        final ProgressDialog pd = new ProgressDialog(this);
        pd.setMessage("Checking converter…"); pd.setCancelable(false); pd.show();
        new Thread(() -> {
            String err = null; File pdf = new File(getCacheDir(), "exact_" + System.nanoTime() + ".pdf");
            final String info = ExactClient.ping(server);
            if (info != null) {
                runOnUiThread(() -> pd.setMessage("Converting with the real layout engine…\n(the first file can take up to a minute)"));
                try { ExactClient.convert(this, server, uri, type, pdf); } catch (Throwable t) { err = t.getMessage() != null ? t.getMessage() : t.toString(); }
            }
            final String ferr = err; final File fpdf = pdf;
            runOnUiThread(() -> {
                try { pd.dismiss(); } catch (Exception ignored) {}
                if (isFinishing()) return;
                if (info == null) showNoServer(server);
                else if (ferr != null) { toast(ferr); fpdf.delete(); }
                else { showExactReader(fpdf); String mf = ExactClient.lastMissingFonts; if (mf != null && !mf.isEmpty()) toast("Fonts not found on the converter machine (the layout may differ):\n" + mf.replace(",", ", ")); }
            });
        }).start();
    }

    private void showNoServer(String server) {
        final String quick = Formats.quickType(currentExt);
        AlertDialog.Builder b = new AlertDialog.Builder(this).setTitle("Converter is not running")
                .setMessage("Exact page layout needs the converter (LibreOffice) at:\n" + server + "\n\nStart it (see converter/README_FA.md)."
                        + (quick != null ? "\n\nOr use the approximate quick preview, which is built into the app but does not keep the real layout."
                        : "\n\nThis format (." + currentExt + ") can only be shown through the converter."))
                .setNeutralButton("Settings", (d, w) -> askServer())
                .setNegativeButton("Cancel", null);
        if (quick != null) b.setPositiveButton("Quick preview", (d, w) -> showViewer(quick));
        b.show();
    }

    // ------------------------------------------------------------ Exact reader (PDF produced by LibreOffice)

    private View readerShell(View content, boolean zoomBar) {
        LinearLayout col = new LinearLayout(this); col.setOrientation(LinearLayout.VERTICAL); col.setBackgroundColor(Color.rgb(245, 246, 249)); col.setFitsSystemWindows(true);
        LinearLayout top = new LinearLayout(this); top.setGravity(Gravity.CENTER_VERTICAL); top.setPadding(dp(6), dp(4), dp(6), dp(4));
        top.addView(smallButton("Home", v -> buildHome()));
        pageBtn = smallButton("…", v -> askPage());
        top.addView(pageBtn, new LinearLayout.LayoutParams(0, -2, 1));
        top.addView(smallButton("Export ▾", v -> showExportMenu()));
        col.addView(top);
        col.addView(content, new LinearLayout.LayoutParams(-1, 0, 1));
        if (zoomBar) {
            LinearLayout bottom = new LinearLayout(this); bottom.setGravity(Gravity.CENTER_VERTICAL); bottom.setPadding(dp(6), dp(4), dp(6), dp(4));
            bottom.addView(smallButton("‹", v -> { if (pageView != null) pageView.goTo(pageView.getPageIndex() - 1, false); }), new LinearLayout.LayoutParams(0, -2, 1));
            bottom.addView(smallButton("−", v -> { if (pageView != null) pageView.zoomBy(1f / 1.25f); }), new LinearLayout.LayoutParams(0, -2, 1));
            zoomBtn = smallButton("100%", v -> askZoom());
            bottom.addView(zoomBtn, new LinearLayout.LayoutParams(0, -2, 1.4f));
            bottom.addView(smallButton("+", v -> { if (pageView != null) pageView.zoomBy(1.25f); }), new LinearLayout.LayoutParams(0, -2, 1));
            bottom.addView(smallButton("Fit", v -> { if (pageView != null) pageView.fitPage(); }), new LinearLayout.LayoutParams(0, -2, 1));
            bottom.addView(smallButton("›", v -> { if (pageView != null) pageView.goTo(pageView.getPageIndex() + 1, false); }), new LinearLayout.LayoutParams(0, -2, 1));
            col.addView(bottom);
        }
        return col;
    }

    private void showExactReader(File pdf) {
        destroyViews();
        ParcelFileDescriptor fd = null;
        try {
            fd = ParcelFileDescriptor.open(pdf, ParcelFileDescriptor.MODE_READ_ONLY);
            PdfRenderer r = new PdfRenderer(fd);
            if (r.getPageCount() == 0) { r.close(); throw new Exception("The converted document has no pages"); }
            exact = true; exactPdf = pdf;
            pageView = new PageView(this);
            pageView.setListener(new PageView.Listener() {
                @Override public void onChanged() { updatePageLabel(); }
                @Override public void onError(String message) { toast(message); }
            });
            setContentView(readerShell(pageView, true));
            pageView.setDocument(fd, r);
        } catch (Exception e) {
            try { if (fd != null) fd.close(); } catch (Exception ignored) {}
            toast("Cannot show the converted file: " + e.getMessage());
            buildHome();
        }
    }

    private void askZoom() {
        if (pageView == null) return;
        final String[] items = {"Fit width (100%)", "Fit page", "150%", "200%", "300%", "500%", "800%", "1500%", "Custom…"};
        final float[] z = {1f, -1f, 1.5f, 2f, 3f, 5f, 8f, 15f};
        new AlertDialog.Builder(this).setTitle("Zoom (up to " + Math.round(pageView.getMaxZoom() * 100) + "% on this page)").setItems(items, (d, i) -> {
            if (pageView == null) return;
            if (i == 1) pageView.fitPage();
            else if (i < z.length) pageView.setZoom(z[i]);
            else {
                final EditText input = new EditText(this); input.setInputType(InputType.TYPE_CLASS_NUMBER); input.setHint("25 - 2000");
                new AlertDialog.Builder(this).setTitle("Zoom percent").setView(wrap(input)).setNegativeButton("Cancel", null)
                        .setPositiveButton("OK", (d2, w2) -> { try { if (pageView != null) pageView.setZoom(Integer.parseInt(input.getText().toString().trim()) / 100f); } catch (NumberFormatException e) { toast("Enter a number"); } }).show();
            }
        }).show();
    }

    // ------------------------------------------------------------ Editing (optional, on a COPY, started only by the user)

    /** Sends the file to the converter as an editing session (the original is only read) and opens the web panel on that session. */
    private void startBrowserEdit() {
        if (Formats.quickType(currentExt) == null) {
            new AlertDialog.Builder(this).setTitle("Editing is not available for this format")
                    .setMessage("Editing works for Word (.docx .docm .dotx .dotm), Excel (.xlsx .xlsm .xltx .xltm) and PowerPoint (.pptx .pptm .ppsx .ppsm .potx .potm). Other formats (.doc, .odt, .rtf, .xls, .ppt ...) would have to be re-written by LibreOffice, which could damage features of the document, so they are view-only here.\n\nTo edit one of them, create a new file of the same kind in the panel and copy the content, or convert it to the OOXML format first.")
                    .setPositiveButton("OK", null).show();
            return;
        }
        final Uri doc = currentUri; final String ext = currentExt, name = currentName;
        final ProgressDialog pd = new ProgressDialog(this);
        pd.setMessage("Sending a copy to the converter…"); pd.setCancelable(false); pd.show();
        new Thread(() -> {
            String sid = null, err = null;
            try {
                EditSession session = new EditSession(serverUrl());
                session.open(this, doc, ext, name);
                sid = session.sid;
            } catch (EditSession.Refused e) { err = e.getMessage(); }
            catch (Exception e) { err = "Cannot reach the converter at " + serverUrl() + "\n(" + (e.getMessage() != null ? e.getMessage() : e.toString()) + ")\n\n" + NET_HELP; }
            final String fsid = sid, ferr = err;
            runOnUiThread(() -> {
                try { pd.dismiss(); } catch (Exception ignored) {}
                if (isFinishing()) return;
                if (ferr != null) new AlertDialog.Builder(this).setTitle("Cannot start editing").setMessage(ferr).setPositiveButton("OK", null).show();
                else openPanel("/panel/#/s/" + fsid);
            });
        }).start();
    }

    // exact-mode exports (the PDF is LibreOffice's own output, so it is identical to what you see)
    private void exactSavePdf() {
        final File pdf = exactPdf; final String name = baseName(currentName) + ".pdf";
        runJob("Saving PDF…", () -> Storage.save(this, pdf, name, "application/pdf", null), "application/pdf");
    }

    private void exactSavePng() {
        final File pdf = exactPdf; final String sub = baseName(currentName) + "_pages_" + stamp();
        runJob("Rendering pages…", () -> {
            int count = 0;
            try (ParcelFileDescriptor fd = ParcelFileDescriptor.open(pdf, ParcelFileDescriptor.MODE_READ_ONLY); PdfRenderer r = new PdfRenderer(fd)) {
                for (int i = 0; i < r.getPageCount(); i++) {
                    File tmp = File.createTempFile("pg_", ".png", getCacheDir());
                    try {
                        try (PdfRenderer.Page p = r.openPage(i)) {
                            float s = Math.min(2.5f, 4096f / Math.max(p.getWidth(), p.getHeight()));
                            Bitmap b = Bitmap.createBitmap(Math.max(1, Math.round(p.getWidth() * s)), Math.max(1, Math.round(p.getHeight() * s)), Bitmap.Config.ARGB_8888);
                            b.eraseColor(Color.WHITE);
                            p.render(b, null, null, PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY);
                            try (FileOutputStream out = new FileOutputStream(tmp)) { b.compress(Bitmap.CompressFormat.PNG, 100, out); }
                            b.recycle();
                        }
                        Storage.save(this, tmp, String.format(Locale.US, "page_%03d.png", i + 1), "image/png", sub);
                    } finally { tmp.delete(); }
                    count++;
                }
            }
            Storage.Saved s = new Storage.Saved(null, Storage.SHOWN_ROOT + "/" + sub); s.count = count; return s;
        }, "image/png");
    }

    private void exactSaveTxt() {
        final File pdf = exactPdf; final String server = serverUrl(); final String name = baseName(currentName) + ".txt";
        final Uri doc = currentUri; final String ext = currentExt; final boolean ooxml = Formats.quickType(ext) != null;
        runJob("Extracting text…", () -> {
            byte[] txt = null;
            if (ooxml) { try { txt = ExactClient.docText(this, server, doc, ext); } catch (Exception ignored) { txt = null; } }   // logical order, best for Persian
            if (txt == null) txt = ExactClient.pdfText(server, pdf);
            File tmp = File.createTempFile("txt_", ".txt", getCacheDir());
            try {
                try (FileOutputStream out = new FileOutputStream(tmp)) { out.write(new byte[]{(byte) 0xEF, (byte) 0xBB, (byte) 0xBF}); out.write(txt); }
                return Storage.save(this, tmp, name, "text/plain", null);
            } finally { tmp.delete(); }
        }, "text/plain");
    }

    private void runJob(final String message, final Job job, final String mime) {
        ensureStorage(() -> {
            final ProgressDialog pd = new ProgressDialog(this);
            pd.setMessage(message); pd.setCancelable(false); pd.show();
            new Thread(() -> {
                Storage.Saved s = null; String err = null;
                try { s = job.run(); } catch (Throwable t) { err = t.getMessage() != null ? t.getMessage() : t.toString(); }
                final Storage.Saved fs = s; final String ferr = err;
                runOnUiThread(() -> {
                    try { pd.dismiss(); } catch (Exception ignored) {}
                    if (isFinishing()) return;
                    if (ferr != null) toast("Failed: " + ferr); else showSaved(fs, mime);
                });
            }).start();
        });
    }

    // ------------------------------------------------------------ Page label / go to page / export menu (both modes)

    private void updatePageLabel() {
        if (pageBtn == null) return;
        if (exact && pageView != null) {
            pageBtn.setText((pageView.getPageIndex() + 1) + " / " + pageView.getPageCount());
            if (zoomBtn != null) zoomBtn.setText(Math.round(pageView.getZoom() * 100) + "%");
        } else {
            pageBtn.setText(ready ? currentPage + " / " + pageCount + "   (quick preview)" : "Rendering…");
        }
    }

    private void askPage() {
        final int count = exact ? (pageView != null ? pageView.getPageCount() : 0) : pageCount;
        if (count == 0 || (!exact && !ready)) return;
        final int cur = exact ? pageView.getPageIndex() + 1 : currentPage;
        final EditText input = new EditText(this); input.setInputType(InputType.TYPE_CLASS_NUMBER); input.setHint("1 - " + count); input.setText(String.valueOf(cur)); input.selectAll();
        new AlertDialog.Builder(this).setTitle("Go to page (1 - " + count + ")").setView(wrap(input)).setNegativeButton("Cancel", null)
                .setPositiveButton("Go", (d, w) -> {
                    try {
                        int n = Integer.parseInt(input.getText().toString().trim());
                        if (n < 1 || n > count) { toast("Page must be between 1 and " + count); return; }
                        if (exact) { if (pageView != null) pageView.goTo(n - 1, false); }
                        else if (webView != null) webView.evaluateJavascript("goPage(" + n + ")", null);
                    } catch (NumberFormatException e) { toast("Enter a page number"); }
                }).show();
    }

    private void showExportMenu() {
        if (exact) {
            final String[] items = {"Save as PDF (exact)", "Save pages as PNG images", "Save text as TXT", "Edit a copy in the browser…", "Open this file in a new window"};
            new AlertDialog.Builder(this).setTitle("Export / Edit").setItems(items, (d, i) -> { if (i == 0) exactSavePdf(); else if (i == 1) exactSavePng(); else if (i == 2) exactSaveTxt(); else if (i == 3) startBrowserEdit(); else openNewWindow(currentUri); }).show();
        } else {
            if (!ready || webView == null) { toast("The document is still loading"); return; }
            final String[] items = {"Save as PDF (approximate)", "Save pages as PNG images", "Save text as TXT", "Print / system PDF (vector)"};
            new AlertDialog.Builder(this).setTitle("Export").setItems(items, (d, i) -> {
                if (i == 0) startPagesExport("pdf"); else if (i == 1) startPagesExport("png"); else if (i == 2) startTextExport(); else printNow();
            }).show();
        }
    }

    // ------------------------------------------------------------ Quick preview (approximate, fully offline, WebView)

    @SuppressLint("SetJavaScriptEnabled")
    private void showViewer(String type) {
        destroyViews();
        exact = false; ready = false; pageCount = 0; currentPage = 1;
        webView = new WebView(this);
        WebSettings ws = webView.getSettings();
        ws.setJavaScriptEnabled(true); ws.setDomStorageEnabled(true); ws.setAllowFileAccess(false); ws.setAllowContentAccess(false);
        ws.setSupportZoom(true); ws.setBuiltInZoomControls(true); ws.setDisplayZoomControls(false);
        ws.setUseWideViewPort(true); ws.setLoadWithOverviewMode(true); ws.setMixedContentMode(WebSettings.MIXED_CONTENT_NEVER_ALLOW);
        webView.setLongClickable(false); webView.setHapticFeedbackEnabled(false); webView.setOnLongClickListener(v -> true); // showcase only: no selecting / copying
        webView.addJavascriptInterface(new Bridge(), "Android");
        webView.setWebViewClient(new WebViewClient() {
            @Override public WebResourceResponse shouldInterceptRequest(WebView v, WebResourceRequest req) { return serve(req.getUrl()); }
        });
        setContentView(readerShell(webView, false));
        updatePageLabel();
        webView.loadUrl("https://" + HOST + "/viewer.html?type=" + type);
    }

    private WebResourceResponse serve(Uri u) {
        try {
            if (!HOST.equals(u.getHost())) return reply(403, "text/plain", new ByteArrayInputStream(new byte[0]));
            String path = u.getPath() == null ? "/" : u.getPath();
            if ("/doc".equals(path)) {
                Uri doc = currentUri;
                if (doc == null) return reply(404, "text/plain", new ByteArrayInputStream(new byte[0]));
                InputStream in = getContentResolver().openInputStream(doc);
                if (in == null) return reply(404, "text/plain", new ByteArrayInputStream(new byte[0]));
                return reply(200, "application/octet-stream", in);
            }
            String asset = path.startsWith("/") ? path.substring(1) : path;
            if (asset.isEmpty() || asset.contains("..")) return reply(403, "text/plain", new ByteArrayInputStream(new byte[0]));
            return reply(200, mimeFor(asset), getAssets().open(asset));
        } catch (Exception e) { return reply(404, "text/plain", new ByteArrayInputStream(new byte[0])); }
    }

    private static WebResourceResponse reply(int code, String mime, InputStream in) {
        Map<String, String> headers = new HashMap<>(); headers.put("Cache-Control", "no-store");
        return new WebResourceResponse(mime, "UTF-8", code, code == 200 ? "OK" : (code == 403 ? "Forbidden" : "Not Found"), headers, in);
    }

    private static String mimeFor(String name) {
        if (name.endsWith(".html")) return "text/html"; if (name.endsWith(".js")) return "application/javascript";
        if (name.endsWith(".css")) return "text/css"; if (name.endsWith(".json")) return "application/json";
        return "application/octet-stream";
    }

    private void printNow() {
        try {
            PrintManager pm = (PrintManager) getSystemService(PRINT_SERVICE);
            String job = baseName(currentName);
            PrintDocumentAdapter adapter = webView.createPrintDocumentAdapter(job);
            pm.print(job, adapter, new PrintAttributes.Builder().build());
        } catch (Exception e) { toast("Print is not available: " + e.getMessage()); }
    }

    private void startPagesExport(final String kind) {
        ensureStorage(() -> {
            if (webView == null) return;
            exportKind = kind; exportCancelled = false;
            synchronized (exportPages) { exportPages.clear(); }
            exportDir = new File(getCacheDir(), "export"); deleteDir(exportDir); exportDir.mkdirs();
            progress = new ProgressDialog(this);
            progress.setMessage("Rendering pages…"); progress.setCancelable(false);
            progress.setButton(DialogInterface.BUTTON_NEGATIVE, "Cancel", (d, w) -> { exportCancelled = true; if (webView != null) webView.evaluateJavascript("cancelExport()", null); });
            progress.show();
            webView.evaluateJavascript("exportPages(2)", null);
        });
    }

    private void startTextExport() { ensureStorage(() -> { if (webView != null) webView.evaluateJavascript("exportText()", null); }); }

    private void finishPagesExport() {
        final String kind = exportKind;
        final List<PageFile> list;
        synchronized (exportPages) { list = new ArrayList<>(exportPages); }
        if (list.isEmpty()) { dismissProgress(); toast("Nothing was rendered"); return; }
        if (progress != null) progress.setMessage("pdf".equals(kind) ? "Building PDF…" : "Saving images…");
        final String base = baseName(currentName);
        new Thread(() -> {
            Storage.Saved saved = null; String err = null;
            try {
                if ("pdf".equals(kind)) {
                    File out = new File(getCacheDir(), "out_" + System.nanoTime() + ".pdf");
                    try { buildPdf(list, out); saved = Storage.save(this, out, base + ".pdf", "application/pdf", null); } finally { out.delete(); }
                } else {
                    String sub = base + "_pages_" + stamp();
                    for (int i = 0; i < list.size(); i++) Storage.save(this, list.get(i).file, String.format(Locale.US, "page_%03d.png", i + 1), "image/png", sub);
                    saved = new Storage.Saved(null, Storage.SHOWN_ROOT + "/" + sub); saved.count = list.size();
                }
            } catch (Throwable t) { err = t.getMessage() != null ? t.getMessage() : t.toString(); }
            deleteDir(exportDir);
            final Storage.Saved fs = saved; final String ferr = err;
            runOnUiThread(() -> { dismissProgress(); if (ferr != null) toast("Export failed: " + ferr); else showSaved(fs, "application/pdf"); });
        }).start();
    }

    private void buildPdf(List<PageFile> list, File out) throws Exception {
        PdfDocument doc = new PdfDocument(); Paint paint = new Paint(Paint.FILTER_BITMAP_FLAG);
        try {
            for (int i = 0; i < list.size(); i++) {
                PageFile pf = list.get(i);
                int wPt = Math.max(1, Math.round(pf.w * 0.75f)), hPt = Math.max(1, Math.round(pf.h * 0.75f));
                Bitmap b = BitmapFactory.decodeFile(pf.file.getAbsolutePath()); if (b == null) continue;
                PdfDocument.Page page = doc.startPage(new PdfDocument.PageInfo.Builder(wPt, hPt, i + 1).create());
                Canvas c = page.getCanvas(); c.drawColor(Color.WHITE); c.drawBitmap(b, null, new RectF(0, 0, wPt, hPt), paint);
                doc.finishPage(page); b.recycle();
            }
            try (FileOutputStream fos = new FileOutputStream(out)) { doc.writeTo(fos); }
        } finally { doc.close(); }
    }

    private void dismissProgress() { try { if (progress != null) progress.dismiss(); } catch (Exception ignored) {} progress = null; }

    private final class Bridge {
        @JavascriptInterface public void onReady(final int pages, final String type) { runOnUiThread(() -> { pageCount = pages; currentPage = 1; ready = true; updatePageLabel(); }); }
        @JavascriptInterface public void onPage(final int page) { runOnUiThread(() -> { currentPage = page; updatePageLabel(); }); }
        @JavascriptInterface public void onError(final String m) { runOnUiThread(() -> { ready = false; if (pageBtn != null) pageBtn.setText("Cannot display this file"); }); }
        @JavascriptInterface public void progress(final int i, final int n) { runOnUiThread(() -> { if (progress != null) progress.setMessage("Rendering page " + i + " / " + n + "…"); }); }
        @JavascriptInterface public void savePage(int index, int w, int h, String base64) {
            try {
                byte[] data = Base64.decode(base64, Base64.DEFAULT);
                File f = new File(exportDir, String.format(Locale.US, "p%04d.png", index + 1));
                try (FileOutputStream out = new FileOutputStream(f)) { out.write(data); }
                synchronized (exportPages) { exportPages.add(new PageFile(f, w, h)); }
            } catch (Throwable t) {
                final String m = t.toString();
                runOnUiThread(() -> { dismissProgress(); toast("Cannot store page: " + m); });
                exportCancelled = true;
            }
        }
        @JavascriptInterface public void exportDone() { runOnUiThread(() -> { if (exportCancelled) dismissProgress(); else finishPagesExport(); }); }
        @JavascriptInterface public void exportFailed(final String m) { runOnUiThread(() -> { dismissProgress(); if (!"cancelled".equals(m)) toast("Export failed: " + m); }); }
        @JavascriptInterface public void saveText(final String text) {
            new Thread(() -> {
                Storage.Saved saved = null; String err = null;
                File tmp = new File(getCacheDir(), "txt_" + System.nanoTime() + ".txt");
                try {
                    try (FileOutputStream out = new FileOutputStream(tmp)) { out.write(new byte[]{(byte) 0xEF, (byte) 0xBB, (byte) 0xBF}); out.write(text.getBytes("UTF-8")); }
                    saved = Storage.save(MainActivity.this, tmp, baseName(currentName) + ".txt", "text/plain", null);
                } catch (Throwable t) { err = t.getMessage() != null ? t.getMessage() : t.toString(); }
                tmp.delete();
                final Storage.Saved fs = saved; final String ferr = err;
                runOnUiThread(() -> { if (ferr != null) toast("Cannot save text: " + ferr); else showSaved(fs, "text/plain"); });
            }).start();
        }
        @JavascriptInterface public void copyText(final String text) { /* disabled: this app is a showcase, not an editor */ }
    }

    // ------------------------------------------------------------ Saved dialog / permission / helpers

    private void showSaved(final Storage.Saved s, final String mime) {
        if (s == null) return;
        AlertDialog.Builder b = new AlertDialog.Builder(this);
        if (s.uri == null) b.setTitle("Done").setMessage(s.count + " page image(s) saved in:\n" + s.where).setPositiveButton("OK", null);
        else b.setTitle("Saved").setMessage(s.where)
                .setPositiveButton("Open", (d, w) -> viewFile(s.uri, mime)).setNeutralButton("Share", (d, w) -> share(s.uri, mime)).setNegativeButton("Close", null);
        b.show();
    }

    private void viewFile(Uri uri, String mime) {
        Intent i = new Intent(Intent.ACTION_VIEW); i.setDataAndType(uri, mime); i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        try { startActivity(i); } catch (Exception e) { toast("No app can open this file"); }
    }

    private void share(Uri uri, String mime) {
        Intent i = new Intent(Intent.ACTION_SEND); i.setType(mime); i.putExtra(Intent.EXTRA_STREAM, uri); i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        try { startActivity(Intent.createChooser(i, "Share")); } catch (Exception e) { toast("No app can share this file"); }
    }

    private void ensureStorage(Runnable r) {
        if (Build.VERSION.SDK_INT < 29 && checkSelfPermission(Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
            pendingAfterPermission = r; requestPermissions(new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, REQ_STORAGE);
        } else r.run();
    }

    @Override public void onRequestPermissionsResult(int code, String[] perms, int[] res) {
        super.onRequestPermissionsResult(code, perms, res);
        if (code != REQ_STORAGE) return;
        Runnable r = pendingAfterPermission; pendingAfterPermission = null;
        if (res.length > 0 && res[0] == PackageManager.PERMISSION_GRANTED && r != null) r.run();
        else toast("Storage permission is required to save files to Downloads");
    }

    private void destroyViews() {
        if (pageView != null) { pageView.release(); pageView = null; }
        if (webView != null) {
            try { webView.stopLoading(); webView.removeJavascriptInterface("Android"); } catch (Exception ignored) {}
            if (webView.getParent() instanceof android.view.ViewGroup) ((android.view.ViewGroup) webView.getParent()).removeView(webView);
            try { webView.destroy(); } catch (Exception ignored) {}
            webView = null;
        }
        editor = null;
        pageBtn = null; zoomBtn = null; ready = false;
    }

    @Override public void onBackPressed() {
        if (editor != null) editor.requestClose();
        else if (pageView != null || webView != null) buildHome();
        else super.onBackPressed();
    }
    @Override protected void onDestroy() { destroyViews(); super.onDestroy(); }

    private View wrap(View v) { FrameLayout f = new FrameLayout(this); f.setPadding(dp(20), dp(8), dp(20), 0); f.addView(v); return f; }

    private Button smallButton(String label, View.OnClickListener l) {
        Button b = new Button(this); b.setText(label); b.setAllCaps(false);
        b.setMinWidth(0); b.setMinimumWidth(0); b.setMinHeight(0); b.setMinimumHeight(dp(44)); b.setPadding(dp(14), 0, dp(14), 0); b.setOnClickListener(l);
        return b;
    }

    private int dp(int v) { return Math.round(v * getResources().getDisplayMetrics().density); }
    private static String stamp() { return new SimpleDateFormat("yyyyMMdd-HHmm", Locale.US).format(new Date()); }

    private static String baseName(String n) {
        if (n == null || n.isEmpty()) return "document";
        int dot = n.lastIndexOf('.'); if (dot > 0) n = n.substring(0, dot);
        n = n.replaceAll("[\\\\/:*?\"<>|]", "_");
        return n.length() > 60 ? n.substring(0, 60) : n;
    }

    private static void deleteDir(File d) {
        if (d == null || !d.exists()) return;
        File[] fs = d.listFiles();
        if (fs != null) for (File f : fs) { if (f.isDirectory()) deleteDir(f); else f.delete(); }
        d.delete();
    }

    private void toast(String s) { Toast.makeText(this, s, Toast.LENGTH_LONG).show(); }
}
