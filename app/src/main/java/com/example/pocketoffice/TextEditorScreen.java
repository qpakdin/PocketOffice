package com.example.pocketoffice;

import android.app.Activity;
import android.app.AlertDialog;
import android.content.Intent;
import android.graphics.Color;
import android.graphics.Typeface;
import android.net.Uri;
import android.text.Editable;
import android.text.InputType;
import android.text.TextUtils;
import android.text.TextWatcher;
import android.util.TypedValue;
import android.view.Gravity;
import android.view.View;
import android.widget.Button;
import android.widget.CheckBox;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.TextView;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.ByteBuffer;
import java.nio.charset.CharacterCodingException;
import java.nio.charset.Charset;
import java.nio.charset.CodingErrorAction;
import java.nio.charset.StandardCharsets;
import java.util.ArrayDeque;
import java.util.Locale;

/**
 * Small text / code editor for .txt .py .sh .json .md ... (any text file).
 * The original file is never touched: the edited text is always written to a NEW file with "Save a copy".
 */
final class TextEditorScreen {
    interface Host { void goHome(); void toast(String s); }

    static final int REQ_SAVE = 30;
    private static final int MAX_BYTES = 5 * 1024 * 1024;

    private final Activity act;
    private final Host host;
    private EditText edit;
    private TextView titleView;
    private String name = "document", ext = "";
    private byte[] raw = new byte[0];
    private Charset charset = StandardCharsets.UTF_8;
    private boolean bom, crlf, dirty, loading, applying, wrap = true, mono = true;
    private float sizeSp = 14f;
    private final ArrayDeque<String> undo = new ArrayDeque<>(), redo = new ArrayDeque<>();
    private long lastChange;

    TextEditorScreen(Activity act, Host host) { this.act = act; this.host = host; }

    boolean hasUnsaved() { return dirty; }

    View build(final Uri uri, String fileName) {
        name = fileName == null ? "document" : fileName;
        ext = Formats.extOf(name);
        LinearLayout col = new LinearLayout(act); col.setOrientation(LinearLayout.VERTICAL); col.setBackgroundColor(Color.rgb(245, 246, 249)); col.setFitsSystemWindows(true);

        LinearLayout top = new LinearLayout(act); top.setGravity(Gravity.CENTER_VERTICAL); top.setPadding(dp(6), dp(4), dp(6), dp(4));
        top.addView(btn("Home", v -> requestClose()));
        titleView = new TextView(act); titleView.setSingleLine(true); titleView.setEllipsize(TextUtils.TruncateAt.MIDDLE); titleView.setTextSize(14); titleView.setPadding(dp(6), 0, dp(6), 0);
        top.addView(titleView, new LinearLayout.LayoutParams(0, -2, 1));
        top.addView(btn("Save a copy", v -> requestSave()));
        top.addView(btn("More ▾", v -> showMenu()));
        col.addView(top);

        edit = new EditText(act);
        edit.setGravity(Gravity.TOP | Gravity.START);
        edit.setInputType(InputType.TYPE_CLASS_TEXT | InputType.TYPE_TEXT_FLAG_MULTI_LINE | InputType.TYPE_TEXT_FLAG_NO_SUGGESTIONS);
        edit.setBackgroundColor(Color.WHITE); edit.setTextColor(Color.rgb(20, 25, 35));
        edit.setPadding(dp(8), dp(8), dp(8), dp(8));
        edit.setHorizontallyScrolling(!wrap);
        applyFont();
        edit.addTextChangedListener(new TextWatcher() {
            @Override public void beforeTextChanged(CharSequence s, int st, int count, int after) {
                if (loading || applying) return;
                long now = System.currentTimeMillis();
                if (now - lastChange > 1000 || Math.abs(after - count) > 20) { pushUndo(s.toString()); redo.clear(); }
                lastChange = now;
            }
            @Override public void onTextChanged(CharSequence s, int st, int before, int count) { }
            @Override public void afterTextChanged(Editable s) { if (!loading && !applying && !dirty) { dirty = true; updateTitle(); } }
        });
        col.addView(edit, new LinearLayout.LayoutParams(-1, 0, 1));

        LinearLayout bottom = new LinearLayout(act); bottom.setGravity(Gravity.CENTER_VERTICAL); bottom.setPadding(dp(6), dp(4), dp(6), dp(4));
        bottom.addView(btn("Undo", v -> doUndo()), new LinearLayout.LayoutParams(0, -2, 1));
        bottom.addView(btn("Redo", v -> doRedo()), new LinearLayout.LayoutParams(0, -2, 1));
        bottom.addView(btn("Find / Replace", v -> showFind()), new LinearLayout.LayoutParams(0, -2, 1.6f));
        col.addView(bottom);

        updateTitle();
        loading = true;
        new Thread(() -> {
            try {
                ByteArrayOutputStream bo = new ByteArrayOutputStream();
                try (InputStream in = act.getContentResolver().openInputStream(uri)) {
                    if (in == null) throw new IOException("cannot open the file");
                    byte[] buf = new byte[65536]; int n;
                    while ((n = in.read(buf)) > 0) {
                        bo.write(buf, 0, n);
                        if (bo.size() > MAX_BYTES) throw new IOException("file is larger than " + (MAX_BYTES / 1024 / 1024) + " MB");
                    }
                }
                final byte[] data = bo.toByteArray();
                act.runOnUiThread(() -> { raw = data; decodeAndShow(null); });
            } catch (Throwable t) {
                final String m = t.getMessage();
                act.runOnUiThread(() -> { host.toast("Cannot open the text file: " + m); host.goHome(); });
            }
        }).start();
        return col;
    }

    // ------------------------------------------------------------ decoding / saving

    private void decodeAndShow(Charset forced) {
        byte[] b = raw; int off = 0; Charset cs = forced; bom = false;
        if (forced == null) {
            if (b.length >= 3 && (b[0] & 0xFF) == 0xEF && (b[1] & 0xFF) == 0xBB && (b[2] & 0xFF) == 0xBF) { cs = StandardCharsets.UTF_8; off = 3; bom = true; }
            else if (b.length >= 2 && (b[0] & 0xFF) == 0xFF && (b[1] & 0xFF) == 0xFE) { cs = StandardCharsets.UTF_16LE; off = 2; bom = true; }
            else if (b.length >= 2 && (b[0] & 0xFF) == 0xFE && (b[1] & 0xFF) == 0xFF) { cs = StandardCharsets.UTF_16BE; off = 2; bom = true; }
        }
        if (cs == null) {
            try {
                StandardCharsets.UTF_8.newDecoder().onMalformedInput(CodingErrorAction.REPORT).onUnmappableCharacter(CodingErrorAction.REPORT)
                        .decode(ByteBuffer.wrap(b, off, b.length - off));
                cs = StandardCharsets.UTF_8;
            } catch (CharacterCodingException e) {
                cs = Charset.forName("windows-1256");        // old Arabic / Persian text files; change it with More > Encoding
            }
        }
        String text = new String(b, off, b.length - off, cs);
        crlf = text.contains("\r\n");
        if (text.indexOf('\r') >= 0) text = text.replace("\r\n", "\n").replace('\r', '\n');
        charset = cs;
        loading = true;
        edit.setText(text);
        edit.setSelection(0);
        loading = false; dirty = false;
        undo.clear(); redo.clear();
        updateTitle();
    }

    void requestSave() {
        Intent i = new Intent(Intent.ACTION_CREATE_DOCUMENT);
        i.addCategory(Intent.CATEGORY_OPENABLE);
        i.setType("application/octet-stream");              // keeps the exact file name (no ".txt" appended)
        int dot = name.lastIndexOf('.');
        String base = dot > 0 ? name.substring(0, dot) : name;
        i.putExtra(Intent.EXTRA_TITLE, base + "_edited" + (ext.isEmpty() ? "" : "." + ext));
        act.startActivityForResult(i, REQ_SAVE);
    }

    void saveTo(final Uri dest) {
        if (dest == null) return;
        final String text = edit.getText().toString();
        final Charset cs = charset; final boolean useBom = bom, useCrlf = crlf;
        new Thread(() -> {
            try {
                String out = useCrlf ? text.replace("\n", "\r\n") : text;
                if (!cs.newEncoder().canEncode(out)) {
                    act.runOnUiThread(() -> host.toast("Some characters cannot be stored as " + cs.displayName() + ". Use More > Encoding > UTF-8."));
                    return;
                }
                try (OutputStream os = act.getContentResolver().openOutputStream(dest, "wt")) {
                    if (os == null) throw new IOException("cannot open the destination");
                    if (useBom) {
                        if (cs.equals(StandardCharsets.UTF_8)) os.write(new byte[]{(byte) 0xEF, (byte) 0xBB, (byte) 0xBF});
                        else if (cs.equals(StandardCharsets.UTF_16LE)) os.write(new byte[]{(byte) 0xFF, (byte) 0xFE});
                        else if (cs.equals(StandardCharsets.UTF_16BE)) os.write(new byte[]{(byte) 0xFE, (byte) 0xFF});
                    }
                    os.write(out.getBytes(cs));
                }
                act.runOnUiThread(() -> { dirty = false; updateTitle(); host.toast("Saved a copy. The original file was not changed."); });
            } catch (Throwable t) {
                final String m = t.getMessage();
                act.runOnUiThread(() -> host.toast("Cannot save: " + m));
            }
        }).start();
    }

    void requestClose() {
        if (!dirty) { host.goHome(); return; }
        new AlertDialog.Builder(act).setTitle("Unsaved changes").setMessage("Leave without saving a copy?")
                .setPositiveButton("Save a copy", (d, w) -> requestSave())
                .setNeutralButton("Discard", (d, w) -> host.goHome())
                .setNegativeButton("Keep editing", null).show();
    }

    // ------------------------------------------------------------ undo / redo

    private void pushUndo(String s) {
        undo.addFirst(s);
        int max = s.length() > 200000 ? 3 : 50;
        while (undo.size() > max) undo.removeLast();
    }

    private void doUndo() {
        if (undo.isEmpty()) { host.toast("Nothing to undo"); return; }
        String cur = edit.getText().toString(), prev = undo.removeFirst();
        redo.addFirst(cur);
        setTextKeepingPosition(prev);
    }

    private void doRedo() {
        if (redo.isEmpty()) { host.toast("Nothing to redo"); return; }
        String cur = edit.getText().toString(), next = redo.removeFirst();
        undo.addFirst(cur);
        setTextKeepingPosition(next);
    }

    private void setTextKeepingPosition(String t) {
        int sel = Math.min(edit.getSelectionStart(), t.length());
        applying = true; edit.setText(t); edit.setSelection(Math.max(0, sel)); applying = false;
        dirty = true; updateTitle();
    }

    // ------------------------------------------------------------ find / replace

    private int indexOf(String t, String q, int from, boolean matchCase) {
        for (int i = Math.max(0, from); i <= t.length() - q.length(); i++) if (t.regionMatches(!matchCase, i, q, 0, q.length())) return i;
        return -1;
    }

    private boolean findNext(String q, boolean matchCase) {
        if (q.isEmpty()) return false;
        String t = edit.getText().toString();
        int from = Math.max(0, edit.getSelectionEnd());
        int i = indexOf(t, q, from, matchCase);
        if (i < 0 && from > 0) i = indexOf(t, q, 0, matchCase);   // wrap around
        if (i < 0) return false;
        edit.requestFocus(); edit.setSelection(i, i + q.length());
        return true;
    }

    private void showFind() {
        LinearLayout v = new LinearLayout(act); v.setOrientation(LinearLayout.VERTICAL); v.setPadding(dp(20), dp(8), dp(20), 0);
        final EditText find = new EditText(act); find.setHint("Find"); find.setSingleLine(true);
        final EditText rep = new EditText(act); rep.setHint("Replace with"); rep.setSingleLine(true);
        final CheckBox cs = new CheckBox(act); cs.setText("Match case");
        v.addView(find); v.addView(rep); v.addView(cs);
        LinearLayout row = new LinearLayout(act);
        row.addView(btn("Find next", x -> { if (!findNext(find.getText().toString(), cs.isChecked())) host.toast("Not found"); }), new LinearLayout.LayoutParams(0, -2, 1));
        row.addView(btn("Replace", x -> {
            String q = find.getText().toString(), r = rep.getText().toString();
            int a = edit.getSelectionStart(), b = edit.getSelectionEnd();
            if (!q.isEmpty() && a >= 0 && b == a + q.length()) {
                String sel = edit.getText().subSequence(a, b).toString();
                if (cs.isChecked() ? sel.equals(q) : sel.equalsIgnoreCase(q)) {
                    edit.getText().replace(a, b, r);
                    edit.setSelection(a + r.length());
                }
            }
            if (!findNext(q, cs.isChecked())) host.toast("Not found");
        }), new LinearLayout.LayoutParams(0, -2, 1));
        row.addView(btn("Replace all", x -> {
            String q = find.getText().toString(), r = rep.getText().toString();
            if (q.isEmpty()) return;
            String t = edit.getText().toString();
            StringBuilder sb = new StringBuilder(); int pos = 0, count = 0;
            while (true) {
                int i = indexOf(t, q, pos, cs.isChecked());
                if (i < 0) break;
                sb.append(t, pos, i).append(r); pos = i + q.length(); count++;
            }
            if (count == 0) { host.toast("Not found"); return; }
            sb.append(t, pos, t.length());
            pushUndo(t); redo.clear(); lastChange = 0;
            setTextKeepingPosition(sb.toString());
            host.toast(count + " replaced");
        }), new LinearLayout.LayoutParams(0, -2, 1.3f));
        v.addView(row);
        new AlertDialog.Builder(act).setTitle("Find / Replace").setView(v).setNegativeButton("Close", null).show();
    }

    // ------------------------------------------------------------ menu

    private void showMenu() {
        final String[] items = {
                "Go to line…", "Word wrap: " + (wrap ? "ON" : "OFF"), "Monospace font: " + (mono ? "ON" : "OFF"),
                "Text size +", "Text size −", "Encoding: " + charset.displayName() + "…", "Line endings: " + (crlf ? "CRLF (Windows)" : "LF (Linux)"), "Select all"};
        new AlertDialog.Builder(act).setTitle("Editor").setItems(items, (d, i) -> {
            if (i == 0) goToLine();
            else if (i == 1) { wrap = !wrap; edit.setHorizontallyScrolling(!wrap); }
            else if (i == 2) { mono = !mono; applyFont(); }
            else if (i == 3) { sizeSp = Math.min(32f, sizeSp + 2f); applyFont(); }
            else if (i == 4) { sizeSp = Math.max(9f, sizeSp - 2f); applyFont(); }
            else if (i == 5) chooseEncoding();
            else if (i == 6) { crlf = !crlf; dirty = true; updateTitle(); }
            else { edit.requestFocus(); edit.selectAll(); }
        }).show();
    }

    private void chooseEncoding() {
        final String[] labels = {"UTF-8", "UTF-16 LE", "UTF-16 BE", "Windows-1256 (Arabic / Persian)", "Windows-1252 (Western)", "ISO-8859-1"};
        final String[] names = {"UTF-8", "UTF-16LE", "UTF-16BE", "windows-1256", "windows-1252", "ISO-8859-1"};
        new AlertDialog.Builder(act).setTitle(dirty ? "Encoding used when saving" : "Read the file as").setItems(labels, (d, i) -> {
            Charset cs = Charset.forName(names[i]);
            if (dirty) { charset = cs; bom = false; host.toast("The copy will be saved as " + labels[i]); }
            else decodeAndShow(cs);
        }).show();
    }

    private void goToLine() {
        final EditText in = new EditText(act); in.setInputType(InputType.TYPE_CLASS_NUMBER); in.setHint("Line number");
        LinearLayout w = new LinearLayout(act); w.setPadding(dp(20), dp(8), dp(20), 0); w.addView(in, new LinearLayout.LayoutParams(-1, -2));
        new AlertDialog.Builder(act).setTitle("Go to line").setView(w).setNegativeButton("Cancel", null).setPositiveButton("Go", (d, x) -> {
            try {
                int line = Integer.parseInt(in.getText().toString().trim());
                String t = edit.getText().toString();
                int pos = 0, cur = 1;
                while (cur < line) { int nl = t.indexOf('\n', pos); if (nl < 0) break; pos = nl + 1; cur++; }
                edit.requestFocus(); edit.setSelection(Math.min(pos, t.length()));
                if (cur < line) host.toast("The file has only " + cur + " lines");
            } catch (NumberFormatException e) { host.toast("Enter a number"); }
        }).show();
    }

    // ------------------------------------------------------------ helpers

    private void applyFont() {
        edit.setTypeface(mono ? Typeface.MONOSPACE : Typeface.DEFAULT);
        edit.setTextSize(TypedValue.COMPLEX_UNIT_SP, sizeSp);
    }

    private void updateTitle() { if (titleView != null) titleView.setText(name + (dirty ? "  *" : "") + "   [" + charset.displayName().toLowerCase(Locale.ROOT) + "]"); }

    private Button btn(String label, View.OnClickListener l) {
        Button b = new Button(act); b.setText(label); b.setAllCaps(false);
        b.setMinWidth(0); b.setMinimumWidth(0); b.setMinHeight(0); b.setMinimumHeight(dp(44)); b.setPadding(dp(12), 0, dp(12), 0); b.setOnClickListener(l);
        return b;
    }

    private int dp(int v) { return Math.round(v * act.getResources().getDisplayMetrics().density); }
}
