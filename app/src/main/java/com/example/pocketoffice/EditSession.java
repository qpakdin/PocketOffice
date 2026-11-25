package com.example.pocketoffice;

import android.content.Context;
import android.net.Uri;

import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;

/**
 * Hands a COPY of a document to the converter as an editing session. The editor itself is the converter's web panel
 * (opened in the browser); this class only uploads the file. The original file is only read, never changed.
 */
final class EditSession {
    /** The converter refused (or could not do) something; the message is meant to be shown to the user. */
    static final class Refused extends Exception { Refused(String m) { super(m); } }

    private final String server;
    String sid;

    EditSession(String server) { this.server = server; }

    /** Uploads the document and starts a session on the converter. The preflight report is shown by the panel and logged by the server. */
    JSONObject open(Context c, Uri doc, String ext, String name) throws Exception {
        File tmp = File.createTempFile("ed_", "." + ext, c.getCacheDir());
        try {
            try (InputStream in = c.getContentResolver().openInputStream(doc); OutputStream o = new FileOutputStream(tmp)) {
                if (in == null) throw new Refused("The file could not be read.");
                copy(in, o);
            }
            HttpURLConnection h = (HttpURLConnection) new URL(server + "/edit/open?ext=" + ext + "&name=" + URLEncoder.encode(name, "UTF-8")).openConnection();
            h.setRequestMethod("POST");
            h.setConnectTimeout(5000); h.setReadTimeout(300000);
            h.setDoOutput(true);
            h.setFixedLengthStreamingMode(tmp.length());
            h.setRequestProperty("Content-Type", "application/octet-stream");
            try (InputStream in = new FileInputStream(tmp); OutputStream out = h.getOutputStream()) { copy(in, out); }
            int code = h.getResponseCode();
            InputStream rs = code < 400 ? h.getInputStream() : h.getErrorStream();
            ByteArrayOutputStream bo = new ByteArrayOutputStream();
            if (rs != null) { try { copy(rs, bo); } finally { rs.close(); } }
            h.disconnect();
            String text = bo.toString("UTF-8");
            JSONObject o;
            try { o = new JSONObject(text); }
            catch (Exception e) { throw new Refused("The converter answered with an unexpected reply (HTTP " + code + "). Is the server v4.0 or newer? " + (text.length() > 200 ? text.substring(0, 200) : text)); }
            if (!o.optBoolean("ok", false)) throw new Refused(o.optString("error", "The converter refused the request."));
            sid = o.getString("sid");
            return o;
        } finally { tmp.delete(); }
    }

    private static void copy(InputStream in, OutputStream out) throws Exception {
        byte[] buf = new byte[65536]; int n;
        while ((n = in.read(buf)) > 0) out.write(buf, 0, n);
        out.flush();
    }
}
