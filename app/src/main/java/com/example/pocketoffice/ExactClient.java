package com.example.pocketoffice;

import android.content.Context;
import android.net.Uri;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLDecoder;

/** Talks to the Pocket Office converter (LibreOffice) running in Termux / on a PC. */
final class ExactClient {
    /** Fonts the last converted file uses that the converter machine does not have (empty = none / unknown). */
    static volatile String lastMissingFonts = "";

    static String ping(String server) {
        try {
            HttpURLConnection h = (HttpURLConnection) new URL(server + "/ping").openConnection();
            h.setConnectTimeout(1500); h.setReadTimeout(3000);
            if (h.getResponseCode() != 200) return null;
            String body = readText(h.getInputStream(), 600);
            return body.startsWith("pocket-office") ? body : null;
        } catch (Exception e) { return null; }
    }

    static File convert(Context c, String server, Uri doc, String ext, File out) throws Exception {
        File tmp = File.createTempFile("up_", "." + ext, c.getCacheDir());
        try {
            try (InputStream in = c.getContentResolver().openInputStream(doc); OutputStream o = new FileOutputStream(tmp)) { copy(in, o); }
            String missing;
            try (OutputStream sink = new FileOutputStream(out)) { missing = post(server + "/convert?ext=" + ext, tmp, sink); }
            lastMissingFonts = missing == null ? "" : URLDecoder.decode(missing, "UTF-8");
            if (out.length() < 100) throw new Exception("Converter returned an empty file");
            return out;
        } catch (Exception e) { out.delete(); throw e; }
        finally { tmp.delete(); }
    }

    /** Text of an OOXML file in logical reading order (right for Persian / RTL). Throws if the server cannot do it for this type. */
    static byte[] docText(Context c, String server, Uri doc, String ext) throws Exception {
        File tmp = File.createTempFile("tx_", "." + ext, c.getCacheDir());
        try {
            try (InputStream in = c.getContentResolver().openInputStream(doc); OutputStream o = new FileOutputStream(tmp)) { copy(in, o); }
            ByteArrayOutputStream bo = new ByteArrayOutputStream();
            post(server + "/text?ext=" + ext, tmp, bo);
            return bo.toByteArray();
        } finally { tmp.delete(); }
    }

    static byte[] pdfText(String server, File pdf) throws Exception {
        ByteArrayOutputStream bo = new ByteArrayOutputStream();
        post(server + "/pdftext", pdf, bo);
        return bo.toByteArray();
    }

    /** Returns the X-PO-Missing-Fonts header (or null). */
    private static String post(String url, File body, OutputStream sink) throws Exception {
        HttpURLConnection h = (HttpURLConnection) new URL(url).openConnection();
        h.setRequestMethod("POST"); h.setDoOutput(true);
        h.setConnectTimeout(5000); h.setReadTimeout(900000);
        h.setFixedLengthStreamingMode(body.length());
        h.setRequestProperty("Content-Type", "application/octet-stream");
        try (InputStream in = new FileInputStream(body); OutputStream out = h.getOutputStream()) { copy(in, out); }
        int code = h.getResponseCode();
        if (code != 200) {
            String msg = readText(h.getErrorStream(), 1500);
            throw new Exception("Converter error " + code + (msg.isEmpty() ? "" : ": " + msg));
        }
        String missing = h.getHeaderField("X-PO-Missing-Fonts");
        try (InputStream in = h.getInputStream()) { copy(in, sink); }
        h.disconnect();
        return missing;
    }

    private static String readText(InputStream in, int max) throws Exception {
        if (in == null) return "";
        ByteArrayOutputStream bo = new ByteArrayOutputStream();
        byte[] buf = new byte[1024]; int n;
        while ((n = in.read(buf)) > 0 && bo.size() < max) bo.write(buf, 0, n);
        in.close();
        return bo.toString("UTF-8").trim();
    }

    private static void copy(InputStream in, OutputStream out) throws Exception {
        byte[] buf = new byte[65536]; int n;
        while ((n = in.read(buf)) > 0) out.write(buf, 0, n);
        out.flush();
    }
}
