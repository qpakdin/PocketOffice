package com.example.pocketoffice;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;

/** Which files are Office documents (shown with the real layout engine) and which are plain text (opened in the text editor). */
final class Formats {
    private static final Set<String> OFFICE = new HashSet<>(Arrays.asList(
            // Word-like
            "docx", "docm", "dotx", "dotm", "doc", "dot", "odt", "ott", "fodt", "rtf", "sxw", "stw",
            // Excel-like
            "xlsx", "xlsm", "xltx", "xltm", "xlsb", "xls", "xlt", "ods", "ots", "fods", "sxc",
            // PowerPoint-like
            "pptx", "pptm", "ppsx", "ppsm", "potx", "potm", "ppt", "pps", "pot", "odp", "otp", "fodp", "sxi"));

    private static final Set<String> QUICK_WORD = new HashSet<>(Arrays.asList("docx", "docm", "dotx", "dotm"));
    private static final Set<String> QUICK_SHEET = new HashSet<>(Arrays.asList("xlsx", "xlsm", "xltx", "xltm"));
    private static final Set<String> QUICK_SLIDES = new HashSet<>(Arrays.asList("pptx", "pptm", "ppsx", "ppsm", "potx", "potm"));

    /** lower-case extension without the dot, or "" if the name has none / odd characters. */
    static String extOf(String name) {
        if (name == null) return "";
        int dot = name.lastIndexOf('.');
        if (dot < 0 || dot == name.length() - 1) return "";
        String e = name.substring(dot + 1).toLowerCase(Locale.ROOT);
        return e.matches("[a-z0-9]{1,8}") ? e : "";
    }

    static boolean isOffice(String ext) { return OFFICE.contains(ext); }

    /** The built-in offline preview only understands the OOXML families. Returns "docx" / "xlsx" / "pptx" or null. */
    static String quickType(String ext) {
        if (QUICK_WORD.contains(ext)) return "docx";
        if (QUICK_SHEET.contains(ext)) return "xlsx";
        if (QUICK_SLIDES.contains(ext)) return "pptx";
        return null;
    }

    /** Content check used for files whose extension is unknown: is this readable text (source code, config, notes ...)? */
    static boolean looksLikeText(byte[] b, int n) {
        if (n == 0) return true;
        if (n >= 5 && b[0] == '%' && b[1] == 'P' && b[2] == 'D' && b[3] == 'F') return false;                    // PDF
        if (n >= 4 && b[0] == 'P' && b[1] == 'K' && b[2] == 3 && b[3] == 4) return false;                         // zip / docx without a known name
        if (n >= 4 && (b[0] & 0xFF) == 0xD0 && (b[1] & 0xFF) == 0xCF && (b[2] & 0xFF) == 0x11) return false;       // old Office (OLE2)
        if (n >= 3 && (b[0] & 0xFF) == 0xEF && (b[1] & 0xFF) == 0xBB && (b[2] & 0xFF) == 0xBF) return true;       // UTF-8 BOM
        if (n >= 2 && (((b[0] & 0xFF) == 0xFF && (b[1] & 0xFF) == 0xFE) || ((b[0] & 0xFF) == 0xFE && (b[1] & 0xFF) == 0xFF))) return true; // UTF-16 BOM
        int ctrl = 0;
        for (int i = 0; i < n; i++) {
            int v = b[i] & 0xFF;
            if (v == 0) return false;
            if (v < 9 || (v > 13 && v < 32 && v != 27)) ctrl++;
        }
        return ctrl * 100 < n * 2;      // fewer than 2 % control characters
    }
}
