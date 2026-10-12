package com.excel2pdf.studio;

import android.app.Activity;
import android.content.ContentResolver;
import android.content.ContentValues;
import android.content.Context;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Environment;
import android.os.Handler;
import android.os.Looper;
import android.provider.MediaStore;
import android.util.Base64;
import android.webkit.JavascriptInterface;
import android.widget.Toast;

import androidx.core.content.FileProvider;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

public class AndroidBridge {

    private final Activity activity;
    private final Handler mainHandler;

    public AndroidBridge(Activity activity) {
        this.activity = activity;
        this.mainHandler = new Handler(Looper.getMainLooper());
    }

    @JavascriptInterface
    public void showToast(final String message) {
        mainHandler.post(() -> Toast.makeText(activity, message, Toast.LENGTH_SHORT).show());
    }

    @JavascriptInterface
    public void savePdf(final String base64Data, final String filename) {
        try {
            byte[] pdfBytes = Base64.decode(base64Data, Base64.DEFAULT);
            String safeName = (filename != null && !filename.isEmpty()) ? filename : "dokumen.pdf";
            if (!safeName.toLowerCase().endsWith(".pdf")) {
                safeName += ".pdf";
            }

            boolean saved = false;

            // Android 10+ (API 29+): Use MediaStore Downloads collection
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                ContentResolver resolver = activity.getContentResolver();
                ContentValues contentValues = new ContentValues();
                contentValues.put(MediaStore.MediaColumns.DISPLAY_NAME, safeName);
                contentValues.put(MediaStore.MediaColumns.MIME_TYPE, "application/pdf");
                contentValues.put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS + "/Excel2PDF");

                Uri uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, contentValues);
                if (uri != null) {
                    try (OutputStream os = resolver.openOutputStream(uri)) {
                        if (os != null) {
                            os.write(pdfBytes);
                            saved = true;
                        }
                    }
                }
            } else {
                // Older Android: Direct file write to external Downloads directory
                File downloadsDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS);
                File targetDir = new File(downloadsDir, "Excel2PDF");
                if (!targetDir.exists()) {
                    targetDir.mkdirs();
                }
                File targetFile = new File(targetDir, safeName);
                try (FileOutputStream fos = new FileOutputStream(targetFile)) {
                    fos.write(pdfBytes);
                    saved = true;
                }
            }

            final String alertMsg = saved 
                ? "PDF berhasil disimpan di folder Download: " + safeName 
                : "Gagal menyimpan PDF.";
            
            mainHandler.post(() -> Toast.makeText(activity, alertMsg, Toast.LENGTH_LONG).show());

        } catch (Exception e) {
            final String errMsg = "Gagal menyimpan PDF: " + e.getMessage();
            mainHandler.post(() -> Toast.makeText(activity, errMsg, Toast.LENGTH_LONG).show());
        }
    }

    @JavascriptInterface
    public void sharePdf(final String base64Data, final String filename) {
        try {
            byte[] pdfBytes = Base64.decode(base64Data, Base64.DEFAULT);
            String safeName = (filename != null && !filename.isEmpty()) ? filename : "dokumen.pdf";
            if (!safeName.toLowerCase().endsWith(".pdf")) {
                safeName += ".pdf";
            }

            File cacheDir = new File(activity.getCacheDir(), "shared_pdfs");
            if (!cacheDir.exists()) {
                cacheDir.mkdirs();
            }

            File targetFile = new File(cacheDir, safeName);
            try (FileOutputStream fos = new FileOutputStream(targetFile)) {
                fos.write(pdfBytes);
            }

            Uri contentUri = FileProvider.getUriForFile(
                activity,
                activity.getPackageName() + ".fileprovider",
                targetFile
            );

            Intent shareIntent = new Intent(Intent.ACTION_SEND);
            shareIntent.setType("application/pdf");
            shareIntent.putExtra(Intent.EXTRA_STREAM, contentUri);
            shareIntent.putExtra(Intent.EXTRA_SUBJECT, safeName);
            shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);

            Intent chooser = Intent.createChooser(shareIntent, "Bagikan PDF melalui...");
            chooser.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);

            mainHandler.post(() -> activity.startActivity(chooser));

        } catch (Exception e) {
            final String errMsg = "Gagal membagikan PDF: " + e.getMessage();
            mainHandler.post(() -> Toast.makeText(activity, errMsg, Toast.LENGTH_SHORT).show());
        }
    }
}
