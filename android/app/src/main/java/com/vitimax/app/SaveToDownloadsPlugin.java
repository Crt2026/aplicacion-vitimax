package com.vitimax.app;

import android.Manifest;
import android.content.ContentResolver;
import android.content.ContentValues;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Environment;
import android.os.Handler;
import android.os.Looper;
import android.provider.MediaStore;
import android.util.Base64;
import android.widget.Toast;

import androidx.core.content.FileProvider;

import com.getcapacitor.JSObject;
import com.getcapacitor.PermissionState;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;
import com.getcapacitor.annotation.Permission;
import com.getcapacitor.annotation.PermissionCallback;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Guarda archivos reales en la carpeta pública Descargas de Android.
 *
 * - Android 10 o superior: usa MediaStore, no necesita permisos.
 * - Android 9 o inferior: pide WRITE_EXTERNAL_STORAGE una sola vez.
 *
 * Soporta cualquier tipo de archivo (Excel, PDF, Word, PowerPoint, HTML,
 * código, imágenes, texto, etc.) porque recibe el nombre y el MIME type.
 */
@CapacitorPlugin(
    name = "SaveToDownloads",
    permissions = {
        @Permission(alias = "storage", strings = { Manifest.permission.WRITE_EXTERNAL_STORAGE })
    }
)
public class SaveToDownloadsPlugin extends Plugin {

    /** Permite pedir el permiso de almacenamiento antes de generar archivos. */
    @PluginMethod
    public void ensurePermission(PluginCall call) {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q
            || getPermissionState("storage") == PermissionState.GRANTED) {
            JSObject result = new JSObject();
            result.put("granted", true);
            result.put("needsPermission", false);
            call.resolve(result);
            return;
        }
        requestPermissionForAlias("storage", call, "storagePermissionCallback");
    }

    @PluginMethod
    public void save(PluginCall call) {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.Q
            && getPermissionState("storage") != PermissionState.GRANTED) {
            requestPermissionForAlias("storage", call, "storagePermissionCallback");
            return;
        }
        saveInternal(call);
    }

    /** Abre el archivo guardado con la aplicación que el teléfono tenga para ese tipo. */
    @PluginMethod
    public void openFile(PluginCall call) {
        String uriString = call.getString("uri");
        String mimeType = call.getString("mimeType", "application/octet-stream");
        if (uriString == null || uriString.isEmpty()) {
            call.reject("Falta la ubicación del archivo. Pulsa Descargar primero.");
            return;
        }
        try {
            Uri uri = Uri.parse(uriString);
            Intent intent = new Intent(Intent.ACTION_VIEW);
            intent.setDataAndType(uri, mimeType);
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
            intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
            getContext().startActivity(intent);
            call.resolve();
        } catch (Exception error) {
            call.reject("No hay una aplicación para abrir este archivo. Búscalo en la carpeta Descargas.", error);
        }
    }

    /** Devuelve datos técnicos para comprobar por qué falla el guardado en un teléfono. */
    @PluginMethod
    public void diagnose(PluginCall call) {
        JSObject result = new JSObject();
        result.put("sdk", Build.VERSION.SDK_INT);
        result.put("android", Build.VERSION.RELEASE);
        result.put("needsPermission", Build.VERSION.SDK_INT < Build.VERSION_CODES.Q);
        result.put("permission", getPermissionState("storage").toString());
        result.put("mediaStore", Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q);
        call.resolve(result);
    }

    @PermissionCallback
    private void storagePermissionCallback(PluginCall call) {
        boolean granted = getPermissionState("storage") == PermissionState.GRANTED;
        if ("ensurePermission".equals(call.getMethodName())) {
            JSObject result = new JSObject();
            result.put("granted", granted);
            result.put("needsPermission", true);
            if (!granted) {
                result.put("message", "Permiso de almacenamiento denegado. Actívalo en Ajustes > Aplicaciones > VITIMAX > Permisos.");
            }
            call.resolve(result);
            return;
        }
        if (granted) {
            saveInternal(call);
        } else {
            call.reject("Permiso de almacenamiento denegado. Actívalo en Ajustes > Aplicaciones > VITIMAX > Permisos y vuelve a intentarlo.");
        }
    }

    private void saveInternal(PluginCall call) {
        String fileName = sanitize(call.getString("fileName", "archivo-vitimax"));
        String mimeType = call.getString("mimeType", "application/octet-stream");
        String data = call.getString("data");
        if (data == null || data.isEmpty()) {
            call.reject("Faltan los datos del archivo.");
            return;
        }

        try {
            byte[] bytes = Base64.decode(data, Base64.DEFAULT);
            JSObject result = new JSObject();
            result.put("size", bytes.length);
            result.put("mimeType", mimeType);

            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                ContentResolver resolver = getContext().getContentResolver();
                ContentValues values = new ContentValues();
                values.put(MediaStore.MediaColumns.DISPLAY_NAME, fileName);
                values.put(MediaStore.MediaColumns.MIME_TYPE, mimeType);
                values.put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS);
                values.put(MediaStore.MediaColumns.IS_PENDING, 1);

                Uri uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
                if (uri == null) {
                    call.reject("No se pudo crear el archivo en Descargas. Revisa el espacio libre del teléfono.");
                    return;
                }
                try (OutputStream out = resolver.openOutputStream(uri)) {
                    if (out == null) {
                        resolver.delete(uri, null, null);
                        call.reject("No se pudo abrir el archivo en Descargas.");
                        return;
                    }
                    out.write(bytes);
                    out.flush();
                }
                ContentValues done = new ContentValues();
                done.put(MediaStore.MediaColumns.IS_PENDING, 0);
                resolver.update(uri, done, null, null);

                result.put("uri", uri.toString());
                result.put("path", "Descargas/" + fileName);
                call.resolve(result);
                toast("Guardado en Descargas: " + fileName);
                return;
            }

            File dir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS);
            if (!dir.exists() && !dir.mkdirs()) {
                call.reject("No se pudo crear la carpeta Descargas.");
                return;
            }
            File file = uniqueFile(dir, fileName);
            try (FileOutputStream out = new FileOutputStream(file)) {
                out.write(bytes);
                out.flush();
            }
            result.put("uri", FileProvider.getUriForFile(
                getContext(),
                getContext().getPackageName() + ".fileprovider",
                file
            ).toString());
            result.put("path", "Descargas/" + file.getName());
            result.put("absolutePath", file.getAbsolutePath());
            call.resolve(result);
            toast("Guardado en Descargas: " + file.getName());
        } catch (Exception error) {
            call.reject("No se pudo guardar el archivo: " + error.getMessage(), error);
        }
    }

    private void toast(String message) {
        new Handler(Looper.getMainLooper()).post(() ->
            Toast.makeText(getContext().getApplicationContext(), message, Toast.LENGTH_LONG).show()
        );
    }

    private static String sanitize(String name) {
        String clean = name == null ? "" : name.replaceAll("[\\\\/:*?\"<>|]+", "-").trim();
        return clean.isEmpty() ? "archivo-vitimax" : clean;
    }

    private static File uniqueFile(File dir, String fileName) {
        File file = new File(dir, fileName);
        if (!file.exists()) return file;
        int dot = fileName.lastIndexOf('.');
        String base = dot > 0 ? fileName.substring(0, dot) : fileName;
        String ext = dot > 0 ? fileName.substring(dot) : "";
        int index = 1;
        while (file.exists()) {
            file = new File(dir, base + " (" + index + ")" + ext);
            index++;
        }
        return file;
    }
}
