package com.vitimax.app;

import android.Manifest;
import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Build;
import android.provider.Settings;

import androidx.core.app.NotificationCompat;
import androidx.core.content.ContextCompat;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;
import com.getcapacitor.annotation.Permission;
import com.getcapacitor.annotation.PermissionCallback;

import java.util.ArrayList;
import java.util.List;

/**
 * Pide y comprueba todos los permisos que usa VITIMAX:
 * notificaciones, cámara, archivos/fotos y almacenamiento.
 *
 * También permite mostrar una notificación real del sistema y abrir los
 * ajustes de la aplicación cuando un permiso fue denegado.
 */
@CapacitorPlugin(
    name = "DevicePermissions",
    permissions = {
        @Permission(alias = "notifications", strings = { Manifest.permission.POST_NOTIFICATIONS }),
        @Permission(alias = "camera", strings = { Manifest.permission.CAMERA }),
        @Permission(alias = "storageRead", strings = { Manifest.permission.READ_EXTERNAL_STORAGE }),
        @Permission(alias = "storageWrite", strings = { Manifest.permission.WRITE_EXTERNAL_STORAGE }),
        @Permission(alias = "media", strings = { Manifest.permission.READ_MEDIA_IMAGES, Manifest.permission.READ_MEDIA_VIDEO, Manifest.permission.READ_MEDIA_AUDIO })
    }
)
public class DevicePermissionsPlugin extends Plugin {

    private static final String CHANNEL_ID = "vitimax_avisos";

    /** Devuelve el estado actual de cada permiso, sin abrir diálogos. */
    @PluginMethod
    public void status(PluginCall call) {
        call.resolve(buildStatus());
    }

    /** Pide de una sola vez todos los permisos que falten. */
    @PluginMethod
    public void requestAll(PluginCall call) {
        List<String> aliases = new ArrayList<>();

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU
            && !isGranted(Manifest.permission.POST_NOTIFICATIONS)) {
            aliases.add("notifications");
        }
        if (!isGranted(Manifest.permission.CAMERA)) {
            aliases.add("camera");
        }
        if (Build.VERSION.SDK_INT <= Build.VERSION_CODES.S_V2
            && !isGranted(Manifest.permission.READ_EXTERNAL_STORAGE)) {
            aliases.add("storageRead");
        }
        if (Build.VERSION.SDK_INT <= Build.VERSION_CODES.P
            && !isGranted(Manifest.permission.WRITE_EXTERNAL_STORAGE)) {
            aliases.add("storageWrite");
        }
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU
            && !isGranted(Manifest.permission.READ_MEDIA_IMAGES)) {
            aliases.add("media");
        }

        if (aliases.isEmpty()) {
            call.resolve(buildStatus());
            return;
        }
        requestPermissionForAliases(aliases.toArray(new String[0]), call, "permissionsCallback");
    }

    @PermissionCallback
    private void permissionsCallback(PluginCall call) {
        call.resolve(buildStatus());
    }

    /** Muestra una notificación real en la barra de estado del teléfono. */
    @PluginMethod
    public void notify(PluginCall call) {
        String title = call.getString("title", "VITIMAX");
        String body = call.getString("body", "Notificación de prueba de VITIMAX.");

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU
            && !isGranted(Manifest.permission.POST_NOTIFICATIONS)) {
            call.reject("Permiso de notificaciones denegado. Actívalo en Ajustes > Aplicaciones > VITIMAX > Notificaciones.");
            return;
        }

        try {
            NotificationManager manager = (NotificationManager) getContext().getSystemService(Context.NOTIFICATION_SERVICE);
            if (manager == null) {
                call.reject("El teléfono no tiene servicio de notificaciones.");
                return;
            }
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
                NotificationChannel channel = new NotificationChannel(
                    CHANNEL_ID,
                    "Avisos de VITIMAX",
                    NotificationManager.IMPORTANCE_DEFAULT
                );
                channel.setDescription("Alertas de stock, pedidos y actividad de VITIMAX");
                manager.createNotificationChannel(channel);
            }

            NotificationCompat.Builder builder = new NotificationCompat.Builder(getContext(), CHANNEL_ID)
                .setSmallIcon(R.drawable.ic_stat_vitimax)
                .setContentTitle(title)
                .setContentText(body)
                .setStyle(new NotificationCompat.BigTextStyle().bigText(body))
                .setPriority(NotificationCompat.PRIORITY_DEFAULT)
                .setAutoCancel(true);

            manager.notify((int) (System.currentTimeMillis() % 100000), builder.build());
            call.resolve();
        } catch (Exception error) {
            call.reject("No se pudo mostrar la notificación: " + error.getMessage(), error);
        }
    }

    /** Abre los ajustes de notificaciones de la app para activarlas manualmente. */
    @PluginMethod
    public void openNotificationSettings(PluginCall call) {
        try {
            Intent intent;
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
                intent = new Intent(Settings.ACTION_APP_NOTIFICATION_SETTINGS);
                intent.putExtra(Settings.EXTRA_APP_PACKAGE, getContext().getPackageName());
            } else {
                intent = new Intent(Settings.ACTION_APPLICATION_DETAILS_SETTINGS);
                intent.setData(Uri.parse("package:" + getContext().getPackageName()));
            }
            intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
            getContext().startActivity(intent);
            call.resolve();
        } catch (Exception error) {
            call.reject("No se pudieron abrir los ajustes de notificaciones.", error);
        }
    }

    /** Abre la pantalla de permisos de VITIMAX en los ajustes del teléfono. */
    @PluginMethod
    public void openAppSettings(PluginCall call) {
        try {
            Intent intent = new Intent(Settings.ACTION_APPLICATION_DETAILS_SETTINGS);
            intent.setData(Uri.parse("package:" + getContext().getPackageName()));
            intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
            getContext().startActivity(intent);
            call.resolve();
        } catch (Exception error) {
            call.reject("No se pudieron abrir los ajustes de la aplicación.", error);
        }
    }

    private JSObject buildStatus() {
        boolean notifications = Build.VERSION.SDK_INT < Build.VERSION_CODES.TIRAMISU
            || isGranted(Manifest.permission.POST_NOTIFICATIONS);
        boolean camera = isGranted(Manifest.permission.CAMERA);
        boolean storage = Build.VERSION.SDK_INT > Build.VERSION_CODES.S_V2
            || (isGranted(Manifest.permission.READ_EXTERNAL_STORAGE)
                && (Build.VERSION.SDK_INT > Build.VERSION_CODES.P || isGranted(Manifest.permission.WRITE_EXTERNAL_STORAGE)));
        boolean media = Build.VERSION.SDK_INT < Build.VERSION_CODES.TIRAMISU
            || isGranted(Manifest.permission.READ_MEDIA_IMAGES);

        JSObject result = new JSObject();
        result.put("sdk", Build.VERSION.SDK_INT);
        result.put("android", Build.VERSION.RELEASE);
        result.put("notifications", notifications);
        result.put("camera", camera);
        result.put("storage", storage);
        result.put("media", media);
        result.put("allGranted", notifications && camera && storage && media);
        return result;
    }

    private boolean isGranted(String permission) {
        return ContextCompat.checkSelfPermission(getContext(), permission) == PackageManager.PERMISSION_GRANTED;
    }
}
