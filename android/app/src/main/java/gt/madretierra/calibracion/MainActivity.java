package gt.madretierra.calibracion;

import android.app.Activity;
import android.content.ClipData;
import android.content.Intent;
import android.graphics.Color;
import android.net.Uri;
import android.os.Bundle;
import android.provider.MediaStore;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;

import androidx.core.content.FileProvider;

import java.io.File;

public class MainActivity extends Activity {

    private static final String URL_APP = "https://cobocuburjosue-glitch.github.io/CALIBRACION-DRONES/";
    private static final int REQ_FOTO = 41;

    private WebView web;
    private ValueCallback<Uri[]> callbackFoto;
    private Uri uriCamara;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        web = new WebView(this);
        web.setBackgroundColor(Color.parseColor("#141618"));
        setContentView(web);

        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setDatabaseEnabled(true);
        s.setLoadWithOverviewMode(true);
        s.setUseWideViewPort(true);
        s.setAllowFileAccess(false);
        s.setMediaPlaybackRequiresUserGesture(true);
        s.setUserAgentString(s.getUserAgentString() + " CalibracionMTApp/" + BuildConfigVersion());

        web.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest req) {
                Uri u = req.getUrl();
                String host = u.getHost() == null ? "" : u.getHost();
                if (host.endsWith("github.io")) return false;          // la app se queda adentro
                try { startActivity(new Intent(Intent.ACTION_VIEW, u)); } catch (Exception ignored) {}
                return true;                                           // fotos de Drive, etc. -> navegador
            }

            @Override
            public void onReceivedError(WebView view, WebResourceRequest req, WebResourceError err) {
                if (req.isForMainFrame()) {
                    view.loadDataWithBaseURL(null,
                        "<html><body style='background:#141618;color:#fff;font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;text-align:center'>"
                        + "<div><h2>Sin conexión</h2><p style='color:#aeb4ba'>Revise su internet e intente de nuevo.</p>"
                        + "<button onclick=\"location.href='" + URL_APP + "'\" style='background:#3f8f3a;color:#fff;border:0;border-radius:10px;padding:14px 26px;font-size:16px'>Reintentar</button></div></body></html>",
                        "text/html", "utf-8", null);
                }
            }
        });

        web.setWebChromeClient(new WebChromeClient() {
            @Override
            public boolean onShowFileChooser(WebView view, ValueCallback<Uri[]> cb, FileChooserParams params) {
                if (callbackFoto != null) callbackFoto.onReceiveValue(null);
                callbackFoto = cb;
                abrirSelectorFoto();
                return true;
            }
        });

        if (savedInstanceState != null) web.restoreState(savedInstanceState);
        else web.loadUrl(URL_APP);
    }

    private String BuildConfigVersion() {
        try { return getPackageManager().getPackageInfo(getPackageName(), 0).versionName; }
        catch (Exception e) { return "1"; }
    }

    private void abrirSelectorFoto() {
        Intent[] extras = new Intent[0];
        uriCamara = null;
        try {
            File dir = new File(getCacheDir(), "fotos");
            if (!dir.exists()) dir.mkdirs();
            File f = File.createTempFile("calibracion_", ".jpg", dir);
            uriCamara = FileProvider.getUriForFile(this, getPackageName() + ".fileprovider", f);
            Intent camara = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
            camara.putExtra(MediaStore.EXTRA_OUTPUT, uriCamara);
            camara.setClipData(ClipData.newRawUri("", uriCamara));
            camara.addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION | Intent.FLAG_GRANT_READ_URI_PERMISSION);
            extras = new Intent[]{camara};
        } catch (Exception ignored) {}

        Intent galeria = new Intent(Intent.ACTION_GET_CONTENT);
        galeria.addCategory(Intent.CATEGORY_OPENABLE);
        galeria.setType("image/*");

        Intent selector = Intent.createChooser(galeria, "Foto del dron");
        selector.putExtra(Intent.EXTRA_INITIAL_INTENTS, extras);
        try {
            startActivityForResult(selector, REQ_FOTO);
        } catch (Exception e) {
            if (callbackFoto != null) callbackFoto.onReceiveValue(null);
            callbackFoto = null;
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode != REQ_FOTO || callbackFoto == null) return;
        Uri[] res = null;
        if (resultCode == RESULT_OK) {
            if (data != null && data.getData() != null) res = new Uri[]{data.getData()};
            else if (uriCamara != null) res = new Uri[]{uriCamara};
        }
        callbackFoto.onReceiveValue(res);
        callbackFoto = null;
    }

    @Override
    protected void onSaveInstanceState(Bundle out) {
        super.onSaveInstanceState(out);
        web.saveState(out);
    }

    @Override
    public void onBackPressed() {
        if (web.canGoBack()) web.goBack();
        else super.onBackPressed();
    }
}
