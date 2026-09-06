package docgen;

import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;

/** Serves the HTML/JS/CSS front-end from the classpath (src/main/resources/web). */
public class StaticHandler implements HttpHandler {

    @Override
    public void handle(HttpExchange ex) throws IOException {
        String path = ex.getRequestURI().getPath();
        if (path == null || path.equals("/")) path = "/index.html";
        if (path.startsWith("/")) path = path.substring(1);
        path = URLDecoder.decode(path, StandardCharsets.UTF_8);
        if (path.contains("..")) { ex.sendResponseHeaders(403, -1); return; }

        String resource = "/web/" + path;
        InputStream in = StaticHandler.class.getResourceAsStream(resource);
        if (in == null) {
            // SPA fallback for unknown paths -> index
            if (!path.contains(".")) {
                in = StaticHandler.class.getResourceAsStream("/web/index.html");
                if (in != null) path = "index.html";
            }
        }
        if (in == null) {
            byte[] msg = "404 Not Found".getBytes(StandardCharsets.UTF_8);
            ex.sendResponseHeaders(404, msg.length);
            try (OutputStream os = ex.getResponseBody()) { os.write(msg); }
            return;
        }
        String ct = contentType(path);
        ex.getResponseHeaders().set("Content-Type", ct + (ct.startsWith("text/") ? "; charset=utf-8" : ""));
        ex.getResponseHeaders().set("Cache-Control", "no-cache");
        byte[] data = in.readAllBytes();
        in.close();
        ex.sendResponseHeaders(200, data.length);
        try (OutputStream os = ex.getResponseBody()) { os.write(data); }
    }

    private String contentType(String path) {
        if (path.endsWith(".html")) return "text/html";
        if (path.endsWith(".js")) return "application/javascript";
        if (path.endsWith(".mjs")) return "application/javascript";
        if (path.endsWith(".css")) return "text/css";
        if (path.endsWith(".json")) return "application/json";
        if (path.endsWith(".svg")) return "image/svg+xml";
        if (path.endsWith(".png")) return "image/png";
        if (path.endsWith(".jpg") || path.endsWith(".jpeg")) return "image/jpeg";
        if (path.endsWith(".gif")) return "image/gif";
        if (path.endsWith(".ico")) return "image/x-icon";
        if (path.endsWith(".woff")) return "font/woff";
        if (path.endsWith(".woff2")) return "font/woff2";
        if (path.endsWith(".ttf")) return "font/ttf";
        if (path.endsWith(".map")) return "application/json";
        if (path.endsWith(".xlsx")) return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        return "application/octet-stream";
    }
}
