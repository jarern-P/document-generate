package docgen;

import com.sun.net.httpserver.HttpServer;
import com.sun.net.httpserver.HttpHandler;
import com.sun.net.httpserver.HttpExchange;

import java.io.*;
import java.net.InetSocketAddress;
import java.net.URI;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.Executors;

/**
 * Project Document Generator — standalone server.
 * Serves the HTML/JS UI and the JSON/REST API; stores templates in SQLite;
 * renders documents with JasperReports (PDF / XLSX).
 */
public class Main {

    public static final int DEFAULT_PORT = 8080;
    public static Path dataDir;
    public static Path dbFile;

    public static void main(String[] args) throws Exception {
        int port = DEFAULT_PORT;
        for (String a : args) {
            if (a.startsWith("--port=")) port = Integer.parseInt(a.substring("--port=".length()));
            else if (a.startsWith("--data=")) dataDir = Paths.get(a.substring("--data=".length()));
        }
        String envPort = System.getenv("DOCGEN_PORT");
        if (envPort != null && !envPort.isBlank()) port = Integer.parseInt(envPort.trim());

        if (dataDir == null) {
            String home = System.getProperty("user.home", ".");
            dataDir = Paths.get(home, ".document-generator");
        }
        Files.createDirectories(dataDir);
        dbFile = dataDir.resolve("docgen.db");
        log("Data directory : " + dataDir);

        Db.init();

        RenderService render = new RenderService();

        HttpServer server = HttpServer.create(new InetSocketAddress("127.0.0.1", port), 0);
        server.setExecutor(Executors.newFixedThreadPool(8));

        Api api = new Api(render, dbFile);

        server.createContext("/api/", api);
        server.createContext("/", new StaticHandler());
        server.start();

        String url = "http://127.0.0.1:" + port + "/";
        log("");
        log("Project Document Generator is running.");
        log("  UI : " + url);
        log("Press Ctrl+C to stop.");

        // Open the default browser after a short delay so the server is up.
        try {
            openBrowser(url);
        } catch (Throwable t) {
            log("Could not auto-open browser: " + t.getMessage());
            log("Please open manually: " + url);
        }

        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            server.stop(0);
            try { Db.close(); } catch (Exception ignored) {}
            log("Server stopped.");
        }));

        // Keep alive
        Thread.currentThread().join();
    }

    private static void openBrowser(String url) {
        String os = System.getProperty("os.name", "").toLowerCase();
        // Prefer java.awt.Desktop when a GUI session is available.
        try {
            if (java.awt.Desktop.isDesktopSupported()) {
                java.awt.Desktop.getDesktop().browse(java.net.URI.create(url));
                return;
            }
        } catch (Throwable ignored) { }
        try {
            if (os.contains("win")) {
                new ProcessBuilder("rundll32", "url.dll,FileProtocolHandler", url).start();
            } else if (os.contains("mac")) {
                new ProcessBuilder("open", url).start();
            } else {
                new ProcessBuilder("xdg-open", url).start();
            }
        } catch (IOException ignored) { }
    }

    static void log(String msg) {
        System.out.println("[docgen] " + msg);
    }
}
