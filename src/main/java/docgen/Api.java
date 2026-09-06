package docgen;

import com.google.gson.*;
import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Base64;
import java.util.List;
import java.util.Map;

/** Handles all /api/* routes as JSON (one big handler + manual routing). */
public class Api implements HttpHandler {

    private final RenderService render;
    private final Gson gson = new GsonBuilder().setPrettyPrinting().create();

    public Api(RenderService render, java.nio.file.Path dbFile) {
        this.render = render;
    }

    @Override
    public void handle(HttpExchange ex) throws IOException {
        try {
            String path = ex.getRequestURI().getPath(); // e.g. /api/templates/3
            String method = ex.getRequestMethod().toUpperCase();
            ex.getResponseHeaders().add("Access-Control-Allow-Origin", "*");
            ex.getResponseHeaders().add("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
            ex.getResponseHeaders().add("Access-Control-Allow-Headers", "Content-Type");
            if (method.equals("OPTIONS")) { ex.sendResponseHeaders(204, -1); return; }

            String rest = path.startsWith("/api/") ? path.substring("/api/".length()) : path;
            rest = rest.split("\\?")[0];

            if (rest.equals("health") && method.equals("GET")) { sendJson(ex, 200, "{\"status\":\"ok\"}"); return; }

            if (rest.equals("templates")) {
                if (method.equals("GET")) { handleList(ex); return; }
                if (method.equals("POST")) { handleSave(ex, -1); return; }
            } else if (rest.startsWith("templates/")) {
                String idStr = rest.substring("templates/".length());
                long id = parseId(idStr);
                if (method.equals("GET")) { handleGet(ex, id); return; }
                if (method.equals("PUT")) { handleSave(ex, id); return; }
                if (method.equals("DELETE")) { handleDelete(ex, id); return; }
            } else if (rest.equals("render") && method.equals("POST")) {
                handleRender(ex); return;
            }

            sendJson(ex, 404, errJson("ไม่พบ endpoint: " + path));
        } catch (RenderService.RenderException e) {
            sendJson(ex, 400, errJson(e.getMessage()));
        } catch (Throwable e) {
            Main.log("API error: " + e);
            e.printStackTrace();
            try { sendJson(ex, 500, errJson(e.getMessage() == null ? e.toString() : e.getMessage())); }
            catch (Throwable ignored) { }
        }
    }

    private long parseId(String s) {
        try {
            String clean = s.contains("/") ? s.substring(0, s.indexOf("/")) : s;
            return Long.parseLong(clean.trim());
        } catch (Exception e) { return -1; }
    }

    // ------------------------------------------------------------------

    private void handleList(HttpExchange ex) throws Exception {
        List<Template> list = Db.listTemplates();
        JsonArray arr = new JsonArray();
        for (Template t : list) arr.add(summary(t));
        sendJson(ex, 200, gson.toJson(arr));
    }

    private void handleGet(HttpExchange ex, long id) throws Exception {
        Template t = Db.getTemplate(id);
        if (t == null) { sendJson(ex, 404, errJson("ไม่พบ template id=" + id)); return; }
        JsonObject o = summary(t);
        o.addProperty("jrxml", t.jrxml);
        sendJson(ex, 200, gson.toJson(o));
    }

    private JsonObject summary(Template t) {
        JsonObject o = new JsonObject();
        o.addProperty("id", t.id);
        o.addProperty("name", t.name);
        o.addProperty("description", t.description == null ? "" : t.description);
        o.addProperty("createdAt", t.createdAt);
        o.addProperty("updatedAt", t.updatedAt);
        try {
            o.add("fields", JsonParser.parseString(t.fields == null || t.fields.isBlank() ? "[]" : t.fields).getAsJsonArray());
        } catch (Exception e) { o.add("fields", new JsonArray()); }
        try {
            o.add("params", JsonParser.parseString(t.params == null || t.params.isBlank() ? "[]" : t.params).getAsJsonArray());
        } catch (Exception e) { o.add("params", new JsonArray()); }
        return o;
    }

    private void handleSave(HttpExchange ex, long id) throws Exception {
        JsonObject body = readJson(ex);
        String name = getStr(body, "name");
        String description = getStr(body, "description");
        String jrxml = getStr(body, "jrxml");

        if (name == null || name.isBlank()) { sendJson(ex, 400, errJson("ต้องระบุชื่อ template")); return; }
        if (jrxml == null || jrxml.isBlank()) { sendJson(ex, 400, errJson("jrxml ว่างเปล่า — กรุณาแปลง Excel ก่อนบันทึก")); return; }

        // Validate & describe (compile) BEFORE storing.
        Template probe = new Template();
        probe.id = id > 0 ? id : Long.MIN_VALUE; // avoid stale cache on create
        probe.jrxml = jrxml;
        JsonObject desc = render.describe(probe);
        String fields = desc.get("fields").toString();
        String params = desc.get("parameters").toString();

        long saved;
        if (id > 0) {
            boolean ok = Db.updateTemplate(id, name, description, jrxml, fields, params);
            if (!ok) { sendJson(ex, 404, errJson("ไม่พบ template id=" + id)); return; }
            saved = id;
        } else {
            saved = Db.insertTemplate(name, description, jrxml, fields, params);
        }
        render.evict(id > 0 ? id : saved);
        JsonObject out = new JsonObject();
        out.addProperty("id", saved);
        out.add("fields", JsonParser.parseString(fields).getAsJsonArray());
        out.add("params", JsonParser.parseString(params).getAsJsonArray());
        sendJson(ex, 200, gson.toJson(out));
    }

    private void handleDelete(HttpExchange ex, long id) throws Exception {
        render.evict(id);
        boolean ok = Db.deleteTemplate(id);
        if (!ok) { sendJson(ex, 404, errJson("ไม่พบ template id=" + id)); return; }
        sendJson(ex, 200, "{\"ok\":true}");
    }

    private void handleRender(HttpExchange ex) throws Exception {
        JsonObject body = readJson(ex);
        long templateId = body.has("templateId") ? body.get("templateId").getAsLong() : -1;
        if (templateId <= 0) { sendJson(ex, 400, errJson("ต้องระบุ templateId")); return; }
        Template t = Db.getTemplate(templateId);
        if (t == null) { sendJson(ex, 404, errJson("ไม่พบ template id=" + templateId)); return; }

        String format = getStr(body, "format");
        String dataB64 = getStr(body, "data");
        if (dataB64 == null || dataB64.isBlank()) { sendJson(ex, 400, errJson("กรุณาอัปโหลดไฟล์ Excel ข้อมูล")); return; }
        byte[] data;
        try {
            data = Base64.getMimeDecoder().decode(dataB64.replaceFirst("^data:[^;]*;base64,", ""));
        } catch (IllegalArgumentException e) {
            sendJson(ex, 400, errJson("ข้อมูล Excel (base64) ไม่ถูกต้อง"));
            return;
        }

        JsonObject params = body.has("params") && body.get("params").isJsonObject()
                ? body.getAsJsonObject("params") : null;

        byte[] out = render.render(t, data, format, params);
        String ext = format != null && format.equalsIgnoreCase("xlsx") ? "xlsx" : "pdf";
        String stamp = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss").withZone(ZoneId.systemDefault()).format(Instant.now());
        // Header values must be ASCII (RFC 7230) — Thai template names cannot go in
        // Content-Disposition, so keep the wire filename ASCII and let the client name it.
        String fname = "document-" + t.id + "-" + stamp + "." + ext;

        if (ext.equals("pdf")) {
            ex.getResponseHeaders().set("Content-Type", "application/pdf");
            ex.getResponseHeaders().set("Content-Disposition", "inline; filename=\"" + fname + "\"");
        } else {
            ex.getResponseHeaders().set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            ex.getResponseHeaders().set("Content-Disposition", "attachment; filename=\"" + fname + "\"");
        }
        ex.sendResponseHeaders(200, out.length);
        try (OutputStream os = ex.getResponseBody()) { os.write(out); }
    }

    // ------------------------------------------------------------------

    private String getStr(JsonObject o, String key) {
        if (o == null || !o.has(key) || o.get(key).isJsonNull()) return null;
        JsonElement e = o.get(key);
        return e.isJsonPrimitive() ? e.getAsString() : e.toString();
    }

    private JsonObject readJson(HttpExchange ex) throws IOException {
        String body = new String(ex.getRequestBody().readAllBytes(), StandardCharsets.UTF_8);
        if (body.isBlank()) return new JsonObject();
        try {
            return JsonParser.parseString(body).getAsJsonObject();
        } catch (Exception e) {
            throw new IOException("JSON body ไม่ถูกต้อง: " + e.getMessage());
        }
    }

    private String errJson(String msg) {
        JsonObject o = new JsonObject();
        o.addProperty("error", msg);
        return o.toString();
    }

    private void sendJson(HttpExchange ex, int status, String json) throws IOException {
        byte[] bytes = json.getBytes(StandardCharsets.UTF_8);
        ex.getResponseHeaders().set("Content-Type", "application/json; charset=utf-8");
        ex.sendResponseHeaders(status, bytes.length);
        try (OutputStream os = ex.getResponseBody()) { os.write(bytes); }
    }
}
