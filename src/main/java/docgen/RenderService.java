package docgen;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import net.sf.jasperreports.engine.*;
import net.sf.jasperreports.engine.data.JRMapCollectionDataSource;
import net.sf.jasperreports.engine.export.ooxml.JRXlsxExporter;
import net.sf.jasperreports.export.SimpleExporterInput;
import net.sf.jasperreports.export.SimpleOutputStreamExporterOutput;
import net.sf.jasperreports.export.SimpleXlsxReportConfiguration;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Compiles stored JRXML and renders documents from an uploaded data Excel.
 * The data file is a plain table: row 1 = column headers, rows 2+ = records.
 * Header names are matched (trim + case-insensitive) against the JRXML
 * field names so values can be coerced into the declared field classes.
 */
public class RenderService {

    private static final Gson GSON = new Gson();
    private static final Pattern FIELD_REF = Pattern.compile("\\$F\\{(.*?)\\}");
    private static final Pattern FONT_ATTR = Pattern.compile("fontName=\"([^\"]*)\"");
    /** Font names whose faces ship with the app (font extension family "Sarabun"). */
    private static final String BUNDLED_FONT = "Sarabun";

    private final Map<Long, JasperReport> compileCache = new HashMap<>();

    public static class RenderException extends Exception {
        public RenderException(String m) { super(m); }
        public RenderException(String m, Throwable c) { super(m, c); }
    }

    // ------------------------------------------------------------------
    // Compile
    // ------------------------------------------------------------------

    public synchronized JasperReport compile(Template t) throws RenderException {
        if (t.id > 0) {
            JasperReport cached = compileCache.get(t.id);
            if (cached != null) return cached;
        }
        try {
            String normalized = normalizeFonts(t.jrxml);
            byte[] xml = normalized.getBytes(StandardCharsets.UTF_8);
            JasperReport rep = JasperCompileManager.compileReport(new ByteArrayInputStream(xml));
            // Only cache real (persisted) templates — probes use id<0 and must not pollute.
            if (t.id > 0) compileCache.put(t.id, rep);
            return rep;
        } catch (JRException e) {
            throw new RenderException("JRXML ผิดพลาด ไม่สามารถ compile ได้: " + e.getMessage(), e);
        }
    }

    public synchronized void evict(long templateId) {
        compileCache.remove(templateId);
    }

    /**
     * Template cells usually name fonts that exist on the author's machine only
     * (Sarabun, Tahoma, Calibri, Angsana...). The app ships Sarabun (Thai + Latin
     * faces) as a Jasper font-extension family, and every font reference is
     * rewritten to Sarabun before compiling so Thai glyphs always render in the
     * PDF regardless of the font the Excel file declared.
     */
    static String normalizeFonts(String jrxml) {
        if (jrxml == null || jrxml.isEmpty()) return jrxml;
        Matcher m = FONT_ATTR.matcher(jrxml);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            m.appendReplacement(sb, Matcher.quoteReplacement("fontName=\"" + BUNDLED_FONT + "\""));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    /** Fields & parameters declared by the compiled report (for template metadata + UI). */
    public JsonObject describe(Template t) throws RenderException {
        JasperReport rep = compile(t);
        JsonObject out = new JsonObject();
        JsonArray fields = new JsonArray();
        for (JRField f : rep.getFields()) {
            JsonObject o = new JsonObject();
            o.addProperty("name", f.getName());
            o.addProperty("type", f.getValueClass().getName());
            fields.add(o);
        }
        JsonArray params = new JsonArray();
        for (JRParameter p : rep.getParameters()) {
            if (p.isSystemDefined()) continue;
            JsonObject o = new JsonObject();
            o.addProperty("name", p.getName());
            o.addProperty("type", p.getValueClass().getName());
            params.add(o);
        }
        out.add("fields", fields);
        out.add("parameters", params);
        return out;
    }

    // ------------------------------------------------------------------
    // Render
    // ------------------------------------------------------------------

    /** @param dataXlsx raw bytes of the data Excel file */
    public byte[] render(Template t, byte[] dataXlsx, String format, JsonObject paramValues) throws RenderException {
        format = (format == null || format.isBlank()) ? "pdf" : format.trim().toLowerCase(Locale.ROOT);
        if (!format.equals("pdf") && !format.equals("xlsx")) {
            throw new RenderException("รูปแบบไม่รองรับ: " + format + " (ใช้ pdf หรือ xlsx)");
        }
        JasperReport rep = compile(t);

        List<Map<String, ?>> records = readDataSheet(dataXlsx, rep);
        sortByGroups(records, rep);

        Map<String, Object> params = buildParams(paramValues, rep);

        try {
            JRDataSource ds = records.isEmpty() ? null : new JRMapCollectionDataSource(records);
            JasperPrint print = JasperFillManager.fillReport(rep, params, ds);
            switch (format) {
                case "pdf": return exportPdf(print);
                case "xlsx": return exportXlsx(print);
                default: throw new RenderException("รูปแบบไม่รองรับ");
            }
        } catch (JRException e) {
            throw new RenderException("ไม่สามารถ render เอกสารได้: " + e.getMessage(), e);
        }
    }

    private byte[] exportPdf(JasperPrint print) throws JRException {
        return JasperExportManager.exportReportToPdf(print);
    }

    private byte[] exportXlsx(JasperPrint print) throws JRException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        JRXlsxExporter exporter = new JRXlsxExporter();
        exporter.setExporterInput(new SimpleExporterInput(print));
        exporter.setExporterOutput(new SimpleOutputStreamExporterOutput(out));
        SimpleXlsxReportConfiguration cfg = new SimpleXlsxReportConfiguration();
        cfg.setWhitePageBackground(false);
        cfg.setRemoveEmptySpaceBetweenColumns(true);
        cfg.setDetectCellType(true);
        cfg.setWrapText(false);
        cfg.setOnePagePerSheet(false);
        exporter.setConfiguration(cfg);
        exporter.exportReport();
        return out.toByteArray();
    }

    // ------------------------------------------------------------------
    // Parameter conversion
    // ------------------------------------------------------------------

    private Map<String, Object> buildParams(JsonObject paramValues, JasperReport rep) throws RenderException {
        Map<String, Object> params = new HashMap<>();
        if (paramValues == null) return params;
        for (JRParameter p : rep.getParameters()) {
            if (p.isSystemDefined()) continue;
            String name = p.getName();
            if (!paramValues.has(name)) continue;
            String raw = paramValues.get(name).isJsonNull() ? null : paramValues.get(name).getAsString();
            if (raw == null || raw.isBlank()) continue;
            params.put(name, convertToClass(raw, p.getValueClass(), name));
        }
        return params;
    }

    // ------------------------------------------------------------------
    // Data sheet -> records (header row + typed data rows)
    // ------------------------------------------------------------------

    private List<Map<String, ?>> readDataSheet(byte[] xlsx, JasperReport rep) throws RenderException {
        // normalized field name -> declared class + exact field name
        Map<String, Class<?>> fieldTypes = new HashMap<>();
        Map<String, String> fieldExact = new HashMap<>();
        for (JRField f : rep.getFields()) {
            fieldTypes.put(norm(f.getName()), f.getValueClass());
            fieldExact.put(norm(f.getName()), f.getName());
        }

        try (Workbook wb = WorkbookFactory.create(new ByteArrayInputStream(xlsx))) {
            Sheet sheet = wb.getNumberOfSheets() > 0 ? wb.getSheetAt(0) : null;
            if (sheet == null) throw new RenderException("ไฟล์ข้อมูลไม่มีชีต (sheet)");

            DataFormatter formatter = new DataFormatter(Locale.ENGLISH);
            Iterator<Row> it = sheet.iterator();
            if (!it.hasNext()) throw new RenderException("ไฟล์ข้อมูลว่างเปล่า");

            Row headerRow = it.next();
            List<String> headers = new ArrayList<>();
            List<Integer> headerCols = new ArrayList<>();
            int firstCell = headerRow.getFirstCellNum();
            int lastCell = headerRow.getLastCellNum();
            if (firstCell < 0) throw new RenderException("ไม่มีแถวหัวคอลัมน์ (row 1 ต้องเป็นชื่อคอลัมน์)");
            for (int c = firstCell; c < lastCell; c++) {
                Cell cell = headerRow.getCell(c);
                String h = cell == null ? "" : formatter.formatCellValue(cell).trim();
                if (!h.isEmpty()) { headers.add(h); headerCols.add(c); }
            }
            if (headers.isEmpty()) throw new RenderException("แถวแรกต้องเป็นชื่อคอลัมน์");

            List<Map<String, ?>> records = new ArrayList<>();
            int rowIdx = 2; // 1-based display row (header = row 1)
            while (it.hasNext()) {
                Row r = it.next();
                rowIdx++;
                if (isRowEmpty(r)) continue;
                Map<String, Object> record = new LinkedHashMap<>();
                boolean any = false;
                for (int i = 0; i < headers.size(); i++) {
                    String h = headers.get(i);
                    Cell cell = r.getCell(headerCols.get(i));
                    Class<?> target = fieldTypes.get(norm(h));
                    Object val = convertCell(cell, target, formatter, h, rowIdx);
                    if (val != null) any = true;
                    // Jasper data source looks up by the exact declared field name.
                    String key = fieldExact.getOrDefault(norm(h), h);
                    record.put(key, val);
                }
                if (any) records.add(record);
            }
            return records;
        } catch (java.io.IOException e) {
            throw new RenderException("ไม่สามารถอ่านไฟล์ Excel ได้: " + e.getMessage(), e);
        }
    }

    private boolean isRowEmpty(Row r) {
        if (r == null) return true;
        for (Cell c : r) {
            if (c != null && c.getCellType() != CellType.BLANK && c.getCellType() != CellType._NONE) return false;
        }
        return true;
    }

    static String norm(String s) {
        return s == null ? "" : s.trim().toLowerCase(Locale.ROOT);
    }

    /** Field names expected in the data file header (exact declared names). */
    public List<String> declaredFieldNames(Template t) throws RenderException {
        JasperReport rep = compile(t);
        List<String> names = new ArrayList<>();
        for (JRField f : rep.getFields()) names.add(f.getName());
        return names;
    }

    private Object convertCell(Cell cell, Class<?> target, DataFormatter formatter, String header, int rowIdx)
            throws RenderException {
        if (cell == null) return null;
        CellType type;
        try {
            type = cell.getCellType();
        } catch (Exception e) {
            return null;
        }
        if (type == CellType.BLANK || type == CellType._NONE) return null;
        if (type == CellType.FORMULA) {
            type = cell.getCachedFormulaResultType();
        }

        // STRING target (default in templates): return the display text
        boolean isString = target == null || target == String.class || target == Object.class;

        try {
            if (type == CellType.BOOLEAN) {
                boolean b = cell.getBooleanCellValue();
                if (isString) return String.valueOf(b);
                return b;
            }
            if (type == CellType.ERROR) return null;

            if (target != null && isDateLike(target)) {
                Date d;
                if (type == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                    d = cell.getDateCellValue();
                } else if (type == CellType.NUMERIC) {
                    d = DateUtil.getJavaDate(cell.getNumericCellValue());
                } else {
                    String s = cell.getStringCellValue();
                    d = parseDate(s);
                }
                if (d == null) return null;
                if (target == Time.class) return new Time(d.getTime());
                if (target == Timestamp.class) return new Timestamp(d.getTime());
                if (target == java.sql.Date.class) return new java.sql.Date(d.getTime());
                return d; // java.util.Date / java.time mapped later if needed
            }

            if (type == CellType.NUMERIC) {
                double num = cell.getNumericCellValue();
                if (isString) {
                    // Integral & "General" style -> plain integer text (e.g. 45000 not 45000.0)
                    if (num == Math.floor(num) && !Double.isInfinite(num)) {
                        return new BigDecimal(num).toBigInteger().toString();
                    }
                    String text = formatter.formatCellValue(cell);
                    return (text == null || text.isEmpty()) ? String.valueOf(num) : text;
                }
                if (target == Integer.class) return (int) Math.round(num);
                if (target == Long.class) return Math.round(num);
                if (target == Short.class) return (short) Math.round(num);
                if (target == Byte.class) return (byte) Math.round(num);
                if (target == Double.class) return num;
                if (target == Float.class) return (float) num;
                if (target == Boolean.class) return num != 0;
                return BigDecimal.valueOf(num);
            }

            // STRING cell
            String s = type == CellType.STRING ? cell.getStringCellValue() : formatter.formatCellValue(cell);
            if (s == null) return null;
            s = s.trim();
            if (s.isEmpty()) return null;
            if (isString) return s;
            return convertToClass(s, target, header + " (แถว " + rowIdx + ")");
        } catch (RenderException re) {
            throw re;
        } catch (Exception e) {
            throw new RenderException("แปลงค่าคอลัมน์ '" + header + "' (แถว " + rowIdx + ") ไม่สำเร็จ: " + e.getMessage(), e);
        }
    }

    /** Parse a free-text date with common Thai/ISO formats. */
    static Date parseDate(String s) {
        if (s == null) return null;
        s = s.trim();
        if (s.isEmpty()) return null;
        String[] patterns = {
            "dd/MM/yyyy HH:mm:ss", "dd/MM/yyyy HH:mm", "dd/MM/yyyy",
            "d/M/yyyy", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy-MM-dd",
            "dd-MM-yyyy HH:mm:ss", "dd-MM-yyyy", "dd.MM.yyyy", "MM/dd/yyyy", "yyyy/MM/dd"
        };
        for (String p : patterns) {
            try {
                SimpleDateFormat f = new SimpleDateFormat(p, Locale.ENGLISH);
                f.setLenient(false);
                return f.parse(s);
            } catch (ParseException ignored) { }
        }
        // Fallback: ISO instant
        try {
            return java.util.Date.from(java.time.Instant.parse(s));
        } catch (Exception ignored) { }
        return null;
    }

    private Object convertToClass(String s, Class<?> target, String where) throws RenderException {
        String clean = s.replace(",", "").replace("฿", "").replace("$", "").replace(" ", "");
        try {
            if (target == Integer.class) return Integer.parseInt(clean);
            if (target == Long.class) return Long.parseLong(clean);
            if (target == Short.class) return Short.parseShort(clean);
            if (target == Byte.class) return Byte.parseByte(clean);
            if (target == Double.class) return Double.parseDouble(clean);
            if (target == Float.class) return Float.parseFloat(clean);
            if (target == Boolean.class) return Boolean.parseBoolean(s);
            if (target == BigDecimal.class) return new BigDecimal(clean);
            if (isDateLike(target)) {
                Date d = parseDate(s);
                if (d == null) return null;
                if (target == Time.class) return new Time(d.getTime());
                if (target == Timestamp.class) return new Timestamp(d.getTime());
                if (target == java.sql.Date.class) return new java.sql.Date(d.getTime());
                return d;
            }
            return s;
        } catch (NumberFormatException e) {
            throw new RenderException("ค่า '" + s + "' ของ " + where + " ไม่ใช่ตัวเลขที่ต้องการ (" + target.getSimpleName() + ")");
        }
    }

    private static boolean isDateLike(Class<?> c) {
        return c == java.util.Date.class || c == java.sql.Date.class || c == Timestamp.class
                || c == Time.class || c == java.time.LocalDate.class || c == java.time.LocalDateTime.class;
    }

    // ------------------------------------------------------------------
    // Group sorting (Jasper groups expect pre-sorted data)
    // ------------------------------------------------------------------

    private void sortByGroups(List<Map<String, ?>> records, JasperReport rep) {
        if (records == null || records.isEmpty()) return;
        JRGroup[] groups = rep.getGroups();
        if (groups == null || groups.length == 0) return;

        List<String> groupFields = new ArrayList<>();
        for (JRGroup g : groups) {
            JRExpression expr = g.getExpression();
            if (expr == null) continue;
            Matcher m = FIELD_REF.matcher(expr.getText());
            if (m.find()) groupFields.add(m.group(1).trim());
        }
        if (groupFields.isEmpty()) return;

        // Case-insensitive key lookup like the column mapping.
        records.sort((a, b) -> {
            for (String gf : groupFields) {
                Object va = findValue(a, gf);
                Object vb = findValue(b, gf);
                int c = compareValues(va, vb);
                if (c != 0) return c;
            }
            return 0;
        });
    }

    private Object findValue(Map<String, ?> record, String name) {
        for (Map.Entry<String, ?> e : record.entrySet()) {
            if (norm(e.getKey()).equals(norm(name))) return e.getValue();
        }
        return null;
    }

    @SuppressWarnings({"unchecked", "rawtypes"})
    private int compareValues(Object a, Object b) {
        if (a == null && b == null) return 0;
        if (a == null) return -1;
        if (b == null) return 1;
        if (a instanceof Comparable && a.getClass().isInstance(b)) {
            return ((Comparable) a).compareTo(b);
        }
        return String.valueOf(a).compareTo(String.valueOf(b));
    }
}
