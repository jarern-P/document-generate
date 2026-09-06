package docgen;

import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public class Db {

    private static Connection conn;

    public static void init() throws Exception {
        Class.forName("org.sqlite.JDBC");
        Path dbPath = Main.dbFile;
        if (dbPath.getParent() != null) Files.createDirectories(dbPath.getParent());
        conn = DriverManager.getConnection("jdbc:sqlite:" + dbPath.toAbsolutePath());
        try (Statement st = conn.createStatement()) {
            st.executeUpdate("""
                CREATE TABLE IF NOT EXISTS templates (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    name        TEXT NOT NULL,
                    description TEXT DEFAULT '',
                    jrxml       TEXT NOT NULL,
                    fields      TEXT DEFAULT '[]',
                    params      TEXT DEFAULT '[]',
                    created_at  TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                    updated_at  TEXT NOT NULL DEFAULT (datetime('now','localtime'))
                )
                """);
        }
    }

    public static Connection c() { return conn; }

    public static void close() throws Exception { if (conn != null) conn.close(); }

    // ---- Template CRUD ----

    public static long insertTemplate(String name, String description, String jrxml, String fields, String params) throws SQLException {
        String sql = "INSERT INTO templates (name, description, jrxml, fields, params) VALUES (?,?,?,?,?)";
        try (PreparedStatement ps = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            ps.setString(1, name);
            ps.setString(2, description == null ? "" : description);
            ps.setString(3, jrxml);
            ps.setString(4, fields == null ? "[]" : fields);
            ps.setString(5, params == null ? "[]" : params);
            ps.executeUpdate();
            try (ResultSet rs = ps.getGeneratedKeys()) {
                return rs.next() ? rs.getLong(1) : -1;
            }
        }
    }

    public static boolean updateTemplate(long id, String name, String description, String jrxml, String fields, String params) throws SQLException {
        String sql = "UPDATE templates SET name=?, description=?, jrxml=?, fields=?, params=?, updated_at=datetime('now','localtime') WHERE id=?";
        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setString(1, name);
            ps.setString(2, description == null ? "" : description);
            ps.setString(3, jrxml);
            ps.setString(4, fields == null ? "[]" : fields);
            ps.setString(5, params == null ? "[]" : params);
            ps.setLong(6, id);
            return ps.executeUpdate() > 0;
        }
    }

    public static boolean deleteTemplate(long id) throws SQLException {
        try (PreparedStatement ps = conn.prepareStatement("DELETE FROM templates WHERE id=?")) {
            ps.setLong(1, id);
            return ps.executeUpdate() > 0;
        }
    }

    public static Template getTemplate(long id) throws SQLException {
        try (PreparedStatement ps = conn.prepareStatement("SELECT * FROM templates WHERE id=?")) {
            ps.setLong(1, id);
            try (ResultSet rs = ps.executeQuery()) {
                if (!rs.next()) return null;
                return rowToTemplate(rs);
            }
        }
    }

    public static List<Template> listTemplates() throws SQLException {
        List<Template> out = new ArrayList<>();
        try (Statement st = conn.createStatement();
             ResultSet rs = st.executeQuery("SELECT * FROM templates ORDER BY updated_at DESC, id DESC")) {
            while (rs.next()) out.add(rowToTemplate(rs));
        }
        return out;
    }

    private static Template rowToTemplate(ResultSet rs) throws SQLException {
        Template t = new Template();
        t.id = rs.getLong("id");
        t.name = rs.getString("name");
        t.description = rs.getString("description");
        t.jrxml = rs.getString("jrxml");
        t.fields = rs.getString("fields");
        t.params = rs.getString("params");
        t.createdAt = rs.getString("created_at");
        t.updatedAt = rs.getString("updated_at");
        return t;
    }
}
