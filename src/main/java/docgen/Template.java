package docgen;

public class Template {
    public long id;
    public String name;
    public String description;
    public String jrxml;
    public String fields;   // JSON array of {name, type, pattern}
    public String params;   // JSON array of {name, type}
    public String createdAt;
    public String updatedAt;
}
