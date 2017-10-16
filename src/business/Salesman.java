package business;

public class Salesman {

    private String tableName;
    private String directory;
    private String export;

    public Salesman (String tableName, String directory, String export) {
        this.tableName = tableName;
        this.directory = directory;
        this.export = export;
    }

    public String getTableName() { return tableName; }

    public void setTableName(String tableName) { this.tableName = tableName; }

    public String getDirectory() { return directory; }

    public void setDirectory(String directory) { this.directory = directory; }

    public String getExport() { return export; }

    public void setExport(String export) { this.export = export; }
}
