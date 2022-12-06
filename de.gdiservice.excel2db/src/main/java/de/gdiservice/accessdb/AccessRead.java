package de.gdiservice.accessdb;
import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import de.gdiservice.excel2db.ArgList;

//import org.postgis.jts.JtsGeometry;
//import org.postgresql.PGConnection;




public class AccessRead {

    Connection targetConnection;
    Connection srcConnection;

    String targetSchema;
    private String url;
    private Properties dbProps;

    public static Connection getConnection(String url, Properties dbProps) throws SQLException {

        Connection con =  DriverManager.getConnection(url+"?ApplicationName=Access2DB", dbProps);
        // PGConnection pgconn = con.unwrap(PGConnection.class);
        // pgconn.addDataType("geometry", JtsGeometry.class);
        return con;
    }

    AccessRead(String targetSchema, String url,  Properties props) {
        this.targetSchema = targetSchema;
        this.url = url;
        this.dbProps = props;
    }

    int getCount(Connection con, String table) throws SQLException {
        ResultSet rs = con.createStatement().executeQuery("select count(*) from "+table);
        if (rs.next()) {
            return rs.getInt(1);
        }
        return -1;
    }

    boolean checkSchema(String schemaName) {
        ResultSet res = null;
        try {
            res = targetConnection.getMetaData().getSchemas();
            System.err.println("checkSchema");

            while (res.next()) {
                String schema = res.getString("TABLE_SCHEM");
                if (schemaName.equals(schema)) {
                    return true;
                }
                System.err.println(res.getString("TABLE_CATALOG")+"  "+schema);
            }
        }
        catch (SQLException ex) {
            ex.printStackTrace();
        }
        finally {
            if (res!=null) {
                try {
                    res.close();
                } catch (SQLException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }
        return false;
    }

    boolean createSchema(String schemaName) {
        Statement stmt = null;
        try {
            stmt = targetConnection.createStatement();
            return stmt.execute("create schema "+schemaName);
        }
        catch (SQLException ex) {
            ex.printStackTrace();
        }
        finally {
            if (stmt!=null) {
                try {
                    stmt.close();
                } catch (SQLException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }
        return false;
    }    

    void cloneTable(TableDescriptor dsd) throws SQLException {

        System.out.println(dsd.tableName);
        //        for (ColumnDescriptor column : dsd.columns) {
        //            System.out.println("\t"+column);
        //        }

        String sqlDropTable = getSQLDropTable(targetSchema, dsd.tableName);
        System.out.println(sqlDropTable);
        targetConnection.createStatement().execute(sqlDropTable);


        String createStmt = getSQLCreateTable(targetSchema, dsd.tableName, dsd.columns);
        System.out.println(createStmt);        
        targetConnection.createStatement().execute(createStmt);


        PreparedStatement insertStatement = targetConnection.prepareStatement(getSQLInsert(targetSchema, dsd.tableName, dsd.columns));

        Statement selectStmt = srcConnection.createStatement();


        String sSelect = getSQLSelect(dsd.tableName, dsd.columns);        
        ResultSet rs = selectStmt.executeQuery(sSelect.toLowerCase()); 
        System.out.println("writing "+dsd.tableName+"  "+getCount(srcConnection, dsd.tableName));
        int count = 0;
        while (rs.next()) {
            for (int i=1; i<=dsd.columns.size(); i++) {
                Object o = rs.getObject(i);
                insertStatement.setObject(i, o);               
            }
            insertStatement.addBatch();;
            count++;
            if (count%1000 == 0) {
                System.out.println(count);
                insertStatement.executeBatch();
            }
        }        
        insertStatement.executeBatch();
    }


    void run(File f) throws IOException {
        
        String fName = f.getCanonicalPath();
//        if (File.separator.equals("/")) {
//            fName = fName.substring(1);
//        }
        String databaseURL = "jdbc:ucanaccess://" + fName;
        System.out.println(databaseURL);


        try (Connection connection = DriverManager.getConnection(databaseURL); ) {

            srcConnection = connection;
            targetConnection = getConnection(this.url, this.dbProps);

            ResultSet rs = connection.getMetaData().getTables(null, null, "%", null);
            List<String> tables = new ArrayList<>();
            while (rs.next()) { 
                String id = rs.getString(3); 
                tables.add(id);
                System.out.println(id);
            }
            rs.close();
            List<TableDescriptor> tableDescriptors = new ArrayList<>();
            for (String tableName : tables) {                
                List<ColumnDescriptor> columnDescriptors = new ArrayList<>();
                rs = connection.getMetaData().getColumns(null,null, tableName, null);
                int i=0;
                while (rs.next()) {
                    columnDescriptors.add(new ColumnDescriptor(i++, rs.getString("COLUMN_NAME"), rs.getString("TYPE_NAME")));
                }
                printRecords(rs, System.out);
                tableDescriptors.add(new TableDescriptor(tableName, columnDescriptors));                
            }

            boolean schemaExists = checkSchema(targetSchema);

            if (!schemaExists) {
                createSchema(targetSchema);
            }

            for (TableDescriptor dsd : tableDescriptors) {
                cloneTable(dsd);
            }

        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }

    static class TableDescriptor {
        String tableName;
        List<ColumnDescriptor> columns;

        TableDescriptor( String tableName, List<ColumnDescriptor> columns) {
            this.columns = columns;
            this.tableName = tableName;
        }        
    }

    static class ColumnDescriptor {
        final int nr;
        String name;
        final String type;

        public ColumnDescriptor(int nr, String colName, String type) {
            this.nr=nr;
            // this.name=(colName!=null ? colName.toLowerCase() : null);
            this.name=colName;
            this.type=type;
        }

        @Override
        public String toString() {
            return "ColumnDescriptor [nr=" + nr + ", name=" + name + ", type=" + type + "]";
        }


    }


    static public String getSQLDropTable(String schema, String tablename) {
        StringBuilder sb = new StringBuilder();
        sb.append("drop TABLE IF EXISTS ");
        if (schema != null) {
            sb.append(schema).append(".");
        }
        sb.append(tablename);
        return sb.toString();
    }

    static public String getSQLCreateTable(String schema, String tablename, List<ColumnDescriptor> columns) {
        StringBuilder sb = new StringBuilder();

        sb.append("CREATE TABLE ").append(schema).append(".").append(tablename);
        sb.append("(");
        // sb.append("id integer NOT NULL");

        for (int i=0, count=columns.size(); i<count; i++) {
            ColumnDescriptor col = columns.get(i);
            if (i>0) {
                sb.append(",\n");
            }
            sb.append("\"").append(col.name).append("\" ");
            sb.append(" varchar");
        }
        sb.append(")");
        return sb.toString();
    }    

    static public String getSQLInsert(String schema, String tablename, List<ColumnDescriptor> columns) {
        StringBuilder sb = new StringBuilder();

        sb.append("INSERT INTO ");
        if (schema != null) {
            sb.append(schema).append(".");
        }

        sb.append(tablename);
        sb.append("(");
        // sb.append("id");


        for (int i=0, count=columns.size(); i<count; i++) {
            if (i>0) {
                sb.append(",");
            }
            sb.append("\"").append(columns.get(i).name).append("\"");
        }
        sb.append(") ");

        sb.append("VALUES (?");
        for (int i=1, count=columns.size(); i<count; i++) {
            sb.append(",?");
        }
        sb.append(") ");
        return sb.toString();
    }    

    static public String getSQLSelect(String tablename, List<ColumnDescriptor> columns) {
        StringBuilder sb = new StringBuilder();
        sb.append(" select ");
        for (int i=0, count=columns.size(); i<count; i++) {
            if (i>0) {
                sb.append(",");
            }
            sb.append("").append(columns.get(i).name).append("");
        }
        sb.append(" from ").append(tablename);
        return sb.toString();
    }

    public static void printRecords(ResultSet rs, PrintStream out) throws SQLException {

        ResultSetMetaData rsMetaData = rs.getMetaData();
        int iColumnCount = rsMetaData.getColumnCount();

        StringBuilder sb = new StringBuilder();
        for (int colNr=1; colNr<=iColumnCount; colNr++) {
            sb.append(rsMetaData.getColumnName(colNr)).append(' ');
            for (int i = sb.length(); i < ((colNr + 1) * 20); i++) {
                sb.append(' ');
            }
        }
        out.println(sb);
        while (rs.next()) {
            sb.delete(0, Integer.MAX_VALUE);
            for (int colNr=1; colNr<=iColumnCount; colNr++) {
                sb.append(rs.getObject(rsMetaData.getColumnName(colNr))).append(' ');
                for (int i = sb.length(); i < ((colNr + 1) * 20); i++) {
                    sb.append(' ');
                }
            }
            out.println(sb);
        }
    }    

    public static void main(String[] args) {

        ArgList argList = new ArgList(args);
        try {
            String host = argList.get("host");
            String port = argList.get("port");
            String database = argList.get("database");
            String user = argList.get("user");
            String password = argList.get("password");
            String targetSchema = argList.get("zielschema");
            String datei = argList.get("datei");

            
            if (host==null) { 
                printVerwendung("host");
            }
            if (port==null) { 
                printVerwendung("port");
            }
            if (database==null) { 
                printVerwendung("database");
            }
            if (user==null) { 
                printVerwendung("user");
            }
            if (password==null) { 
                printVerwendung("password");
            }
            if (datei==null) { 
                printVerwendung("datei");
            }


            Properties dbProps = new Properties();
            dbProps.put("user", user);
            dbProps.put("password", password);                    
            String url = "jdbc:postgresql://"+host+":"+port+"/"+database+"?ApplicationName=Access2DB";
            System.out.println("URL=DB-Verbindung Ziel: \""+ url+ "\"");


            AccessRead accessRead = new AccessRead(targetSchema, url, dbProps);
            // accessRead.run("C:\\Users\\Ralf\\Nextcloud\\Austausch\\rtr\\mvb998.mdb");
            File f = new File(datei);
            System.out.println("Datei: \""+datei+"\"");
            System.out.println(f.getAbsoluteFile());
            System.out.println(f.getCanonicalPath());
            
            if (!f.exists()) {
                System.out.println("Datei \""+datei+"\" existiert nicht");
                System.exit(1);
            }
            
            if (!Files.isReadable(Paths.get(f.getCanonicalPath()))) {
                System.out.println("Datei \""+datei+"\" isReadable=false");
                System.exit(1);
            }
            accessRead.run(f);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    
    static void printVerwendung(String p) {
        System.out.println("Es fehlt ein Parameter:"+p);
        System.out.println("Verwendung: host=<ip oder dns> port=5432 database=mydb user=me password=### zielschema=test_schema datei=..\\myfile.mdb");
        System.out.println("\thost");
        System.out.println("\tport");
        System.out.println("\tdatabase");
        System.out.println("\tuser");
        System.out.println("\tpassword");
        System.out.println("\tzielschema");
        System.out.println("\tdatei");
        System.exit(1);
    }

}
