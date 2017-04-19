import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.sql.*;
import java.util.Properties;

public class DBRead {
//a
    static Properties prop;

    static {
        prop = new Properties();
        InputStream in = DBRead.class.getResourceAsStream("DBRead.properties");
        try {
            prop.load(in);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                in.close();
            } catch (IOException e) {}
        }
    }

    // DB CONSTANTS READ FROM PROPERTY FILE 'DBRead.properties'
    private static final String DB_DRIVER_NAME = prop.getProperty("DB_DRIVER_NAME");
    private static final String DB_URL = prop.getProperty("DB_URL");
    private static final String DB_USER = prop.getProperty("DB_USER");
    private static final String DB_PASSWORD = prop.getProperty("DB_PASSWORD");

    // FILENAME INCLUDING THE PATH READ FROM PROPERTY FILE 'DBRead.properties'
    private static final String QUERIES_FILENAME = prop.getProperty("QUERIES_FILENAME");

    // APPLICATION CONSTANTS
    private static final String LINE_SEPARATOR = System.getProperty("line.separator");
    private static final String COLUMN_SEPARATOR = "\t";

    // INSTANCE VARIABLES
    private static StringBuffer queriesSb;
    private static Connection con;
    private static Statement stmt;

    public static void main(String args[]) {

        // Read the input file
        try {
            queriesSb = (readFile(QUERIES_FILENAME));
            if (queriesSb.length() < 10) {
                System.err.println("The input file is empty");
                System.exit(-1);
            }
        } catch (IOException e) {
            System.err.println(e.getMessage());
            System.exit(-1);
        }

        // Make a connection to DB and print out the results
        try {
            Class.forName(DB_DRIVER_NAME);
            con = DriverManager.getConnection(DB_URL, DB_USER, DB_PASSWORD);
            stmt = con.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_UPDATABLE);

            ResultSet rs;
            String[] queries = queriesSb.toString().split(LINE_SEPARATOR);
            for (String query : queries) {
                try {
                    if (query.toUpperCase().trim().startsWith("SELECT")) {
                        executeSelect(query);
                    } else if (query.toUpperCase().trim().startsWith("UPDATE")) {
                        executeUpdate(query);
                    } else if (query.toUpperCase().trim().startsWith("INSERT")) {
                        executeInsert(query);
                    } else if (query.toUpperCase().trim().startsWith("DELETE")) {
                        executeDelete(query);
                    } else {
                        System.out.print("Invalid or not supported: " + query);
                    }
                } catch (SQLException e) {
                    System.out.println(query);
                    System.out.print(e.getMessage());
                }
                System.out.println(LINE_SEPARATOR);
            }
        } catch (ClassNotFoundException e) {
            System.err.println("Database driver was not found");
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            try {
                stmt.close();
                con.close();
            } catch (Exception e) {
            }
        }
    }

    private static void executeSelect(String query) throws SQLException {
        ResultSet rs;
        rs = stmt.executeQuery(query);
        ResultSetMetaData rsmd = rs.getMetaData();

        int totalColumns = rsmd.getColumnCount();

        int totalRows = 0;
        if (rs.last()) {
            totalRows = rs.getRow();
            rs.beforeFirst();
        }

        for (int r = 1; r <= totalColumns; r++) {
            System.out.print(rsmd.getColumnName(r));
            if (r != totalColumns) {
                System.out.print(COLUMN_SEPARATOR);
            } else {
                System.out.print(LINE_SEPARATOR);
            }
        }

        if (totalRows > 0) {
            int row = 0;
            while (rs.next()) {
                row++;
                for (int c = 1; c <= totalColumns; c++) {
                    String columnValue = rs.getString(c);
                    System.out.print(columnValue);
                    if (c != totalColumns) {
                        System.out.print(COLUMN_SEPARATOR);
                    } else if (row < totalRows) {
                        System.out.print(LINE_SEPARATOR);
                    }
                }
            }
        } else {
            System.out.print("No rows found");
        }
    }

    private static void executeUpdate(String query) throws SQLException {
        int rowsUpdated = stmt.executeUpdate(query);
        System.out.println(query);
        System.out.print("Rows updated: " + rowsUpdated);
    }

    private static void executeInsert(String query) throws SQLException {
        int rowsInserted = stmt.executeUpdate(query);
        System.out.println(query);
        System.out.print("Rows inserted: " + rowsInserted);
    }

    private static void executeDelete(String query) throws SQLException {
        int rowsDeleted = stmt.executeUpdate(query);
        System.out.println(query);
        System.out.print("Rows deleted: " + rowsDeleted);
    }

    private static StringBuffer readFile(String fileName) throws IOException {
        BufferedReader br = new BufferedReader(new FileReader(fileName));
        try {
            StringBuffer sb = new StringBuffer();
            String line = br.readLine();

            while (line != null) {
                sb.append(line);
                sb.append(LINE_SEPARATOR);
                line = br.readLine();
            }
            return sb;
        } finally {
            br.close();
        }
    }
}