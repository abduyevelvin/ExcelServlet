import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSetMetaData;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class LoadDataExcelFromOracl {

    Connection conn = null;
    Statement stmt = null;
    ResultSet rs = null;

    private Connection ConnectOraclDatabse(String aUserName, String aPassword, String Driver, String Url) throws ClassNotFoundException, SQLException {

        String driver = Driver;
        String url = Url;
        String username = aUserName;
        String password = aPassword;

        Class.forName(driver);
        conn = DriverManager.getConnection(url, username, password);
        return conn;
    }

    public String LoadDataExcel(String aQuery, String aSheetName, String UserName, String Password, String Driver, String Url, String outFolder, String outFileName) {

        long i = 1;

        try {
            //conn = ConnectOraclDatabse("test_user", "123", "oracle.jdbc.driver.OracleDriver", "jdbc:oracle:thin:@localhost:1521:orcl");
            conn = ConnectOraclDatabse(UserName, Password, Driver, Url);
            System.out.println("conn=" + conn);
            String query = aQuery;

            stmt = conn.createStatement();
            rs = stmt.executeQuery(query);
            stmt.setFetchSize(1000);

            SXSSFWorkbook workbook = new SXSSFWorkbook(-1);
            Sheet sheet = workbook.createSheet();

            ResultSetMetaData rsMetaData = rs.getMetaData();
            int numberOfColumns = rsMetaData.getColumnCount();
            //XSSFRow rowhead = sheet.createRow((int) 0);
            Row rowhead = sheet.createRow(0);

            String columnName = "";
            for (int j = 1; j < numberOfColumns + 1; j++) {
                columnName = rsMetaData.getColumnName(j);
                rowhead.createCell((short) j - 1).setCellValue(columnName);

            }

            Row row;
            int cnt = 1;
            while (rs.next()) {
                row = sheet.createRow((int) i);
                for (int j = 1; j < numberOfColumns + 1; j++) {
                    row.createCell((int) j - 1).setCellValue(rs.getString(j));
                }
                i = i + 1;
                System.out.println(cnt++);

                if (i % 100 == 0) {
                    ((SXSSFSheet) sheet).flushRows(100);
                }

            }

            try {
                FileOutputStream out = new FileOutputStream(new File(outFolder + outFileName));
                workbook.write(out);

                out.close();
                System.out.println("Excel written successfully..");

            } catch (FileNotFoundException e) {
                e.printStackTrace();
                System.out.println("Excel not written successfully..");
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("Excel not written successfully..");
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                rs.close();
                stmt.close();
                conn.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }

        return outFileName;

    }
}
