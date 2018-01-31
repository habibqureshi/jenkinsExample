

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

public class Main {
    private static final String FILE_NAME = "MyFirstExcel.xlsx";

      public static void main(String[] args) {
          XSSFWorkbook workbook = new XSSFWorkbook();
          XSSFSheet sheet = workbook.createSheet("Datatypes in Java");
          XSSFSheet sheet2 = workbook.createSheet("my second sheet");

          Object[][] datatypes = {
                  {"name", "fname", "city"},
                  {"moeed", "habib", "lahore"},
                  {"adeel", "habib", "lahore"},
                  {"hamshim", "habib", "lahore"},
                  {"lukman", "habib", "lahore"},
                  {"tyab", "habib", "lahore"}
          };

          int rowNum = 0;
          System.out.println("Creating excel");
          Row row,row1;
          Cell cell,cell1;

          for (Object[] datatype : datatypes) {
               row = sheet.createRow(rowNum++);
               row1=sheet2.createRow(rowNum++);

              int colNum = 0;
              for (Object field : datatype) {
                   cell = row.createCell(colNum++);
                   cell1=row1.createCell(colNum++);
                  if (field instanceof String) {
                      cell.setCellValue((String) field);
                      cell1.setCellValue((String)field+"-1234");
                  } else if (field instanceof Integer) {
                      cell.setCellValue((Integer) field);
                  }
              }
          }

          try {
              File file = new File(FILE_NAME);
              FileOutputStream outputStream;
             // if(!file.exists())
                 outputStream = new FileOutputStream(file);
              //else
              workbook.write(outputStream);
              workbook.close();
          } catch (FileNotFoundException e) {
              e.printStackTrace();
          } catch (IOException e) {
              e.printStackTrace();
          }

          System.out.println("Done");
      }


//          try {
//              Class.forName("com.mysql.jdbc.Driver");
//              Connection connection= DriverManager.getConnection("jdbc:mysql://localhost:3306/jubilee_webapp?useUnicode=yes&characterEncoding=UTF-8","root","root");
//              System.out.println("ho gya connection"+connection);
//              PreparedStatement stmt = null;
//              stmt=connection.prepareStatement("select * from policies");
//              ResultSet rs=stmt.executeQuery();
//              while(rs.next()){
//                  for (int i=1;i<=rs.getMetaData().getColumnCount();i++)
//                      System.out.println(rs.getString(i)+"   ");
//                  System.out.println("\n");
//
//              }
//             // System.out.println(
//
//          } catch (SQLException e) {
//              e.printStackTrace();
//
//          } catch (ClassNotFoundException e) {
//              e.printStackTrace();
//          }
//          int i=0;
//
//          System.out.println("i="+i++);
//          System.out.println("i="+i);


}
