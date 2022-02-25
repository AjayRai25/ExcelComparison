package com;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.sqlite.SQLiteException;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

public class Main {
    private static final String USERNAME = "root";
    private static final String PASSWORD = "root";
    private static final String CONN = "jdbc:mysql://localhost/root";
    private static final String SQLLCONN = "jdbc:sqlite:ExcelComparisonDatabase.db";
    static String firstExcelPath ="D:\\Riversand\\Test1.xlsx";
    static String secondExcelPath ="D:\\Riversand\\Test2.xlsx";

    public static void main(String[] args) throws SQLException, IOException {
        Connection con = DriverManager.getConnection(SQLLCONN);
        Statement statement = con.createStatement();
        //Creating table for the Main excel sheet
        getTableHeader(statement,firstExcelPath,"ATTRIBUTES","CompareFromData");
        getTableHeader(statement,secondExcelPath,"ATTRIBUTES","CompareToData");
        populateData(statement);
        con.close();
    }
    public static void getTableHeader(Statement statement, String firstExcelPath, String sheetName, String tableName) throws IOException, SQLException {
        statement.execute("DROP TABLE IF EXISTS "+tableName);
        StringBuilder query = new StringBuilder("CREATE TABLE ExcelOne (");
        FileInputStream excelFile1 = new FileInputStream(firstExcelPath);
        Workbook workbook1 = new XSSFWorkbook(excelFile1);
        Sheet sheet = workbook1.getSheet(sheetName);
        Row attributeHeaderRow;
        attributeHeaderRow = sheet.getRow(0);
        System.out.println(attributeHeaderRow.getLastCellNum());
        for(int j = 0; j< attributeHeaderRow.getLastCellNum(); j++) {
            String attributeId = attributeHeaderRow.getCell(j).getStringCellValue().trim();
            query.append("[").append(attributeId).append("] STRING (50),");
        }
        query = new StringBuilder(query.substring(0, query.length() - 1));
        query.append(");");
        excelFile1.close();
        statement.execute(String.valueOf(query));
    }
    public static void populateData(Statement statement,String firstExcelPath,String sheetName, String tableName) throws IOException, SQLException {
        FileInputStream excelFile1 = new FileInputStream(firstExcelPath);
        Workbook workbook1 = new XSSFWorkbook(excelFile1);
        Sheet sheet = workbook1.getSheet("ATTRIBUTES");
        int rows = sheet.getLastRowNum();
        int headerRow = sheet.getRow(0).getLastCellNum();
        for(int i=1; i<=rows;i++){
            StringBuilder query = new StringBuilder("insert into ExcelOne values( ");
            Row row = sheet.getRow(i);
            System.out.println(headerRow);
            for(int j = 0; j< headerRow; j++) {
                String cellData = ExcelUtil.getCellValue(row.getCell(j));
                query.append("\""+cellData+"\", ");
            }
            query = new StringBuilder(query.substring(0, query.length() - 2)).append(");");
            System.out.println(query);
            try{
                statement.execute(String.valueOf(query));
            }catch (SQLiteException e){
                e.printStackTrace();
            }
        }
    }
}
