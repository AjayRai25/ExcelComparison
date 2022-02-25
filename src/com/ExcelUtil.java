package com;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;


public class ExcelUtil {
    static DataFormatter df = new DataFormatter();
    public static String getCellValue(Cell cell) {
        String cellValue;
        if(cell==null){
            return null;//undefined cell with no data
        }
        if(cell.getCellType()== CellType.NUMERIC){
            cellValue = df.formatCellValue(cell);
        }
        else if (cell.getCellType() == CellType.BLANK) {
            cellValue = df.formatCellValue(cell);
        }
        else if (cell.getCellType() == CellType.BOOLEAN) {
            cellValue = df.formatCellValue(cell);
        }
        else{
            cellValue = (cell.getStringCellValue());
        }
        return cellValue;
    }
}
