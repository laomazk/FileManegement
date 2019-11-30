package cn.cnsuh.read;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import static org.apache.poi.ss.usermodel.CellType.STRING;

/**
 * @author Magic
 * @create 2019-11-30 21:18
 */
public class CopyExcel {
    private static String filePath = "d:/work/";
    public static void copy(String sheetName, String copyFileName) throws Exception {
        FileInputStream fis = new FileInputStream("d:/work/成绩表.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheet(sheetName);

        FileInputStream copyFis = new FileInputStream(filePath + copyFileName + ".xlsx");
        XSSFWorkbook copyWb = new XSSFWorkbook(copyFis);
        XSSFSheet copySheet = copyWb.getSheet("Sheet1");

        for (Row row : sheet) {
            //如果当前行没有数据，跳出循环
            if (row.getCell(0).toString().equals("")) {
                return;
            }
            int end = row.getLastCellNum();
            int copyEnd = copySheet.getLastRowNum() ;
            XSSFRow copySheetRow = copySheet.createRow(copyEnd);
            for (int i = 0; i < end; i++) {
                Cell cell = row.getCell(i);
                Object obj;
                if (cell.getCellTypeEnum() == STRING) {
                    obj = cell.getStringCellValue();
                    copySheetRow.createCell(i).setCellValue((String)obj);
                } else {
                    obj = cell.getNumericCellValue();
                    copySheetRow.createCell(i).setCellValue((Double) obj);
                }
            }
        }
        FileOutputStream fileOut = new FileOutputStream(filePath + copyFileName + ".xlsx");
        copyWb.write(fileOut);
        fileOut.close();
        copyWb.close();
        wb.close();
    }
}
