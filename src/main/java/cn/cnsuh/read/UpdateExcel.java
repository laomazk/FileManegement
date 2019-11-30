package cn.cnsuh.read;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;


/**
 * @author Magic
 * @create 2019-11-30 21:05
 */
public class UpdateExcel {
    private static String filePath = "d:/work/";

    //private static String sheetName = "Sheet0";
    public static void update(String sheetName, String number, int grade) throws Exception {
        //Scanner sc = new Scanner(System.in);
        FileInputStream fis = new FileInputStream("d:/work/成绩表.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        //String sheetName = sc.next();
        XSSFSheet sheet = wb.getSheet(sheetName);
        for (Row row : sheet) {
            if (row.getCell(0).toString().equals("")) {
                return;
            }
            Cell cell = row.getCell(0);
            String id = cell.getStringCellValue();
            if (id.equals(number)) {
                String name = row.getCell(2).getStringCellValue();
                XSSFRow sheetRow = sheet.createRow(row.getRowNum());
                sheetRow.createCell(0).setCellValue(number);
                sheetRow.createCell(1).setCellValue(name);
                sheetRow.createCell(2).setCellValue(grade);
                break;
            }
        }
        FileOutputStream fileOut = new FileOutputStream("d:/work/成绩表.xlsx");
        wb.write(fileOut);
        fileOut.close();
        wb.close();
    }
}
