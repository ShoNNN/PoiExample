import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class CreateWorkBook {

    static FileOutputStream fileOutputStream;
    static FileInputStream fileInputStream;
    static String filePath = "Example.xlsx";
    static String sheetName = "test";

    public static void main(String[] args) {

        try {
            fileOutputStream = new FileOutputStream(new File(filePath));
            fileInputStream = new FileInputStream(new File(filePath));
            write(fileOutputStream);
            read(fileInputStream);

        } catch (IOException ex) {
            ex.printStackTrace();
        }

    }

    public static void write(FileOutputStream filepath) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet(sheetName);

        XSSFRow row = sheet.createRow(0);

        XSSFCell cell = row.createCell(0);

        cell.setCellValue("Hello World");

        workbook.write(filepath);

        workbook.close();

    }

    public static void read(FileInputStream filepath) throws  IOException {

        XSSFWorkbook workbook = new XSSFWorkbook(filepath);

        XSSFSheet sheet = workbook.getSheet(sheetName);

        XSSFRow row = sheet.getRow(0);

        XSSFCell cell = row.getCell(0);

        System.out.println(cell.getStringCellValue());

    }

}
