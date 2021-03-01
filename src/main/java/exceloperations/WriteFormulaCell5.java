package exceloperations;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteFormulaCell5 {
    public static void main(String[] args) throws IOException {
        String path="datafiles/books.xlsx";

        FileInputStream fis=new FileInputStream(path);

        XSSFWorkbook workbook=new XSSFWorkbook(fis);

        XSSFSheet sheet=workbook.getSheetAt(0);

        sheet.getRow(7).getCell(2).setCellFormula("SUM(C2:C6)");

        fis.close();

        FileOutputStream fos=new FileOutputStream(path);
        workbook.write(fos);

        workbook.close();
        fos.close();

        System.out.println("Done!!!");

    }
    }

