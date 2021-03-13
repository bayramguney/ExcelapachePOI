package exceloperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class WriteAndRead {
    public static void main(String[] args) throws IOException {


        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Emp Info");

        Object empdata[][] = {{"EmpID", "Name", "Job"},
                {101, "David", "Engineer"},
                {102, "Smith", "Manager"},
                {103, "Scott", "Analyst"}
        };

        int rowCount = 0;

        for (Object emp[] : empdata) {
            XSSFRow row = sheet.createRow(rowCount++);
            int columnCount = 0;
            for (Object value : emp) {
                XSSFCell cell = row.createCell(columnCount++);

                if (value instanceof String)
                    cell.setCellValue((String) value);
                if (value instanceof Integer)
                    cell.setCellValue((Integer) value);
                if (value instanceof Boolean)
                    cell.setCellValue((Boolean) value);

            }
        }

        String filePath = "datafiles/Test.xlsx";
        FileOutputStream outstream = new FileOutputStream(filePath);
        workbook.write(outstream);

        outstream.close();

        System.out.println("Employee.xls file written successfully...");

        System.out.println("----------------------------------------");
        System.out.println("Excel Reading Part----------------------");


        String excelFilePath="datafiles/Test.xlsx";
        FileInputStream inputstream=new FileInputStream(excelFilePath);

        XSSFWorkbook workbookRead=new XSSFWorkbook(inputstream);
        XSSFSheet sheetRead=workbookRead.getSheet("Emp Info");

        Iterator iterator=sheet.iterator();

        while(iterator.hasNext())
        {
            XSSFRow row=(XSSFRow) iterator.next();

            Iterator cellIterator=row.cellIterator();

            while(cellIterator.hasNext())
            {
                XSSFCell cell=(XSSFCell) cellIterator.next();

                switch(cell.getCellType())
                {
                    case STRING: System.out.print(cell.getStringCellValue()); break;
                    case NUMERIC: System.out.print(cell.getNumericCellValue());break;
                    case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
                }
                System.out.print(" |  ");
            }
            System.out.println();
        }

        inputstream.close();
    }





    }
