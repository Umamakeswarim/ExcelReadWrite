package excel.exceldemo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelSheet {
    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\elcot\\eclipse-workspace\\exceldemo\\ExcelFiles\\Sheet1.xlsx";

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            // Write some data to the sheet
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");
            headerRow.createCell(2).setCellValue("Location");

            Row dataRow1 = sheet.createRow(1);
            dataRow1.createCell(0).setCellValue("John");
            dataRow1.createCell(1).setCellValue(30);
            dataRow1.createCell(2).setCellValue("New York");

            Row dataRow2 = sheet.createRow(2);
            dataRow2.createCell(0).setCellValue("Alice");
            dataRow2.createCell(1).setCellValue(25);
            dataRow2.createCell(2).setCellValue("London");

            // Write the workbook to a file
            try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
                workbook.write(outputStream);
                System.out.println("Excel file created successfully.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
