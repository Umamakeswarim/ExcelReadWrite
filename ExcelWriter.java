package excel.exceldemo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {

    public static void main(String[] args) {
        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();

        // Create a blank sheet
        Sheet sheet = workbook.createSheet("Sheet1");

        // Create some data rows
        Row headerRow = sheet.createRow(0);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue("Name");

        Row dataRow = sheet.createRow(1);
        Cell dataCell = dataRow.createCell(0);
        dataCell.setCellValue("John Doe");

        // Write the workbook to a file
        try (FileOutputStream fileOut = new FileOutputStream("C:\\Users\\elcot\\eclipse-workspace\\exceldemo\\ExcelFiles\\Book2.xlsx")) {
            workbook.write(fileOut);
            System.out.println("Excel file has been created successfully.");
        } catch (IOException e) {
            System.out.println("Error writing Excel file: " + e.getMessage());
        } finally {
            // Close the workbook
            try {
                workbook.close();
            } catch (IOException e) {
                System.out.println("Error closing workbook: " + e.getMessage());
            }
        }
    }
}