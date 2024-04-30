package excel.exceldemo;

import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExceldataWriter {
    public static void main(String[] args) {
        String[] columnHeaders = {"Name", "Age", "Email"};
        Object[][] data = {
            {"John Doe", 30, "john@test.com"},
            {"Jane Doe", 28, "jane@test.com"},
            {"Bob Smith", 35, "bob@example.com"},
            {"Swapnil", 37, "swapnil@example.com"}
        };

        String filePath = "C:\\Users\\elcot\\eclipse-workspace\\exceldemo\\ExcelFiles\\Sheet2.xlsx";

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet2");

            // Writing column headers
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < columnHeaders.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columnHeaders[i]);
            }

            // Writing data rows
            for (int rowIndex = 0; rowIndex < data.length; rowIndex++) {
                Row row = sheet.createRow(rowIndex + 1);
                for (int colIndex = 0; colIndex < data[rowIndex].length; colIndex++) {
                    Cell cell = row.createCell(colIndex);
                    if (data[rowIndex][colIndex] instanceof String) {
                        cell.setCellValue((String) data[rowIndex][colIndex]);
                    } else if (data[rowIndex][colIndex] instanceof Integer) {
                        cell.setCellValue((Integer) data[rowIndex][colIndex]);
                    }
                }
            }

            // Writing to file
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
                System.out.println("Excel file has been created successfully!");
            }

        } catch (Exception e) {
            System.out.println("An error occurred: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
