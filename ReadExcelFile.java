package excel.exceldemo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;

public class ReadExcelFile {
    public static void main(String[] args) {
        try {
            // Load the Excel file
            FileInputStream file = new FileInputStream(new File("C:\\Users\\elcot\\eclipse-workspace\\exceldemo\\ExcelFiles\\Sheet2.xlsx"));

            // Create Workbook instance holding reference to .xlsx file
            Workbook workbook = WorkbookFactory.create(file);

            // Get the first sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through each row of the sheet
            for (Row row : sheet) {
                // Iterate through each cell of the row
                for (Cell cell : row) {
                    // Check the cell type and print the value accordingly
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        case BLANK:
                            System.out.print("[BLANK]\t");
                            break;
                        default:
                            System.out.print("[UNKNOWN]\t");
                    }
                }
                System.out.println(); // Move to the next line after printing each row
            }
            workbook.close();
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
