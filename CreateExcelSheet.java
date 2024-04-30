package excel.exceldemo;

import java.io.File;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class CreateExcelSheet {

	public static void main(String[] args) throws Exception {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet();
		sheet.createRow(0);
		sheet.getRow(0).createCell(0).setCellValue("Hello");
		sheet.getRow(0).createCell(1).setCellValue("World");
		
		sheet.createRow(1);
		sheet.getRow(1).createCell(0).setCellValue("Uma");
		sheet.getRow(1).createCell(1).setCellValue("Gokul");
		
		File file = new File("C:\\Users\\elcot\\eclipse-workspace\\exceldemo\\ExcelFiles\\CreateFile.xlsx");
		
		workbook.write(file);
		// TODO Auto-generated method stub

	}

}
