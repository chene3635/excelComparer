package utils;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class excelUtils {

	public static void main (String[] args) throws IOException{
		String excelFilePath = ".\\data\\capitals.xlsx";
		FileInputStream inputStream = new FileInputStream (excelFilePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		//using for loop to read rows and columns
		
	int rows = sheet.getLastRowNum();
	int cols = sheet.getRow(1).getLastCellNum();
	
	for (int r = 0; r <=rows; r++) 
	{ 
		XSSFRow row = sheet.getRow(r);
		for (int c = 0; c < cols; c++) {
			XSSFCell cell = row.getCell(c);
			
			switch (cell.getCellType()) {
			case STRING: System.out.println(cell.getStringCellValue()); break;
			case NUMERIC: System.out.println(cell.getNumericCellValue()); break;
			case BOOLEAN: System.out.println(cell.getBooleanCellValue()); break;
			}
		}
		
		System.out.println();
	}
	
	}
}
