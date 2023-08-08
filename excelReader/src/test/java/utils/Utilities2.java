package utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Utilities2 {
	
	public static HashMap <String, ArrayList<String>> map1 (String file) throws IOException {
		
		HashMap<String, ArrayList<String>> map = new HashMap<String, ArrayList<String>>();
    	
    	 FileInputStream file1 = new FileInputStream("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\Employee.xlsx");
         XSSFWorkbook wb = new XSSFWorkbook(file1);
         XSSFSheet sh = wb.getSheet("employees");
         
         int rowcount=sh.getLastRowNum();
         for(int i=0;i<rowcount+1;i++) {
             //GET CELL
             Cell cell1 = sh.getRow(i).getCell(0);   
             //SET AS STRING TYPE
             cell1.setCellValue(String);
         }
         int firstRow = 0;
         int lastRow = 0;
         int firstCol = 0;
         int lastCol = 2;
         
         sh.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
         
         String key = null;
     	 ArrayList<String> value = null;
         Iterator<Row> rowIterator = sh.iterator();

         while(rowIterator.hasNext()) {

        	    Row row = rowIterator.next();

        	    // for each row, iterate through each columns
        	    Iterator<Cell> cellIterator = row.cellIterator(); 
        	    key = null;  
        	    value = new ArrayList<String>();     

        	    while(cellIterator.hasNext()) {
        	        Cell cell = cellIterator.next();
        	        int columnIndex = cell.getColumnIndex();
        	        if(columnIndex == 0) {      	        	
        	            key = cell.getStringCellValue();
        	        } else {
        	            value.add(cell.getStringCellValue());
        	        	
        	        }
        	        
        	    }

        	    if(key != null && value != null) {
        	        map.put(key, value);
        	        key = null;
        	        value = null;
        	    }
        	}
         wb.close();
         file1.close();
		return map; 
         
      
	}
	public static HashMap <String, ArrayList<String>> map2 (String string) throws IOException {
		
		HashMap<String, ArrayList<String>> map = new HashMap<String, ArrayList<String>>();
        String key2 = null;
        ArrayList<String> value2 = null;
        
        FileInputStream file2 = new FileInputStream("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\Employee2.xlsx");
        XSSFWorkbook wb2 = new XSSFWorkbook(file2);
        XSSFSheet sh2 = wb2.getSheet("employees");
        
        int firstRow = 0;
        int lastRow = 0;
        int firstCol = 0;
        int lastCol = 2;
        
        sh2.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        
    	Iterator<Row> rowIterator2 = sh2.iterator();
        
        while(rowIterator2.hasNext()) {

       	    Row row = rowIterator2.next();

       	    // for each row, iterate through each columns
       	    Iterator<Cell> cellIterator = row.cellIterator(); 
       	    key2 = null;  
       	    value2 = new ArrayList<String>();     

       	    while(cellIterator.hasNext()) {
       	    	
       	        Cell cell = cellIterator.next();
       	        int columnIndex2 = cell.getColumnIndex();

       	        if(columnIndex2 == 0) {
       	            key2 = cell.getStringCellValue();
       	        } else {
       	            value2.add(cell.getStringCellValue());
       	            
       	        }
       	        
       	    }

       	    if(key2 != null && value2 != null) {
       	        map.put(key2, value2);
       	        key2 = null;
       	        value2 = null;
       	    }
       	}
        wb2.close();
        file2.close();
        return map;
		
	}
	
	public static HashMap<String, String> map3 (HashMap<String, ArrayList<String>> map, HashMap<String, ArrayList<String>> map2) {
		 HashMap<String, String> mapFinal = new HashMap<String,String>();
		 for (String key3: map.keySet()) {
	        	if(map2.containsKey(key3)){
	        		int file1ValueSize = map.get(key3).size();
	        		String strResult="";
	        		for(int i=0; i<file1ValueSize; i++) {
	        			if(map.get(key3).get(i).equals(map2.get(key3).get(i))) {
	        				String Result = "Matched";
	        				strResult=strResult+";"+Result;
	        				//map3.put(key3, value3);
	        			} else {
	        				String Result = map.get(key3).get(i)+"|"+map2.get(key3).get(i);
	        				strResult=strResult+";"+Result;
	        				//map3.put(key3, map.get(key3).get(i)+"|"+map2.get(key3).get(i));
	        			}
	        			
	        		}
	        		strResult = strResult.substring(1);
	        		mapFinal.put(key3,strResult);
	        		map2.remove(key3);
	        	} else {
	         	   mapFinal.put(key3, key3+" - is NOT present in file2");
	         	   
	        	}
	       }
		
		return mapFinal;
		
	}
	
	
	
	public static void print (HashMap<String, String> map3) throws IOException {
		int rowno = 0;
		XSSFWorkbook workbook = new XSSFWorkbook();
	    XSSFSheet sheet = workbook.createSheet("matchedEmployees");   
        for (String strKey: map3.keySet()) {
        	
        	XSSFRow row = sheet.createRow(rowno++);
        	row.createCell(0).setCellValue(strKey);
        	int ColoumSize = map3.get(strKey).split(";").length;
        	for (int i = 0; i < ColoumSize; i++) {
        		row.createCell(i+1).setCellValue(map3.get(strKey).split(";")[i]);
        	}
        }
        
        FileOutputStream finalFile = new FileOutputStream(".\\data\\employeesMatched.xlsx");
        workbook.write(finalFile);
        finalFile.close();
        workbook.close(); 
	}
	
	public static void main (String args []) throws IOException {
		print(map3(map1("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\Employee.xlsx"), map2("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\Employee2.xlsx")));
		System.out.println("Finished");
		
	}
	
}
