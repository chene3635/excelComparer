package utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Utilities {

	public static HashMap <String, ArrayList<String>> map1 (String file) throws IOException {
		
		HashMap<String, ArrayList<String>> map = new HashMap<String, ArrayList<String>>();
    	
    	 FileInputStream file1 = new FileInputStream("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\test_zip.xlsx");
         XSSFWorkbook wb = new XSSFWorkbook(file1);
         XSSFSheet sh = wb.getSheet("Test_ZIP");
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
        
        FileInputStream file2 = new FileInputStream("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\test_zip2.xlsx");
        XSSFWorkbook wb2 = new XSSFWorkbook(file2);
        XSSFSheet sh2 = wb2.getSheet("Test_ZIP");
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
	    XSSFSheet sheet = workbook.createSheet("matchedAmenities");   
        for (String strKey: map3.keySet()) {
        	
        	XSSFRow row = sheet.createRow(rowno++);
        	row.createCell(0).setCellValue(strKey);
        	int ColoumSize = map3.get(strKey).split(";").length;
        	for (int i = 0; i < ColoumSize; i++) {
        		row.createCell(i+1).setCellValue(map3.get(strKey).split(";")[i]);
        	}
        }
        
        FileOutputStream finalFile = new FileOutputStream(".\\data\\matchedGasStationeds4.xlsx");
        workbook.write(finalFile);
        finalFile.close();
        workbook.close(); 
	}
	
	public static void main (String args []) throws IOException {
		print(map3(map1("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\test_zip.xlsx"), map2("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\test_zip2.xlsx")));
		System.out.println("Finished");
		
	}
	
}
