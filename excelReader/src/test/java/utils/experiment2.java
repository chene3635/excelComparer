package utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class experiment2 {
	public static void main (String args[]) throws IOException {

        FileInputStream file = new FileInputStream("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\test_zip.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(file);
        XSSFSheet sh = wb.getSheet("Test_ZIP");
  
        HashMap<String, String> map = new HashMap<String, String>();
  
        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            String key = sh.getRow(r).getCell(0).getStringCellValue();
            String value = sh.getRow(r).getCell(0).getStringCellValue();
            map.put(key, value);
        }
        
        FileInputStream file2 = new FileInputStream("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\test_zip2.xlsx");
        XSSFWorkbook wb2 = new XSSFWorkbook(file2);
        XSSFSheet sh2 = wb2.getSheet("Test_ZIP");
        
        HashMap<String,String> map2 = new HashMap<String,String>();
        
        for(int u = 1; u <= sh2.getLastRowNum(); u++) {
        	String key2 = sh2.getRow(u).getCell(0).getStringCellValue();
        	String value2 = sh2.getRow(u).getCell(0).getStringCellValue();
        	map2.put(key2, value2);
        }
        
        
        //compares the two hashmaps and stores it into a third one
        
        HashMap<String, String> map3 = new HashMap<String,String>();
        
        for (String key: map.keySet()) {
        	if(map2.containsKey(key)){
        		if(map.get(key).equals(map2.get(key))) {
            		//System.out.println("Matched " + key);
            		//String keys = key;
            		String value = "Matched";
            		map3.put(key, value);
            		System.out.println("Matched : "+key);
            	}
            	else {
            		map3.put(key, map.get(key)+"|"+map2.get(key));
            		System.out.println("Not matched "+ map.get(key)+"|"+map2.get(key));
            	}
        	}else {
        		map3.put(key, map.get(key)+"|"+"Value is NOT present in file2");
        	}
        	
       }
        
       for(String key2: map2.keySet() ) {
    	   if(map.containsKey(key2)) {
    		  if(map2.get(key2).equals(map.get(key2))){
    			  String value = "Matched";
    			  map3.put(key2, value);
    			  System.out.println("Matched : " + key2);
    		  }
    	   } else {
    		   map3.put(key2, "Value is NOT present in file1" + "|" + map2.get(key2));
    		   
    	   }
       }
        
        System.out.println("***********************************************");
        
        for(String strKey : map3.keySet()) {
        	System.out.println("Key : "+strKey+", Value : "+map3.get(strKey));
        }
        
        
//        Iterator<Entry<String, String>> new_Iterator = map3.entrySet().iterator();
//        
//        while (new_Iterator.hasNext()) {
//            Map.Entry<String, String> new_Map = (Map.Entry<String, String>) new_Iterator.next();
//            //System.out.println(new_Map.getKey() + "|" + new_Map.getValue());
//        } 
        
        /*HashSet<String> combinedKeys = new HashSet<String>(map.keySet());
        combinedKeys.addAll(map2.keySet());
        combinedKeys.removeAll(map.keySet());*/
        
        wb.close();
        file.close();
        wb2.close();
        file2.close();
        
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("matchedCapitals");
        
        int rowno = 0;
        for(Map.Entry entry:map3.entrySet()) {
        	XSSFRow row = sheet.createRow(rowno++);
        	row.createCell(0).setCellValue((String)entry.getKey());
        	row.createCell(1).setCellValue((String)entry.getValue());
        	
        }
        
        FileOutputStream finalFile = new FileOutputStream(".\\data\\matchedCapital.xlsx");
        workbook.write(finalFile);
        finalFile.close();
        workbook.close();
       
	}

}

