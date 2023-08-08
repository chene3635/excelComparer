package utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class experiment3 {
	
	public static void main (String args[]) throws IOException {
  
        	HashMap<String, ArrayList<String>> map = new HashMap<String, ArrayList<String>>();
        	String key = null;
        	ArrayList<String> value = null;
       
        
        	 FileInputStream file = new FileInputStream("C:\\Users\\EChen\\eclipse-workspace\\excelReader\\data\\test_zip.xlsx");
             XSSFWorkbook wb = new XSSFWorkbook(file);
             XSSFSheet sh = wb.getSheet("Test_ZIP");
             
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
             
        HashMap<String, ArrayList<String>> map2 = new HashMap<String, ArrayList<String>>();
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
       	        map2.put(key2, value2);
       	        key2 = null;
       	        value2 = null;
       	    }
       	}
        
        
        
        
        //compares the two hashmaps and stores it into a third one
        
        HashMap<String, String> map3 = new HashMap<String,String>();
        
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
        		map3.put(key3,strResult);
        		map2.remove(key3);
        	} else {
         	   map3.put(key3, key3+" - is NOT present in file2");
         	   
        	}
       }
       
       
       for(String Key4: map2.keySet()) {
    	   map3.put(Key4,Key4+" - is NOT present in File1");
       }
        		/* if(map.get(key3).equals(map2.get(key3))) {
            		//System.out.println("Matched " + key);
            		//String keys = key;
            		String value3 = "Matched";
            		map3.put(key3, value3);
            		System.out.println("Matched : "+key3);
            	}
            	else {
            		map3.put(key3, map.get(key3)+"|"+map2.get(key3));
            		System.out.println("Not matched "+ map.get(key3)+"|"+map2.get(key3));
            	}
        	}else {
        		map3.put(key3, map.get(key3)+"|"+"Value is NOT present in file2");
        		
        	} */
        	
       
      
       /*for(String key4: map2.keySet() ) {
    	   if(map.containsKey(key4)) {
    		   int file2ValueSize = map.get(key4).size();
    		   for(int i = 0; i < file2ValueSize; i++) {
    			   if(map.get(key4).get(i).equals(map2.get(key4).get(i))) {
    				   String value4 = "Matched";
    	    			  map3.put(key4, value4);
    	    			  //System.out.println("Matched : " + key4);
       			} else {
       				map3.put(key4, map.get(key4).get(i)+"|"+map2.get(key4).get(i));
            		//System.out.println("Not matched "+ map.get(key4)+"|"+map2.get(key4));
       			}
    			   
    		 }
 
       } else {
    	   map3.put(key4, key4+" - is NOT present in file2");
       }
       }*/
       
        /*for(String key4: map2.keySet() ) {
 	   if(map.containsKey(key4)) {
 		  if(map2.get(key4).equals(map.get(key4))){
 			  String value4 = "Matched";
 			  map3.put(key4, value4);
 			  System.out.println("Matched : " + key4);
 		  }
 	   } else {
 		   map3.put(key2, "Value is NOT present in file1" + "|" + map2.get(key));
 		   
 	   }
    } */
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
        
        wb.close();
        file.close();
        wb2.close();
        file2.close();
        
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("matchedAmenities");
        
        
        int rowno = 0;
       
        for (String strKey: map3.keySet()) {
        	
        	XSSFRow row = sheet.createRow(rowno++);
        	row.createCell(0).setCellValue(strKey);
        	int ColoumSize = map3.get(strKey).split(";").length;
        	for (int i = 0; i < ColoumSize; i++) {
        		row.createCell(i+1).setCellValue(map3.get(strKey).split(";")[i]);
        	}
        }
        
        FileOutputStream finalFile = new FileOutputStream(".\\data\\matchedGasStation.xlsx");
        workbook.write(finalFile);
        finalFile.close();
        workbook.close(); 
       
}
}
