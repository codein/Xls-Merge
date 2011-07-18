package edu.binghamton.cs.g6.tools.xlssync;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;


/**
 * 
 * This class has static methods implemented to Write to a xls file.
 * @author E0F295
 *
 */
public class XlsWriter {

	/**
	 * Given a {@link HashMap} and {@link String} fileName
	 * This method renders the given {@link HashMap} parameter into a .xls file with the specified filename
	 * @param xlsSheet this is the in memory representation of the .xls file to be rendered
	 * @param fileName the file name in which the .xls file needs to be rendered 
	 * @throws IOException 
	 */

	
	public static void writeToXls(HashMap<String,ArrayList<Cell>> xlsSheet,String fileName) throws IOException
	{
		
		   HSSFWorkbook workbook = new HSSFWorkbook();
		   Sheet sheet = workbook.createSheet();
		
		   Set entries = xlsSheet.entrySet();
		    Iterator it = entries.iterator();
		    int rowNumber=0;
		    while (it.hasNext()) {
			      Map.Entry entry = (Map.Entry) it.next();
			      
			      
			      ArrayList<Cell> xlsRow1 = (ArrayList<Cell>) entry.getValue();
			      Row row = sheet.createRow(rowNumber++);
			      Cell cell = null;
			      for(int a=0;a<xlsRow1.size();a++)
			      {
			    	  
			    	  cell = xlsRow1.get(a);
			    	  Cell newCell = row.createCell(a);
			    	  if(cell!=null)
			    	  {
			    		  if(a==0&&cell.toString().length()>0)
			    		  {
			    			  newCell.setCellValue(cell.toString().substring(0,cell.toString().length()-2));
			    		  }
			    		  else
			    		  {
			    			  newCell.setCellValue(cell.toString());
			    		  }
			    		  
			    	  }
			    	  else
			    	  {
			    		  newCell.setCellValue("");
			    	  }
			      }
		    }
		    
		 // Write the output to a file
	        FileOutputStream fileOut = new FileOutputStream(fileName);
	        workbook.write(fileOut);
	        fileOut.close();
		    
		    
	}
	
//	public static void writeToXls(HashMap<String,ArrayList<String>> xlsSheet, String fileName) throws IOException
//	{
//		
//		   HSSFWorkbook workbook = new HSSFWorkbook();
//		   Sheet sheet = workbook.createSheet();
//		
//		   Set entries = xlsSheet.entrySet();
//		    Iterator it = entries.iterator();
//		    int rowNumber=0;
//		    while (it.hasNext()) {
//			      Map.Entry entry = (Map.Entry) it.next();
//			      
//			      
//			      ArrayList<String> xlsRow1 = (ArrayList<String>) entry.getValue();
//			      Row row = sheet.createRow(rowNumber++);
//			      String cell = null;
//			      for(int a=0;a<xlsRow1.size();a++)
//			      {
//			    	  
//			    	  cell = xlsRow1.get(a);
//			    	  Cell newCell = row.createCell(a);
//			    	  if(cell!=null)
//			    	  {
//
//			    			  newCell.setCellValue(cell.toString());
//			    		  
//			    		  
//			    	  }
//			    	  else
//			    	  {
//			    		  newCell.setCellValue("");
//			    	  }
//			      }
//		    }
//		    
//		 // Write the output to a file
//	        FileOutputStream fileOut = new FileOutputStream(fileName);
//	        workbook.write(fileOut);
//	        fileOut.close();
//		    
//		    
//	}

}
