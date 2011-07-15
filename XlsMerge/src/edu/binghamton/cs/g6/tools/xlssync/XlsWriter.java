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

public class XlsWriter {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub

	}
	
	public static void writeToXls(HashMap<String,ArrayList<Cell>> xlsSheet) throws IOException
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
	        FileOutputStream fileOut = new FileOutputStream("c:/Robin local/gap analysis/merge/merge.xls");
	        workbook.write(fileOut);
	        fileOut.close();
		    
		    
	}

}
