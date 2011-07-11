package edu.binghamton.cs.g6.tools.xlssync;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;



public class XlsReader {

	
	private HSSFWorkbook wb;
	public static void main(String[] args) throws IOException
	{
	       XlsReader xlsReader = new XlsReader("/home/codein/Desktop/project/xlsMerge/A 1.xls");
	       System.out.print(xlsReader.getXlsSheet());
	       
	}
	
	public XlsReader(String fileName) throws IOException 
	{
		InputStream inp = new FileInputStream(fileName);
	        wb = new HSSFWorkbook(new POIFSFileSystem(inp));
	       
	}
	
	public HashMap<String,ArrayList<Cell>> getXlsSheet()
	{
		 HashMap<String,ArrayList<Cell>> xlsSheet = new HashMap<String,ArrayList<Cell>> ();
	       Sheet sheet = wb.getSheetAt(0);
	       int columnNumber;
	       for (Iterator<Row> rit = sheet.rowIterator(); rit.hasNext();) 
	       {
	    	   ArrayList<Cell> xlsRow = new ArrayList<Cell>();
	    	   Row row = rit.next();
	    	   columnNumber=0;
	    	   
	    	   for (Iterator<Cell> cit = row.cellIterator(); cit.hasNext(); columnNumber++) 
	    	   {
	    		   Cell cell = cit.next();
	    		   if(cell.getColumnIndex()==columnNumber)
	    		   {
	    			   xlsRow.add(cell);	
	    		   }
	    		   else
	    		   {
	    			   xlsRow.add(null);
	    			   columnNumber++;
	    			   xlsRow.add(cell);
	    		   }
	    	   }
	    	   xlsSheet.put(xlsRow.get(0).toString(), xlsRow);	    	   
	       }
	       return xlsSheet;
	}
	
}
