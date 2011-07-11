package edu.binghamton.cs.g6.tools.xlssync;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;

public class XlsMerger {
	
	public static void main(String[] args) throws IOException 
	{
		XlsReader xlsReader1 = new XlsReader("C:/Robin local/old/A 1.xls");
		 HashMap<String,ArrayList<Cell>> file1 = xlsReader1.getXlsSheet();
		 
		 XlsReader xlsReader2 = new XlsReader("C:/Robin local/old/A 2.xls");
		 HashMap<String,ArrayList<Cell>> file2 = xlsReader2.getXlsSheet();
		
		 Set entries = file1.entrySet();
		    Iterator it = entries.iterator();
		    while (it.hasNext()) {
		    	
		      Map.Entry entry = (Map.Entry) it.next();
		      
		      
		      ArrayList<Cell> xlsRow1 = (ArrayList<Cell>) entry.getValue();
		      ArrayList<Cell> xlsRow2 = file2.get(entry.getKey());
		      System.out.println(entry.getKey() + "-->");
		      System.out.println(xlsRow1);
		      System.out.println(xlsRow2);
		      
		      ArrayList<Cell> xlsRowNew = new ArrayList<Cell>();
//		      for(Cell cell1:xlsRow1)
//		      {
//		    	  if(cell1!=null)
//		    	  {
//		    		  System.out.println(cell1.getColumnIndex());
//		    	  }
//		    	 
//		      }
//		      
//		      for(Cell cell1:xlsRow2)
//		      {
//		    	  if(cell1!=null)
//		    	  {
//		    		  System.out.println(cell1.getColumnIndex());
//		    	  }
//		      }
		      
		    } 
		
	}

}
