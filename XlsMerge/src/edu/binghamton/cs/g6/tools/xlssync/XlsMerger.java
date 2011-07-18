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
//		XlsReader xlsReader1 = new XlsReader("C:/Robin local/old/A 1.xls");
		XlsReader xlsReader1 = new XlsReader("c:/Robin local/gap analysis/merge/Chris.xls");
		 HashMap<String,ArrayList<Cell>> file1 = xlsReader1.getXlsSheet();
		 
//		 XlsReader xlsReader2 = new XlsReader("C:/Robin local/old/A 2.xls");
		XlsReader xlsReader2 = new XlsReader("c:/Robin local/gap analysis/merge/clinical.xls");
		 HashMap<String,ArrayList<Cell>> file2 = xlsReader2.getXlsSheet();
		
		 HashMap<String,ArrayList<Cell>> mergeFile = new  HashMap<String, ArrayList<Cell>>();
		 Set entries = file1.entrySet();
		    Iterator it = entries.iterator();
		    while (it.hasNext()) {
		    	
		      Map.Entry entry = (Map.Entry) it.next();
		      
		      
		      ArrayList<Cell> xlsRow1 = (ArrayList<Cell>) entry.getValue();
		      ArrayList<Cell> xlsRow2 = file2.get(entry.getKey());
		      System.out.println(entry.getKey() + "-->");
		      System.out.println(xlsRow1);
		      System.out.println(xlsRow2);
		      
		      ArrayList<Cell> xlsRowNew = mergeCellArray(xlsRow1, xlsRow2);
		      System.out.println(xlsRowNew);
		      mergeFile.put(xlsRowNew.get(0).toString(), xlsRowNew);
		      
		    } 
		    
		    XlsWriter.writeToXls(mergeFile,"merge.xls");
		
	}
	
	public static ArrayList<Cell> mergeCellArray(ArrayList<Cell> cellArray1,ArrayList<Cell>  cellArray2)
	{
		ArrayList<Cell> cellArrayNew = new ArrayList<Cell>();	
		if(cellArray1.size()!=cellArray2.size()) //if row size does not match
		{
//			System.out.println("Error Cell Array size missmatch at row "+cellArray1.get(0));
//		      System.out.println(cellArray1);
//		      System.out.println(cellArray2);
		      if(cellArray1.size()>cellArray2.size()) //return the largest one
		      {
		    	  return cellArray1;
		      }
		      else
		      {
		    	  return cellArray2;
		      }
		}
		else // if row size match then 
		{
			for(int a = 0;a<cellArray1.size();a++) //compare each cell and select the best
			{
				Cell newCell = compare(cellArray1.get(a), cellArray2.get(a));
				cellArrayNew.add(newCell);
			}
		}
		return cellArrayNew;
	}
	
	public static Cell compare(Cell cell1, Cell cell2)
	{
		if(cell1 == null || cell2 ==null)
		{
			 if(cell2 == null)
			 {
				return cell1;				
			 }
			 else if(cell1 == null)
			 {
				 return cell2;
			 }
		}
		else
		{
			if(cell1.toString().compareTo(cell2.toString())>1)
			{
				return cell1;
			}
			else
			{
				return cell2;
			}
		}
		return null;

	}
	


}
