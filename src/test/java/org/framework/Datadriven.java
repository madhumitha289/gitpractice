package org.framework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.helper.DataUtil;



public class Datadriven {
	public static void main(String[] args) throws IOException {
		 File f = new File("C:\\Users\\pc\\Videos\\Selenium-Java Course (16 feb 2022)\\jdk and eclipse software\\eclipse\\FrameworkPractice\\target\\Book1.xlsx");
		 FileInputStream Stream = new FileInputStream(f);
		 Workbook w = new XSSFWorkbook(Stream);
		 Sheet s = w.getSheet("Sheet1");
		 int physicalNumberOfRows = s.getPhysicalNumberOfRows();
		 for(int i=0; i<physicalNumberOfRows; i++)
		 {
			 Row row = s.getRow(i);
			 int physicalNumberOfCells = row.getPhysicalNumberOfCells();
		     for(int j=0; j<physicalNumberOfCells; j++)
		     {
		      Cell cell= row.getCell(j);
		      int cellType= cell.getCellType();
		      if(cellType==1)
		      {
		    	  String stringCellValue= cell.getStringCellValue();
		          System.out.println(stringCellValue);
		      }
		      else if(DateUtil.isCellDateFormatted(cell))
		      {
		    	  Date date = cell.getDateCellValue();
		          System.out.println(date);
		          SimpleDateFormat s1= new SimpleDateFormat("DD/MM/YYYY");
		          String format = s1.format(date);
		          System.out.println(format);
		      }
		      else
		      {
		    	  double numericCell1= cell.getNumericCellValue();
		    	  System.out.println(numericCell1);
		    	  long l =(long)numericCell1;
		    	  System.out.println(l);  
		      }
		     }
		 }		 
	}
}
