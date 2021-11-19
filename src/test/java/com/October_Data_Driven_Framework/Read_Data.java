package com.October_Data_Driven_Framework;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_Data {
	
// ---> To get Particular Data
	
	public static void Particular_Data() throws IOException {         
		
		File f = new File
        ("C:\\Users\\M.Rajkamal\\apache-maven-3.8.3\\apache-maven-3.8.3\\bin\\October_Data_Driven_Framework\\User_Data.xlsx");
	
	    FileInputStream fis = new FileInputStream(f);
	
	    Workbook w = new XSSFWorkbook(fis);  // ---> up casting 
	 
	    Sheet sheetAt = w.getSheetAt(0);
	    
	    Row row = sheetAt.getRow(3);
	    
	    Cell cell = row.getCell(0);
	    
	    CellType cellType = cell.getCellType();
	    
	    
	    if (cellType.equals(CellType.STRING)) {  // ---> refname.equals(Enum.String)
	    	
	    	String stringCellValue = cell.getStringCellValue();
	    
		    System.out.println(stringCellValue);
	    }
        
	    else if (cellType.equals(CellType.NUMERIC)) {
	    	
	    	double numericCellValue = cell.getNumericCellValue();
	    	
	    	int value = (int) numericCellValue;
	    	
	    	System.out.println(value);
	    }
	}
	
	
// ---> To get All Data
	
	public static void All_Data() throws IOException {
		
		File f = new File
		         ("C:\\Users\\M.Rajkamal\\apache-maven-3.8.3\\apache-maven-3.8.3\\bin\\October_Data_Driven_Framework\\User_Data.xlsx");
			
        FileInputStream fis = new FileInputStream(f);
			
	    Workbook w = new XSSFWorkbook(fis);  // ---> up casting 
			 
	    Sheet sheetAt = w.getSheetAt(0);
		
	    int numberOfRows = sheetAt.getPhysicalNumberOfRows();
	    
	    for (int i = 0; i < numberOfRows; i++) {
	    	
	    	Row row = sheetAt.getRow(i);
	    	
	    	int numberofCells = row.getPhysicalNumberOfCells();
	    	
	    	for (int j = 0; j < numberofCells; j++) {
	    		
	    		Cell cell = row.getCell(j);
	    		
	    		CellType cellType = cell.getCellType();
	    		

	    	    if (cellType.equals(CellType.STRING)) {  // ---> refname.equals(Enum.String)
	    	    	
	    	    	String stringCellValue = cell.getStringCellValue();
	    	    
	    		    System.out.print(stringCellValue+"   ");
	    		    
	    	    }
	
	    	    else if (cellType.equals(CellType.NUMERIC)) {
	    	    	
	    	    	double numericCellValue = cell.getNumericCellValue();
	    	    	
	    	    	int value = (int) numericCellValue;
	    	    	
	    	    	System.out.print(value);
	    	    }
	    	    
			}
	    	
	    	System.out.println();
				
		}
			
    }
	
// ---> To get Particular Row	
	
    public static void Particular_Row() throws IOException {
		
		File f = new File
		         ("C:\\Users\\M.Rajkamal\\apache-maven-3.8.3\\apache-maven-3.8.3\\bin\\October_Data_Driven_Framework\\User_Data.xlsx");
			
        FileInputStream fis = new FileInputStream(f);
			
	    Workbook w = new XSSFWorkbook(fis);  // ---> up casting 
			 
	    Sheet sheetAt = w.getSheetAt(0);
	    
        Row row = sheetAt.getRow(2);
	    
        int numberOfCells = row.getPhysicalNumberOfCells();
        
        for (int i = 0; i < numberOfCells; i++) {
        	
        	Cell cell = row.getCell(i);
        	
        	CellType type = cell.getCellType();
        	
        	if (type.equals(CellType.STRING)) {
        		
        		String stringCellValue = cell.getStringCellValue();
        		System.out.print(stringCellValue+ "");
        		
			}
        	
        	else if(type.equals(CellType.NUMERIC)) {
        		
        		double numericCellValue = cell.getNumericCellValue();
    	    	
    	    	int value = (int) numericCellValue;
    	    	System.out.println(value); 	
			}
		}
    }
    
// ---> To get Particular Column
    
    public static void Particular_Column() throws IOException {
		
		File f = new File
		         ("C:\\Users\\M.Rajkamal\\apache-maven-3.8.3\\apache-maven-3.8.3\\bin\\October_Data_Driven_Framework\\User_Data.xlsx");
			
        FileInputStream fis = new FileInputStream(f);
			
	    Workbook w = new XSSFWorkbook(fis);  // ---> up casting 
			 
	    Sheet sheetAt = w.getSheetAt(0);
	    
	    int numberOfRows = sheetAt.getPhysicalNumberOfRows();
        
        for (int i = 0; i < numberOfRows; i++) {
        	
        	Row row = sheetAt.getRow(i);
        	
        	Cell cell = row.getCell(0);
        	
        	CellType type = cell.getCellType();
        	
        	if (type.equals(CellType.STRING)) {
        		
        		String stringCellValue = cell.getStringCellValue();
        		System.out.println(stringCellValue+ "");
        		
			}
        	
        	else if(type.equals(CellType.NUMERIC)) {
        		
        		double numericCellValue = cell.getNumericCellValue();
    	    	
    	    	int value = (int) numericCellValue;
    	    	System.out.println(value); 			
			}
		}
    }

	public static void main(String[] args) throws IOException {
		
		System.out.println("*****************Particular_Data*****************");
		System.out.println();
		
		Particular_Data();
		
		System.out.println();
		System.out.println("*********************All_Data********************");
		System.out.println();
		
		All_Data();
		
		System.out.println();
		System.out.println("*****************Particular_Row***********8******");
		System.out.println();
		
		Particular_Row();
		
		System.out.println();
		System.out.println();
		System.out.println("*****************Particular_Column****************");
		System.out.println();
		
		Particular_Column();
		
	}
	
}
