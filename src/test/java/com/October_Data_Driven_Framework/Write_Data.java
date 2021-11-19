package com.October_Data_Driven_Framework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Data {
	
    public static void Write_Data() throws IOException {         
		
		File f = new File
        ("C:\\Users\\M.Rajkamal\\Desktop\\User_Details.xlsx");
	
	    FileInputStream fis = new FileInputStream(f);
	
	    Workbook wb = new XSSFWorkbook(fis);  
	    
        Sheet createSheet = wb.createSheet("Credentials");
                       
        
        
//      Row createRow = createSheet.createRow(0);  -----------------|   
//	                                                                |
//      Cell createCell = createRow.createCell(0);  ----------------|  can also be mentioned as below 
//	                                                                |
//      createCell.setCellValue("Username");  ----------------------|
//                                                                  |	    
//	                                                                V
	 
	   wb.getSheet("Credentials").createRow(1).createCell(1).setCellValue("Name");
	    
	   wb.getSheet("Credentials").getRow(1).createCell(2).setCellValue("Pincode");
	   
//     wb.getSheet("Credentials").createRow(2).createCell(1).setCellValue(" ");
	    
//     wb.getSheet("Credentials").getRow(2).createCell(2).setCellValue(" ");
	   
	   wb.getSheet("Credentials").createRow(3).createCell(1).setCellValue("Kavin");
	   
	   wb.getSheet("Credentials").getRow(3).createCell(2).setCellValue("600002");
	   
//     wb.getSheet("Credentials").createRow(4).createCell(1).setCellValue(" ");
	    
//     wb.getSheet("Credentials").getRow(4).createCell(2).setCellValue(" ");
	   
	   wb.getSheet("Credentials").createRow(5).createCell(1).setCellValue("Kartthik");
	    
	   wb.getSheet("Credentials").getRow(5).createCell(2).setCellValue("607803");
	   
//     wb.getSheet("Credentials").createRow(6).createCell(1).setCellValue(" ");
	    
//     wb.getSheet("Credentials").getRow(6).createCell(2).setCellValue(" ");
	   
	   wb.getSheet("Credentials").createRow(7).createCell(1).setCellValue("Ajith");
	    
	   wb.getSheet("Credentials").getRow(7).createCell(2).setCellValue("600024");
	   
	   
	   FileOutputStream fos = new FileOutputStream(f);
	   
	   wb.write(fos);
	   
	   wb.close();
	   
	   
	   System.out.println("Write_Data Successfull"); 
	   
       }

       public static void main(String[] args) throws IOException {
    	   
    	   Write_Data();
	
      }
}
