package com.ibm.util.excel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;

public class WriteToExcel {
	
	private static Workbook xlsFile = null;
	private static String fileN = "";
	
	public WriteToExcel (String fileToWrite) {
        File fileH = new File(fileToWrite);
        if (fileH.exists()) {
        	fileH.delete();
        }
        
        fileN = fileToWrite;
        
        if(fileToWrite.endsWith("xlsx")){
            xlsFile = new XSSFWorkbook();
        }else if(fileToWrite.endsWith("xls")){
            xlsFile = new HSSFWorkbook();
        }    
        
	}
	
	public String closeXLS () {
		
		FileOutputStream fileOut;

        try {
        	System.out.println("Closing " + fileN);
            fileOut = new FileOutputStream(fileN);
        	xlsFile.write(fileOut);
        	fileOut.close();
        	System.out.println("Closed " + fileN);
        	return "";
        } catch (IOException i) {
        	return("File " + fileN + " not open.");
        } catch (Exception p) {
        	return(p.toString());
        } 
	}
	
	public void createSheet(String tabName) throws Exception{       
        xlsFile.createSheet(java.net.URLDecoder.decode(tabName, "UTF-8"));
	}
	
	public void writeSheet(String tabName, String[] rowData) throws Exception{
         
		Row row = null;
		Cell cell = null;
		CellStyle cs = null;
		
        Sheet sheet = xlsFile.getSheet(java.net.URLDecoder.decode(tabName, "UTF-8"));
        cs = xlsFile.createCellStyle();
        cs.setWrapText(true);
        int rowCount = sheet.getLastRowNum();
        row = sheet.createRow(rowCount + 1);
        for (int k=0;k < rowData.length;k++) {
        	cell = row.createCell(k);
        	cell.setCellStyle(cs);
        	if (rowData[k] != null) {
        		cell.setCellValue(java.net.URLDecoder.decode(rowData[k], "UTF-8").replace("^", "\n").replace("|", "%"));
        	} else {
        		cell.setCellValue("");
        	}
        }
    }
	
	public void writeHeader(String tabName, String[] rowData) throws Exception{
        
		Row row = null;
		Cell cell = null;
		
        Sheet sheet = xlsFile.getSheet(java.net.URLDecoder.decode(tabName, "UTF-8"));
        row = sheet.createRow(0);
        for (int k=0;k < rowData.length;k++) {
        	cell = row.createCell(k);
        	if (rowData[k] != null) {
        		cell.setCellValue(java.net.URLDecoder.decode(rowData[k], "UTF-8").replace("^", "\n").replace("|", "%"));
        	} else {
        		cell.setCellValue("");
        	}
        }
    }
     

}


