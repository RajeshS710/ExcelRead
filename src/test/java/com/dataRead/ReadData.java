package com.dataRead;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	
	public static void main(String[] args) throws IOException {
		
		
		String excel = "C:\\Users\\Admin\\eclipse-workspace\\ExcelRead\\Data\\Data.xlsx";
		
		File f = new File(excel);
		
		FileInputStream file = new FileInputStream(f);
		
		XSSFWorkbook w = new XSSFWorkbook(file);
		
		XSSFSheet sheet = w.getSheet("Sheet1");
		int lastRowNum = sheet.getLastRowNum();
		short cell = sheet.getRow(1).getLastCellNum();

		for(int i = 0; i<lastRowNum; i++) {
			
			
			XSSFRow row = sheet.getRow(i);
			for(int j = 0; j<cell; j++) {
				switch(row.getCell(j).getCellType()) {
				case STRING: System.out.print(row.getCell(j));break;
				case NUMERIC: System.out.print(row.getCell(j));break;
				case BOOLEAN: System.out.print(row.getCell(j));break;
				case FORMULA: System.out.println(row.getCell(j));break;
				}
				System.out.print("|");
			}
			System.out.println();
		}
	}
	}