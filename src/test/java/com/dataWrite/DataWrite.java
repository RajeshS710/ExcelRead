package com.dataWrite;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataWrite {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet = workbook.createSheet("My workbook");
		
		Object  empdata[][] = {
				{"EmpId","Name","Job"},
				{101,"Rajesh","Engineer"},
				{102,"Mani","mechonic"},
				{103,"kdkd","notworked"}
		};
		
	int row = empdata.length;	
	int cell = empdata[0].length;
		
	System.out.println(row);
	System.out.println(cell);

	
	for(int i=0; i<row; i++) {
		
		XSSFRow r = sheet.createRow(i);
		
		for(int j = 0; j<cell; j++) {
			
			XSSFCell c = r.createCell(j);
	
			Object value = empdata[i][j];
			
		if(value instanceof String)
			c.setCellValue((String)value);
		
		if(value instanceof Integer)
			c.setCellValue((Integer)value);
		
		if(value instanceof Boolean)
			c.setCellValue((Boolean)value);
		}
		
		
	}
	
String filePath = "C:\\Users\\Admin\\eclipse-workspace\\ExcelRead\\Data\\output.xlsx";

FileOutputStream ou = new FileOutputStream(filePath);

workbook.write(ou);

ou.close();
	}
	
}
