package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelreading_using_for_loop {
static	String excelpath="C:\\Users\\USER\\Desktop\\Excel_Test_Data.xlsx";

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		File f=new File("C:\\Users\\USER\\Desktop\\Excel_Test_Data.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook xsw=new XSSFWorkbook(fis);
		XSSFSheet xss=xsw.getSheet("Sheet1");
		int r=xss.getLastRowNum();
		System.out.println("Row Counts  "+r);
		short c=xss.getRow(0).getLastCellNum();
		System.out.println(" cell "+c);
		for(int i=0;i<=r;i++) {
			
			for(int j=0;j<=c-1;j++) {
				
				/////////////Below 3 lines are cell data is converted in to string
				XSSFCell value=xss.getRow(i).getCell(j);
				value.setCellType(Cell.CELL_TYPE_STRING);
				String d=value.getStringCellValue();
				
				System.out.println("   values of excel cell  "+d);
				
			}
		}
	}

}
