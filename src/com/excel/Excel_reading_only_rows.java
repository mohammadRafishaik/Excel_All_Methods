package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_reading_only_rows {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		File f=new File("C:\\Users\\USER\\Desktop\\Excel_Test_Data.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook xsw=new XSSFWorkbook(fis);
		XSSFSheet xss=xsw.getSheet("Sheet1");
		int rows=xss.getLastRowNum();
		System.out.println("Rows Count "+rows);
		for(int j=0;j<=rows;j++) {
			
			
			String values=xss.getRow(0).getCell(j).getStringCellValue();
			System.out.println(" all rows values "+values);
		}
		

	}

}
