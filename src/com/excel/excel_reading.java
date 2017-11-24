package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excel_reading {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		
		File f=new File("C:\\Users\\USER\\Desktop\\Excel_Test_Data.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook xsw=new XSSFWorkbook(fis);
		XSSFSheet xss=xsw.getSheet("Sheet1");
		String r=xss.getRow(0).getCell(0).getStringCellValue();
		System.out.println("text "+r);		
		
		

	}

}
