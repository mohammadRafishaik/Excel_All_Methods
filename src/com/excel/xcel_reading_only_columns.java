package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xcel_reading_only_columns {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub

		File f=new File("C:\\Users\\USER\\Desktop\\Excel_Test_Data.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook xsw=new XSSFWorkbook(fis);
		XSSFSheet xss=xsw.getSheet("Sheet1");
		
		String value=xss.getRow(0).getCell(0).getStringCellValue();
			
		
		
	}

}
