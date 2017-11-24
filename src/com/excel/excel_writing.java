package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excel_writing {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		File f=new File("C:\\Users\\USER\\Desktop\\Excel_Test_Data.xlsx");
		
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook xsw=new XSSFWorkbook(fis);
		XSSFSheet xss=xsw.getSheet("Sheet1");
		Cell c1=xss.getRow(0).getCell(5);
		if(c1==null) {
			
			c1=xss.getRow(0).createCell(5);
		}
		
c1.setCellValue("passsword");



System.out.println(" enter value  "+c1);
FileOutputStream fos=new FileOutputStream(f);
xsw.write(fos);
xsw.close();
fos.close();



	}

}