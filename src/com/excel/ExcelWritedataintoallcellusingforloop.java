package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWritedataintoallcellusingforloop {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		

		File f=new File("C:\\Users\\USER\\Desktop\\Todayexcelsampledata.xlsx");
		
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook xsw=new XSSFWorkbook(fis);
		XSSFSheet xss=xsw.getSheet("Sheet1");
		
		int row=xss.getLastRowNum();
		int collmns=xss.getRow(0).getLastCellNum();
				
		for(int i=0;i<=row-1;i++) {
			for(int j=0;j<=collmns-1;j++) {
				
				Cell c1=xss.getRow(i).getCell(j);
				if(c1==null) {
					c1=xss.getRow(i).createCell(j);
				}
	
				c1.setCellValue("mehubooba");
				System.out.println("  enter Text  "+c1);
			
				
				FileOutputStream fos=new FileOutputStream(f);
				xsw.write(fos);
				xsw.close();
				fos.close();
				
				
			}
		}

	}

}
