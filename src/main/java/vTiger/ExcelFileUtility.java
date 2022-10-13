package vTiger;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtility {

	public String getDataFromExcel(String sheetName,int row,int cell) throws EncryptedDocumentException, InvalidFormatException, IOException {
		FileInputStream fis=new FileInputStream(ConstantUtility.excelPath);
		Workbook wb=WorkbookFactory.create(fis);
		Sheet sh=wb.getSheet(sheetName);
		Row r=sh.getRow(row);
		Cell c=r.getCell(cell);
		String value=c.getStringCellValue();
		wb.close();
		return value;
	}
	
	public int getLastRowNumber(String SheetName) throws IOException, EncryptedDocumentException, InvalidFormatException {
		FileInputStream fis=new FileInputStream(ConstantUtility.excelPath);
		Workbook wb=WorkbookFactory.create(fis);
		Sheet sh=wb.getSheet(SheetName);
		int r=sh.getLastRowNum();
		wb.close();
		return r;
	}
	public void writeDataIntoExcel(String SheetName,int row,int cell, String value) throws EncryptedDocumentException, InvalidFormatException, IOException {
		FileInputStream fis=new FileInputStream(ConstantUtility.excelPath);
		Workbook wb=WorkbookFactory.create(fis);
		Sheet sh=wb.getSheet(SheetName);
		Row r=sh.getRow(row);
		Cell ce=r.createCell(cell);
		ce.setCellValue(value);
		
		FileOutputStream fos=new FileOutputStream(ConstantUtility.excelPath);
		wb.write(fos);
		wb.close();
		
	}
	
	public Object[][] readMultipleData(String sheetName) throws EncryptedDocumentException, InvalidFormatException, IOException{
			FileInputStream fis=new FileInputStream(ConstantUtility.excelPath);
			Workbook wb=WorkbookFactory.create(fis);
			Sheet sh=wb.getSheet(sheetName);
			int lastrow=sh.getLastRowNum();
			int lastcell=sh.getRow(0).getLastCellNum();
			
			Object[][] data=new Object[lastrow][lastcell];
			for(int i=0;i<lastrow;i++) {
				for(int j=0;j<lastcell;j++) {
					data[i][j]=sh.getRow(i+1).getCell(j).getStringCellValue();
				}
			}
			return data;
	}
}
