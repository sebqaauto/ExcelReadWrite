package demoExcelReadWrite;
import java.io.*;

import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ExcelDemo1 {

	public static void main(String[] args) {
		
		ExcelDemo1 excel = new ExcelDemo1();
		excel.writeToExcelSheet();
		excel.readFromExcelSheet();
		
	}
	
	public void writeToExcelSheet() {
		
		try {
			// Set the path of the file and open the Stream to connect to it 
			FileInputStream fis = new FileInputStream("/Users/Sebastian/Desktop/EclipseWorkSpace/demo/demoExcelReadWrite/Book1.xlsx");
			//Open the workbook
			XSSFWorkbook workbk = new XSSFWorkbook(fis);
			//Open the Sheet
			XSSFSheet sheet = workbk.getSheetAt(0);
			//Access the sheet's Rows and Columns 
			Row row = sheet.createRow(1);
			Cell cell = row.createCell(0);
			cell.setCellValue("Dhoni");
			cell = row.createCell(1);
			cell.setCellValue("007");
			cell = row.createCell(2);
			cell.setCellValue("Wkt Batsmen");
			row = sheet.createRow(2);
			cell = row.createCell(0);
			cell.setCellValue("Yuvi");
			cell = row.createCell(1);
			cell.setCellValue("05");
			cell = row.createCell(2);
			cell.setCellValue("All - Rounder");
			
			
			FileOutputStream fos = new FileOutputStream("/Users/Sebastian/Desktop/EclipseWorkSpace/demo/demoExcelReadWrite/Book1.xlsx");
			workbk.write(fos);
			fis.close();
			fos.close();
			workbk.close();
			
			
						
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}catch (IOException e) {
			
			e.printStackTrace();
		}
		
		
	}
	
	public void readFromExcelSheet() {
		try {
			// Set the path of the file and open the Stream to connect to it 
			FileInputStream fis = new FileInputStream("/Users/Sebastian/Desktop/EclipseWorkSpace/demo/demoExcelReadWrite/Book1.xlsx");
			//Open the workbook
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			//Open the Sheet
			XSSFSheet sheet = workbook.getSheetAt(0);
			//Access the sheet's Rows and Columns 
			//Row row = sheet.getRow(0);
			//Cell cell = row.getCell(0);
			//String strCell1 = cell.getStringCellValue();
			//System.out.println(strCell1);
			
			int lastRow = sheet.getLastRowNum();
		
			for(int i = 0; i<=lastRow; i++) {
				Row row = sheet.getRow(i);
				int lastColumn = row.getLastCellNum();
				for(int k=0; k<lastColumn; k++) {
					Cell cell = row.getCell(k);
					String strCell1 = cell.getStringCellValue();
					System.out.print(strCell1+ "  ");
				}
				
				System.out.println("  " );
			}
			fis.close();
			workbook.close();
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		

	}

}
