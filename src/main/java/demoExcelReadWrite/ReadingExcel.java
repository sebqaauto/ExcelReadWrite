package demoExcelReadWrite;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) {
		
		ReadingExcel read = new ReadingExcel();
		try {
			read.readExcelFile();
		} catch (IOException e) {
			
			e.printStackTrace();
		}
	}
	
	
	public void readExcelFile() throws IOException {
		//Set the path to open the Stream 
		FileInputStream fis = new FileInputStream ("/Users/Sebastian/Desktop/EclipseWorkSpace/demo/demoExcelReadWrite/Book2.xlsx");
		//Open the Workbook
		XSSFWorkbook xlWorkbook = new XSSFWorkbook(fis);
		//Open the Sheet
		XSSFSheet xlSheet = xlWorkbook.getSheetAt(0);
		//Get hold of the Rows in the particular
		XSSFRow xlRow = xlSheet.getRow(0);
		//Go to the Cell and read its value
		XSSFCell xlCell = xlRow.getCell(0);
		String cellValue = xlCell.getStringCellValue();
		System.out.println(cellValue+ "  ");
		xlCell = xlRow.getCell(1);
		cellValue = xlCell.getStringCellValue();
		System.out.println(cellValue+ "  ");
		xlCell = xlRow.getCell(2);
		cellValue = xlCell.getStringCellValue();
		System.out.println(cellValue+ "  ");
		
		//Get the number of rows in the excel, accordingly iterate the column values and print them 
		int lastRow = xlSheet.getLastRowNum();
		//Iterate through the Rows 
		for(int i=0; i<=lastRow; i++) {
			xlRow = xlSheet.getRow(i);
			//Get the last column and iterate through the columns 
			int lastColumn = xlRow.getLastCellNum();
			for(int k=0; k<lastColumn; k++) {
				xlCell = xlRow.getCell(k);
				cellValue = xlCell.getStringCellValue();
				System.out.print(cellValue+ "  ");
			}
			System.out.println(" ");
		}
		
		fis.close();
		xlWorkbook.close();
	}

}
