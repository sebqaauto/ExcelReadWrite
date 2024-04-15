package demoExcelReadWrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingToExcel {

	public static void main(String[] args) {
		
		WritingToExcel excel = new WritingToExcel();
		excel.writeToExcel();
	}
	
	public void writeToExcel() {
		//Set the Stream to connect to excel file - To open the file
		try {
			FileInputStream fis = new FileInputStream ("/Users/Sebastian/Desktop/EclipseWorkSpace/demo/demoExcelReadWrite/Book2.xlsx");
			//Open the Workbook
			XSSFWorkbook xlWorkbook = new XSSFWorkbook(fis);
			//Open the Sheet
			XSSFSheet xlSheet = xlWorkbook.getSheetAt(0);
			//Get hold of the Rows in the particular
			XSSFRow xlRow = xlSheet.createRow(0);
			//Now use Cells and write your data in to the cell
			XSSFCell xlCell = xlRow.createCell(0);
			//Row number 1 is created and updated with values
			xlCell.setCellValue("Team_Name");
			xlCell = xlRow.createCell(1);
			xlCell.setCellValue("Titles_Count");
			xlCell = xlRow.createCell(2);
			xlCell.setCellValue("Captain");
			//Row number 2 is created and updated with values
			xlRow = xlSheet.createRow(1);
			xlCell = xlRow.createCell(0);
			xlCell.setCellValue("CSK");
			xlCell = xlRow.createCell(1);
			xlCell.setCellValue("5");
			xlCell = xlRow.createCell(2);
			xlCell.setCellValue("Dhoni");
			//Row number 3 is created and updated with values
			xlRow = xlSheet.createRow(2);
			xlCell = xlRow.createCell(0);
			xlCell.setCellValue("MI");
			xlCell = xlRow.createCell(1);
			xlCell.setCellValue("5");
			xlCell = xlRow.createCell(2);
			xlCell.setCellValue("Rohit");
			
			//OutStream to write the values to the destination file
			FileOutputStream fos = new FileOutputStream("/Users/Sebastian/Desktop/EclipseWorkSpace/demo/demoExcelReadWrite/Book2.xlsx");
			xlWorkbook.write(fos);
			fis.close();
			fos.close();
			xlWorkbook.close();
			
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}
