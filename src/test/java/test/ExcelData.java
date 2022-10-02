package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelData {
	public static void main(String[] args) throws IOException {
		File file =new File("C:\\Users\\admin\\eclipse-workspace\\test\\Excel Sheets\\ExcelPractice.xlsx");
		FileInputStream stream= new FileInputStream(file);
		Workbook workbook =new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Sheet1");
		Row row = sheet.getRow(1);	
		Cell cell2 = row.getCell(1);
		System.out.println(cell2);
		for (int i = 0; i < sheet.getPhysicalNumberOfRows()  ; i++) {
			Row row1 = sheet.getRow(i);
			for (int j = 0; j <row1.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				System.out.println(cell);
				
			}
		}
		
	}

}
