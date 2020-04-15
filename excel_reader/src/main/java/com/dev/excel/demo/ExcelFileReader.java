package com.dev.excel.demo;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileReader {

	public static void main(String[] args) {

		File inputFile = new File("Excel/Persons.xlsx");

		readExcelFile(inputFile);
	}

	public static void readExcelFile(File input) {

		try {
			
			FileInputStream inputStream = new FileInputStream(input);
			
			XSSFWorkbook workBook = new XSSFWorkbook(inputStream);
			
			XSSFSheet workSheet = workBook.getSheetAt(0);
			
			Row row;
			
			Cell cell;
			
			Iterator<Row> itr = workSheet.iterator();
			
			while(itr.hasNext()) {
				
				row = itr.next();
				
				Iterator<Cell> cellItr = row.cellIterator();
				
				while(cellItr.hasNext()) {
					
					cell = cellItr.next();
					
					switch (cell.getCellType()) {
					case STRING:
						System.out.print(cell.getStringCellValue() + "\t\t");
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t\t");
						break;
					default:
						break;
					}
				}
				System.out.println("");
			}
			workBook.close();

		} catch (Exception e) {

			e.printStackTrace();
		}
	}
}
