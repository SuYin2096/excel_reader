package com.dev.excel.demo;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReaderDemo {

	public static void main(String[] args) {

		File file = new File("Excel/Persons.xlsx");
		
		readExcelFile(file);
	}

	public static void readExcelFile(File input) {

		try {
			FileInputStream inputStream = new FileInputStream(input);

			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

			XSSFSheet sheet = workbook.getSheetAt(0);

			Row row;

			Cell cell;

			Iterator<Row> itr = sheet.rowIterator();

			while (itr.hasNext()) {

				row = itr.next();

				Iterator<Cell> cellItr = row.cellIterator();

				while (cellItr.hasNext()) {

					cell = cellItr.next();

					switch (cell.getCellTypeEnum()) {
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
			
			workbook.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
