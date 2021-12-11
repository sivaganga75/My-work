package com.excel.read;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ExcelCode {
	public static FileInputStream file;
	public static HSSFWorkbook workbook;
	public static HSSFSheet sheet;

	public static String readStringData(int i, int j) throws IOException {

		file = new FileInputStream("C:\\Users\\Sivaganga\\Desktop\\New folder\\Book1.xls");
		workbook = new HSSFWorkbook(file);
		sheet = workbook.getSheet("Sheet1");
		Row row = sheet.getRow(i);
		Cell cell = row.getCell(j);
		for (Row r : sheet) {
			for (Cell c : r) {
				System.out.print(c + "   ");
			}
			System.out.println();
		}

		return cell.getStringCellValue();
	}

	public static String readIntegerData(int i, int j) throws IOException {

		file = new FileInputStream("C:\\Users\\Sivaganga\\Desktop\\New folder\\Book1.xls");
		workbook = new HSSFWorkbook(file);
		sheet = workbook.getSheet("Sheet1");
		Row row = sheet.getRow(i);
		Cell cell = row.getCell(j);
		int var = (int) cell.getNumericCellValue();
		return String.valueOf(var);
	}
}


