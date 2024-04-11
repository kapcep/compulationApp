package com.computation.estimate;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcelFileExample {
	public static void main(String[] args) {
		try (Workbook workbook = new XSSFWorkbook()) { 
			Sheet sheet = workbook.createSheet("Sheet1");

			
			Row row = sheet.createRow(0);
			Cell cell = row.createCell(0);
			cell.setCellValue("Hello, World!");

			try (FileOutputStream outputStream = new FileOutputStream(
					"files/new_excel_file.xlsx")) {
				workbook.write(outputStream);
			}

			System.out.println("Excel файл створено успішно!");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
