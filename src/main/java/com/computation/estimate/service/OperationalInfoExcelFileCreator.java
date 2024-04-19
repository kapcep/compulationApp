package com.computation.estimate.service;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OperationalInfoExcelFileCreator {
	public static void main(String[] args) {
		OperationalInfoExcelFileCreator operationalInfoExcelFileCreator = new OperationalInfoExcelFileCreator();
		try (Workbook workbook = new XSSFWorkbook()) {
			Sheet sheet = workbook.createSheet("відомість обємів робіт");

			operationalInfoExcelFileCreator.createHeaderToTable(sheet,
					workbook);

			try (FileOutputStream outputStream = new FileOutputStream(
					"files/operational_info_file.xlsx")) {
				workbook.write(outputStream);
			}

			System.out.println("Excel файл створено успішно!");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void createHeaderToTable(Sheet sheet, Workbook workbook) {

		// centerAlignStyleAndBoldFont

		CellStyle centerAlignStyleAndBoldFont = workbook.createCellStyle();
		centerAlignStyleAndBoldFont.setAlignment(HorizontalAlignment.CENTER);

		Font fontBold = workbook.createFont();
		fontBold.setBold(true);
		fontBold.setFontHeightInPoints((short) 12);
		// fontBold.setFontName("Times New Roman");
		centerAlignStyleAndBoldFont.setFont(fontBold);

		// fontWith10heightStyle

		CellStyle fontWith10HeightStyle = workbook.createCellStyle();
		fontWith10HeightStyle.setAlignment(HorizontalAlignment.CENTER);
		Font fontWith10height = workbook.createFont();
		fontWith10height.setFontHeightInPoints((short) 10);
		fontWith10HeightStyle.setFont(fontWith10height);

		// add cells to sheet

		// row one
		Row row0 = sheet.createRow(0);
		Cell cell00 = row0.createCell(0);
		cell00.setCellValue("ВІДОМІСТЬ ОБ`ЄМІВ РОБІТ");
		cell00.setCellStyle(centerAlignStyleAndBoldFont);

		// row two
		Row row1 = sheet.createRow(1);

		Cell cell10 = row1.createCell(0);
		cell10.setCellStyle(fontWith10HeightStyle);
		cell10.setCellValue("№ пп");

		Cell cell11 = row1.createCell(1);
		cell11.setCellStyle(fontWith10HeightStyle);
		cell11.setCellValue("Найменування робiт і витрат");

		Cell cell12 = row1.createCell(2);
		cell12.setCellStyle(fontWith10HeightStyle);
		cell12.setCellValue("Вимірник");

		Cell cell13 = row1.createCell(3);
		cell13.setCellStyle(fontWith10HeightStyle);
		cell13.setCellValue("Загальний обсяг робіт ДЦ");

		Cell cell16 = row1.createCell(6);
		cell16.setCellStyle(fontWith10HeightStyle);
		cell16.setCellValue("Виконано");

		Cell cell18 = row1.createCell(8);
		cell18.setCellStyle(fontWith10HeightStyle);
		cell18.setCellValue("Залишок");

		// set width of columns
		sheet.setColumnWidth(0, 4 * 300);
		sheet.setColumnWidth(1, 552 * 30);
		sheet.setColumnWidth(2, 85 * 30);
		sheet.setColumnWidth(3, 100 * 30);
		sheet.setColumnWidth(4, 100 * 30);
		sheet.setColumnWidth(5, 160 * 30);
		sheet.setColumnWidth(6, 100 * 30);
		sheet.setColumnWidth(7, 160 * 30);
		sheet.setColumnWidth(8, 100 * 30);

		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 5));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 7));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 8, 9));

	}
}