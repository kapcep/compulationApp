package com.computation.estimate.service;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OperationalInfoExcelFileCreator {
	public static void main(String[] args) {
		OperationalInfoExcelFileCreator operationalInfoExcelFileCreator = new OperationalInfoExcelFileCreator();
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			XSSFSheet registerSheet = (XSSFSheet) workbook
					.createSheet("Реєстр");
			Sheet operationalInfoSheet = workbook
					.createSheet("відомість обємів робіт");

			operationalInfoExcelFileCreator.createRegisterTable(registerSheet,
					workbook);
			operationalInfoExcelFileCreator
					.createHeaderToTable(operationalInfoSheet, workbook);

			try (FileOutputStream outputStream = new FileOutputStream(
					"files/operational_info_file.xlsx")) {
				workbook.write(outputStream);
			}

			System.out.println("Excel файл створено успішно!");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void createRegisterTable(XSSFSheet sheet, XSSFWorkbook workbook) {
		CellStyle registerRowCenteredStyle = getFontWith10HeightStyle(workbook);

		CellStyle registerRowRightStyle = getFontWith10HeightStyle(workbook);
		registerRowRightStyle.setAlignment(HorizontalAlignment.LEFT);

		CellStyle registerHeaderStyle = getCenterAlignAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle(
				workbook);

		registerHeaderStyle.setFillForegroundColor(
				IndexedColors.GREY_25_PERCENT.getIndex());
		registerHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		// Header
		XSSFRow row0 = sheet.createRow(0);
		final XSSFCell cell00 = row0.createCell(0);
		cell00.setCellValue("Дата");
		cell00.setCellStyle(registerHeaderStyle);

		final XSSFCell cell01 = row0.createCell(1);
		cell01.setCellValue("Статус");
		cell01.setCellStyle(registerHeaderStyle);

		final XSSFCell cell02 = row0.createCell(2);
		cell02.setCellValue("№ п/п");
		cell02.setCellStyle(registerHeaderStyle);

		final XSSFCell cell03 = row0.createCell(3);
		cell03.setCellValue("Кошторис");
		cell03.setCellStyle(registerHeaderStyle);

		final XSSFCell cell04 = row0.createCell(4);
		cell04.setCellValue("Розділ");
		cell04.setCellStyle(registerHeaderStyle);

		final XSSFCell cell05 = row0.createCell(5);
		cell05.setCellValue("Підрозділ");
		cell05.setCellStyle(registerHeaderStyle);

		final XSSFCell cell06 = row0.createCell(6);
		cell06.setCellValue("Назва робіт");
		cell06.setCellStyle(registerHeaderStyle);

		final XSSFCell cell07 = row0.createCell(7);
		cell07.setCellValue("Од.вим.");
		cell07.setCellStyle(registerHeaderStyle);

		final XSSFCell cell08 = row0.createCell(8);
		cell08.setCellValue("К-сть");
		cell08.setCellStyle(registerHeaderStyle);

		final XSSFCell cell09 = row0.createCell(9);
		cell09.setCellValue("Ділянка");
		cell09.setCellStyle(registerHeaderStyle);

		final XSSFCell cell010 = row0.createCell(10);
		cell010.setCellValue("Примітка");
		cell010.setCellStyle(registerHeaderStyle);

		// Rows
		XSSFRow row1 = sheet.createRow(1);

		for (int i = 0; i <= 10; i++) {
			final XSSFCell cell = row1.createCell(i);
			cell.setCellValue(i);

			if (i == 0 || i == 2 || i == 7 || i == 8) {
				cell.setCellStyle(registerRowCenteredStyle);
			} else {
				cell.setCellStyle(registerRowRightStyle);
			}

		}

		AreaReference area = workbook.getCreationHelper().createAreaReference(
				new CellReference(row0.getCell(0)),
				new CellReference(row1.getCell(10)));

		var table = sheet.createTable(area);
		table.setDisplayName("registerOfWorks");

		sheet.setColumnWidth(0, 85 * 30);
		sheet.setColumnWidth(1, 77 * 30);
		sheet.setColumnWidth(2, 80 * 30);
		sheet.setColumnWidth(3, 154 * 30);
		sheet.setColumnWidth(4, 188 * 30);
		sheet.setColumnWidth(5, 100 * 30);
		sheet.setColumnWidth(6, 188 * 30);
		sheet.setColumnWidth(7, 87 * 30);
		sheet.setColumnWidth(8, 105 * 30);
		sheet.setColumnWidth(9, 105 * 30);
		sheet.setColumnWidth(10, 105 * 30);

	}

	private void createHeaderToTable(Sheet sheet, Workbook workbook) {

		// centerAlignStyleAndBoldAndHorizontalAlignmentAndVerticalAlignmanetFont

		CellStyle centerAlignAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle = getCenterAlignAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle(
				workbook);

		// fontWith10heightStyle

		CellStyle fontWith10HeightStyle = getFontWith10HeightStyle(workbook);

		// add cells to sheet

		// row one
		Row row0 = sheet.createRow(0);
		row0.setHeight((short) 900);
		Cell cell00 = row0.createCell(0);
		cell00.setCellValue(
				"НАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА НАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТАНАЗВА ОБЄКТА");
		cell00.setCellStyle(
				centerAlignAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle);

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

		Cell cell14 = row1.createCell(4);
		cell14.setCellStyle(fontWith10HeightStyle);
		Cell cell15 = row1.createCell(5);
		cell15.setCellStyle(fontWith10HeightStyle);

		Cell cell16 = row1.createCell(6);
		cell16.setCellStyle(fontWith10HeightStyle);
		cell16.setCellValue("Виконано");

		Cell cell17 = row1.createCell(7);
		cell17.setCellStyle(fontWith10HeightStyle);

		Cell cell18 = row1.createCell(8);
		cell18.setCellStyle(fontWith10HeightStyle);
		cell18.setCellValue("Залишок");

		Cell cell19 = row1.createCell(9);
		cell19.setCellStyle(fontWith10HeightStyle);

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

		// union cells
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 5));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 7));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 8, 9));

		// row 3
		Row row2 = sheet.createRow(2);

		// add horizontal numbering
		for (int i = 1; i <= 10; i++) {
			Cell cell = row2.createCell(i - 1);
			cell.setCellStyle(fontWith10HeightStyle);
			cell.setCellValue(i);
		}

	}

	private CellStyle getFontWith10HeightStyle(Workbook workbook) {
		CellStyle fontWith10HeightStyle = workbook.createCellStyle();
		setBordersToCell(fontWith10HeightStyle);
		fontWith10HeightStyle = returnStyleWithHorizontalAndVerticalAlignment(
				fontWith10HeightStyle);

		Font fontWith10height = workbook.createFont();
		fontWith10height.setFontHeightInPoints((short) 10);
		fontWith10HeightStyle.setFont(fontWith10height);

		return fontWith10HeightStyle;
	}

	private CellStyle getCenterAlignAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle(
			Workbook workbook) {
		CellStyle centerAlignStyleAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle = workbook
				.createCellStyle();
		centerAlignStyleAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle = returnStyleWithHorizontalAndVerticalAlignment(
				centerAlignStyleAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle);
		centerAlignStyleAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle
				.setWrapText(true);
		setBordersToCell(
				centerAlignStyleAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle);

		Font fontBold = workbook.createFont();
		fontBold.setBold(true);
		fontBold.setFontHeightInPoints((short) 12);
		centerAlignStyleAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle
				.setFont(fontBold);

		return centerAlignStyleAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle;
	}

	private CellStyle returnStyleWithHorizontalAndVerticalAlignment(
			CellStyle fontWith10HeightStyle) {
		fontWith10HeightStyle.setAlignment(HorizontalAlignment.CENTER);
		fontWith10HeightStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		return fontWith10HeightStyle;
	}

	private void setBordersToCell(CellStyle style) {
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
	}
}
