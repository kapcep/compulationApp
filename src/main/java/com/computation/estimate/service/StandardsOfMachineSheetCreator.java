package com.computation.estimate.service;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.computation.estimate.utils.CellStyleCreatorUtil;

public class StandardsOfMachineSheetCreator {

	private final static String standardsOfMachineSheetName = "Таблиця маш.год";

	public XSSFSheet createStandardsOfMachineSheet(
			XSSFWorkbook operationalInfoWorkbook) {
		XSSFSheet standardsOfMachineSheet = operationalInfoWorkbook
				.createSheet(standardsOfMachineSheetName);
		XSSFTable table = createTable(standardsOfMachineSheet);

		return standardsOfMachineSheet;
	}

	private XSSFTable createTable(XSSFSheet sheet) {

		XSSFWorkbook workbook = sheet.getWorkbook();
		CellStyle registerHeaderStyle = CellStyleCreatorUtil
				.getCenterAlignAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle(
						workbook);

		registerHeaderStyle.setFillForegroundColor(
				IndexedColors.GREY_25_PERCENT.getIndex());
		registerHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		CellStyle registerRowRightStyle = CellStyleCreatorUtil
				.getFontWith10HeightStyle(workbook);
		registerRowRightStyle.setAlignment(HorizontalAlignment.LEFT);

		// Header
		XSSFRow row0 = sheet.createRow(0);
		final XSSFCell cell00 = row0.createCell(0);
		cell00.setCellValue("№ п/п");
		cell00.setCellStyle(registerHeaderStyle);

		final XSSFCell cell01 = row0.createCell(1);
		cell01.setCellValue("Назва розцінки");
		cell01.setCellStyle(registerHeaderStyle);

		final XSSFCell cell02 = row0.createCell(2);
		cell02.setCellValue("Од.вим.");
		cell02.setCellStyle(registerHeaderStyle);

		// Rows
		XSSFRow row1 = sheet.createRow(1);

		for (int i = 0; i <= 2; i++) {
			final XSSFCell cell = row1.createCell(i);
			cell.setCellValue("Value" + i);

			cell.setCellStyle(registerRowRightStyle);

		}

		AreaReference area = workbook.getCreationHelper().createAreaReference(
				new CellReference(row0.getCell(0)),
				new CellReference(row1.getCell(2)));

		// create table
		var table = sheet.createTable(area);
		table.setDisplayName("tbl_standardsOfMachine");
		table.setName("tbl_standardsOfMachine");

		return null;
	}

}
