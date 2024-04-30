package com.computation.estimate.service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.computation.estimate.entity.ComputationPosition;

public class OperationalInfoExcelFileCreator {
	final static String outboxExcelFilePath = "files/451_du.xlsx";

	public static void main(String[] args) {
		OperationalInfoExcelFileCreator operationalInfoExcelFileCreator = new OperationalInfoExcelFileCreator();
		GetComputationInfoFromExcelFile computationInfoFromExcelFile = new GetComputationInfoFromExcelFile();
		try (XSSFWorkbook operationalInfoWorkbook = new XSSFWorkbook()) {
			XSSFSheet registerSheet = (XSSFSheet) operationalInfoWorkbook
					.createSheet("Реєстр");
			Sheet operationalInfoSheet = operationalInfoWorkbook
					.createSheet("відомість обємів робіт");

			// freeze 4 row
			operationalInfoSheet.createFreezePane(2, 4);

			// create register table
			operationalInfoExcelFileCreator.createRegisterTable(registerSheet,
					operationalInfoWorkbook);

			// create operational info headers in operationalInfoSheet
			operationalInfoExcelFileCreator.createOperationalInfoHeaderToTable(
					operationalInfoSheet, operationalInfoWorkbook);

			// get computationPositions
			Workbook outboxExcelWorkbook = operationalInfoExcelFileCreator
					.getOutboxExcelWorkbook(outboxExcelFilePath);
			List<ComputationPosition> computationPositions = computationInfoFromExcelFile
					.getComputationPosition(outboxExcelWorkbook);

			// write сomputation positions to operational info sheet
			operationalInfoExcelFileCreator
					.writeComputationPositionsToOperationalInfoSheet(
							operationalInfoSheet, computationPositions,
							operationalInfoWorkbook);

			// add conditionalFormatting rules

			SheetConditionalFormatting conditionalFormatting = operationalInfoSheet
					.getSheetConditionalFormatting();

			operationalInfoExcelFileCreator.createConditionalFormattingRule(
					operationalInfoSheet, conditionalFormatting,
					ComparisonOperator.GT, "0", IndexedColors.YELLOW.index,
					"G8:G" + (operationalInfoSheet.getLastRowNum() + 1));
			operationalInfoExcelFileCreator.createConditionalFormattingRule(
					operationalInfoSheet, conditionalFormatting,
					ComparisonOperator.LT, "0", IndexedColors.RED.index,
					"I8:I" + (operationalInfoSheet.getLastRowNum() + 1));

			try (FileOutputStream outputStream = new FileOutputStream(
					"files/operational_info_file.xlsx")) {
				operationalInfoWorkbook.write(outputStream);
			}

			System.out.println("Excel файл створено успішно!");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private SheetConditionalFormatting createConditionalFormattingRule(
			Sheet operationalInfoSheet,
			SheetConditionalFormatting conditionalFormatting,
			byte comparisonOperator, String formula, short colorIndex,
			String regionsString) {

		ConditionalFormattingRule rule = conditionalFormatting
				.createConditionalFormattingRule(comparisonOperator, formula);
		PatternFormatting fill = rule.createPatternFormatting();
		fill.setFillBackgroundColor(colorIndex);
		fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

		CellRangeAddress[] regions = {
				CellRangeAddress.valueOf(regionsString) };
		conditionalFormatting.addConditionalFormatting(regions, rule);
		return conditionalFormatting;
	}

	private void writeComputationPositionsToOperationalInfoSheet(
			Sheet operationalInfoSheet,
			List<ComputationPosition> computationPositions, Workbook workbook) {

		CellStyle computationPositionNameStyle = getFontWith10HeightStyle(
				workbook);
		computationPositionNameStyle.setAlignment(HorizontalAlignment.LEFT);
		computationPositionNameStyle.setWrapText(true);

		CellStyle centeredStyle = getFontWith10HeightStyle(workbook);

		int countNumberingOfComputationPositions = 1;
		String computationNumberAndNameTemp = "";

		int rowCount = 4;
		int countSection = 1;
		for (int i = 0; i < computationPositions.size(); i++) {

			final ComputationPosition computationCurrentPosition = computationPositions
					.get(i);

			ComputationPosition computationPreviousPosition = null;
			String computationPreviousTypeName = "";
			String computationPreviousPositionName = "";

			if (i != 0) {
				computationPreviousPosition = computationPositions.get(i - 1);
				computationPreviousTypeName = computationPreviousPosition
						.getComputationTypePosition().getComputationTypeName();
				computationPreviousPositionName = computationPreviousPosition
						.getComputationPositionName();
			}

			final String computationCurrentTypeName = computationCurrentPosition
					.getComputationTypePosition().getComputationTypeName();

			String computationCurrentPositionName = computationCurrentPosition
					.getComputationPositionName();
			String computationCurrentNumberAndName = computationCurrentPosition
					.getComputation().getComputationNumber() + " "
					+ computationCurrentPosition.getComputation()
							.getComputationName();

			if (computationCurrentTypeName == "H") {
				continue;
			}

			if (!computationNumberAndNameTemp
					.equals(computationCurrentNumberAndName)) {
				countSection = 1;
				Row row = operationalInfoSheet.createRow(rowCount++);
				Cell cell = row.createCell(1);
				cell.setCellValue("ЛК " + computationCurrentNumberAndName);
				computationNumberAndNameTemp = computationCurrentNumberAndName
						.toString();

				CellStyle computationNumberAndNameStyle = returnHeaderRowStyle(
						workbook, IndexedColors.SEA_GREEN.getIndex());

				cell.setCellStyle(computationNumberAndNameStyle);

				formatRow(row, computationNumberAndNameStyle);
			}

			if (computationCurrentTypeName == "R") {

				Row row = operationalInfoSheet.createRow(rowCount++);
				Cell cell = row.createCell(1);
				cell.setCellValue("Розділ " + countSection + ". "
						+ computationCurrentPositionName);
				CellStyle sectionStyle = returnHeaderRowStyle(workbook,
						IndexedColors.LIGHT_ORANGE.getIndex());

				cell.setCellStyle(sectionStyle);

				countSection++;
				formatRow(row, sectionStyle);

				continue;
			}

			if (computationPreviousTypeName == "R"
					&& computationCurrentTypeName != "Y") {
				rowCount = createSubSectionRow(operationalInfoSheet, workbook,
						rowCount, computationPreviousPositionName);
			}

			if (computationCurrentTypeName == "Y") {
				rowCount = createSubSectionRow(operationalInfoSheet, workbook,
						rowCount, computationCurrentPositionName);
				continue;
			}

			Row row = operationalInfoSheet.createRow(rowCount++);

			// set numbering
			if (computationCurrentTypeName == " Поз. Л.С. ") {
				Cell cell0 = row.createCell(0);
				cell0.setCellValue(countNumberingOfComputationPositions++);
				cell0.setCellStyle(centeredStyle);
			} else {
				continue;
			}

			// set computationPositionName
			Cell cell1 = row.createCell(1);
			cell1.setCellValue(
					computationCurrentPosition.getComputationPositionName());
			cell1.setCellStyle(computationPositionNameStyle);

			// set computationPositionUnitOfMeasurement
			Cell cell2 = row.createCell(2);
			cell2.setCellValue(computationCurrentPosition
					.getComputationPositionUnitOfMeasurement()
					.getComputationPositionUnitOfMeasurementName());
			cell2.setCellStyle(centeredStyle);

			// set amount
			Cell cell3 = row.createCell(3);
			cell3.setCellValue(computationCurrentPosition.getAmount());
			cell3.setCellStyle(centeredStyle);

			// set amount unit cost
			Cell cell4 = row.createCell(4);
			cell4.setCellFormula("F" + rowCount + "/" + "D" + rowCount);
			cell4.setCellStyle(centeredStyle);

			// set computationPositionContractPrice
			Cell cell5 = row.createCell(5);
			cell5.setCellValue(computationCurrentPosition
					.getComputationPositionContractPrice());
			cell5.setCellStyle(centeredStyle);

			// set the sum of the number of works per period
			Cell cell6 = row.createCell(6);
			cell6.setCellFormula(
					"SUM(K" + rowCount + ":" + "M" + rowCount + ")");
			cell6.setCellStyle(centeredStyle);

			// the cost of the work performed
			Cell cell7 = row.createCell(7);
			cell7.setCellFormula("E" + rowCount + "*" + "G" + rowCount);
			cell7.setCellStyle(centeredStyle);

			// the remainder of the completed works
			Cell cell8 = row.createCell(8);
			cell8.setCellFormula("D" + rowCount + "-" + "G" + rowCount);
			cell8.setCellStyle(centeredStyle);

			// the price of the remaining work performed
			Cell cell9 = row.createCell(9);
			cell9.setCellFormula("I" + rowCount + "*" + "E" + rowCount);
			cell9.setCellStyle(centeredStyle);

			// blank cells
			Cell cell10 = row.createCell(10);
			cell10.setCellStyle(centeredStyle);
			Cell cell11 = row.createCell(11);
			cell11.setCellStyle(centeredStyle);
			Cell cell12 = row.createCell(12);
			cell12.setCellStyle(centeredStyle);
		}

	}

	private int createSubSectionRow(Sheet operationalInfoSheet,
			Workbook workbook, int rowCount, String computationPositionName) {
		Row row = operationalInfoSheet.createRow(rowCount++);
		Cell cell = row.createCell(1);
		cell.setCellValue(computationPositionName);

		CellStyle subSectionStyle = returnHeaderRowStyle(workbook,
				IndexedColors.LEMON_CHIFFON.getIndex());
		cell.setCellStyle(subSectionStyle);
		formatRow(row, subSectionStyle);
		return rowCount;
	}

	private void formatRow(Row row, CellStyle sectionStyle) {
		for (int j = 0; j < 13; j++) {
			if (j == 1) {
				continue;
			}
			Cell rowStyleCell = row.createCell(j);
			rowStyleCell.setCellStyle(sectionStyle);
		}
	}

	private CellStyle returnHeaderRowStyle(Workbook workbook,
			short colorIndex) {
		CellStyle sectionStyle = workbook.createCellStyle();
		Font font = getFontWith10height(workbook);
		font.setBold(true);
		sectionStyle.setFont(font);

		sectionStyle.setFillForegroundColor(colorIndex);
		sectionStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		return sectionStyle;
	}

	private Workbook getOutboxExcelWorkbook(String outboxExcelFilePath) {

		try (Workbook workbook = WorkbookFactory
				.create(new FileInputStream(outboxExcelFilePath))) {

			return workbook;

		} catch (IOException e) {
			return null;
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
		sheet.setColumnWidth(5, 120 * 30);
		sheet.setColumnWidth(6, 188 * 30);
		sheet.setColumnWidth(7, 87 * 30);
		sheet.setColumnWidth(8, 105 * 30);
		sheet.setColumnWidth(9, 105 * 30);
		sheet.setColumnWidth(10, 105 * 30);

	}

	private void createOperationalInfoHeaderToTable(Sheet sheet,
			Workbook workbook) {

		// centerAlignStyleAndBoldAndHorizontalAlignmentAndVerticalAlignmanetFont
		CellStyle centerAlignAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle = getCenterAlignAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle(
				workbook);

		// fontWith10heightStyle
		CellStyle fontWith10HeightStyle = getFontWith10HeightStyle(workbook);

		// modifiedFontWith10Height
		CellStyle modifiedFontWith10HeightBoldWrapTextStyle = workbook
				.createCellStyle();
		modifiedFontWith10HeightBoldWrapTextStyle
				.cloneStyleFrom(fontWith10HeightStyle);
		Font font = getFontWith10height(workbook);
		font.setBold(true);
		modifiedFontWith10HeightBoldWrapTextStyle.setFont(font);
		modifiedFontWith10HeightBoldWrapTextStyle.setWrapText(true);

		// add cells to sheet

		// row one
		Row row0 = sheet.createRow(0);
		row0.setHeight((short) 900);
		Cell cell00 = row0.createCell(0);
		cell00.setCellValue("НАЗВА ОБЄКТА");
		cell00.setCellStyle(
				centerAlignAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle);

		// row two
		Row row1 = sheet.createRow(1);

		Cell cell_1_0 = row1.createCell(0);
		cell_1_0.setCellStyle(modifiedFontWith10HeightBoldWrapTextStyle);
		cell_1_0.setCellValue("№ пп");

		Cell cell_1_1 = row1.createCell(1);
		cell_1_1.setCellStyle(modifiedFontWith10HeightBoldWrapTextStyle);
		cell_1_1.setCellValue("Найменування робiт і витрат");

		Cell cell_1_2 = row1.createCell(2);
		cell_1_2.setCellStyle(modifiedFontWith10HeightBoldWrapTextStyle);
		cell_1_2.setCellValue("Вимірник");

		Cell cell_1_3 = row1.createCell(3);
		cell_1_3.setCellStyle(modifiedFontWith10HeightBoldWrapTextStyle);
		cell_1_3.setCellValue("Загальний обсяг робіт ДЦ");

		Cell cell_1_4 = row1.createCell(4);
		cell_1_4.setCellStyle(fontWith10HeightStyle);
		Cell cell_1_5 = row1.createCell(5);
		cell_1_5.setCellStyle(fontWith10HeightStyle);

		Cell cell_1_6 = row1.createCell(6);
		cell_1_6.setCellStyle(modifiedFontWith10HeightBoldWrapTextStyle);
		cell_1_6.setCellValue("Виконано");

		Cell cell_1_7 = row1.createCell(7);
		cell_1_7.setCellStyle(fontWith10HeightStyle);

		Cell cell_1_8 = row1.createCell(8);
		cell_1_8.setCellStyle(modifiedFontWith10HeightBoldWrapTextStyle);
		cell_1_8.setCellValue("Залишок");

		Cell cell_1_9 = row1.createCell(9);
		cell_1_9.setCellStyle(fontWith10HeightStyle);

		Cell cell_1_10 = row1.createCell(10);
		Calendar calendarCurrentDate = plusDaysToCurrentDate(0);
		cell_1_10.setCellValue(calendarCurrentDate.getTime());
		CellStyle dateFormatStyle = getDateFormatStyle(workbook);
		cell_1_10.setCellStyle(dateFormatStyle);

		Cell cell_1_11 = row1.createCell(11);
		Calendar calendarPlusOneDay = plusDaysToCurrentDate(1);
		cell_1_11.setCellValue(calendarPlusOneDay.getTime());
		cell_1_11.setCellStyle(dateFormatStyle);

		Cell cell_1_12 = row1.createCell(12);
		Calendar calendarPlusTwoDays = plusDaysToCurrentDate(2);
		cell_1_12.setCellValue(calendarPlusTwoDays.getTime());
		cell_1_12.setCellStyle(dateFormatStyle);

		// row 3
		Row row2 = sheet.createRow(2);
		Cell cell_2_10 = row2.createCell(10);
		cell_2_10.setCellStyle(dateFormatStyle);

		Cell cell_2_11 = row2.createCell(11);
		cell_2_11.setCellStyle(dateFormatStyle);

		Cell cell_2_12 = row2.createCell(12);
		cell_2_12.setCellStyle(dateFormatStyle);

		// add cells to row 3

		for (int i = 1; i <= 10; i++) {
			Cell cell = row2.createCell(i - 1);
			cell.setCellStyle(modifiedFontWith10HeightBoldWrapTextStyle);
			if (i == 4) {
				cell.setCellValue("Обсяг");
			} else if (i == 5) {
				cell.setCellValue("Вартість одиниці");
			} else if (i == 6) {
				cell.setCellValue("Вартість робіт з ПДВ");
			} else if (i == 7) {
				cell.setCellValue("Обсяг");
			} else if (i == 8) {
				cell.setCellValue("Вартість робіт з ПДВ");
			} else if (i == 9) {
				cell.setCellValue("Обсяг");
			} else if (i == 10) {
				cell.setCellValue("Залишкова вартість робіт з ПДВ");
			}

		}

		// row 4
		Row row3 = sheet.createRow(3);

		// add horizontal numbering
		for (int i = 1; i <= 13; i++) {
			if (i <= 10) {
				Cell cell = row3.createCell(i - 1);
				cell.setCellStyle(modifiedFontWith10HeightBoldWrapTextStyle);
				cell.setCellValue(i);
			} else {
				Cell cell = row3.createCell(i - 1);
				cell.setCellStyle(fontWith10HeightStyle);
			}
		}

		// set width of columns
		sheet.setColumnWidth(0, 4 * 300);
		sheet.setColumnWidth(1, 552 * 30);
		sheet.setColumnWidth(2, 85 * 30);
		sheet.setColumnWidth(3, 100 * 30);
		sheet.setColumnWidth(4, 100 * 30);
		sheet.setColumnWidth(5, 160 * 30);
		sheet.setColumnWidth(6, 100 * 30);
		sheet.setColumnWidth(7, 100 * 30);
		sheet.setColumnWidth(8, 100 * 30);
		sheet.setColumnWidth(9, 160 * 30);

		// union cells
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9));
		sheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
		sheet.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 5));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 6, 7));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 8, 9));
		sheet.addMergedRegion(new CellRangeAddress(1, 3, 10, 10));
		sheet.addMergedRegion(new CellRangeAddress(1, 3, 11, 11));
		sheet.addMergedRegion(new CellRangeAddress(1, 3, 12, 12));

	}

	private Calendar plusDaysToCurrentDate(int days) {
		Calendar calendarPlusOneDay = Calendar.getInstance();
		calendarPlusOneDay.setTime(new Date());
		calendarPlusOneDay.add(Calendar.DAY_OF_MONTH, days);
		calendarPlusOneDay.set(Calendar.HOUR_OF_DAY, 0);
		calendarPlusOneDay.set(Calendar.MINUTE, 0);
		calendarPlusOneDay.set(Calendar.SECOND, 0);
		calendarPlusOneDay.set(Calendar.MILLISECOND, 0);
		return calendarPlusOneDay;
	}

	private CellStyle getDateFormatStyle(Workbook workbook) {
		CellStyle dateFormatStyle = workbook.createCellStyle();
		DataFormat format = workbook.createDataFormat();
		dateFormatStyle.setDataFormat(format.getFormat("dd.MM"));
		setBordersToCell(dateFormatStyle);
		dateFormatStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		dateFormatStyle.setAlignment(HorizontalAlignment.CENTER);

		Font font = getFontWith10height(workbook);
		font.setBold(true);
		dateFormatStyle.setFont(font);

		return dateFormatStyle;
	}

	private CellStyle getFontWith10HeightStyle(Workbook workbook) {
		CellStyle fontWith10HeightStyle = workbook.createCellStyle();
		setBordersToCell(fontWith10HeightStyle);
		fontWith10HeightStyle = returnStyleWithHorizontalAndVerticalAlignment(
				fontWith10HeightStyle);

		var fontWith10height = getFontWith10height(workbook);
		fontWith10HeightStyle.setFont(fontWith10height);

		return fontWith10HeightStyle;
	}

	private Font getFontWith10height(Workbook workbook) {
		Font fontWith10height = workbook.createFont();
		fontWith10height.setFontHeightInPoints((short) 10);
		fontWith10height.setFontName("Arial");

		return fontWith10height;
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
		fontBold.setFontName("Arial");
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
