package com.computation.estimate.utils;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class CellStyleCreatorUtil {

	public static CellStyle getCenterAlignAndBoldAndHorizontalAlignmentAndVerticalAlignmanetStyle(
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

	public static CellStyle setBordersToCell(CellStyle style) {
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);

		return style;
	}

	public static CellStyle returnStyleWithHorizontalAndVerticalAlignment(
			CellStyle fontWith10HeightStyle) {
		fontWith10HeightStyle.setAlignment(HorizontalAlignment.CENTER);
		fontWith10HeightStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		return fontWith10HeightStyle;
	}

	public static CellStyle getFontWith10HeightStyle(Workbook workbook) {
		CellStyle fontWith10HeightStyle = workbook.createCellStyle();
		CellStyleCreatorUtil.setBordersToCell(fontWith10HeightStyle);
		fontWith10HeightStyle = CellStyleCreatorUtil
				.returnStyleWithHorizontalAndVerticalAlignment(
						fontWith10HeightStyle);

		var fontWith10height = getFontWith10height(workbook);
		fontWith10HeightStyle.setFont(fontWith10height);

		return fontWith10HeightStyle;
	}

	public static Font getFontWith10height(Workbook workbook) {
		Font fontWith10height = workbook.createFont();
		fontWith10height.setFontHeightInPoints((short) 10);
		fontWith10height.setFontName("Arial");

		return fontWith10height;
	}
}
