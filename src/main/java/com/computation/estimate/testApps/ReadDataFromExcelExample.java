package com.computation.estimate.testApps;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadDataFromExcelExample {
	public static void main(String[] args) {
		try (Workbook workbook = WorkbookFactory
				.create(new FileInputStream("files/U_392_ДЦ_КК_2.xls"))) {
			Sheet sheet = workbook.getSheetAt(0);
			DecimalFormat df = new DecimalFormat("#,##0.00");

			int count = 1;

			for (int i = 14; i <= sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);

				Cell firstCell = row.getCell(0);

				if (row != null && firstCell.getCellType() == CellType.NUMERIC
						&& (int) firstCell.getNumericCellValue() == count) {

					Cell computationPositionNameCell = row.getCell(2);
					Cell computationPositionCell = row.getCell(8);
					System.out
							.println(
									"(" + count + ") "
											+ computationPositionNameCell
													.getStringCellValue()
													.replaceAll("\\s+", " ")
											+ " ціна: "
											+ df.format(computationPositionCell
													.getNumericCellValue()
													* 1.2)
											+ "грн.");
					count++;
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
