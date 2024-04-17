package com.computation.estimate.testApps;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadDataFromContractPriceExcelExample {
	public static void main(String[] args) {
		try (Workbook workbook = WorkbookFactory
				.create(new FileInputStream("files/U_392_ДЦ_КК_2.xls"))) {
			Sheet sheet = workbook.getSheetAt(0);
			DecimalFormat df = new DecimalFormat("#,##0.00");

			int count = 1;
			Map<Integer, Double> computationPositionContractPriceMap = new HashMap<>();

			for (int i = 14; i <= sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				Cell firstCell = row.getCell(0);

				if (row != null && firstCell.getCellType() == CellType.NUMERIC
						&& (int) firstCell.getNumericCellValue() == count) {

					Cell computationPositionContractPriceCell = row.getCell(8);

					Double computationPositionContractPrice = Double.valueOf(df
							.format(computationPositionContractPriceCell
									.getNumericCellValue() * 1.2)
							.replaceAll(",", ".").replaceAll(" ", ""));

					computationPositionContractPriceMap.put(count,
							computationPositionContractPrice);
					count++;
				}
			}
			computationPositionContractPriceMap
					.forEach((k, v) -> System.out.println(k + " : " + v));
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}
