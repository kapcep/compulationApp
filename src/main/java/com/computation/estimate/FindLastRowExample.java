package com.computation.estimate;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.computation.estimate.entity.Compulation;

public class FindLastRowExample {
	public static void main(String[] args) {
		try (Workbook workbook = WorkbookFactory
				.create(new FileInputStream("files/451_du.xlsx"))) {
			getCompulation(workbook);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void getCompulation(Workbook workbook) {
		Sheet sheet = workbook.getSheet("5_Загальні_дані_про_ЛК");

		int lastRowNum = sheet.getLastRowNum();
		HashMap<Integer, Compulation> cellMap = new HashMap<>();

		for (int i = 1; i < lastRowNum + 1; i++) {

			Row row = sheet.getRow(i);

			Compulation compulation = new Compulation(
					row.getCell(5).getStringCellValue(),
					row.getCell(6).getStringCellValue(),
					row.getCell(7).getNumericCellValue());

			cellMap.put((int) row.getCell(0).getNumericCellValue(),
					compulation);
		}

		cellMap.forEach((k,
				v) -> System.out.println(k + " " + v.getCompulationName() + " "
						+ v.getCompulationNumber() + " "
						+ v.getCompulationValue()));
	}
}