package com.computation.estimate;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.computation.estimate.entity.Compulation;

public class GetDataFromExcelFile {
	public static void main(String[] args) {
		try (Workbook workbook = WorkbookFactory
				.create(new FileInputStream("files/451_du.xlsx"))) {
			var compulations = getCompulation(workbook);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static List<Compulation> getCompulation(Workbook workbook) {
		Sheet sheet = workbook.getSheet("5_Загальні_дані_про_ЛК");

		int lastRowNum = sheet.getLastRowNum();
		List<Compulation> compulations = new ArrayList<>();

		for (int i = 1; i < lastRowNum + 1; i++) {

			Row row = sheet.getRow(i);

			Compulation compulation = new Compulation(
					(int) row.getCell(0).getNumericCellValue(),
					row.getCell(5).getStringCellValue(),
					row.getCell(6).getStringCellValue(),
					row.getCell(7).getNumericCellValue());

			compulations.add(compulation);
		}

		compulations.forEach(c -> System.out.println(c.getCompulationId() + " "
				+ c.getCompulationNumber() + " " + c.getCompulationName() + " "
				+ c.getCompulationValue()));
		return compulations;
	}
}