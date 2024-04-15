package com.computation.estimate;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.computation.estimate.entity.Compulation;
import com.computation.estimate.entity.CompulationPosition;
import com.computation.estimate.entity.CompulationPositionUnitOfMeasurement;

public class GetDataFromExcelFile {
	public static void main(String[] args) {
		try (Workbook workbook = WorkbookFactory
				.create(new FileInputStream("files/451_du.xlsx"))) {
//			var compulations = getCompulations(workbook);
			getCompulationPositionUnitOfMeasurements(workbook);
//			var compulationPositions = getCompulationPosition(workbook);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static List<Compulation> getCompulations(Workbook workbook) {
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

	private static List<CompulationPosition> getCompulationPosition(
			Workbook workbook) {
		return null;
	}

	private static List<CompulationPositionUnitOfMeasurement> getCompulationPositionUnitOfMeasurements(
			Workbook workbook) {
		Sheet sheet = workbook.getSheet("6_Позиція_ЛК");
		Set<String> unitOfMeasurementNameSet = new HashSet<>();
		List<CompulationPositionUnitOfMeasurement> compulationPositionUnitOfMeasurements = new ArrayList<>();
		int lastRowNum = sheet.getLastRowNum();

		for (int i = 1; i < lastRowNum + 1; i++) {
			Row row = sheet.getRow(i);
			unitOfMeasurementNameSet.add(row.getCell(8).getStringCellValue());
		}

		unitOfMeasurementNameSet
				.forEach(u -> compulationPositionUnitOfMeasurements
						.add(new CompulationPositionUnitOfMeasurement(u)));
		compulationPositionUnitOfMeasurements.forEach(System.out::println);

		return compulationPositionUnitOfMeasurements;

	}
}