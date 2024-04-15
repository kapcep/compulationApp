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

import com.computation.estimate.entity.Computation;
import com.computation.estimate.entity.ComputationPosition;
import com.computation.estimate.entity.ComputationPositionUnitOfMeasurement;
import com.computation.estimate.entity.ComputationTypePosition;

public class GetDataFromExcelFile {
	final static String generalDataOfComputationSheetName = "5_Загальні_дані_про_ЛК";
	final static String computationPositionSheetName = "6_Позиція_ЛК";

	public static void main(String[] args) {
		try (Workbook workbook = WorkbookFactory
				.create(new FileInputStream("files/451_du.xlsx"))) {

			var computationPositions = getComputationPosition(workbook);
			computationPositions.forEach(System.out::println);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static List<Computation> getComputations(Workbook workbook) {

		Sheet sheet = workbook.getSheet(generalDataOfComputationSheetName);

		int lastRowNum = sheet.getLastRowNum();
		List<Computation> computations = new ArrayList<>();

		for (int i = 1; i < lastRowNum + 1; i++) {

			Row row = sheet.getRow(i);

			Computation computation = new Computation(
					(int) row.getCell(0).getNumericCellValue(),
					row.getCell(5).getStringCellValue(),
					row.getCell(6).getStringCellValue(),
					row.getCell(7).getNumericCellValue());

			computations.add(computation);
		}

		return computations;
	}

	private static List<ComputationPosition> getComputationPosition(
			Workbook workbook) {
		Sheet sheet = workbook.getSheet(computationPositionSheetName);
		int lastRowNum = sheet.getLastRowNum();
		List<ComputationPosition> computationPositions = new ArrayList<>();
		var computations = getComputations(workbook);
		var computationPositionUnitOfMeasurements = getComputationPositionUnitOfMeasurements(
				workbook);

		for (int i = 1; i < lastRowNum + 1; i++) {
			Row row = sheet.getRow(i);
			// get computationPositionId
			final int computationPositionId = (int) row.getCell(0)
					.getNumericCellValue();

			// get Computation
			final int computationId = (int) row.getCell(1)
					.getNumericCellValue();
			final Computation computation = computations.stream()
					.filter(c -> c.getComputationId() == computationId)
					.findFirst().get();

			// get ComputationTypePosition
			final String computationTypePositionName = row.getCell(2)
					.getStringCellValue();
			final ComputationTypePosition computationTypePosition = getComputationTypePositionByName(
					computationTypePositionName);

			// get positionNumberInComputation
			final int positionNumberInComputation = (int) row.getCell(4)
					.getNumericCellValue();

			/// get positionCode
			final String positionCode = row.getCell(5).getStringCellValue();

			// get explanation
			final String explanation = row.getCell(6).getStringCellValue();

			// get computationPositionName
			final String computationPositionName = row.getCell(7)
					.getStringCellValue();

			// get computationPositionUnitOfMeasurement
			String computationPositionUnitOfMeasurementCellValue = row
					.getCell(8).getStringCellValue();
			var computationPositionUnitOfMeasurement = computationPositionUnitOfMeasurements
					.stream()
					.filter(c -> c.getComputationPositionUnitOfMeasurementName()
							.equals(computationPositionUnitOfMeasurementCellValue))
					.findFirst().get();

			// get amount
			double amount = row.getCell(9).getNumericCellValue();

			// get computationPositionPriceValue
			double computationPositionPriceValue = row.getCell(12)
					.getNumericCellValue();

			// get ComputationPosition
			ComputationPosition computationPosition = new ComputationPosition(
					computationPositionId, computation, computationTypePosition,
					positionNumberInComputation, positionCode, explanation,
					computationPositionName,
					computationPositionUnitOfMeasurement, amount,
					computationPositionPriceValue);

			// add ComputationPosition to collection
			computationPositions.add(computationPosition);
		}

		return computationPositions;
	}

	private static List<ComputationPositionUnitOfMeasurement> getComputationPositionUnitOfMeasurements(
			Workbook workbook) {

		Sheet sheet = workbook.getSheet(computationPositionSheetName);
		Set<String> unitOfMeasurementNameSet = new HashSet<>();
		List<ComputationPositionUnitOfMeasurement> computationPositionUnitOfMeasurements = new ArrayList<>();
		int lastRowNum = sheet.getLastRowNum();

		for (int i = 2; i < lastRowNum + 1; i++) {
			Row row = sheet.getRow(i);
			unitOfMeasurementNameSet.add(row.getCell(8).getStringCellValue());
		}

		unitOfMeasurementNameSet
				.forEach(u -> computationPositionUnitOfMeasurements
						.add(new ComputationPositionUnitOfMeasurement(u)));

		return computationPositionUnitOfMeasurements;

	}

	private static ComputationTypePosition getComputationTypePositionByName(
			String computationTypePositionName) {
		for (ComputationTypePosition name : ComputationTypePosition.values()) {
			if (name.getComputationTypeName()
					.equals(computationTypePositionName)) {
				return name;
			}
		}
		return null;
	}
}