package com.computation.estimate.service;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.computation.estimate.entity.Computation;
import com.computation.estimate.entity.ComputationPosition;
import com.computation.estimate.entity.ComputationPositionUnitOfMeasurement;
import com.computation.estimate.entity.ComputationTypePosition;
import com.computation.estimate.entity.MarkOfResourceDescrition;
import com.computation.estimate.entity.ResourceDescription;
import com.computation.estimate.entity.ResourceDescriptionUnitOfMeasurement;

public class ComputationInfoFromExcelFileGetter {

	private final static String generalDataOfComputationSheetName = "5_Загальні_дані_про_ЛК";
	private final static String computationPositionSheetName = "6_Позиція_ЛК";
	private final static String resourceDescriptionSheetName = "7_Опис_ресурсу";
	private final static String generalDataSheetName = "2_загальні_дані_про_будову";
	private final static String computationPositionFilePath = "files/computationPosition_";
	private final static String resourceDescriptionFilePath = "files/resourceDescription_";
	private String objectName;

	final static String computationPositionContractPriceExcelFilePath = "files/contract_price_file.xls";
	private final Logger logger = LoggerFactory.getLogger(this.getClass());

	public ComputationInfoFromExcelFileGetter(String workbookName)
			throws EncryptedDocumentException, FileNotFoundException,
			IOException {
		HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory
				.create(new FileInputStream(workbookName));
		this.objectName = getObjectName(workbook);
	}

	public String getObjectName() {
		return objectName;
	}

	private List<Computation> getComputations(Workbook workbook) {

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

	public List<ComputationPosition> getComputationPosition(Workbook workbook) {
		Sheet computationPositionSheet = workbook
				.getSheet(computationPositionSheetName);
		int lastRowNum = computationPositionSheet.getLastRowNum();
		List<ComputationPosition> computationPositions = new ArrayList<>();
		var computations = getComputations(workbook);
		var computationPositionUnitOfMeasurements = getComputationPositionUnitOfMeasurements(
				workbook);
		var computationPositionContractPrices = getComputationPositionContractPrice();
		int count = 0;

		for (int i = 1; i < lastRowNum + 1; i++) {
			Row row = computationPositionSheet.getRow(i);
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

			// get computationPositionContractPrice
			double computationPositionContractPrice = 0;

			if (computationPositionContractPrices.size() > count
					&& computationTypePosition.getComputationTypeName()
							.equals(" Поз. Л.С. ")) {
				count++;
				computationPositionContractPrice = computationPositionContractPrices
						.get(count);
			}

			// get ComputationPosition
			ComputationPosition computationPosition = new ComputationPosition(
					computationPositionId, computation, computationTypePosition,
					positionNumberInComputation, positionCode, explanation,
					computationPositionName,
					computationPositionUnitOfMeasurement, amount,
					computationPositionPriceValue,
					computationPositionContractPrice);

			// add ComputationPosition to collection
			computationPositions.add(computationPosition);
		}

		return computationPositions;
	}

	private String getObjectName(Workbook workbook) {
		Sheet generalDataSheet = workbook.getSheet(generalDataSheetName);
		Row row = generalDataSheet.getRow(1);
		Cell cell = row.getCell(5);
		objectName = cell.getStringCellValue();
		return objectName;
	}

	public List<ResourceDescription> getResourceDescription(Workbook workbook) {

		Sheet sheet = workbook.getSheet(resourceDescriptionSheetName);
		int lastRowNum = sheet.getLastRowNum();
		List<ResourceDescription> resourceDescriptions = new ArrayList<>();

		var computationPositions = getComputationPosition(workbook);
		var resourceDescriptionUnitOfMeasurements = getResourceDescriptionUnitOfMeasurements(
				workbook);

		for (int i = 1; i < lastRowNum + 1; i++) {

			Row row = sheet.getRow(i);

			// get resourceDescriptionId
			final int resourceDescriptionId = (int) row.getCell(0)
					.getNumericCellValue();

			// get ComputationPosition
			final int computationPositionId = (int) row.getCell(1)
					.getNumericCellValue();
			final ComputationPosition computationPosition = computationPositions
					.stream()
					.filter(c -> c
							.getComputationPositionId() == computationPositionId)
					.findFirst().get();

			// get MarkOfResourceDescrition
			final String markOfResourceDescritionName = row.getCell(3)
					.getStringCellValue();
			final MarkOfResourceDescrition markOfResourceDescrition = getMarkOfResourceDescritionByName(
					markOfResourceDescritionName);

			// get resourceCode
			final String resourceCode = row.getCell(4).getStringCellValue();

			// get computationResourcePrice
			final double computationResourcePrice = row.getCell(7)
					.getNumericCellValue();

			// get standardConsumptionOfTheResource
			Cell standardConsumptionOfTheResourceCell = row.getCell(11);
			double standardConsumptionOfTheResource = 0;
			if (standardConsumptionOfTheResourceCell.getCellType()
					.equals(CellType.NUMERIC)) {
				standardConsumptionOfTheResource = standardConsumptionOfTheResourceCell
						.getNumericCellValue();
			}

			// get resourceDescriptionName
			final String resourceDescriptionName = row.getCell(13)
					.getStringCellValue();

			// get resourceDescriptionUnitOfMeasurement

			final Cell resourceDescriptionUnitOfMeasurementCell = row
					.getCell(14);

			String resourceDescriptionUnitOfMeasurementCellValue = "";

			if (resourceDescriptionUnitOfMeasurementCell.getCellType()
					.equals(CellType.STRING)) {
				resourceDescriptionUnitOfMeasurementCellValue = resourceDescriptionUnitOfMeasurementCell
						.getStringCellValue();
			}

			final String resourceDescriptionUnitOfMeasurementCellValueFinal = resourceDescriptionUnitOfMeasurementCellValue;

			var resourceDescriptionUnitOfMeasurement = resourceDescriptionUnitOfMeasurements
					.stream()
					.filter(c -> c.getResourceDescriptionUnitOfMeasurementName()
							.equals(resourceDescriptionUnitOfMeasurementCellValueFinal))
					.findFirst().get();

			// get ResourceDescription
			ResourceDescription resourceDescription = new ResourceDescription(
					resourceDescriptionId, computationPosition,
					markOfResourceDescrition, resourceCode,
					computationResourcePrice, standardConsumptionOfTheResource,
					resourceDescriptionName,
					resourceDescriptionUnitOfMeasurement);

			resourceDescriptions.add(resourceDescription);

		}

		return resourceDescriptions;
	}

	private List<ComputationPositionUnitOfMeasurement> getComputationPositionUnitOfMeasurements(
			Workbook workbook) {

		Sheet sheet = workbook.getSheet(computationPositionSheetName);
		Set<String> unitOfMeasurementNameSet = new HashSet<>();
		List<ComputationPositionUnitOfMeasurement> computationPositionUnitOfMeasurements = new ArrayList<>();
		int lastRowNum = sheet.getLastRowNum();

		for (int i = 1; i < lastRowNum + 1; i++) {
			Row row = sheet.getRow(i);
			unitOfMeasurementNameSet.add(row.getCell(8).getStringCellValue());
		}

		unitOfMeasurementNameSet
				.forEach(u -> computationPositionUnitOfMeasurements
						.add(new ComputationPositionUnitOfMeasurement(u)));

		return computationPositionUnitOfMeasurements;

	}

	private List<ResourceDescriptionUnitOfMeasurement> getResourceDescriptionUnitOfMeasurements(
			Workbook workbook) {

		Sheet sheet = workbook.getSheet(resourceDescriptionSheetName);
		Set<String> unitOfMeasurementNameSet = new HashSet<>();
		List<ResourceDescriptionUnitOfMeasurement> resourceDescriptionUnitOfMeasurements = new ArrayList<>();
		int lastRowNum = sheet.getLastRowNum();

		for (int i = 1; i < lastRowNum + 1; i++) {
			Row row = sheet.getRow(i);
			unitOfMeasurementNameSet.add(row.getCell(14).getStringCellValue());
		}

		unitOfMeasurementNameSet
				.forEach(u -> resourceDescriptionUnitOfMeasurements
						.add(new ResourceDescriptionUnitOfMeasurement(u)));

		return resourceDescriptionUnitOfMeasurements;

	}

	private ComputationTypePosition getComputationTypePositionByName(
			String computationTypePositionName) {
		for (ComputationTypePosition name : ComputationTypePosition.values()) {
			if (name.getComputationTypeName()
					.equals(computationTypePositionName)) {
				return name;
			}
		}
		return null;
	}

	private MarkOfResourceDescrition getMarkOfResourceDescritionByName(
			String markOfResourceDescritionName) {
		for (MarkOfResourceDescrition name : MarkOfResourceDescrition
				.values()) {
			if (name.toString().equals(markOfResourceDescritionName)) {
				return name;
			}
		}
		return null;
	}

	public Map<Integer, Double> getComputationPositionContractPrice() {

		Map<Integer, Double> computationPositionContractPriceMap = new HashMap<>();

		try (Workbook workbook = WorkbookFactory.create(new FileInputStream(
				computationPositionContractPriceExcelFilePath))) {
			Sheet sheet = workbook.getSheetAt(0);
			DecimalFormat df = new DecimalFormat("#,##0.00");

			int count = 1;

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
		} catch (IOException e) {
			e.printStackTrace();
		}

		return computationPositionContractPriceMap;
	}

	public void writeComputationPositionsToTextFile(
			List<ComputationPosition> computationPositions) {

		SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy_HH_mm_ss");

		try (BufferedWriter writer = new BufferedWriter(
				new FileWriter(computationPositionFilePath
						+ sdf.format(new Date()) + ".txt"))) {

			writer.write(
					"|Номер п/п|Кошторис|Тип позиції|№ у ЛК|Шифр позиції|Обґрунтування"
							+ "|Найменування|Одиниця виміру|Кількість|Договірна ціна|Кошторисна ціна|");
			writer.newLine();

			for (ComputationPosition computationPosition : computationPositions) {
				writer.write(computationPosition.toString());
				writer.newLine();
			}
			writer.flush();

			logger.info("File was written to: " + computationPositionFilePath
					+ sdf.format(new Date()) + ".txt");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void writeResourceDescriptionsToTextFile(
			List<ResourceDescription> resourceDescriptions) {

		SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy_HH_mm_ss");

		try (BufferedWriter writer = new BufferedWriter(
				new FileWriter(resourceDescriptionFilePath
						+ sdf.format(new Date()) + ".txt"))) {

			writer.write(
					"|Номер п/п|Назва роботи|Мітка|Шифр ресурсу|Найменування ресурсу"
							+ "|Одиниця виміру|Нормативна витрата ресурсу|Ціна ресурсу|");
			writer.newLine();

			for (ResourceDescription resourceDescription : resourceDescriptions) {
				writer.write(resourceDescription.toString());
				writer.newLine();
			}
			writer.flush();

			logger.info("File was written to: " + resourceDescriptionFilePath
					+ sdf.format(new Date()) + ".txt");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}