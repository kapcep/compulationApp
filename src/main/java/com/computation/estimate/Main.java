package com.computation.estimate;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.computation.estimate.service.GetComputationInfoFromExcelFile;

public class Main {

	final static String outboxExcelFilePath = "files/451_du.xlsx";

	public static void main(String[] args) {
		try (Workbook workbook = WorkbookFactory
				.create(new FileInputStream(outboxExcelFilePath))) {

			GetComputationInfoFromExcelFile getDataFromExcelFile = new GetComputationInfoFromExcelFile();

			var computationPositions = getDataFromExcelFile
					.getComputationPosition(workbook);
			var resourceDescriptions = getDataFromExcelFile
					.getResourceDescription(workbook);
			getDataFromExcelFile.getComputationPositionContractPrice();

			getDataFromExcelFile
					.writeComputationPositionsToTextFile(computationPositions);
			getDataFromExcelFile
					.writeResourceDescriptionsToTextFile(resourceDescriptions);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
