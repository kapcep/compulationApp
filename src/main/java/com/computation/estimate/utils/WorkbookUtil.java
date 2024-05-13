package com.computation.estimate.utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookUtil {

	public static XSSFWorkbook returnWorkbook(String workbookPath) {

		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			
			return workbook;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}
}
