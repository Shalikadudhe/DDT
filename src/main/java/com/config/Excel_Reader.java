package com.config;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Reader {

	public Object[][] loadSheet(String fileNm, String sheetNm) throws IOException {
		Object[][] obj = null;
		Workbook wb = null;

		FileInputStream fis = new FileInputStream(fileNm);

		/*
		 * String extension = fileNm.split("\\.")[1]; if (extension.equals("xls")) { wb
		 * = new HSSFWorkbook(fis); } else if (extension.equals("xlsx")) { wb = new
		 * XSSFWorkbook(fis); } Sheet sh = wb.getSheet(sheetNm); int rows =
		 * sh.getLastRowNum(); obj = new Object[rows][sh.getRow(0).getLastCellNum() -
		 * 1]; // size of excel sheet data
		 * 
		 * for (int i = 1; i <= rows; i++) { Row row = sh.getRow(i); int n_cells =
		 * row.getLastCellNum(); for (int j = 1; j < n_cells; j++) {
		 * 
		 * Cell cell = row.getCell(j); obj[i - 1][j - 1] = cell.getStringCellValue(); }
		 * }
		 * 
		 */
		if (fileNm.endsWith(".xlsx")) {
			wb = new XSSFWorkbook(fis);
		} else if (fileNm.endsWith(".xls")) {
			wb = new HSSFWorkbook(fis);
		}
		Sheet sheet = wb.getSheet(sheetNm);
		int N_row = sheet.getLastRowNum();
		System.out.println(N_row);
		obj = new Object[N_row][sheet.getRow(1).getLastCellNum() - 1];

//		System.out.println(obj.length);
		for (int i = 1; i <= N_row; i++) {

			Row row = sheet.getRow(i);
			int N_cell = row.getLastCellNum();

			for (int j = 1; j < N_cell; j++) {
				Cell cell = row.getCell(j);
				obj[i - 1][j - 1] = cell.getStringCellValue();

				// cell=row.createCell(j, cell.getCellType());
				// cell.setCellValue(result);
			}
		}

		return obj;
	}

/*	public void writeResult(String fileNm, String sheetNm, String result)
			throws EncryptedDocumentException, IOException {
		try {
			FileInputStream fis = new FileInputStream(fileNm);
			Workbook wb = WorkbookFactory.create(fis);
			Sheet sh = wb.getSheet(sheetNm);

			Row r = sh.getRow(1);
			Cell cc = r.createCell(3);
			cc.setCellValue(result);
			FileOutputStream fos = new FileOutputStream(fileNm);
			wb.write(fos);
			fos.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
*/
	}

