package com.charter.excelapi;

import java.util.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class CharterExcelCreation {

	public static void main(String args[]) {
		CharterExcelCreation cec = new CharterExcelCreation();
		cec.charterValidationSheet("CHG123", "CST1234", "ACID12345",
				"SCNAHSD123", "SERVTYPE1234", "ServiceID12345", "CMACK12345",
				"MTAMAC12345", "UIMPASS");

	}

	private void charterValidationSheet(String csgOrder, String customerId,
			String accounId, String scenario, String servicetype,
			String ServiceId, String cmmac, String mtamac, String validationType) {
		boolean isCSGOrderExist = false;
		boolean isCustomerIdExist = false;
		boolean isAccounIdExist = false;
		boolean isServiceTypeExist = false;
		String checkRowExistence = null;
		Workbook workbook = null;
		Sheet sheet = null;
		Row headerRow = null;
		Row startContentRow = null;
		Cell cell = null;
		String FileNameAsCurrentDate = new SimpleDateFormat(
				"'_'dd-MM-yyyy'.xls'").format(new Date());
		String fileName = csgOrder.concat(FileNameAsCurrentDate);

		File isFileExist = new File("C:\\" + fileName);

		// Check the File already Exist execute if block otherwise execute else
		// block
		try {
			if (isFileExist.exists() && isFileExist.isFile()) {
				// boolean checkRowExistence=true;
				// List<String> listForRowExistenceCheck = ArrayList<String> ();

				// Workbook is Advance interface which can work both .xls or
				// .xlsx
				// File Format.
				workbook = WorkbookFactory.create(isFileExist);
				sheet = workbook.getSheet("Sheet0");
				headerRow = sheet.getRow(0);
				// Update sheet and add New Row if row not exist.
				Iterator<Row> rowIterator = sheet.iterator();
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						Cell duplicateCellCheacker = cellIterator.next();
						if (duplicateCellCheacker.getStringCellValue().equals(
								csgOrder)) {
							isCSGOrderExist = true;
						}
						if (duplicateCellCheacker.getStringCellValue().equals(
								customerId)) {
							isCustomerIdExist = true;
						}
						if (duplicateCellCheacker.getStringCellValue().equals(
								accounId)) {
							isAccounIdExist = true;
						}
						if (duplicateCellCheacker.getStringCellValue().equals(
								servicetype)) {
							isServiceTypeExist = true;
						}

					}
					if (isCSGOrderExist && isCustomerIdExist && isAccounIdExist
							&& isServiceTypeExist) {
						// listForRowExistenceCheck.add(“Yes”);
						checkRowExistence = "Row Found";
					}

				}

				/*
				 * for(int k=0,k<= listForRowExistenceCheck.size(),k++) {
				 * If(!listForRowExistenceCheck.get(k).equals(“Yes”)) {
				 * checkRowExistence =false; break; } }
				 */
				// If Row Already exist (checkRowExistence=true) then do below
				// execution otherwise execute else block
				if (checkRowExistence.equals("Row Found")) {
					for (int i = 1; i <= sheet.getLastRowNum(); i++) {
						startContentRow = sheet.getRow(i);
						for (int j = 0; j <= startContentRow.getLastCellNum(); j++) {
							// Update the value of cell
							// cell = sheet.getRow(i).getCell(j);
							if (validationType != null) {
								// Check header Already Exist put the value .if
								// not
								// then create the New column header and its
								// value
								if ((headerRow.getCell(j).getStringCellValue())
										.equals(validationType.substring(0, 3)
												+ "Validation")) {
									if (cell.getStringCellValue() == null) {
										cell.setCellValue(validationType
												.substring(3, 7));
									}

								} else {
									// creating new column by Increasing header
									// by
									// +1
									cell = headerRow.createCell(headerRow
											.getLastCellNum() + 1);
									// Set New Column Header Name
									cell.setCellValue(validationType.substring(
											0, 3) + "Validation");
									// Set Value
									cell = startContentRow
											.createCell(startContentRow
													.getLastCellNum() + 1);
									cell.setCellValue(validationType.substring(
											3, 7));
								}

							}
						}
					}
				} else {
					// Creating New Row in Existing Sheet
					createNewRow(sheet, csgOrder, customerId, accounId,
							scenario, servicetype, ServiceId, cmmac, mtamac,
							validationType);

				}
			} else {
				// if Worksheet not exist create new it with Header and value

				workbook = WorkbookFactory.create(isFileExist);

				sheet = workbook.createSheet("Sheet0");
				startContentRow = sheet.createRow(1);
				// Creating New Header Row inside the New Sheet
				createSheetHeader(sheet, cell, headerRow);
				// Creating New Content Row inside the New Sheet
				createNewRow(sheet, csgOrder, customerId, accounId, scenario,
						servicetype, ServiceId, cmmac, mtamac, validationType);
			}

			FileOutputStream outFile = new FileOutputStream(new File("C:\\"
					+ fileName));
			workbook.write(outFile);
			outFile.close();
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}

		catch (EncryptedDocumentException e3) {
			// TODO Auto-generated catch block
			e3.printStackTrace();
		}

		catch (InvalidFormatException e4) {
			// TODO Auto-generated catch block
			e4.printStackTrace();
		}

		writeOverallTestStatus(fileName);
	}

	// Creating Sheet Header

	public void createSheetHeader(Sheet sheet, Cell cell, Row headerRow) {
		headerRow = sheet.createRow(0);
		boolean newColumnFlag = true;
		int initval;
		for (initval = 0; initval <= OSMHeaderValues.values().length; initval++) {
			cell = headerRow.createCell(initval);
			if (OSMHeaderValues.values()[initval].toString().equals(
					"CSG Order ID")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].toString().equals(
					"Customer ID")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}
			if (OSMHeaderValues.values()[initval].toString().equals(
					"Account ID ")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].toString().equals("Scenario")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].toString().equals(
					"Service Type")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].toString().equals(
					"Service ID")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].toString().equals("CMMAC")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].toString().equals("MTAMAC")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].toString().equals(
					"UIM Validation")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}

			if (OSMHeaderValues.values()[initval].toString().equals(
					"PRO Validatoin")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}
			if (OSMHeaderValues.values()[initval].toString().equals(
					"Overall Status")) {
				cell.setCellValue(OSMHeaderValues.values()[initval].toString());
				newColumnFlag = false;
			}
		}
		if (newColumnFlag) {
			cell.setCellValue(OSMHeaderValues.values()[initval].toString());

		}

	}

	// Creating New Row into the Sheet
	public void createNewRow(Sheet sheet, String csgOrder, String customerId,
			String accounId, String scenario, String serviceType,
			String ServiceId, String cmmac, String mtamac, String validationType) {

		// Create a new row in current sheet
		Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
		Row headerRow = sheet.getRow(0);
		Cell cell = null;
		for (int l = 0; l <= newRow.getLastCellNum(); l++) {
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("CSG Order ID")) {

				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(csgOrder);

			}

			if ((headerRow.getCell(l).getStringCellValue())
					.equals("CSG Order ID")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(csgOrder);
			}
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Customer ID")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(customerId);
			}
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Account ID")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(accounId);
			}
			if ((headerRow.getCell(l).getStringCellValue()).equals("Scenario")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(scenario);
			}
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Service Type")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(serviceType);
			}
			if ((headerRow.getCell(l).getStringCellValue())
					.equals("Service ID")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(ServiceId);
			}
			if ((headerRow.getCell(l).getStringCellValue()).equals("CMMAC")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(cmmac);
			}
			if ((headerRow.getCell(l).getStringCellValue()).equals("MTAMAC")) {
				// Create a new cell in current row
				cell = newRow.createCell(l);
				// Set value for new cell
				cell.setCellValue(mtamac);

			}

			if (validationType != null) {
				// Check header Already Exist put the value .if not then create
				// the New column header and its value
				if ((headerRow.getCell(l).getStringCellValue())
						.equals(validationType.substring(0, 3) + "Validation")) {
					cell = newRow.createCell(l);
					cell.setCellValue(validationType.substring(3, 7));
				} else {
					// creating new column by Increasing header by +1
					cell = headerRow.createCell(headerRow.getLastCellNum() + 1);
					// Set New Column Header Name
					cell.setCellValue(validationType.substring(0, 3)
							+ "Validation");
					// Set Value
					cell = newRow.createCell(newRow.getLastCellNum() + 1);
					cell.setCellValue(validationType.substring(3, 7));

				}

			}

		}

	}

	// Method Write The OverAll Test Status
	public void writeOverallTestStatus(String fileName) {

		List<String> validationList = new ArrayList<String>();
		StringBuilder testCaseStatus = new StringBuilder();
		FileInputStream excelFile;
		Workbook workbook;
		try {
			excelFile = new FileInputStream(new File("C:\\" + fileName));
			workbook = WorkbookFactory.create(excelFile);

			Sheet sheet = workbook.getSheetAt(0);
			Row headerRow = sheet.getRow(0);
			Cell cell = null;

			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				Row startContentRow = sheet.getRow(i);
				for (int j = 0; j <= startContentRow.getLastCellNum(); j++) {
					if ((headerRow.getCell(j).getStringCellValue().substring(4,
							12)).equals("Validation")) {
						if (startContentRow.getCell(j).getStringCellValue() != null) {

							validationList.add(startContentRow.getCell(j)
									.getStringCellValue());
						}
					}

				}
				for (int inlist = 0; inlist < validationList.size(); inlist++) {
					if (validationList.get(inlist).equalsIgnoreCase("Pass")) {
						testCaseStatus = new StringBuilder("Pass");
					} else {
						testCaseStatus = new StringBuilder("Fail");
						break;
					}

				}
				// Clear List
				validationList.clear();
				// creating new column Overall Status Increasing header by +1
				if (i == 1) {
					cell = headerRow.createCell(headerRow.getLastCellNum() + 1);
					cell.setCellValue("Overall Status");
				}
				cell = startContentRow.createCell(startContentRow
						.getLastCellNum() + 1);
				cell.setCellValue(testCaseStatus.toString());

				FileOutputStream outFile = new FileOutputStream(new File("C:\\"
						+ fileName));
				workbook.write(outFile);
				outFile.close();

			}
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}

		catch (EncryptedDocumentException e3) {
			// TODO Auto-generated catch block
			e3.printStackTrace();
		}

		catch (InvalidFormatException e4) {
			// TODO Auto-generated catch block
			e4.printStackTrace();
		}

	}
}