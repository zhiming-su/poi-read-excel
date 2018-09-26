package com.sprint.batch.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Pattern;

public class readExcel {
	public static final String SAMPLE_XLSX_FILE_PATH = "./a.xlsx";


	public static void read() throws FileNotFoundException, IOException {

		// Creating a Workbook from an Excel file (.xls or .xlsx)
		Workbook workbook = new XSSFWorkbook(new FileInputStream(SAMPLE_XLSX_FILE_PATH));

		// Retrieving the number of sheets in the Workbook
		// System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets
		// : ");

		/*
		 * ((Iterable<XSSFSheet>) workbook).forEach(sheet -> { System.out.println("=> "
		 * + sheet.getSheetName()); });
		 */
		/*
		 * ============================================================= Iterating over
		 * all the sheets in the workbook (Multiple ways)
		 * =============================================================
		 */
		int num = workbook.getNumberOfSheets();
		for (int page = 0; page < num; page++) {
			Sheet sheet = workbook.getSheetAt(page);
			// 遍历全部非空行
			//DataFormatter dataFormatter = new DataFormatter();
			String fileName = "G:\\sqlFIle\\" + workbook.getSheetName(page) + ".sql";
			File file = new File(fileName);
			if (!file.exists()) {
				// System.out.println(fileName);
				file.createNewFile();
			}

			FileOutputStream fot = new FileOutputStream(file);
			StringBuilder sbPk = new StringBuilder();
			StringBuilder sb1 = new StringBuilder();
			
			for (Row row : sheet) {				
				 String clumnName = null;
				 String clumnType = null;
				 String clumnZhushi = null;
				 String pk =null;
				for (Cell cell : row) {
				//	StringBuilder sb = new StringBuilder();
					if (row.getRowNum() == 1 && cell.getColumnIndex() == 1) {
						String tableName = cell.toString();
						String dropInfo = "DROP TABLE IF EXISTS " + tableName + " ;";
						String createTitle = "CREATE TABLE " + tableName + " (\n";
						sb1.append(dropInfo);
						sb1.append("\n");
						sb1.append(createTitle);
						fot.write(sb1.toString().getBytes());
					}
					if (row.getRowNum() >= 3 && cell.getColumnIndex() == 0) {
						clumnZhushi = cell.toString();
					}
					if (row.getRowNum() >= 3 && cell.getColumnIndex() == 1) {
						clumnName = cell.toString();
					}
					if (row.getRowNum() >= 3 && cell.getColumnIndex() == 2) {
						clumnType = cell.toString();
					}
					if (row.getRowNum() >= 3 && cell.getColumnIndex() == 5) {
						pk = cell.toString();						
						if (pk.equals("主键")) {
							sbPk.append(clumnName + ",");
						}						
					}	
				}
				if (clumnName != null) {

					if (Pattern.compile(clumnName).matcher(sbPk.toString()).find()) {
						String lineInfo = clumnName + "\t\t" + clumnType + " NOT NULL COMMENT '" + clumnZhushi
								+ "',\n";
						fot.write(lineInfo.getBytes());
					} else if (clumnName.equals("INDATE")) {
						String lineInfo = clumnName + "\t\t" + clumnType + " default CURRENT_TIMESTAMP COMMENT '"
								+ clumnZhushi + "',\n";
						fot.write(lineInfo.getBytes());						
					} else if (clumnName.equals("MODIFY_TIME")) {
						String lineInfo = clumnName + "\t\t" + clumnType
								+ " default CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '" + clumnZhushi
								+ "',\n";
						fot.write(lineInfo.getBytes());						
					} else {
						String lineInfo = clumnName + "\t\t" + clumnType + " COMMENT '" + clumnZhushi + "',\n";
						fot.write(lineInfo.getBytes());
					}
					
					
				}
				//fot.write(sb.toString().getBytes());
				//fot.write(sb1.toString().getBytes());
				if (sheet.getLastRowNum() == row.getRowNum()) {
					
					String primaryKey = "primary key (" + sbPk.toString().replaceAll(",$", "") + ")" + "\n" + ");";
					fot.write(primaryKey.getBytes());
					fot.close();
					
				}
				
			}
			

		}
	}
}