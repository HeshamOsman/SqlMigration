package com.sqlmig.SQLGen;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PrimarySource {
	
	public static final int SHEET_NUMBER = 2;
	public static final int FIRST_NAME_CELL_NUMBER = 0;
	public static final int SECOND_NAME_CELL_NUMBER = 1;
	public static final int SPECIALITY_CELL_NUMBER = 2;
//	public static final int QUALIFICATION_CELL_NUMBER = 3;
	public static final int CITY_CELL_NUMBER = 3;
	public static final int MEDICAL_REP_AFTER_CELL_NUMBER = 4;
	
	
	public static final int ADMINSTRATOR_WORKSPACE_MODULE_USER_ID = 25;
	public static final int QUALIFICATION_ID = 1;
	public static final int COUNTRY_ID = 61;
	public static final String STATUS = "active";
	
public static void readXLSXFile() throws IOException {
		
		
		
		
		InputStream ExcelFileToRead = new FileInputStream("PV Data Migration Sheet-CNS.xlsx");
		
		FileWriter fileWriter = new FileWriter("generatedPrimarySourceSQL.txt");
		PrintWriter printWriter = new PrintWriter(fileWriter);
		
		
		
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

		XSSFWorkbook test = new XSSFWorkbook();

		XSSFSheet sheet = wb.getSheetAt(SHEET_NUMBER);
		XSSFRow row;
		XSSFCell cell;

		Iterator<Row> rows = sheet.rowIterator();
		rows.next();
        int rowCount = 0;
		while (rows.hasNext()) {
			
			row = (XSSFRow) rows.next();
			Iterator<Cell> cells = row.cellIterator();
			int cellCount = 0;
			String firstName = null;
			String lastName = null;
			String speciality = null;
			String city = null;
			String reps = "";
			while (cells.hasNext()) {
				cell = (XSSFCell) cells.next();
                switch (cellCount) {
				case FIRST_NAME_CELL_NUMBER:
					firstName = cell.getStringCellValue();
					break;
				case SECOND_NAME_CELL_NUMBER:
					lastName = cell.getStringCellValue();
					break;
				case SPECIALITY_CELL_NUMBER:
					speciality = cell.getStringCellValue();
					break;
//				case QUALIFICATION_CELL_NUMBER:
////					region = cell.getStringCellValue();
//					break;
				case CITY_CELL_NUMBER:
					city = cell.getStringCellValue();
					break;
				default:
					if(cell.getStringCellValue()!=null && !cell.getStringCellValue().trim().isEmpty())
                	    reps += ""+cell.getStringCellValue()+",";
			    }
//                if(cellCount>3) {
//                	
//                }
			
					cellCount++;
			}
			String freps=null;
			if(reps.length()>1) {
//				 reps += "migration@migration.com,";
				 freps=  reps.substring(0, reps.length()-1);
			}
			printWriter.printf("INSERT INTO tbl_pv_primary_source (`reporter_givename`,`reporter_familyname`,`specialization`,`reporter_city`,`qualification_id`,`workspace_module_user_id`,`reporter_country_id`,`status`,`temporary_assigned_users_emails`)\n" + 
			"    VALUES('%s','%s','%s','%s','%s','%s','%s','%s','%s');\n\n\n",firstName,lastName,speciality,city,QUALIFICATION_ID,ADMINSTRATOR_WORKSPACE_MODULE_USER_ID,COUNTRY_ID,STATUS,freps);
			
		
			System.out.println();
			rowCount++;
		}
		printWriter.close();

	}
}
