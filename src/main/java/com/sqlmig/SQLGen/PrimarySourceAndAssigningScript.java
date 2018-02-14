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

public class PrimarySourceAndAssigningScript {
public static void readXLSXFile() throws IOException {
		
		
		
		
		InputStream ExcelFileToRead = new FileInputStream("PV Data Migration Sheet-CNS.xlsx");
		
		FileWriter fileWriter = new FileWriter("newfsql.txt");
		PrintWriter printWriter = new PrintWriter(fileWriter);
		
		
		
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

		XSSFWorkbook test = new XSSFWorkbook();

		XSSFSheet sheet = wb.getSheetAt(2);
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
			String region = null;
			String reps = "";
			while (cells.hasNext()) {
				cell = (XSSFCell) cells.next();
                switch (cellCount) {
				case 0:
					firstName = cell.getStringCellValue();
					break;
				case 1:
					lastName = cell.getStringCellValue();
					break;
				case 2:
					speciality = cell.getStringCellValue();
					break;
				case 3:
					region = cell.getStringCellValue();
					break;
			    }
                if(cellCount>3) {
                	if(cell.getStringCellValue()!=null && !cell.getStringCellValue().trim().isEmpty())
                	    reps += "'"+cell.getStringCellValue()+"',";
                }
			
					cellCount++;
			}
			
			printWriter.printf("INSERT INTO tbl_pv_primary_source (`reporter_givename`,`reporter_familyname`,`specialization`,`reporter_city`,`qualification_id`,`workspace_module_user_id`,`reporter_country_id`,`status`)\n" + 
			"    VALUES('%s','%s','%s','%s',1,2,61,'active');\n",firstName,lastName,speciality,region);
			
			
			System.out.println();
			if(reps.length()>1) {
				String freps=  reps.substring(0, reps.length()-1);
				printWriter.printf("INSERT INTO tbl_pv_workspace_module_user_primary_source (`workspace_module_user_id`,`primary_source_id`) \n" + 
						"SELECT  wmu.id,LAST_INSERT_ID() \n" + 
						"FROM tbl_continuum_workspace_module_user as wmu inner JOIN tbl_continuum_user as couser \n" + 
						"on wmu.user_id = couser.id \n" + 
						"WHERE wmu.workspace_module_id = 2 and  wmu.status = 'active' and wmu.role_id = 2  and couser.email in (%s) ;\n\n\n",freps);
			}
		
			System.out.println();
			rowCount++;
		}
		printWriter.close();

	}
}
