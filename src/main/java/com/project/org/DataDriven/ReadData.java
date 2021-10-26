package com.project.org.DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	public static void main(String[] args) throws Throwable {
		File f = new File("C:\\Users\\User\\eclipse-workspace1\\DataDriven\\Book1.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);
		
		for (int i = 0; i < sheetAt.getPhysicalNumberOfRows(); i++) {
			Row row = sheetAt.getRow(i);
		for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
			Cell cell = row.getCell(j);
			
		CellType cellType = cell.getCellType();
		if (cellType.equals(cellType.STRING)) {
			String name = cell.getStringCellValue();
			System.out.println(name);
		}
		else if (cellType.equals(cellType.STRING)) {
			String head = cell.getStringCellValue();
			System.out.println(head);
		}
		else if (cellType.equals(cellType.NUMERIC)) {
			double numeric = cell.getNumericCellValue();
			System.out.println(numeric);
			
		}
		else {
			System.out.println("invalid input");
		}
		
		}
			
		}
	}

}
