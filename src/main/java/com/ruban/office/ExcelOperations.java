package com.ruban.office;

import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelOperations {

	private static final String FileName = "Grade_Report.xlsx";

	private static final List<StudentModel> modelList = new ArrayList<StudentModel>();

	private static final List<StudentModel> modelGrade = new ArrayList<StudentModel>();

	private static final List<String> headings = new ArrayList<String>();

	public static void main(String[] args) {

		try {
			int row = 0;
			FileInputStream excelFile = new FileInputStream(new File(FileName));

			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = datatypeSheet.iterator();

			while (iterator.hasNext()) {
				Row currentRow = iterator.next();
				if (row != 0) {
					Iterator<Cell> cellIterator = currentRow.iterator();
					StudentModel model = new StudentModel();
					while (cellIterator.hasNext()) {
						Cell currentCell = cellIterator.next();
						if (currentCell.getCellType() == Cell.CELL_TYPE_STRING) {
							model.setName(currentCell.getStringCellValue());
							System.out.print(currentCell.getStringCellValue() + "--");
						} else if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							model.setMarks(currentCell.getNumericCellValue());
							System.out.print(currentCell.getNumericCellValue() + "--");
						}
					}
					modelList.add(model);
				} else {
					Iterator<Cell> cellIterator = currentRow.iterator();
					while (cellIterator.hasNext()) {
						Cell currentCell = cellIterator.next();
						headings.add(currentCell.getStringCellValue());
						System.out.print(currentCell.getStringCellValue() + "--");
					}
				}
				row++;
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		for (StudentModel model : modelList) {
			if (model.getMarks() > 75) {
				model.setGrade("A");
			} else if (model.getMarks() > 50) {
				model.setGrade("B");
			} else if (model.getMarks() > 30) {
				model.setGrade("C");
			} else {
				model.setGrade("F");
			}
			modelGrade.add(model);
		}

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("After Grading");

		int rowNum = 0;

		Row row = sheet.createRow(rowNum++);
		Cell cells1 = row.createCell(0);
		cells1.setCellValue((String) headings.get(0));
		Cell cells2 = row.createCell(1);
		cells2.setCellValue((String) headings.get(1));
		Cell cells3 = row.createCell(2);
		cells3.setCellValue((String) headings.get(2));

		System.out.println("Creating excel");

		for (StudentModel model : modelGrade) {

			Row rows = sheet.createRow(rowNum++);
			Cell cell1 = rows.createCell(0);
			cell1.setCellValue((String) model.getName());
			Cell cell2 = rows.createCell(1);
			cell2.setCellValue((Double) model.getMarks());
			Cell cell3 = rows.createCell(2);
			cell3.setCellValue((String) model.getGrade());

		}

		try {
			FileOutputStream outputStream = new FileOutputStream(FileName);
			workbook.write(outputStream);
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		System.out.println("Done");
	}

}
