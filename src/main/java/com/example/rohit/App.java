package com.example.rohit;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Rohit Chakraborty
 *
 */
public class App {

	private static final int ROW_GAP = 20;
	private static final int TABLE_SIZE = 16;

	private static final String CREDITOR = "Creditor";
	private static final String TABLE_STYLE = "TableStyleMedium2";

	private static final List<String> TABLE_HEADERS = Arrays.asList("Agreement_No.", "Project_Name", "Due_Date",
			"Principal_Amount_(Rs.)", "Interest_Amount_(Rs.)", "Other_Charges_(Rs.)", "Total_(Rs.)");

	public static void main(String[] args) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet mainSheet = workbook.createSheet("Main");
		XSSFSheet creditorSheet = workbook.createSheet(CREDITOR);
		List<Creditor> creditors = getCreditors();
		createMainSheet(workbook, mainSheet, creditors);
		createCreditorTable(workbook, creditorSheet, creditors);

		try (FileOutputStream outputStream = new FileOutputStream("JavaBooks.xlsx")) {
			workbook.write(outputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void createMainSheet(XSSFWorkbook workbook, XSSFSheet sheet, List<Creditor> creditors) {
		int columnIndex = 0;
		int rowIndex = 0;
		setMainColumnWidth(sheet);
		for (int i = 0; i < creditors.size(); i++) {
			XSSFRow row = sheet.createRow(rowIndex);
			XSSFCell cell = row.createCell(columnIndex);
			setCellFont(workbook, cell);
			cell.setCellValue(CREDITOR);
			cell = row.createCell(columnIndex + 1);
			setCellFont(workbook, cell);
			cell.setCellValue(creditors.get(i).getCreditorName());

			row = sheet.createRow(rowIndex + 1);
			XSSFFont font = setFont(workbook);
			font.setFontHeightInPoints((short) 12);
			XSSFCellStyle tableRowStyle = workbook.createCellStyle();
			tableRowStyle.setFont(font);
			XSSFCellStyle style = setHeaderStyle(workbook);
			for (int j = 0; j < TABLE_HEADERS.size(); j++) {
				cell = row.createCell(j);
				cell.setCellStyle(style);
				cell.setCellValue(TABLE_HEADERS.get(j));
			}
			for (int j = rowIndex + 2; j < rowIndex + TABLE_SIZE + 1; j++) {
				row = sheet.createRow(j);
				for (int k = 0; k < TABLE_HEADERS.size(); k++) {
					cell = row.createCell(k);
					cell.setCellStyle(tableRowStyle);
				}
			}
			// Set which area the table should be placed in
			AreaReference reference = workbook.getCreationHelper().createAreaReference(
					new CellReference(rowIndex + 1, 0),
					new CellReference(rowIndex + TABLE_SIZE, TABLE_HEADERS.size() - 1));
			XSSFTable table = sheet.createTable(reference);
			table.getCTTable().addNewTableStyleInfo();
			table.getCTTable().getTableStyleInfo().setName(TABLE_STYLE);
			// Style the table
			setTableStyle(table);
			rowIndex += ROW_GAP;
		}
	}

	private static void createCreditorTable(XSSFWorkbook workbook, XSSFSheet sheet, List<Creditor> creditors) {
		// Set which area the table should be placed in
		AreaReference reference = workbook.getCreationHelper().createAreaReference(new CellReference(0, 0),
				new CellReference(creditors.size(), 1));
		XSSFTable table = sheet.createTable(reference);
		sheet.setColumnWidth(0, 25 * 256);
		table.setName("Creditors");
		table.setDisplayName("creditors");
		table.getCTTable().addNewTableStyleInfo();
		table.getCTTable().getTableStyleInfo().setName(TABLE_STYLE);

		// Style the table
		setTableStyle(table);

		// Set the values for the table
		XSSFRow row;
		XSSFCell cell;
		for (int i = 0; i < creditors.size() + 1; i++) {
			// Create row
			row = sheet.createRow(i);
			for (int j = 0; j < 2; j++) {
				// Create cell
				cell = row.createCell(j);
				if (i == 0 && j == 0) {
					XSSFCellStyle style = setHeaderStyle(workbook);
					cell.setCellStyle(style);
					cell.setCellValue(CREDITOR);
				} else if (i == 0 && j == 1) {
					XSSFCellStyle style = setHeaderStyle(workbook);
					cell.setCellStyle(style);
					cell.setCellValue("ID");
				} else if (j == 0) {
					cell.setCellValue(creditors.get(i - 1).getCreditorName());
				} else if (j == 1) {
					cell.setCellValue(creditors.get(i - 1).getId());
				}
			}
		}
	}

	private static void setTableStyle(XSSFTable table) {
		// Style the table
		XSSFTableStyleInfo style = (XSSFTableStyleInfo) table.getStyle();
		style.setName(TABLE_STYLE);
		style.setShowColumnStripes(false);
		style.setShowRowStripes(true);
		style.setFirstColumn(false);
		style.setLastColumn(false);
	}

	private static List<Creditor> getCreditors() {
		Creditor creditor1 = new Creditor("Test1", "A1231");
		Creditor creditor2 = new Creditor("Test2", "B2231");
		Creditor creditor3 = new Creditor("Test3", "C3231");
		Creditor creditor4 = new Creditor("Test4", "D4231");
		return Arrays.asList(creditor1, creditor2, creditor3, creditor4);
	}

	private static void setCellFont(XSSFWorkbook workbook, XSSFCell cell) {
		XSSFFont font = setFont(workbook);
		font.setFontHeightInPoints((short) 13);
		XSSFCellStyle style = workbook.createCellStyle();
		style.setFont(font);
		cell.setCellStyle(style);
	}

	private static XSSFCellStyle setHeaderStyle(XSSFWorkbook workbook) {
		XSSFFont font = setFont(workbook);
		font.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
		font.setFontHeightInPoints((short) 12);
		XSSFCellStyle style = workbook.createCellStyle();
		style.setFont(font);
		return style;
	}

	private static XSSFFont setFont(XSSFWorkbook workbook) {
		XSSFFont font = workbook.createFont();
		font.setFontName("Times New Roman");
		return font;
	}

	private static void setMainColumnWidth(XSSFSheet sheet) {
		sheet.setColumnWidth(0, 25 * 256);
		sheet.setColumnWidth(1, 35 * 256);
		sheet.setColumnWidth(2, 15 * 256);
		sheet.setColumnWidth(3, 30 * 256);
		sheet.setColumnWidth(4, 30 * 256);
		sheet.setColumnWidth(5, 30 * 256);
		sheet.setColumnWidth(6, 30 * 256);
	}

}
