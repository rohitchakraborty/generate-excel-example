package com.example.rohit;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFBorderFormatting;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Rohit Chakraborty
 *
 */
public class App {

	private static final int ROW_GAP = 30;
	private static final int TABLE_SIZE = 25;

	private static final String CREDITOR = "Creditor";
	private static final String TABLE_STYLE = "TableStyleMedium2";
	private static final String DATE_FORMAT = "dd-mm-yyyy";
	private static final String NUMBER_FORMAT = "#,##0;[RED]#,##0";
	private static final String FIRST_NUMERIC_COLUMN = "D";
	private static final String SECOND_NUMERIC_COLUMN = "E";
	private static final String THIRD_NUMERIC_COLUMN = "F";

	private static final List<String> TABLE_HEADERS = Arrays.asList("Agreement_No.", "Project_Name", "Due_Date",
			"Principal_Amount_(Rs.)", "Interest_Amount_(Rs.)", "Other_Charges_(Rs.)", "Total_Claim_(Rs.)", "Status",
			"Reason", "Timestamp");

	public static void main(String[] args) {
		var workbook = new XSSFWorkbook();
		XSSFSheet mainSheet = workbook.createSheet("Main");
		XSSFSheet creditorSheet = workbook.createSheet(CREDITOR);
		List<Creditor> creditors = getCreditors();
		createMainSheet(workbook, mainSheet, creditors);
		createCreditorTable(workbook, creditorSheet, creditors);
		addConditionalFormatting(mainSheet);
		addNumericConditionalFormatting(mainSheet, FIRST_NUMERIC_COLUMN);
		addNumericConditionalFormatting(mainSheet, SECOND_NUMERIC_COLUMN);
		addNumericConditionalFormatting(mainSheet, THIRD_NUMERIC_COLUMN);

		try (var outputStream = new FileOutputStream("JavaBooks.xlsx")) {
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
		XSSFDataFormat dataFormat = workbook.createDataFormat();
		XSSFFont font = setFont(workbook);
		Map<String, XSSFCellStyle> styleMap = Map.of("style", workbook.createCellStyle(), "tableRowStyle",
				workbook.createCellStyle(), "tableHeaderStyle", workbook.createCellStyle(), "dateCellStyle",
				workbook.createCellStyle(), "numberCellStyle", workbook.createCellStyle(), "lockedCellStyle",
				workbook.createCellStyle());
		setStyles(workbook, styleMap, dataFormat, font);
		var validationHelper = new XSSFDataValidationHelper(sheet);
		for (int i = 0; i < creditors.size(); i++) {
			// Creditor name row
			XSSFRow row = sheet.createRow(rowIndex);
			XSSFCell cell = row.createCell(columnIndex);
			setCellFont(workbook, cell);
			cell.setCellValue(CREDITOR);
			cell = row.createCell(columnIndex + 1);
			setCellFont(workbook, cell);
			styleMap.get("style").setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			styleMap.get("style").setFillPattern(FillPatternType.SOLID_FOREGROUND);
			setBorder(styleMap.get("style"));
			cell.setCellStyle(styleMap.get("style"));
			cell.setCellValue(creditors.get(i).getCreditorName());

			// Set which area the table should be placed in
			var reference = workbook.getCreationHelper().createAreaReference(new CellReference(rowIndex + 1, 0),
					new CellReference(rowIndex + TABLE_SIZE, TABLE_HEADERS.size() - 1));
			XSSFTable table = sheet.createTable(reference);
			table.getCTTable().addNewTableStyleInfo();
			table.getCTTable().getTableStyleInfo().setName(TABLE_STYLE);
			// Style the table
			setTableStyle(table);

			// Define data validation
			var dataValidation = addDataValidation(validationHelper, rowIndex, FIRST_NUMERIC_COLUMN, 3);
			sheet.addValidationData(dataValidation);
			dataValidation = addDataValidation(validationHelper, rowIndex, SECOND_NUMERIC_COLUMN, 4);
			sheet.addValidationData(dataValidation);
			dataValidation = addDataValidation(validationHelper, rowIndex, THIRD_NUMERIC_COLUMN, 5);
			sheet.addValidationData(dataValidation);

			// Add rows and columns
			row = sheet.createRow(rowIndex + 1);
			font.setFontHeightInPoints((short) 12);
			for (int j = 0; j < TABLE_HEADERS.size(); j++) {
				cell = row.createCell(j);
				cell.setCellStyle(styleMap.get("tableHeaderStyle"));
				cell.setCellValue(TABLE_HEADERS.get(j));
			}
			for (int j = rowIndex + 2; j < rowIndex + TABLE_SIZE + 1; j++) {
				row = sheet.createRow(j);
				for (int k = 0; k < TABLE_HEADERS.size(); k++) {
					cell = row.createCell(k);
					if (k == 2) {
						cell.setCellStyle(styleMap.get("dateCellStyle"));
					} else if (k > 2 && k < 7) {
						cell.setCellStyle(styleMap.get("numberCellStyle"));
						if (k == TABLE_HEADERS.size() - 4) {
							// SUM formula
							cell.setCellFormula("SUM(" + FIRST_NUMERIC_COLUMN + (j + 1) + ":" + THIRD_NUMERIC_COLUMN
									+ (j + 1) + ")");
							cell.setCellType(CellType.FORMULA);
						}
					} else if (k > 6) {
						cell.setCellStyle(styleMap.get("lockedCellStyle"));
					} else {
						cell.setCellStyle(styleMap.get("tableRowStyle"));
					}
				}
			}
			rowIndex += ROW_GAP;
		}
		sheet.lockSelectLockedCells(true);
		sheet.lockAutoFilter(false);
		sheet.enableLocking();
	}

	private static void createCreditorTable(XSSFWorkbook workbook, XSSFSheet sheet, List<Creditor> creditors) {
		// Set which area the table should be placed in
		var reference = workbook.getCreationHelper().createAreaReference(new CellReference(0, 0),
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
		XSSFCellStyle style = workbook.createCellStyle();
		setHeaderStyle(workbook, style);
		for (int i = 0; i < creditors.size() + 1; i++) {
			// Create row
			row = sheet.createRow(i);
			for (int j = 0; j < 2; j++) {
				// Create cell
				cell = row.createCell(j);
				if (i == 0 && j == 0) {
					cell.setCellStyle(style);
					cell.setCellValue(CREDITOR);
				} else if (i == 0 && j == 1) {
					cell.setCellStyle(style);
					cell.setCellValue("ID");
				} else if (j == 0) {
					cell.setCellValue(creditors.get(i - 1).getCreditorName());
				} else if (j == 1) {
					cell.setCellValue(creditors.get(i - 1).getId());
				}
			}
		}
		XSSFName name = workbook.createName();
		name.setNameName("creditorValues");
		name.setRefersToFormula(CREDITOR + "!B2:B" + (creditors.size() + 1));
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
		var creditor1 = new Creditor("Test1", "A1231");
		var creditor2 = new Creditor("Test2", "B2231");
		var creditor3 = new Creditor("Test3", "C3231");
		var creditor4 = new Creditor("Test4", "D4231");
		return Arrays.asList(creditor1, creditor2, creditor3, creditor4);
	}

	private static void setCellFont(XSSFWorkbook workbook, XSSFCell cell) {
		XSSFFont font = setFont(workbook);
		font.setFontHeightInPoints((short) 13);
		XSSFCellStyle style = workbook.createCellStyle();
		style.setFont(font);
		cell.setCellStyle(style);
	}

	private static void setHeaderStyle(XSSFWorkbook workbook, XSSFCellStyle style) {
		XSSFFont font = setFont(workbook);
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setFontHeightInPoints((short) 12);
		style.setFont(font);
		style.setAlignment(HorizontalAlignment.CENTER);
	}

	private static XSSFFont setFont(XSSFWorkbook workbook) {
		XSSFFont font = workbook.createFont();
		font.setFontName("Times New Roman");
		return font;
	}

	private static void setBorder(XSSFCellStyle style) {
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
	}

	private static DataValidation addDataValidation(XSSFDataValidationHelper validationHelper, int rowIndex,
			String column, int columnIndex) {
		var constraint = validationHelper.createCustomConstraint("=ISNUMBER(" + column + (rowIndex + 3) + ")");
		var cellRange = new CellRangeAddressList(rowIndex + 2, rowIndex + TABLE_SIZE, columnIndex, columnIndex);
		var dataValidation = validationHelper.createValidation(constraint, cellRange);
		dataValidation.setSuppressDropDownArrow(true);
		dataValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
		dataValidation.createErrorBox("Invalid Amount", "Please enter valid amount");
		dataValidation.setShowErrorBox(true);
		return dataValidation;
	}

	private static void setMainColumnWidth(XSSFSheet sheet) {
		sheet.setColumnWidth(0, 25 * 256);
		sheet.setColumnWidth(1, 35 * 256);
		sheet.setColumnWidth(2, 15 * 256);
		sheet.setColumnWidth(3, 30 * 256);
		sheet.setColumnWidth(4, 30 * 256);
		sheet.setColumnWidth(5, 30 * 256);
		sheet.setColumnWidth(6, 30 * 256);
		sheet.setColumnWidth(7, 20 * 256);
		sheet.setColumnWidth(8, 35 * 256);
		sheet.setColumnWidth(9, 30 * 256);
	}

	private static void setStyles(XSSFWorkbook workbook, Map<String, XSSFCellStyle> styleMap, XSSFDataFormat dataFormat,
			XSSFFont font) {
		styleMap.get("tableRowStyle").setFont(font);
		styleMap.get("dateCellStyle").setFont(font);
		styleMap.get("numberCellStyle").setFont(font);
		styleMap.get("lockedCellStyle").setFont(font);
		styleMap.get("dateCellStyle").setDataFormat(dataFormat.getFormat(DATE_FORMAT));
		styleMap.get("numberCellStyle").setDataFormat(dataFormat.getFormat(NUMBER_FORMAT));
		setHeaderStyle(workbook, styleMap.get("tableHeaderStyle"));
		setBorder(styleMap.get("tableHeaderStyle"));
		setBorder(styleMap.get("tableRowStyle"));
		setBorder(styleMap.get("dateCellStyle"));
		setBorder(styleMap.get("numberCellStyle"));
		setBorder(styleMap.get("lockedCellStyle"));
		styleMap.get("tableRowStyle").setLocked(false);
		styleMap.get("dateCellStyle").setLocked(false);
		styleMap.get("numberCellStyle").setLocked(false);
	}

	private static void addConditionalFormatting(XSSFSheet sheet) {
		var rowIndex = 0;
		List<XSSFTable> tables = sheet.getTables();
		for (var i = 0; i < tables.size(); i++) {
			XSSFSheetConditionalFormatting conditionalFormatting = sheet.getSheetConditionalFormatting();
			XSSFConditionalFormattingRule rule = conditionalFormatting.createConditionalFormattingRule(
					"=ISBLANK(B" + (rowIndex + 3) + ":B" + (rowIndex + TABLE_SIZE + 1) + ")");
			XSSFBorderFormatting borderFormatting = rule.createBorderFormatting();
			borderFormatting.setBorderBottom(BorderStyle.THIN);
			borderFormatting.setBottomBorderColor(IndexedColors.RED.getIndex());
			borderFormatting.setBorderRight(BorderStyle.THIN);
			borderFormatting.setRightBorderColor(IndexedColors.RED.getIndex());
			borderFormatting.setBorderLeft(BorderStyle.THIN);
			borderFormatting.setLeftBorderColor(IndexedColors.RED.getIndex());
//			borderFormatting.setBorderTop(BorderStyle.THIN);
//			borderFormatting.setTopBorderColor(IndexedColors.RED.getIndex());
			CellRangeAddress[] regions = {
					CellRangeAddress.valueOf("B" + (rowIndex + 3) + ":B" + (rowIndex + TABLE_SIZE + 1)) };
			conditionalFormatting.addConditionalFormatting(regions, rule);
			rowIndex += ROW_GAP;
		}
	}
	
	private static void addNumericConditionalFormatting(XSSFSheet sheet, String columnRef) {
		var rowIndex = 0;
		List<XSSFTable> tables = sheet.getTables();
		for (var i = 0; i < tables.size(); i++) {
			XSSFSheetConditionalFormatting conditionalFormatting = sheet.getSheetConditionalFormatting();
			XSSFConditionalFormattingRule rule = conditionalFormatting.createConditionalFormattingRule(
					"=AND(D" + (rowIndex + 3) + "=\"\", E" + (rowIndex + 3) + "=\"\", F" + (rowIndex + 3) + "=\"\")");
			XSSFBorderFormatting borderFormatting = rule.createBorderFormatting();
			borderFormatting.setBorderBottom(BorderStyle.THIN);
			borderFormatting.setBottomBorderColor(IndexedColors.RED.getIndex());
			borderFormatting.setBorderRight(BorderStyle.THIN);
			borderFormatting.setRightBorderColor(IndexedColors.RED.getIndex());
			borderFormatting.setBorderLeft(BorderStyle.THIN);
			borderFormatting.setLeftBorderColor(IndexedColors.RED.getIndex());
//			borderFormatting.setBorderTop(BorderStyle.THIN);
//			borderFormatting.setTopBorderColor(IndexedColors.RED.getIndex());
			CellRangeAddress[] regions = {
					CellRangeAddress.valueOf(columnRef + (rowIndex + 3) + ":" + columnRef + (rowIndex + TABLE_SIZE + 1)) };
			conditionalFormatting.addConditionalFormatting(regions, rule);
			rowIndex += ROW_GAP;
		}
	}

}
