/**
 * File: EmeryRptSheet.java
 * Description: A wrapper for HSSFSheet, works in tandem with EmeryRptWorkbook.java
 * 				Manages the creation of Excel compatible spreadsheet reports with
 * 				access to predefined cell styles.
 * 
 * Notes: Look into setting the cell formatting by Column instead of for each individual
 * cell.  Common formatting types could set default column widths (ex. Description - 40 chars).
 * But this may be confusing and cause more work for the User. -epearson
 *
 * @author Erik Pearson
 *
 * Create Date: 07/16/2010
 */

package com.emerywaterhouse.rpt.helper;

import java.util.Arrays;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import com.emerywaterhouse.rpt.helper.EmeryRptWorkbook.Format;

public class EmeryRptSheet {
	/**
	 * Represents a Column in a spreadsheet
	 */
	public static class Column{
		private final String m_Field; // field name
		private final int m_Width; // column width in number of chars
		
		// Constructor
		public Column(String field, int width){
			m_Field = field;
			m_Width = width;
		}

		/**
		 * @return the m_Field
		 */
		public String getField() {
			return m_Field;
		}

		/**
		 * @return the m_Format
		 */
		public int getWidth() {
			return m_Width;
		}
	}
	
	/**
	 * Exception thrown when user tries to set a cell
	 * outside the range of the original column
	 * headers (fields)
	 */
	@SuppressWarnings("serial")
	public class NoSuchFieldException extends RuntimeException{
		NoSuchFieldException(String s){
			super(s);
		}
	}

	// cell width in POI is calculated as 1/256 of a character unit; a value of 295
	// seems to provide the best conversion results
	private static final int CHAR_WIDTH = 295;

	// a reference to the EmeryRptWorkbook is required for access to cell styles
	private EmeryRptWorkbook m_Workbook;
	private HSSFSheet m_Sheet;
	private Column[] m_Columns; // column headers
	private HSSFRow m_Row;
	private int m_NumOfRows;
	private String m_ReportTitle;

	/**
	 *  Constructor
	 *  
	 *  EmeryRptSheet requires that the user define all the fields/column headers
	 *  up-front. This limits the spreadsheet to act more like a table in order to
	 *  reduce errors; A runtime-exception will be thrown if the user attempts to
	 *  write outside the predefined columns. 
	 *  
	 * @param workbook		a reference to the sheets container for access to cell styles
	 * @param sheet			an HSSFSheet
	 * @param reportTitle	the title of the report printed on the first line of the report
	 * @param fields	`	these are used for the column headers and represent the 'fields' in
	 * 						a query
	 */
	protected EmeryRptSheet(EmeryRptWorkbook workbook, HSSFSheet sheet, 
							String reportTitle, Column[] columns){
		m_Workbook = workbook;
		m_Sheet = sheet;
		m_ReportTitle = reportTitle;

		m_Columns = Arrays.copyOf(columns, columns.length);
		m_NumOfRows = 0;

		init();
	}

	/**
	 * Initialize workbook
	 */
	private void init() {
		initHeaderRow();
	}

	/**
	 * Set the column titles
	 */
	private void initHeaderRow() {
		
		if(m_ReportTitle != null){
			// Set Report Title
			addRow(m_NumOfRows++);
			addCell(0, m_ReportTitle, Format.HeaderLeft);

			m_NumOfRows++;
			
			// Add an empty row
			addRow(m_NumOfRows++);
		}
		
		// Set Headers
		for(int i = 0; i < m_Columns.length; i++){
			// Set Column width to length of the header string
			m_Sheet.setColumnWidth(i, (m_Columns[i].getWidth() * CHAR_WIDTH));

			addCell(i, m_Columns[i].getField(), Format.HeaderLeftWrap);
		}
	}

	/**
	 * adds a text type cell to current row at the specified column
	 *
	 * @param col     0-based column number of spreadsheet cell
	 * @param value   text value to be stored in cell
	 * @param style   cell style to be used
	 */
	private void addCell(int col, String value, Format format)
	{
		HSSFRichTextString text = new HSSFRichTextString(value);
		HSSFCell cell = m_Row.createCell(col);

		cell.setCellType(HSSFCell.CELL_TYPE_STRING);
		cell.setCellValue(text);
		cell.setCellStyle(m_Workbook.getCellStyle(format));      
	}

	/**
	 * adds a numeric type cell to current row at specified column
	 *
	 * @param col     0-based column number of spreadsheet cell
	 * @param value   numeric value to be stored in cell
	 * @param style   cell style to be used
	 */
	private void addCell(int col, double value, Format format)
	{
		HSSFCell cell = m_Row.createCell(col);
		cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
		cell.setCellStyle(m_Workbook.getCellStyle(format));

		//
		// Math.round(V * D) / D is required to store the
		// true decimal value rather than the Java float value
		// (for example, 2.82 rather than 2.81999993324279)
		switch(format){
		case Number2d:
		case Percent2d:
		case Currency2d:
			cell.setCellValue(Math.round(value * 100d) / 100d); break;
		case Number3d:
		case Percent3d:
		case Currency3d:
			cell.setCellValue(Math.round(value * 1000d) / 1000d); break;
		case Number4d:
		case Percent4d:
		case Currency4d:
			cell.setCellValue(Math.round(value * 10000d) / 10000d); break;
		default:
			cell.setCellValue(value);
		}
	}   

	/**
	 * Adds a row to the end of a worksheet.   
	 */
	public void addRow(){
		m_Row = m_Sheet.createRow(m_NumOfRows++);
	}

	/**
	 * adds row to the current sheet
	 *
	 * @param row  0-based row number of row to be added
	 */
	private void addRow(int row){
		m_Row = m_Sheet.createRow(row);
	}

	/**
	 * Adds a text type cell to the current row at the given column
	 * 
	 * @param column	0-based column number of spreadsheet cell
	 * @param value		text value to be stored in cell
	 * @param wrap		true - wrap text
	 * @throws Exception 
	 */
	public void setField(int column, String value, Boolean wrap){
		validateColumn(column);
		if(wrap){
			addCell(column, value, Format.TextLeftWrap);
		} else {
			addCell(column, value, Format.TextLeft);
		}
	}

	/**
	 * Adds a numeric value to the cell at the given column in
	 * the current row
	 * 
	 * @param column
	 * @param value
	 */
	public void setField(int column, int value){
		validateColumn(column);
		addCell(column, value, Format.Number);
	}
	
	/**
	 * Adds a numeric value to the cell at the given column in
	 * the current row
	 * 
	 * @param column
	 * @param value
	 */
	public void setField(int column, double value){
		validateColumn(column);
		addCell(column, value, Format.Number);
	}

	/**
	 * Adds a numeric value to the cell at the given column in
	 * the current row
	 * 
	 * @param column	0-based column number of spreadsheet cell
	 * @param value		value to be stored
	 * @param format	cell format; EmeryRptWorkbook.Format enum
	 */
	public void setField(int column, double value, Format format){
		validateColumn(column);
		addCell(column, value, format);
	}

	/**
	 * Set the width of specified column.
	 * 
	 * @param column	0-based column index	
	 * @param width		width in number of characters (approximate)
	 */
	public void setColumnWidth(int column, int width){
		validateColumn(column);
		m_Sheet.setColumnWidth(column, width * CHAR_WIDTH);
	}

	/**
	 * Checks to see that the column number is within the range
	 * of the column headers; throws a runtime error
	 * 
	 * @param column	0-based column index
	 */
	private void validateColumn(int column){
		if(column > m_Columns.length - 1 || column < 0){
			throw new NoSuchFieldException("The field/column does not exist. " +
					"Verify the column index is correct, or add the field " +
			"to the sheet.");
		}
	}
}
