/**
 * File: EmeryReportWorkbook.java
 * Description: A wrapper for the HSSFWorkbook data of a report.
 * 				Acts as a factory for and keeps track of multiple sheets, 
 * 				has built in cell styles customized for Emery reports.
 * 				See EmeryRptSheet.java for sheet functionality.  
 *
 * @author Erik Pearson
 *
 * Create Date: 07/19/2010
 */
package com.emerywaterhouse.rpt.helper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import com.emerywaterhouse.rpt.helper.EmeryRptSheet.Column;

public class EmeryRptWorkbook {
	/**
	 * Basic cell formatting options used in Emery reports.
	 * 
	 * @author epearson
	 */
	public enum Format{
		General, HeaderLeft, HeaderLeftWrap, HeaderCtr, HeaderCtrWrap,
		HeaderRightWrap, TextLeft, TextLeftWrap, TextRight, TextCtr,
		Number, Number2d, Number3d, Number4d, 
		Percent, Percent2d, Percent3d, Percent4d, 
		Currency, Currency2d, Currency3d, Currency4d
	}
	
	private HSSFWorkbook m_Workbook;
	private List<EmeryRptSheet> m_SheetList;
	private HSSFDataFormat m_CustomDataFormat;
	
	// Styles
	private HSSFFont m_FontNorm;
	private HSSFFont m_FontBold;
	private HSSFCellStyle m_StyleHdrLeft;
	private HSSFCellStyle m_StyleHdrLeftWrap;
	private HSSFCellStyle m_StyleHdrCntr;
	private HSSFCellStyle m_StyleHdrCntrWrap;
	private HSSFCellStyle m_StyleHdrRghtWrap;
	private HSSFCellStyle m_StyleDtlLeft;
	private HSSFCellStyle m_StyleDtlLeftWrap;
	private HSSFCellStyle m_StyleDtlCntr;
	private HSSFCellStyle m_StyleDtlRght;
	private HSSFCellStyle m_StyleNumber2d;
	private HSSFCellStyle m_StyleNumber3d;
	private HSSFCellStyle m_StyleNumber4d;
	private HSSFCellStyle m_StylePercent;
	private HSSFCellStyle m_StylePercent2d;
	private HSSFCellStyle m_StylePercent3d;
	private HSSFCellStyle m_StylePercent4d;
	private HSSFCellStyle m_StyleCurrency;
	private HSSFCellStyle m_StyleCurrency2d;
	private HSSFCellStyle m_StyleCurrency3d;
	private HSSFCellStyle m_StyleCurrency4d;
	
	/**
	 * Constructor
	 */
	public EmeryRptWorkbook(){
		m_Workbook = new HSSFWorkbook();
		m_SheetList = new ArrayList<EmeryRptSheet>();
		initStyles();
	}
	
	/**
	 * A worksheet factory.
	 * 
	 * @param name		name of the sheet
	 * @param title		title of the sheet (e.g. the report name and date)
	 * @param fields	column headers (e.g. names for the fields in a query)
	 * @return			the sheet created
	 */
	public EmeryRptSheet createSheet(String name, String title, Column[] fields){
		EmeryRptSheet sheet = new EmeryRptSheet(this, m_Workbook.createSheet(name), title, fields);
		m_SheetList.add(sheet); // keep track of the sheets
		
		return sheet;
	}
	
	/**
	 * Get a sheet by index.
	 * 
	 * @param index
	 * @return			the sheet at the given index
	 */
	public EmeryRptSheet getSheet(int index){
		return m_SheetList.get(index);
	}
	
	/**
	 * Initialize the styles
	 */
	private void initStyles() {
		// m_CustomDataFormat is used to define a non-standard data format when
		// defining a style. For example, "0.00" is a built-in format, but "0.000"
		// and "0.0000" are custom formats.		
		m_CustomDataFormat = m_Workbook.createDataFormat();
		
		// defines normal font
		m_FontNorm = m_Workbook.createFont();
		m_FontNorm.setFontName("Arial");
		m_FontNorm.setFontHeightInPoints((short)10);

		// defines bold font
		m_FontBold = m_Workbook.createFont();
		m_FontBold.setFontName("Arial");
		m_FontBold.setFontHeightInPoints((short)10);
		m_FontBold.setBold(true);

		// defines style column header, left-justified
		m_StyleHdrLeft = m_Workbook.createCellStyle();
		m_StyleHdrLeft.setFont(m_FontBold);
		m_StyleHdrLeft.setAlignment(HorizontalAlignment.LEFT);
		m_StyleHdrLeft.setVerticalAlignment(VerticalAlignment.TOP);

		// defines style column header, left-justified, wrap text
		m_StyleHdrLeftWrap = m_Workbook.createCellStyle();
		m_StyleHdrLeftWrap.setFont(m_FontBold);
		m_StyleHdrLeftWrap.setAlignment(HorizontalAlignment.LEFT);
		m_StyleHdrLeftWrap.setVerticalAlignment(VerticalAlignment.TOP);
		m_StyleHdrLeftWrap.setWrapText(true);

		// defines style column header, center-justified
		m_StyleHdrCntr = m_Workbook.createCellStyle();
		m_StyleHdrCntr.setFont(m_FontBold);
		m_StyleHdrCntr.setAlignment(HorizontalAlignment.CENTER);
		m_StyleHdrCntr.setVerticalAlignment(VerticalAlignment.TOP);

		// defines style column header, center-justified, wrap text
		m_StyleHdrCntrWrap = m_Workbook.createCellStyle();
		m_StyleHdrCntrWrap.setFont(m_FontBold);
		m_StyleHdrCntrWrap.setAlignment(HorizontalAlignment.CENTER);
		m_StyleHdrCntrWrap.setVerticalAlignment(VerticalAlignment.TOP);
		m_StyleHdrCntrWrap.setWrapText(true);

		// defines style column header, right-justified, wrap text
		m_StyleHdrRghtWrap = m_Workbook.createCellStyle();
		m_StyleHdrRghtWrap.setFont(m_FontBold);
		m_StyleHdrRghtWrap.setAlignment(HorizontalAlignment.RIGHT);
		m_StyleHdrRghtWrap.setVerticalAlignment(VerticalAlignment.TOP);
		m_StyleHdrRghtWrap.setWrapText(true);

		// defines style detail data cell, left-justified
		m_StyleDtlLeft = m_Workbook.createCellStyle();
		m_StyleDtlLeft.setFont(m_FontNorm);
		m_StyleDtlLeft.setAlignment(HorizontalAlignment.LEFT);
		m_StyleDtlLeft.setVerticalAlignment(VerticalAlignment.TOP);

		// defines style detail data cell, left-justified, wrap text
		m_StyleDtlLeftWrap = m_Workbook.createCellStyle();
		m_StyleDtlLeftWrap.setFont(m_FontNorm);
		m_StyleDtlLeftWrap.setAlignment(HorizontalAlignment.LEFT);
		m_StyleDtlLeftWrap.setVerticalAlignment(VerticalAlignment.TOP);
		m_StyleDtlLeftWrap.setWrapText(true);

		// defines style detail data cell, center-justified
		m_StyleDtlCntr = m_Workbook.createCellStyle();
		m_StyleDtlCntr.setFont(m_FontNorm);
		m_StyleDtlCntr.setAlignment(HorizontalAlignment.CENTER);
		m_StyleDtlCntr.setVerticalAlignment(VerticalAlignment.TOP);

		// defines style detail data cell, right-justified
		m_StyleDtlRght = m_Workbook.createCellStyle();
		m_StyleDtlRght.setFont(m_FontNorm);
		m_StyleDtlRght.setAlignment(HorizontalAlignment.RIGHT);
		m_StyleDtlRght.setVerticalAlignment(VerticalAlignment.TOP);
		
		/*
		 * Define Styles for floating point values: numbers, percents, currency
		 * All are right justified with varying decimal places
		 */
		
		// defines style for floating-point number, 2 decimal places
		m_StyleNumber2d = createNumberCellStyle(Format.Number2d);		

		// defines style for floating-point number, 3 decimal places
		m_StyleNumber3d = createNumberCellStyle(Format.Number3d);

		// defines style for floating-point number, 4 decimal places
		m_StyleNumber4d = createNumberCellStyle(Format.Number4d);
		
		// define style for Percent
		m_StylePercent = createNumberCellStyle(Format.Percent);
		
		// define style for Percent, 2 decimal places
		m_StylePercent2d = createNumberCellStyle(Format.Percent2d);
		
		// define style for Percent, 3 decimal places
		m_StylePercent3d = createNumberCellStyle(Format.Percent3d);
		
		// define style for Percent, 4 decimal places
		m_StylePercent4d = createNumberCellStyle(Format.Percent4d);
		
		// define style for Currency
		m_StyleCurrency = createNumberCellStyle(Format.Currency);
		
		// define style for Currency
		m_StyleCurrency2d = createNumberCellStyle(Format.Currency2d);
		
		// define style for Currency
		m_StyleCurrency3d = createNumberCellStyle(Format.Currency3d);
		
		// define style for Currency
		m_StyleCurrency4d = createNumberCellStyle(Format.Currency4d);
	}
	
	/**
	 * Returns a cell style with default styles for number formats.
	 * 
	 * @return	default number cell style
	 */
	private HSSFCellStyle createNumberCellStyle(Format format) {
		HSSFCellStyle style = m_Workbook.createCellStyle();
		
		style.setFont(m_FontNorm);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		
		switch (format){
		case Number2d: 
			style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00")); break;
		case Number3d: 
			style.setDataFormat(m_CustomDataFormat.getFormat("0.000")); break;	
		case Number4d: 
			style.setDataFormat(m_CustomDataFormat.getFormat("0.0000")); break;
		case Percent: 
			style.setDataFormat(m_CustomDataFormat.getFormat("0%")); break;
		case Percent2d: 
			style.setDataFormat(m_CustomDataFormat.getFormat("0.00%")); break;
		case Percent3d: 
			style.setDataFormat(m_CustomDataFormat.getFormat("0.000%")); break;	
		case Percent4d:
			style.setDataFormat(m_CustomDataFormat.getFormat("0.0000%")); break;
		case Currency: 
			style.setDataFormat(m_CustomDataFormat.getFormat("$#,##0")); break;
		case Currency2d: 
			style.setDataFormat(m_CustomDataFormat.getFormat("$#,##0.00")); break;
		case Currency3d: 
			style.setDataFormat(m_CustomDataFormat.getFormat("$#,##0.000")); break;	
		case Currency4d:
			style.setDataFormat(m_CustomDataFormat.getFormat("$#,##0.0000")); break;	
		default: break;
		}
		
		return style;
	}
	
	/**
	 * Gets a cell style based on format type
	 * (Used by EmeryRptSheet.java)
	 * 
	 * @param format	Format of the cell
	 * @return	style	the HSSFCellStyle requested	
	 */
	protected HSSFCellStyle getCellStyle(Format format){
		HSSFCellStyle style;
		
		switch(format){
		case HeaderLeft: 
			style = m_StyleHdrLeft; break;
		case HeaderLeftWrap: 
			style = m_StyleHdrLeftWrap; break;
		case HeaderRightWrap: 
			style = m_StyleHdrRghtWrap; break;
		case HeaderCtr: 
			style = m_StyleHdrCntr; break;
		case HeaderCtrWrap: 
			style = m_StyleHdrCntrWrap; break;
		case TextRight:
		case Number:
			style = m_StyleDtlRght; break;
		case Number2d: 
			style = m_StyleNumber2d; break;
		case Number3d: 
			style = m_StyleNumber3d; break;	
		case Number4d: 
			style = m_StyleNumber4d; break;
		case Percent: 
			style = m_StylePercent; break;
		case Percent2d: 
			style = m_StylePercent2d; break;
		case Percent3d: 
			style = m_StylePercent3d; break;	
		case Percent4d:
			style = m_StylePercent4d; break;
		case Currency: 
			style = m_StyleCurrency; break;
		case Currency2d: 
			style = m_StyleCurrency2d; break;
		case Currency3d: 
			style = m_StyleCurrency3d; break;	
		case Currency4d:
			style = m_StyleCurrency4d; break;
		default:
			style = m_StyleDtlLeft;
		}
		
		return style;
	}

	/**
	 * Writes workbook to a given output stream.
	 * 
	 * @param out	output stream
	 * @throws IOException
	 */
	public void write(FileOutputStream out) throws IOException{
		m_Workbook.write(out);
	}
}
