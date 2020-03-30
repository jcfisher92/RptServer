/**
 * File: SimpleReport.java
 * Description: This class extends the Report base class and automatically
 * handles the creation of the spreadsheet report.  The user will need
 * to determine how it uses the parameters and implement a method that
 * produces the ResultSet of data to populate the spreadsheet.  If you
 * want to display a report title in the spreadsheet, make sure to set
 * the report name member variable.
 * 
 * Formatting limitations: This simple report treats all data as either
 * an integer, a floating-point value, or a string, based on the Oracle
 * data types (NUMBER, VARCHAR2, etc).  Floating-point values
 * are formatted with either 2, 3, 4, or more decimal places showing.
 * Currency and percentage formatting are not included.  
 *
 * @author Erik Pearson
 *
 * Create Date: 08/25/2010
 * Last Update: $Id: SimpleReport.java,v 1.1 2010/08/29 00:27:47 epearson Exp $
 * 
 * History
 *    $Log: SimpleReport.java,v $
 *    Revision 1.1  2010/08/29 00:27:47  epearson
 *    Initial add
 *
 *    
 */
package com.emerywaterhouse.rpt.server;

import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

import com.emerywaterhouse.rpt.helper.EmeryRptSheet;
import com.emerywaterhouse.rpt.helper.EmeryRptWorkbook;
import com.emerywaterhouse.rpt.helper.EmeryRptSheet.Column;

public abstract class SimpleReport extends Report {

	private static final int MAX_COL_WIDTH = 30; // max column width in characters
	private static final int MIN_COL_WIDTH = 13; // min column width in characters
	private static final int MAX_ROWS = 65000; // max rows in Excel 2003 (actually 65536)
	
	private String m_ReportTitle; // title printed at the top of the spreadsheet (optional)
	
	/**
	 * Constructor
	 */
	public SimpleReport(){
		super();
		m_ReportTitle = null;
	}

	/** 
	 * Create the report.
	 * @see com.emerywaterhouse.rpt.server.Report#createReport()
	 */
	@Override
	public boolean createReport() {
		m_Status = RptServer.RUNNING;	
		boolean created = false;
		EmeryRptWorkbook rptWorkbook = null;
		
		try{
			m_OraConn = m_RptProc.getOraConn();	
			rptWorkbook = buildWorkbook();
			m_FileNames.add(buildFileName());
			rptWorkbook.write(new FileOutputStream(m_FilePath + m_FileNames.get(0), false));
			created = true;
		} catch (Exception e){
			log.fatal("Exception: ", e);
		} finally {
			m_Status = RptServer.STOPPED;
		}
		
		return created;
	}
	
	/**
	 * Builds the file name.  Descendant class must implement.
	 * 
	 * @return	the filename
	 */
	protected abstract String buildFileName();

	/**
	 * Build the spreadsheet report using the helper class
	 * EmeryRptWorkbook and EmeryRptSheet.
	 * 
	 * @return	an EmeryRptWorkbook, a wrapper for Apache POI-HSSF Workbook
	 * @throws SQLException 
	 */
	private EmeryRptWorkbook buildWorkbook() throws SQLException {
		PreparedStatement reportQuery = null;
		ResultSet reportData = null;
		ResultSetMetaData reportMetaData = null; // used to format columns and cells
		
		EmeryRptWorkbook rptWorkbook = new EmeryRptWorkbook();
		EmeryRptSheet rptSheet = null;
		EmeryRptSheet.Column[] colNames = null;
		int numSheets = 0; // keep track of number of sheets and use for sheet labels
		
		// Note: The Statement, ResultSet, and ResultSetMetaData
		// will automatically be closed when they go out of the scope
		// of this method.
		reportQuery = buildReportQuery();
		reportData = reportQuery.executeQuery();
		reportMetaData = reportData.getMetaData();
		
		colNames = buildColumns(reportMetaData);
		
		// create first work-sheet
		rptSheet = rptWorkbook.createSheet("Sheet 1", m_ReportTitle, colNames);
		numSheets++;
		
		// iterate through result-set and populate the spreadsheet with item data
		while(reportData.next()){
			rptSheet.addRow();
			
			// loop through the fields and get the data based on type
			for(int i = 0; i < reportMetaData.getColumnCount(); i++){
				// if the type is NUMBER(p,s), format the cells based on the number of
				// decimal places (the scale s); otherwise format the data as a String
				if(reportMetaData.getColumnTypeName(i+1).equals("NUMBER")){
					switch(reportMetaData.getScale(i+1)){
					case 0: 
						rptSheet.setField(i, reportData.getInt(i+1));
						break;
					case 2: 
						rptSheet.setField(i, reportData.getDouble(i+1), EmeryRptWorkbook.Format.Number2d);
						break;
					case 3: 
						rptSheet.setField(i, reportData.getDouble(i+1), EmeryRptWorkbook.Format.Number3d);
						break;
					case 4:
						rptSheet.setField(i, reportData.getDouble(i+1), EmeryRptWorkbook.Format.Number4d);
						break;
					default: 
						rptSheet.setField(i, reportData.getDouble(i+1));
						break;
					}
				} else {
					rptSheet.setField(i, reportData.getString(i+1), false);
				}
			}
			
			// This report has the potential for exceeding Excel 2003's row limit,
			// if it goes over, create a new work-sheet and continue adding data
			if( (reportData.getRow() + 1) > (MAX_ROWS * numSheets) ){
				numSheets++;
				rptSheet = rptWorkbook.createSheet("Sheet " + numSheets, m_ReportTitle, colNames);
			}
		}
				
		return rptWorkbook;
	}

	/**
	 * Build the columns for the EmeryRptSheet using result-set meta-data.
	 * 
	 * @param rsmd			result-set meta-data with information about columns
	 * @return colNames		column information including name and width in characters
	 * @throws SQLException 
	 */
	private Column[] buildColumns(ResultSetMetaData rsmd) throws SQLException {
		EmeryRptSheet.Column[] colNames = null;
		
		// build column names from result-set meta-data
		colNames = new EmeryRptSheet.Column[rsmd.getColumnCount()];
		
		for(int i = 0; i < colNames.length; i++){
			int width = rsmd.getPrecision(i+1);
			
			if(width > MAX_COL_WIDTH)
				width = MAX_COL_WIDTH;
			else if (width < MIN_COL_WIDTH){
				width = MIN_COL_WIDTH;
			}
			
			colNames[i] = new EmeryRptSheet.Column(rsmd.getColumnLabel(i+1), width);
		}
		
		return colNames;
	}

	/**
	 * Build the report query.  The user must implement this class.
	 * The prepared statement must have all variables bound before
	 * being returned.  The column headers for the spreadsheet are
	 * taken from either the column labels in the query or the
	 * column names.
	 * 
	 * @return	the query used to get the data for the report
	 * @throws SQLException
	 */
	protected abstract PreparedStatement buildReportQuery() throws SQLException;
	
	/**
	 * Get the title of the report.
	 * 
	 * @return the title of the report
	 */
	public String getReportTitle(){
		return m_ReportTitle;
	}
	
	/**
	 * Set the report title.
	 * 
	 * @param reportTitle	the title you want displayed on the report
	 */
	public void setReportTitle(String reportTitle){
		m_ReportTitle = reportTitle;
	}
}
