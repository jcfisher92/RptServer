/**
 * File: ItemVelocityProject.java
 * Description: Item Velocity Project report based on inv_dtl table
 *
 * @author Tony Li
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class ItemVelocityProject extends Report
{

	private static int BASE_COLS = 15;

	private String m_Velocity = "";
	private String m_InvoiceBeginDate = "";
	private String m_InvoiceEndDate = "";
	private String m_CustomerId = "";
	
	//
	// Log4j logger
	private Logger m_Log;

	//
	// workbook entries.
	private Workbook m_Wrkbk;
	private Sheet m_Sheet;

	//
	// The cell styles for each of the base columns in the spreadsheet.
	private CellStyle[] m_CellStyles;

	private PreparedStatement m_ItemVelocity;

	/**
	 * default constructor
	 */
	public ItemVelocityProject()
	{
		super();
		m_Log = Logger.getLogger(RptServer.class);
		m_Wrkbk = new XSSFWorkbook();
		m_Sheet = m_Wrkbk.createSheet();
		setupWorkbook();
	}


	@Override
	public boolean createReport() {
		boolean created = false;
		m_Status = RptServer.RUNNING;

		try {
			m_EdbConn = m_RptProc.getEdbConn();
			prepareStatements();
			created = buildOutputFile();
		}

		catch ( Exception ex ) {
			m_Log.fatal(this.getClass().getName() + " exception:", ex);
		}

		finally {
			closeStatements();

			if ( m_Status == RptServer.RUNNING )
				m_Status = RptServer.STOPPED;
		}
		return created;
	}


	/**
	 * Creates the report title and the captions.
	 */
	private int createCaptions()
	{
		Font fontTitle;
		CellStyle styleTitle;   // Bold, centered
		CellStyle styleTitleLeft;   // Bold, Left Justified
		Row row = null;
		Cell cell = null;
		int rowNum = 0;
		StringBuffer caption = new StringBuffer("Item Velocity Project Report: ");

		if ( m_Sheet != null ) {
			fontTitle = m_Wrkbk.createFont();
			fontTitle.setFontHeightInPoints((short) 10);
			fontTitle.setFontName("Arial");
			fontTitle.setBold(true);

			styleTitle = m_Wrkbk.createCellStyle();
			styleTitle.setFont(fontTitle);
			styleTitle.setAlignment(HorizontalAlignment.CENTER);

			styleTitleLeft = m_Wrkbk.createCellStyle();
			styleTitleLeft.setFont(fontTitle);
			styleTitleLeft.setAlignment(HorizontalAlignment.LEFT);

			//
			// set the report title
			row = m_Sheet.createRow(rowNum);
			cell = row.createCell(0);
			cell.setCellType(CellType.STRING);
			cell.setCellStyle(styleTitleLeft);

			if ( m_Velocity != null && m_Velocity.length() > 0){
				if(!m_Velocity.isEmpty()) {
					caption.append("Velocity: ");
					caption.append(m_Velocity);
				}
				
				if(!m_CustomerId.isEmpty()) {
					caption.append(" Customers: ");
					caption.append(m_CustomerId);
				}
				
				caption.append(" between ").append(m_InvoiceBeginDate);
				caption.append(" and ").append(m_InvoiceEndDate);
			}

			cell.setCellValue(new XSSFRichTextString(caption.toString()));

			rowNum = 2;
			row = m_Sheet.createRow(rowNum);

			try {
				if ( row != null ) {
					for ( int i = 0; i < BASE_COLS; i++ ) {
						cell = row.createCell(i);
						cell.setCellStyle(styleTitleLeft);
					}

					row.getCell(0).setCellValue(new XSSFRichTextString("Item ID"));
					m_Sheet.setColumnWidth(0, 3000);
					row.getCell(1).setCellValue(new XSSFRichTextString("Item Description"));
					m_Sheet.setColumnWidth(1, 14000);
					row.getCell(2).setCellValue(new XSSFRichTextString("Vendor Name"));
					m_Sheet.setColumnWidth(2, 14000);
					row.getCell(3).setCellValue(new XSSFRichTextString("Setup Date"));
					row.getCell(4).setCellValue(new XSSFRichTextString("Quantity on Hand"));
					m_Sheet.setColumnWidth(3, 2000);
					row.getCell(5).setCellValue(new XSSFRichTextString("Project Days of Supply"));
					m_Sheet.setColumnWidth(4, 2000);
					row.getCell(6).setCellValue(new XSSFRichTextString("Units Ordered"));
					m_Sheet.setColumnWidth(5, 2000);
					row.getCell(7).setCellValue(new XSSFRichTextString("Units Sold"));
					m_Sheet.setColumnWidth(6, 2000);
					row.getCell(8).setCellValue(new XSSFRichTextString("Units Cut"));
					m_Sheet.setColumnWidth(7, 2000);
					row.getCell(9).setCellValue(new XSSFRichTextString("Fill Rate %"));
					m_Sheet.setColumnWidth(8, 3000);	
					row.getCell(10).setCellValue(new XSSFRichTextString("Units on Order"));
					m_Sheet.setColumnWidth(10, 2000);
					row.getCell(11).setCellValue(new XSSFRichTextString("Emery on Order"));
					m_Sheet.setColumnWidth(11, 3000);					                
					row.getCell(12).setCellValue(new XSSFRichTextString("Department Name"));
					m_Sheet.setColumnWidth(12, 2000);
					row.getCell(13).setCellValue(new XSSFRichTextString("Warehouse"));
					row.getCell(14).setCellValue(new XSSFRichTextString("SOQ Comment"));
				}
			}
			finally {
				row = null;
				cell = null;
				fontTitle = null;
				styleTitle = null;
				caption = null;
			}
		}

		return ++rowNum;
	}

	/**
	 * Executes the queries and builds the output file
	 *
	 * @return true if the file was built, false if not.
	 * @throws FileNotFoundException
	 */

	private boolean buildOutputFile() throws FileNotFoundException
	{
		Row row = null;
		FileOutputStream outFile = null;
		ResultSet itemVelocity = null;
		ResultSet DontShipBefore = null;
		int colCnt = BASE_COLS;
		int rowNum = 1;
		boolean result = false;

		outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

		try {
			rowNum = createCaptions();
			
			m_ItemVelocity.setString(1, m_InvoiceEndDate);
			m_ItemVelocity.setString(2, m_InvoiceBeginDate);						
			m_ItemVelocity.setInt(3, calculateSaturdays(m_InvoiceEndDate, m_InvoiceBeginDate));
			m_ItemVelocity.setString(4, m_InvoiceBeginDate);
			m_ItemVelocity.setString(5, m_InvoiceEndDate);
			m_ItemVelocity.setString(6, m_InvoiceBeginDate);
			m_ItemVelocity.setString(7, m_InvoiceEndDate);		
			
			itemVelocity = m_ItemVelocity.executeQuery();

			while ( itemVelocity.next() && m_Status == RptServer.RUNNING ) {				

				row = createRow(rowNum, colCnt);
				row.getCell(0).setCellValue(new XSSFRichTextString(itemVelocity.getString("item_id")));
				row.getCell(1).setCellValue(new XSSFRichTextString(itemVelocity.getString("description")));
				row.getCell(2).setCellValue(new XSSFRichTextString(itemVelocity.getString("vendor_name")));
				row.getCell(3).setCellValue(new XSSFRichTextString(itemVelocity.getString("setup_date")));
				row.getCell(4).setCellValue(new XSSFRichTextString(itemVelocity.getString("quantity_on_hand")));					
				row.getCell(5).setCellValue(new XSSFRichTextString(itemVelocity.getString("project_days_of_supply")));
				row.getCell(6).setCellValue(new XSSFRichTextString(itemVelocity.getString("units_ordered")));
				row.getCell(7).setCellValue(new XSSFRichTextString(itemVelocity.getString("units_sold")));
				row.getCell(8).setCellValue(new XSSFRichTextString(itemVelocity.getString("units_cut")));
				row.getCell(9).setCellValue(new XSSFRichTextString(itemVelocity.getString("fill_rate_pct")));
				row.getCell(10).setCellValue(new XSSFRichTextString(itemVelocity.getString("customer_order_units")));
				row.getCell(11).setCellValue(new XSSFRichTextString(itemVelocity.getString("em_order_units")));
				row.getCell(12).setCellValue(new XSSFRichTextString(itemVelocity.getString("dept_name")));
				row.getCell(13).setCellValue(new XSSFRichTextString(itemVelocity.getString("warehouse_id")));
				row.getCell(14).setCellValue(new XSSFRichTextString(itemVelocity.getString("soq")));

				rowNum++;
			}

			m_Wrkbk.write(outFile);
			result = true;
		}

		catch ( Exception ex ) {
			m_ErrMsg.append(ex.getClass().getName() + "\r\n");
			m_ErrMsg.append(ex.getMessage());
			m_Log.error(this.getClass().getName(), ex);
		}

		finally {
			row = null;
			closeRSet(itemVelocity);
			closeRSet(DontShipBefore);

			try {
				outFile.close();
			}

			catch( Exception e ) {
				m_Log.error(this.getClass().getName(), e);
			}

			outFile = null;
		}

		return result;
	}
	
	/*
	 * When we do project calculation, we need to get rid of Saturday since no order release that day.  -TLI 04212014
	 */	
	@SuppressWarnings("deprecation")
	private int calculateSaturdays(String invoiceEndDate, String invoiceBeginDate){
	    int weekendDayCount = 0;

	    Date beginDate = null;
	    Date endDate = null;
	    try {
			beginDate = new SimpleDateFormat("MM/dd/yyyy").parse(invoiceBeginDate);
			endDate = new SimpleDateFormat("MM/dd/yyyy").parse(invoiceEndDate);
		} catch (ParseException e) {
			m_Log.error(this.getClass().getName(), e);
		}
	    
	    while(beginDate.compareTo(endDate) < 0){
	    	beginDate.setDate(beginDate.getDate() + 1);
	        if(beginDate.getDay() == 6){
	            ++weekendDayCount ;
	        }
	    }

	    return weekendDayCount ;
	}
	

	/**
	 * Creates a row in the worksheet.
	 * @param rowNum The row number.
	 * @param colCnt The number of columns in the row.
	 *
	 * @return The formatted row of the spreadsheet.
	 */
	private Row createRow(int rowNum, int colCnt)
	{
		Row row = null;
		Cell cell = null;

		if ( m_Sheet != null ) {
			row = m_Sheet.createRow(rowNum);

			//
			// set the type and style of the cell.
			if ( row != null ) {
				for ( int i = 0; i < colCnt; i++ ) {
					cell = row.createCell(i);
					cell.setCellStyle(m_CellStyles[i]);
				}
			}
		}

		return row;
	}

	/**
	 * Builds the sql based on the type of filter requested by the user.
	 * @return A complete sql statement.
	 * 	    * 
	 */
	private void prepareStatements() throws Exception
	{
		if ( m_EdbConn != null ) {
			StringBuffer sql = new StringBuffer();

			sql.append(
			"select " +
			"   project2.item_id, project2.description, project2.vendor_name,  project2.setup_date, project2.quantity_on_hand, " +
			"   project2.project_days_of_supply, project2.units_ordered, project2.units_sold,  project2.units_cut, project2.fill_rate_pct, " +
			"   coalesce(sum(order_line.qty_ordered), 0) customer_order_units," +
			"   project2.em_order_units, project2.dept_name, project2.warehouse_id, project2.soq " + 	      
			"from ( " +
			"	select " +
			"	project.*, " +
			"   coalesce(sum(po_dtl.qty_ordered - po_dtl.qty_put_away), 0) em_order_units " +
			"	from ( " +
		    "   select item_entity_attr.item_id, item_entity_attr.item_ea_id, item_entity_attr.description, setup_date, " +
		    "         ejd_item_warehouse.qoh as quantity_on_hand, " +
            "       decode (coalesce (sum (inv_dtl.qty_shipped), 0 ), 0, 'n/a', " +
			"				round(qoh / sum (inv_dtl.qty_shipped) * (@(extract(day from to_date (?, 'mm/dd/yyyy') - to_date(?, 'mm/dd/yyyy')))-?))" +
			"			) project_days_of_supply, " +
            "        coalesce(sum(inv_dtl.qty_ordered), 0) units_ordered, " +
            "        coalesce(sum(inv_dtl.qty_shipped), 0) units_sold, " +
            "        coalesce(sum(inv_dtl.qty_ordered - inv_dtl.qty_shipped), 0) units_cut, " +
			"   		round(decode(sum(inv_dtl.qty_ordered), 0, 100, " +
			"   		sum(inv_dtl.qty_shipped) / sum(inv_dtl.qty_ordered)) * 100, 3) fill_rate_pct, " +
			"   		emery_dept.name as dept_name, vendor.name as vendor_name, " +
			"   		warehouse.warehouse_id, ejd_item.soq_comment as soq " +
            "    from item_entity_attr " +
            "    join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id " +
            "    join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id " +
            "    join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id " +
            "    join emery_dept on ejd_item.dept_id = emery_dept.dept_id " +
            "    join warehouse on warehouse.warehouse_id = ejd_item_warehouse.warehouse_id "+
            "    join vendor on item_entity_attr.vendor_id = vendor.vendor_id " +
            "    inner join inv_dtl on item_entity_attr.item_ea_id = inv_dtl.item_ea_id and " +			
			"   		inv_dtl.sale_type = 'WAREHOUSE' and " + 
			"   		inv_dtl.tran_type = 'SALE' and " +  
			"   		inv_dtl.qty_ordered > 0 and " +       
			"   		inv_dtl.invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') and " + 
			"   		inv_dtl.warehouse = warehouse.name " +
			"where " +
			"   setup_date between to_Date(?, 'mm/dd/yyyy') - 120 and to_Date(?, 'mm/dd/yyyy') - 30 "
			);
			
			if(!m_Velocity.isEmpty()|| !m_CustomerId.isEmpty() )
				sql.append(buildWhereClause(m_Velocity, m_CustomerId));
			
			sql.append(
			"   group by " +
            "    item_entity_attr.item_id, item_entity_attr.item_ea_id, setup_date, item_entity_attr.description, emery_dept.name, " +
            "    warehouse.warehouse_id, qoh, ejd_item.soq_comment, vendor.name " +	
			"   )  project " +  
			"   left outer join po_dtl on project.item_ea_id = po_dtl.item_ea_id and po_dtl.status = 'OPEN' " +
			"   left outer join warehouse on warehouse.fas_facility_id = po_dtl.warehouse  and  project.warehouse_id = warehouse.warehouse_id " +
			
            "   group by   project.item_id, project.item_ea_id, project.description, project.setup_date, project.quantity_on_hand, " +
			"              project.project_days_of_supply, project.units_sold, project.units_ordered, project.units_cut, project.fill_rate_pct, " +
			"              project.dept_name, project.warehouse_id, project.soq, project.vendor_name " +
			"   ) project2 " +       
			"left outer join order_line on order_line.invoice_num is null and order_status_id = 1 and project2.item_id = order_line.item_id " +
			"left outer join order_header on order_line.order_id = order_header.order_id and order_header.warehouse_id = project2.warehouse_id " +
			"group by   project2.item_id, project2.description, project2.setup_date, project2.quantity_on_hand, " +
			"           project2.project_days_of_supply, project2.units_sold, project2.units_ordered, project2.units_cut, project2.fill_rate_pct, " +
			"           project2.dept_name, project2.warehouse_id, project2.soq, project2.em_order_units, project2.vendor_name");	
			
			m_ItemVelocity = m_EdbConn.prepareStatement(sql.toString());
		}
	}

	private String buildWhereClause(String velocities, String customerId){

		StringBuffer whereClause = new StringBuffer (" and ");

		if(!velocities.isEmpty() )
		{
			StringBuffer sql = new StringBuffer (" item_velocity.velocity in ( ");

			String Plist[];
			Plist = velocities.split(",");     

			//for each velocity
			for (int i = 0 ; i < Plist.length ; i++) {  		
				sql.append("'").append(Plist[i]).append("', ");
			}
			whereClause.append(sql.substring(0, sql.toString().length() -2) + ") ");
		}

		if( !customerId.isEmpty() )
		{
			if(!velocities.isEmpty() )
				whereClause.append(" and ");

			StringBuffer sql = new StringBuffer (" inv_dtl.cust_nbr in (");

			String Plist[];
			Plist = customerId.split(",");     

			//for each customer
			for (int i = 0 ; i < Plist.length ; i++) {  		
				sql.append("'").append(Plist[i]).append("', ");
			}
			whereClause.append(sql.substring(0, sql.toString().length() -2) + ") ");
		}
		
		return whereClause.toString();
	}

	
	/**
	 *  Closes all the sql statements so they release the db cursors.
	 */
	private void closeStatements()
	{
		closeStmt(m_ItemVelocity);
	}


	/**
	 * Cleanup any allocated resources.
	 * @throws Throwable
	 */
	@Override
	public void finalize() throws Throwable
	{
		if ( m_CellStyles != null ) {
			for ( int i = 0; i < m_CellStyles.length; i++ )
				m_CellStyles[i] = null;
		}

		m_Sheet = null;
		m_Wrkbk = null;
		m_CellStyles = null;

		super.finalize();
	}


	/**
	 * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
	 *
	 * The parameters are transferred from EIS
	 */
	public void setParams(ArrayList<Param> params)
	{
		StringBuffer fname = new StringBuffer();
		String tm = Long.toString(System.currentTimeMillis()).substring(3);
		int pcount = params.size();
		Param param = null;

		for ( int i = 0; i < pcount; i++ ) {
			param = params.get(i);

			if ( param.name.equals("velocity") )
				m_Velocity = param.value.trim();

			if ( param.name.equals("invoiceBeginDate") )
				m_InvoiceBeginDate = param.value.trim();
			
			if ( param.name.equals("invoiceEndDate") )
				m_InvoiceEndDate = param.value.trim();

			if ( param.name.equals("custId") )
				m_CustomerId = param.value.trim();
		}

		//
		// Build the file name.
		fname.append(tm);
		fname.append("-");
		fname.append(m_RptProc.getUid());
		fname.append("itemvelocity.xlsx");

		m_FileNames.add(fname.toString());
	}


	/**
	 * Sets up the styles for the cells based on the column data.  Does any other initialization
	 * needed by the workbook.
	 */
	private void setupWorkbook()
	{
		CellStyle styleText;      // Text right justified
		CellStyle styleInt;       // Style with 0 decimals
		CellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00)
		CellStyle stylePct;       // Style with 0 decimals + %

		styleText = m_Wrkbk.createCellStyle();
		styleText.setAlignment(HorizontalAlignment.LEFT);

		styleInt = m_Wrkbk.createCellStyle();
		styleInt.setAlignment(HorizontalAlignment.RIGHT);
		styleInt.setDataFormat((short)3);

		styleMoney = m_Wrkbk.createCellStyle();
		styleMoney.setAlignment(HorizontalAlignment.RIGHT);
		styleMoney.setDataFormat((short)8);

		stylePct = m_Wrkbk.createCellStyle();
		stylePct.setAlignment(HorizontalAlignment.RIGHT);
		stylePct.setDataFormat((short)2);

		m_CellStyles = new CellStyle[] {
				styleText,    // col 0 item id
				styleText,    // col 1 description
				styleText,    // col 2 vendor name
				styleText,    // col 3 setup date
				styleInt,     // col 4 quantity on hand
				styleText,    // col 5 project days of supply
				styleInt,     // col 6 units ordered
				styleInt,     // col 7 units sold
				styleInt,     // col 8 units cut
				stylePct,     // col 9 fill rate pct				
				styleInt,     // col 10 customer on ordered
				styleInt,     // col 11 emery on order
				styleText,    // col 12 department name
				styleInt,     // col 13 warehouse Id		
				styleText,    // col 14 soq_comment
		};

		styleText = null;
		styleInt = null;
		styleMoney = null;
		stylePct = null;
	}
}