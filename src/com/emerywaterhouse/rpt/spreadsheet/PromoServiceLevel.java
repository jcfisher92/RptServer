/**
 * File: PromoServiceLevel.java
 * Description: Promotion service level report based on packet table
 *
 * @author Tony Li
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;

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

public class PromoServiceLevel extends Report
{

	private static int BASE_COLS = 13;

	private String m_PacketId = "";
	private String m_CustomerId = "";
	private String m_InvoiceBeginDate = "";
	private String m_InvoiceEndDate = "";

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

	private PreparedStatement m_PromoPacket;

	/**
	 * default constructor
	 */
	public PromoServiceLevel()
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
		StringBuffer caption = new StringBuffer("Promotion Service Level Report: ");

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

			if ( m_PacketId != null && m_PacketId.length() > 0){
				if(!m_PacketId.isEmpty()) {
					caption.append("Packet: ");
					caption.append(m_PacketId);
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

					row.getCell(0).setCellValue(new XSSFRichTextString("Packet #"));
					row.getCell(1).setCellValue(new XSSFRichTextString("Title"));
					row.getCell(2).setCellValue(new XSSFRichTextString("Vendor Name"));
					m_Sheet.setColumnWidth(2, 7000);
					row.getCell(3).setCellValue(new XSSFRichTextString("Item ID"));
					m_Sheet.setColumnWidth(2, 3000);
					row.getCell(4).setCellValue(new XSSFRichTextString("Item Description"));
					m_Sheet.setColumnWidth(4, 14000);
					row.getCell(5).setCellValue(new XSSFRichTextString("Units on Order"));
					m_Sheet.setColumnWidth(5, 2000);
					row.getCell(6).setCellValue(new XSSFRichTextString("Units Invoiced"));
					m_Sheet.setColumnWidth(6, 2000);
					row.getCell(7).setCellValue(new XSSFRichTextString("Fill Rate %"));
					m_Sheet.setColumnWidth(7, 3000);	               
					row.getCell(8).setCellValue(new XSSFRichTextString("Emery on Order"));
					m_Sheet.setColumnWidth(8, 3000);
					row.getCell(9).setCellValue(new XSSFRichTextString("Customer Account"));
					row.getCell(10).setCellValue(new XSSFRichTextString("Customer Number"));	               
					row.getCell(11).setCellValue(new XSSFRichTextString("Customer Name"));
					m_Sheet.setColumnWidth(11, 2000);
					row.getCell(12).setCellValue(new XSSFRichTextString("Warehouse"));	               
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
		ResultSet promopacket = null;
		ResultSet DontShipBefore = null;
		int colCnt = BASE_COLS;
		int rowNum = 1;
		boolean result = false;

		outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

		try {
			rowNum = createCaptions();


			m_PromoPacket.setString(1, m_InvoiceBeginDate);
			m_PromoPacket.setString(2, m_InvoiceEndDate);
			promopacket = m_PromoPacket.executeQuery();

			while ( promopacket.next() && m_Status == RptServer.RUNNING ) {				

				row = createRow(rowNum, colCnt);
				row.getCell(0).setCellValue(new XSSFRichTextString(promopacket.getString("packet_id")));
				row.getCell(1).setCellValue(new XSSFRichTextString(promopacket.getString("title")));
				row.getCell(2).setCellValue(new XSSFRichTextString(promopacket.getString("item_id")));
				row.getCell(3).setCellValue(new XSSFRichTextString(promopacket.getString("vendor_name")));					
				row.getCell(4).setCellValue(new XSSFRichTextString(promopacket.getString("description")));
				row.getCell(5).setCellValue(new XSSFRichTextString(promopacket.getString("units_ordered")));
				row.getCell(6).setCellValue(new XSSFRichTextString(promopacket.getString("units_invoiced")));
				row.getCell(7).setCellValue(new XSSFRichTextString(promopacket.getString("fill_rate_pct")));
				row.getCell(8).setCellValue(new XSSFRichTextString(promopacket.getString("emery_on_order")));
				row.getCell(9).setCellValue(new XSSFRichTextString(promopacket.getString("customer_acct")));
				row.getCell(10).setCellValue(new XSSFRichTextString(promopacket.getString("customer_nbr")));
				row.getCell(11).setCellValue(new XSSFRichTextString(promopacket.getString("customer_name")));
				row.getCell(12).setCellValue(new XSSFRichTextString(promopacket.getString("warehouse")));

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
			closeRSet(promopacket);
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

			sql.append("select ");
			sql.append("   packet_id, title, vendor_name, item_id, description, units_ordered, units_invoiced, fill_rate_pct, ");
			sql.append("   decode(warehouse, 'PORTLAND', pos_util.get_on_order(item_id, '01'), ");
			sql.append("      pos_util.get_on_order(item_id, '02')) emery_on_order, ");
			sql.append("   customer_acct, customer_nbr, customer_name, warehouse  "); 	      
			sql.append("from ( ");
			sql.append("   select ");
			sql.append("      packet.packet_id, packet.title, inv_dtl.vendor_name, promo_item.item_id, ");
			sql.append("      item.description, sum(inv_dtl.qty_ordered) units_ordered, ");
			sql.append("   	sum(inv_dtl.qty_shipped) units_invoiced,  ");
			sql.append("   	decode(sum(qty_ordered), 0, 0, round(sum(inv_dtl.qty_shipped)/sum(qty_ordered) * 100, 1)) fill_rate_pct, ");
			sql.append("   	inv_dtl.cust_acct customer_acct, inv_dtl.cust_nbr customer_nbr, ");
			sql.append("   	customer.name customer_name, inv_dtl.warehouse ");
			sql.append("   from packet ");
			sql.append("   join promotion on promotion.packet_id = packet.packet_id ");
			sql.append("   join promo_item on promo_item.promo_id = promotion.promo_id ");
			sql.append("   join item on item.item_id = promo_item.item_id ");
			sql.append("   join inv_dtl on inv_dtl.item_nbr = promo_item.item_id and inv_dtl.promo_nbr = promo_item.promo_id and ");
			sql.append("      inv_dtl.tran_type = 'SALE' and inv_dtl.invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') ");
			sql.append("   join customer on customer.customer_id = inv_dtl.cust_nbr ");
			if(!m_PacketId.isEmpty() || !m_CustomerId.isEmpty() )
				sql.append(buildWhereClause(m_PacketId, m_CustomerId));
			sql.append("   group by ");
			sql.append("      packet.packet_id, packet.title, inv_dtl.cust_acct, inv_dtl.cust_nbr, ");
			sql.append("      inv_dtl.warehouse, promo_item.item_id, item.description, inv_dtl.vendor_name, ");
			sql.append("      customer.name ");
			sql.append("   order by ");
			sql.append("  	packet.packet_id, inv_dtl.cust_acct, inv_dtl.cust_nbr, warehouse, ");
			sql.append("  	promo_item.item_id ");
			sql.append(") packet_svc_lvl");

			// our main big honking query gets put together in buildsql
			m_PromoPacket = m_EdbConn.prepareStatement(sql.toString());
		}
	}

	private String buildWhereClause(String packetId, String customerId){

		StringBuffer whereClause = new StringBuffer (" where ");
		if(!packetId.isEmpty() )
		{
			StringBuffer sql = new StringBuffer (" packet.packet_id in (");

			String Plist[];
			Plist = packetId.split(",");     

			//for each packet
			for (int i = 0 ; i < Plist.length ; i++) {  		
				sql.append("'").append(Plist[i]).append("', ");
			}
			whereClause.append(sql.substring(0, sql.toString().length() -2) + ") ");
		}

		if( !customerId.isEmpty() )
		{
			if(!packetId.isEmpty() )
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
		closeStmt(m_PromoPacket);
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

			if ( param.name.equals("packetId") )
				m_PacketId = param.value.trim();

			if ( param.name.equals("custId") )
				m_CustomerId = param.value.trim();

			if ( param.name.equals("invoiceBeginDate") )
				m_InvoiceBeginDate = param.value.trim();

			if ( param.name.equals("invoiceEndDate") )
				m_InvoiceEndDate = param.value.trim();
		}

		//
		// Build the file name.
		fname.append(tm);
		fname.append("-");
		fname.append(m_RptProc.getUid());
		fname.append("promoservicelevel.xlsx");

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
				styleText,    // col 0 packet id
				styleText,    // col 1 title
				styleText,    // col 2 vendor name
				styleText,    // col 3 item id
				styleText,    // col 4 description
				styleInt,     // col 5 units ordered
				styleInt,     // col 6 units invoiced
				stylePct,     // col 7 fill rate pct				
				styleInt,     // col 8 emery on order
				styleText,    // col 9 customer acct
				styleText,    // col 10 customer number
				styleText,    // col 11 customer name
				styleText,    // col 12 warehouse				
		};

		styleText = null;
		styleInt = null;
		styleMoney = null;
		stylePct = null;
	}
}
