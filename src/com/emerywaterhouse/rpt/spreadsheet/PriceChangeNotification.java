/**
 * Title:			PriceChangeNotification.java
 * Description:   Creates a spreadsheet from PAR data and sends it as an email
 * 					attachment
 * Company:			Emery-Waterhouse
 * @author			prichter
 * @version			1.0
 * <p>
 * Create Date:	Nov 3, 2010
 * Last Update:   $Id: PriceChangeNotification.java,v 1.2 2012/05/22 08:20:01 prichter Exp $
 * <p>
 * History:
 *		$Log: PriceChangeNotification.java,v $
 *		Revision 1.2  2012/05/22 08:20:01  prichter
 *		Added support for other data sources (e.g. DPC)
 *
 *		Revision 1.1  2010/12/20 10:05:11  prichter
 *		Initial add
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.GregorianCalendar;

import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class PriceChangeNotification extends Report {
	
	// Parameters
	private String m_CustId;
	private Date m_RunDate;
	private int m_Months = 12;
	private String m_Source = "PAR";
	
	// Statements
	private PreparedStatement m_DpcData;
	private PreparedStatement m_ParData;
	
	// POI classes
   XSSFWorkbook m_Wrkbk = new XSSFWorkbook();
   XSSFSheet m_Sheet = m_Wrkbk.createSheet();
   private XSSFFont m_Font;
   private XSSFFont m_FontSubtitle;
   private XSSFFont m_FontTitle;
   private XSSFFont m_FontBold;
   private XSSFFont m_FontData;

   private XSSFCellStyle m_StyleTitle;    // size 12 Bold centered
   private XSSFCellStyle m_StyleSubtitle; // size 10 left-alligned
   private XSSFCellStyle m_StyleBoldText; //size 8 bold left-alligned
   private XSSFCellStyle m_StyleBoldNbr;  // size 8 bold right-alligned
   private XSSFCellStyle m_StyleBoldCtr;  // size 8 bold centered
   private XSSFCellStyle m_StyleText;  	// Text right justified
   private XSSFCellStyle m_StyleDec;      // Style with 2 decimals
   private XSSFCellStyle m_StyleTextCtr;	// Text style centered
   private XSSFCellStyle m_StyleInt;      // Style with 0 decimals
   private XSSFCellStyle m_StylePct;      // Style with 0 decimals + %
   private XSSFCellStyle m_StyleLabel;    // Text labels, right justify, 8pt
   
   private int m_Row = 0;
   private int m_Col = 0;
	
	FileOutputStream m_OutFile;  
	
	/**
	 * Constructor
	 */
	public PriceChangeNotification()
	{
		super();
		XSSFDataFormat format = m_Wrkbk.createDataFormat();
		
      //
      // Create the default font for this workbook
      m_Font = m_Wrkbk.createFont();
      m_Font.setFontHeightInPoints((short) 8);
      m_Font.setFontName("Arial");

      //
      // Create a font that is normal size & bold
      m_FontBold = m_Wrkbk.createFont();
      m_FontBold.setFontHeightInPoints((short)8);
      m_FontBold.setFontName("Arial");
      m_FontBold.setBold(true);

      //
      // Create a font that is normal size
      m_FontData = m_Wrkbk.createFont();
      m_FontData.setFontHeightInPoints((short)8);
      m_FontData.setFontName("Arial");
      
      //
      // Create a font for sub titles
      m_FontSubtitle = m_Wrkbk.createFont();
      m_FontSubtitle.setFontHeightInPoints((short)10);
      m_FontSubtitle.setFontName("Arial");
      m_FontSubtitle.setBold(true);

      //
      // Create a font for titles
      m_FontTitle = m_Wrkbk.createFont();
      m_FontTitle.setFontHeightInPoints((short)12);
      m_FontTitle.setFontName("Arial");
      m_FontTitle.setBold(true);

      //
      // Setup the cell styles used in this report
      m_StyleBoldCtr = m_Wrkbk.createCellStyle();
      m_StyleBoldCtr.setFont(m_FontBold);
      m_StyleBoldCtr.setAlignment(HorizontalAlignment.CENTER);
      m_StyleBoldCtr.setWrapText(true);

      m_StyleBoldText = m_Wrkbk.createCellStyle();
      m_StyleBoldText.setFont(m_FontBold);
      m_StyleBoldText.setAlignment(HorizontalAlignment.LEFT);
      m_StyleBoldText.setWrapText(true);

      m_StyleDec = m_Wrkbk.createCellStyle();
      m_StyleDec.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDec.setFont(m_FontData);
      m_StyleDec.setDataFormat((short)4);

      m_StyleInt = m_Wrkbk.createCellStyle();
      m_StyleInt.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleInt.setFont(m_FontData);
      m_StyleInt.setDataFormat((short)3);

      m_StyleLabel = m_Wrkbk.createCellStyle();
      m_StyleLabel.setFont(m_Font);
      m_StyleLabel.setAlignment(HorizontalAlignment.RIGHT);

      m_StylePct = m_Wrkbk.createCellStyle();
      m_StylePct.setAlignment(HorizontalAlignment.RIGHT);
      m_StylePct.setFont(m_FontData);
      m_StylePct.setDataFormat(format.getFormat("0.00%"));
      //m_StylePct.setDataFormat((short)4);


      m_StyleText = m_Wrkbk.createCellStyle();
      m_StyleText.setFont(m_FontData);
      m_StyleText.setAlignment(HorizontalAlignment.LEFT);
      m_StyleText.setWrapText(true);

      m_StyleTextCtr = m_Wrkbk.createCellStyle();
      m_StyleTextCtr.setFont(m_FontData);
      m_StyleTextCtr.setAlignment(HorizontalAlignment.CENTER);
      m_StyleTextCtr.setWrapText(true);

      m_StyleSubtitle = m_Wrkbk.createCellStyle();
      m_StyleSubtitle.setFont(m_FontSubtitle);
      m_StyleSubtitle.setAlignment(HorizontalAlignment.LEFT);
      
      m_StyleTitle = m_Wrkbk.createCellStyle();
      m_StyleTitle.setFont(m_FontTitle);
      m_StyleTitle.setAlignment(HorizontalAlignment.LEFT);
      m_StyleBoldNbr = m_Wrkbk.createCellStyle();
      m_StyleBoldNbr.setFont(m_FontBold);
      m_StyleBoldNbr.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleBoldNbr.setWrapText(true);
	}
	
	/**
	 * Clean up resources
	 */
	public void close() 
	{
		DbUtils.closeDbConn(null, m_DpcData, null);
		DbUtils.closeDbConn(null, m_ParData, null);
		DbUtils.closeDbConn(m_EdbConn, null, null);
		
		m_EdbConn = null;		
	}

   /**
    * Convenience method that adds a new String type cell with no borders and the specified alignment.
    *
    * @param rowNum int - the row index.
    * @param colNum short - the column index.
    * @param val String - the cell value.
    *
    * @return XSSFCell - the newly added String type cell, or a reference to the existing one.
    */
   private XSSFCell addCell(int rowNum, int colNum, String val, XSSFCellStyle style)
   {
      XSSFCell cell = addCell(rowNum, colNum);

      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(val));
      cell.setCellStyle(style);

      return cell;
   }

   /**
    * Convenience method that adds a new numeric type cell with no borders and the specified alignment.
    *
    * @param rowNum - the row index.
    * @param colNum short - the column index.
    * @param val double - the cell value.
    * @param style XSSFCellStyle - the cell style and format
    *
    * @return XSSFCell - the newly added numeric type cell, or a reference to the existing one.
    */
   private XSSFCell addCell(int rowNum, int colNum, double val, XSSFCellStyle style)
   {
      XSSFCell cell = addCell(rowNum, colNum);

      cell.setCellType(CellType.NUMERIC);
      cell.setCellStyle(style);
      cell.setCellValue(val);

      return cell;
   }

   /**
    * Adds a new cell with the specified borders and horizontal alignment.
    *
    * @param rowNum - the row index.
    * @param colNum short - the column index.
    *
    * @return XSSFCell - the newly added cell, or a reference to the existing one.
    */
   private XSSFCell addCell(int rowNum, int colNum)
   {
      XSSFRow row = addRow(rowNum);
      XSSFCell cell = row.getCell(colNum);

      if ( cell == null )
         cell = row.createCell(colNum);

      row = null;

      return cell;
   }

   /**
    * Adds a new row or returns the existing one.
    *
    * @param rowNum int - the row index.
    * @return XSSFRow - the row object added, or a reference to the existing one.
    */
   private XSSFRow addRow(int rowNum)
   {
      XSSFRow row = m_Sheet.getRow(rowNum);

      if ( row == null )
         row = m_Sheet.createRow(rowNum);

      return row;
   }

	/**
	 * Creates the output file 
	 * @return boolean - true if successful
	 * @throws FileNotFoundException 
	 * @throws SQLException 
	 */
	public boolean buildOutputFile() throws FileNotFoundException, SQLException
	{
		boolean built = false;		
      StringBuffer fileName = new StringBuffer(); 
      int itemCnt = 0;
      ResultSet rs = null;
      SimpleDateFormat fmt = new SimpleDateFormat("MM/dd/yyyy");
      String title = "PRICE CHANGE NOTIFICATION for " + fmt.format(m_RunDate);
      String custName = getCustName(m_CustId);
      
      fileName.append("PriceChgNotification");
      fileName.append("-");
      fileName.append(m_CustId);
      fileName.append("-");
      fileName.append(new SimpleDateFormat("MM-dd-yyyy").format(m_RunDate));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
      m_OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      fileName = null;

      m_Sheet.getPrintSetup().setLandscape(true);
      m_Sheet.setMargin(XSSFSheet.LeftMargin, .25);
      m_Sheet.setMargin(XSSFSheet.RightMargin, .25);
      m_Sheet.setMargin(XSSFSheet.BottomMargin, .5);
      m_Sheet.getHeader().setLeft(HeaderFooter.font("Arial", "Bold") + HeaderFooter.fontSize((short) 10) + HeaderFooter.date());
      m_Sheet.getHeader().setCenter(HeaderFooter.font("Arial", "Bold") + HeaderFooter.fontSize((short) 12) + title);
      m_Sheet.getHeader().setRight(HeaderFooter.font("Arial", "Bold") + HeaderFooter.fontSize((short) 12) + HeaderFooter.page() + " of " + HeaderFooter.numPages());
      
      m_Row = 0;
      
      m_RptProc.setEmailMsg(
      	"Attached is the Price Change Notification Report for " +
      	m_CustId + " " + custName + " for the reporting period ending " +
      	new SimpleDateFormat("MM-dd-yyyy").format(m_RunDate)
      );
      
   	addCell(m_Row, 1, "Customer: " + m_CustId, m_StyleBoldText);
   	addCell(m_Row, 2, custName, m_StyleBoldText);
   	m_Row++;
   	m_Row++;
   	
   	// Column headings and widths
   	m_Col = 0;
   	m_Sheet.setColumnWidth(m_Col, 1500);
   	addCell(m_Row, m_Col++, "Dept#", m_StyleBoldCtr);

   	m_Sheet.setColumnWidth(m_Col, 3500);
   	addCell(m_Row, m_Col++, "Department", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 6000);
   	addCell(m_Row, m_Col++, "Vendor Name", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 1500);
   	addCell(m_Row, m_Col++, "FLC", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 2000);
   	addCell(m_Row, m_Col++, "Item Id", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 10000);
   	addCell(m_Row, m_Col++, "Item Description", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 3000);
   	addCell(m_Row, m_Col++, "UPC", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 1800);
   	addCell(m_Row, m_Col++, "New Cost", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 1800);
   	addCell(m_Row, m_Col++, "Old Cost", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 1400);
   	addCell(m_Row, m_Col++, "UOM", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 1800);
   	addCell(m_Row, m_Col++, "New Retail", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 1800);
   	addCell(m_Row, m_Col++, "Old Retail", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 1200);
   	addCell(m_Row, m_Col++, "CRP", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 1800);
   	addCell(m_Row, m_Col++, "New Margin", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 1800);
   	addCell(m_Row, m_Col++, "Old Margin", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 2200);
   	addCell(m_Row, m_Col++, "Last Buy", m_StyleBoldCtr);
   	
   	m_Sheet.setColumnWidth(m_Col, 1500);
   	addCell(m_Row, m_Col++, "Last Qty", m_StyleBoldCtr);
   	
   	m_Row++;
      
      try {
      	if ( m_Source.equals("PAR") ) {
      		m_ParData.setString(1, m_CustId);
         	m_ParData.setInt(2, m_Months);
         	m_ParData.setString(3, m_CustId);
         	m_ParData.setDate(4, m_RunDate);
         	
         	rs = m_ParData.executeQuery();
      	}
      	
      	else {
      		m_DpcData.setString(1, m_CustId);
      		m_DpcData.setInt(2, m_Months);
      		m_DpcData.setString(3, m_CustId);
      		m_DpcData.setDate(4, m_RunDate);
      		
      		rs = m_DpcData.executeQuery();
      	}
      	
      	while ( rs.next() ) {
      		m_Col = 0;
      		setCurAction("Processing customer " + m_CustId + " item " + rs.getString("item_id"));
      		addCell(m_Row, m_Col++, rs.getString("dept_num"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("department"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("vendor_name"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("flc_id"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("item_id"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("item_descr"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("upc_code"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getDouble("new_sell"), m_StyleDec);
      		addCell(m_Row, m_Col++, rs.getDouble("old_sell"), m_StyleDec);
      		addCell(m_Row, m_Col++, rs.getString("unit"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getDouble("new_retail"), m_StyleDec);
      		addCell(m_Row, m_Col++, rs.getDouble("old_retail"), m_StyleDec);
      		addCell(m_Row, m_Col++, rs.getString("crp"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getDouble("new_margin"), m_StylePct);
      		addCell(m_Row, m_Col++, rs.getDouble("old_margin"), m_StylePct);
      		
      		if ( rs.getDate("invoice_date") != null ) {
      			addCell(m_Row, m_Col++, fmt.format(rs.getDate("invoice_date")), m_StyleText);
      			addCell(m_Row, m_Col++, rs.getInt("qty_shipped"), m_StyleInt);
      		}
      		
      		m_Row++;
      		itemCnt++;
      	}
      	
      	m_Row++;
      	addCell(m_Row, 1, "Items reported: " + itemCnt, m_StyleBoldText);
      	
			m_Wrkbk.write(m_OutFile);
			m_OutFile.close();
			
			built = true;
      }
      
      catch ( Exception e ) {
         log.fatal(this.getClass().getName() + " exception:", e);
      }
      
      finally {
      	m_Wrkbk = null;
      	m_OutFile = null;
      }

		return built;
	}

	/* (non-Javadoc)
	 * @see com.emerywaterhouse.rpt.server.Report#createReport()
	 */
	@Override
	public boolean createReport()
	{
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         if ( m_EdbConn == null && m_RptProc != null )
         	m_EdbConn = m_RptProc.getEdbConn();
         
         if ( prepareStatements() )
      	  created = buildOutputFile();            
      }
      
      catch ( Exception ex ) {
         log.fatal(this.getClass().getName() + " exception:", ex);
      }
      
      finally {
         close(); 
         
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }
      
      return created;
	}
	
	/**
	 * Returns the name of the customer the spreadsheet it run for
	 * 
	 * @param custId String - the customer id
	 * @return String - the custome name
	 * @throws SQLException
	 */
	private String getCustName(String custId) throws SQLException
	{
		String name = "Invalid Customer#";
		Statement stmt = null;
		ResultSet rs = null;
		
		try {
			stmt = m_EdbConn.createStatement();
			rs = stmt.executeQuery("select name from customer where customer_id = '" + custId + "'");
			
			if ( rs.next() )
				name = rs.getString("name");
		}
		
		finally {
			DbUtils.closeDbConn(null, stmt, rs);
			rs = null;
			stmt = null;
		}
		
		return name;
	}
	
	/**
	 * Prepares queries
	 * 
	 * @return boolean - true if successfull
	 * @throws Exception
	 */
	private boolean prepareStatements() throws Exception
	{
		StringBuffer sql = new StringBuffer();
		
		try {
         sql.setLength(0);
         sql.append("select dpc.dpc_date,  ");
         sql.append("       emery_dept.dept_num, ");
         sql.append("       emery_dept.name Dedpctment, vendor.name vendor_name,  ");
         sql.append("       ejd_item.flc_id, item_entity_attr.item_id, item_entity_attr.description item_descr, ");
         sql.append("       ejd_item_whs_upc.upc_code, dpc.new_sell, dpc.old_sell, ");
         sql.append("       retail_unit.unit, dpc.new_retail, dpc.old_retail, ");
         sql.append("       decode(ejd_price_procs.get_retail_price(dpc.customer_id, item_entity_attr.item_ea_id), -1, ' ', 'Y') crp, ");
         sql.append("       decode(dpc.new_retail, 0, 0, round((dpc.new_retail - dpc.new_sell) / dpc.new_retail,3)) new_margin, ");
         sql.append("       decode(dpc.old_retail, 0, 0, round((dpc.old_retail - dpc.old_sell) / dpc.old_retail,3)) old_margin, ");
         sql.append("       last_sale.invoice_date, ");
         sql.append("       last_sale.qty_shipped ");
         sql.append("from dpc ");         
         sql.append("join item_entity_attr on item_entity_attr.item_id = dpc.item_id ");
         sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
         sql.append("join emery_dept on emery_dept.dept_id = ejd_item.dept_id ");
         sql.append("join retail_unit on retail_unit.unit_id = item_entity_attr.ret_unit_id ");
         sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
         sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = item_entity_attr.ejd_item_id and ");
         sql.append("                            ejd_item_whs_upc.primary_upc = 1 ");                
         sql.append("left outer join ( ");
         sql.append("  select inv_dtl.cust_acct, inv_dtl.item_nbr, inv_dtl.invoice_date, sum(inv_dtl.qty_shipped) qty_shipped ");
         sql.append("  from inv_dtl ");
         sql.append("  where inv_dtl.tran_type = 'SALE' and ");
         sql.append("        inv_dtl.sale_type = 'WAREHOUSE' and ");
         sql.append("        inv_dtl.cust_acct = ? and ");
         sql.append("        inv_dtl.qty_shipped > 0 and ");
         sql.append("        inv_dtl.invoice_date = ( ");
         sql.append("           select max(dtl.invoice_date) ");
         sql.append("           from inv_dtl dtl ");
         sql.append("           where dtl.invoice_date > add_months(current_date, ? * -1) and ");                            
         sql.append("                 dtl.cust_acct = inv_dtl.cust_acct and ");
         sql.append("                 dtl.item_nbr = inv_dtl.item_nbr and ");
         sql.append("                 dtl.sale_type = 'WAREHOUSE' and ");
         sql.append("                 dtl.tran_type = 'SALE' and ");
         sql.append("                 dtl.qty_shipped > 0) ");
         sql.append("  group by inv_dtl.cust_acct, inv_dtl.item_nbr, inv_dtl.invoice_date ");
         sql.append(") last_sale on last_sale.cust_acct = dpc.customer_id and ");
         sql.append("               last_sale.item_nbr = item_entity_attr.item_id ");      
         sql.append("where dpc.customer_id = ? and ");
         sql.append("      dpc.dpc_date = ? and ");
         sql.append("      (dpc.new_sell <> dpc.old_sell or dpc.new_retail <> dpc.old_retail ) ");
         sql.append("order by emery_dept.name, ejd_item.flc_id, dpc.item_id ");                     
         m_DpcData = m_EdbConn.prepareStatement(sql.toString());

			sql.setLength(0);
			sql.append("select par.par_date,  ");
			sql.append("       emery_dept.dept_num, ");
			sql.append("       emery_dept.name Department, vendor.name vendor_name,  ");
			sql.append("       ejd_item.flc_id, item_entity_attr.item_id, item_entity_attr.description item_descr, ");
            sql.append("       ejd_item_whs_upc.upc_code, par.new_sell, par.old_sell, ");
			sql.append("       retail_unit.unit, par.new_retail, par.old_retail, ");
            sql.append("       decode(ejd_price_procs.get_retail_price(par.customer_id, item_entity_attr.item_ea_id), -1, ' ', 'Y') crp, ");			
            sql.append("       decode(par.new_retail, 0, 0, round((par.new_retail - par.new_sell) / par.new_retail,3)) new_margin, ");
			sql.append("       decode(par.old_retail, 0, 0, round((par.old_retail - par.old_sell) / par.old_retail,3)) old_margin, ");
			sql.append("       last_sale.invoice_date, ");
			sql.append("       last_sale.qty_shipped ");
			sql.append("from par ");
            sql.append("join item_entity_attr on item_entity_attr.item_ea_id = par.item_ea_id ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join emery_dept on emery_dept.dept_id = ejd_item.dept_id ");
            sql.append("join retail_unit on retail_unit.unit_id = item_entity_attr.ret_unit_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = item_entity_attr.ejd_item_id and ");
            sql.append("                            ejd_item_whs_upc.primary_upc = 1 ");
			sql.append("left outer join ( ");
			sql.append("  select inv_dtl.cust_acct, inv_dtl.item_nbr, inv_dtl.invoice_date, sum(inv_dtl.qty_shipped) qty_shipped ");
			sql.append("  from inv_dtl ");
			sql.append("  where inv_dtl.tran_type = 'SALE' and ");
			sql.append("        inv_dtl.sale_type = 'WAREHOUSE' and ");
			sql.append("        inv_dtl.cust_acct = ? and ");
			sql.append("        inv_dtl.qty_shipped > 0 and ");
			sql.append("        inv_dtl.invoice_date = ( ");
			sql.append("           select max(dtl.invoice_date) ");
			sql.append("           from inv_dtl dtl ");
            sql.append("           where dtl.invoice_date > add_months(current_date, ? * -1) and ");
            sql.append("                 dtl.cust_acct = inv_dtl.cust_acct and ");
			sql.append("                 dtl.item_nbr = inv_dtl.item_nbr and ");
			sql.append("                 dtl.sale_type = 'WAREHOUSE' and ");
			sql.append("                 dtl.tran_type = 'SALE' and ");
			sql.append("                 dtl.qty_shipped > 0) ");
			sql.append("  group by inv_dtl.cust_acct, inv_dtl.item_nbr, inv_dtl.invoice_date ");
			sql.append(") last_sale on last_sale.cust_acct = par.customer_id and ");
            sql.append("               last_sale.item_nbr = item_entity_attr.item_id ");      			
            sql.append("where par.customer_id = ? and ");
            sql.append("      par.par_date = ? and ");
			sql.append("      (par.new_sell <> par.old_sell or par.new_retail <> par.old_retail ) ");
            sql.append("order by emery_dept.name, ejd_item.flc_id, par.item_id ");
            m_ParData = m_EdbConn.prepareStatement(sql.toString());
			
			return true;
		}
		
		finally {
			sql = null;
		}
	}
	
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params) 
   {
   	String name;
   	int yr, mo, dy;
   	String tmp;

      for ( int i = 0; i < params.size(); i++ ) {
      	name = params.get(i).name;
      	
      	if ( name.equals("custid")) { 
      		m_CustId = params.get(i).value.trim();
      		continue;
      	}
      	
      	if ( name.equals("pardate") ) {   // format yyyymmdd
      		tmp = params.get(i).value.trim();
      		yr = Integer.valueOf(tmp.substring(0, 4));
      		mo = Integer.valueOf(tmp.substring(4, 6)) - 1;
      		dy = Integer.valueOf(tmp.substring(6));
      		
      		m_RunDate = new Date(new GregorianCalendar(yr, mo, dy).getTimeInMillis());
      		continue;
      	}
      	
      	if ( name.equals("historymonths") ) {
      		m_Months = Integer.valueOf(params.get(i).value.trim() );
      		continue;
      	}
      	
      	if ( name.equals("sourcesystem") ) {
      		tmp = params.get(i).value.trim();
      		
      		if ( tmp.equals("PAR") || tmp.equals("DPC") )
      			m_Source = tmp;
      		
      		else
      			log.error("PriceChangeNotification.setParams(). Invalid source system " + tmp + ". Assuming PAR" );
      	}
      }
   }

   /*
   public static void main(String... args) throws SQLException 
   {
        System.out.println(Calendar.getInstance().getTime());
        BasicConfigurator.configure();
        PriceChangeNotification pcn = new PriceChangeNotification();
        pcn.log = Logger.getLogger(Report.class);

        Param[] parms = new Param[] {
              new Param("string", "057924", "custid"), 
              new Param("string", "20180420", "pardate"),
              new Param("integer", "1", "historymonths"),
              new Param("string", "PAR", "sourcesystem")
        };
        
        ArrayList<Param> parmslist = new ArrayList<Param>();
        for (Param p : parms) {
           parmslist.add(p);
        }
        
        pcn.setParams(parmslist);

        Connection conn;
        Properties connProps = new Properties();
        connProps.put("user", "ejd");
        connProps.put("password", "boxer");
        conn = DriverManager.getConnection("jdbc:edb://172.30.1.33/emery_jensen", connProps);
        pcn.m_EdbConn = conn;

        pcn.m_FilePath = "C:/Users/bcornwell/temp/";
        boolean res = pcn.createReport();
        System.out.println(res);
        System.out.println(Calendar.getInstance().getTime());
   }
   */
}

