/**
 * File: CustShowOrderSummary.java
 * Description: Summary of customer orders by TM for the Emery market place show.
 *
 * @author Stephen Martel
 * 
 * Create Data: 03/02/2016
 * Last Update: 
 *
 * History:
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.LinkedList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

/**
 * 
 */
public class CustShowOrderSummary extends Report 
{
   private String m_FName;
   private String m_LName;
   
   private XSSFWorkbook m_Wrkbk;
   private Sheet m_Sheet;
   private Row m_Row;
   private Font m_FontNorm;
   private Font m_FontBold;
   private XSSFCellStyle m_StyleHdrLeft = null;
   private XSSFCellStyle m_StyleTxtC = null;      // Text centered
   private XSSFCellStyle m_StyleTxtL = null;      // Text left justified
   private XSSFCellStyle m_StyleTxtLB = null;     // Text left justified bold
   private XSSFCellStyle m_StyleInt = null;       // Style with 0 decimals
   private XSSFCellStyle m_StyleDouble = null;    // numeric #,##0.00
   private XSSFCellStyle m_StyleDoubleB = null;   // numeric #,##0.00 bold
   
   private PreparedStatement m_OrderData = null;
   private PreparedStatement m_PrevShowTotal = null;
   private PreparedStatement m_CurrShowTotal = null;
   private PreparedStatement m_CurrShowShipped = null;
   private PreparedStatement m_CurrShowUnshipped = null;
   private PreparedStatement m_CurrShowCart = null;
   private PreparedStatement m_GetTMNames = null;
   private PreparedStatement m_GetDirectorNames = null;
   
   /**
    * 
    */
   public CustShowOrderSummary() 
   {
      super();
      
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet("Show Orders");
      
      setupWorkbook();
   }

   /**
    * adds a numeric type cell to current row at col p_Col in current sheet
    *
    * @param col 0-based column number of spreadsheet cell
    * @param value numeric value to be stored in cell
    * @param style Excel style to be used to display cell
    */
   private void addCell(int col, double value, CellStyle style)
   {
      Cell cell = m_Row.createCell(col);
      cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
      cell.setCellStyle(style);
      cell.setCellValue(value);
      
      cell = null;
   }
   
   /**
    * adds a text type cell to current row at col p_Col in current sheet
    *
    * @param col 0-based column number of spreadsheet cell
    * @param value text value to be stored in cell
    * @param style Excel style to be used to display cell
    */
   private void addCell(int col, String value, CellStyle style)
   {
      Cell cell = m_Row.createCell(col);
      cell.setCellType(Cell.CELL_TYPE_STRING);
      cell.setCellValue(new XSSFRichTextString(value));
      cell.setCellStyle(style);
      cell = null;
   }
   
   /**
    * adds row to the current sheet
    *
    * @param row  0-based row number of row to be added
    */
   private void addRow(int row) 
   {
      m_Row = m_Sheet.createRow(row);
   }
   
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      boolean result = false;
      ResultSet orderData = null, tmNames = null, dirNames = null;
      int curRow = 10;
      String custId = null;
      boolean isDirector = true; // whether this is a single tm or a director
      boolean isAll = true; // whether or not this report is running for all divisions / tms
      int currentSummaryLine = 3; // the current row to write summary lines on (used when doing an 'All' report)
      
      try {
    	  
    	 // if the first name is 'All', we need to loop over every director, and every tm.
    	 //    this will be 4 loops over the division code, and then output the total of all divisions.
    	 // otherwise, if it isn't 'All', just put the name in the directorNames collection anyway
    	 //    so that we'll loop over the division code once, and if it is a director,
    	 //    it will grab every tm under them, and if it is a tm, it will grab only the single tm.
    	 
    	 LinkedList<String> directorNames = new LinkedList<String>();
    	 // if 'All', populate the director lists with every director name
    	 if (m_FName.equalsIgnoreCase("All")) {
    		 dirNames = m_GetDirectorNames.executeQuery();
    		 while (dirNames.next()) {
    			 directorNames.add(dirNames.getString(1));
    		 }
    	 } else { // just use the supplied name
    		 directorNames.add(m_FName);
    		 isAll = false;
    	 } 
    	 
    	 // column sums for all divisions
   	     double sumAllPrevShowTotals = 0;
   	     double sumAllCurrShowTotals = 0;
   	     double sumAllCarts = 0;
   	     double sumAllUnshipped = 0;
   	     double sumAllShipped = 0;
   	  
   	     for (int i = 0; i < directorNames.size(); ++i) {
   	    	 
	    	 LinkedList<String> tmFirstNames = new LinkedList<String>();
	    	 LinkedList<String> tmLastNames = new LinkedList<String>();
	    	 String division = "";
	    	  
	    	 // column sums for the Division (sum of every TM and every customer's 2015 total, 2016 total, show carts, ...)
		     double sumDivPrevShowTotals = 0;
		     double sumDivCurrShowTotals = 0;
		     double sumDivCarts = 0;
		     double sumDivUnshipped = 0;
		     double sumDivShipped = 0;
	    	 
	    	 // run the directors query with the supplied names.
	    	 // if there are results, they are a director, loop over all of their tms and grab data
	    	 // otherwise, if there are no results, then the passed in name is just a single tm, get their data
	    	 m_GetTMNames.setString(1, directorNames.get(i));
	    	 tmNames = m_GetTMNames.executeQuery();
	    	 while(tmNames.next()) {
	    		tmFirstNames.add(tmNames.getString(1));
	    		tmLastNames.add(tmNames.getString(2));
	  		    if (division.isEmpty())
				   division = tmNames.getString(3);
	    	 }
	    	 if (tmFirstNames.size() == 0) {
	    		tmFirstNames.add(m_FName);
	    	  	tmLastNames.add(m_LName);
	  		    isDirector = false;
	    	 } 
	    	  
	    	 // loop over all appropriate tms
	    	 for (int x = 0; x < tmFirstNames.size(); ++x) {
	
		         m_OrderData.setString(1, tmFirstNames.get(x));
		         m_OrderData.setString(2, tmLastNames.get(x));
		         orderData = m_OrderData.executeQuery();
	
		         String name = tmFirstNames.get(x) + " " + tmLastNames.get(x);
		         
		         // column sums for the TM (sum of every customer's 2015 total, 2016 total, show carts, ...)
		         double sumTmPrevShowTotals = 0;
		         double sumTmCurrShowTotals = 0;
		         double sumTmCarts = 0;
		         double sumTmUnshipped = 0;
		         double sumTmShipped = 0;
		         
		         while ( orderData.next() && m_Status == RptServer.RUNNING ) {
		            custId = orderData.getString(1);
		            
		            // TODO hardcoded the 5 fields below here to pull 2015/2016 show data
		            double prevShowTotal = 0, currShowTotal = 0, cartTotal = 0, unshippedTotal = 0, shippedTotal = 0;
		            prevShowTotal = getShowTotal(custId, name, 2015, 30);  // previous show, show id
		            currShowTotal = getShowTotal(custId, name, 2016, 30);  // current show, show id
		            cartTotal = getShowCart(custId, 30);             // show id
		            unshippedTotal = getUnshipped(custId, 30);       // current show
		            shippedTotal = getShipped(custId, 30);           // current show
		            
		            // add to TM sums
		            sumTmPrevShowTotals += prevShowTotal;
		            sumTmCurrShowTotals += currShowTotal;
		            sumTmCarts += cartTotal;
		            sumTmUnshipped += unshippedTotal;
		            sumTmShipped += shippedTotal;

			        if (!isAll) { // they decided they only want tm totals in the all report, not details
			            addRow(curRow);
			            addCell(0, custId, m_StyleTxtL);
			            addCell(1, orderData.getString(2), m_StyleTxtL);
			            addCell(2, tmFirstNames.get(x) + " " + tmLastNames.get(x), m_StyleTxtL);
			            addCell(3, prevShowTotal, m_StyleDouble);
			            addCell(4, currShowTotal, m_StyleDouble);
			            addCell(5, cartTotal, m_StyleDouble);
			            addCell(6, unshippedTotal, m_StyleDouble);
			            addCell(7, shippedTotal, m_StyleDouble);
			    		if (division.isEmpty())
			    			division = orderData.getString(3);
				        addCell(8, division, m_StyleTxtL);
			            ++curRow;
			        }
		            
		         }
		         
		         // add TM total row
		         addRow(curRow);
		         addCell(0, "Totals", m_StyleTxtLB);
		         addCell(1, "All Customers", m_StyleTxtLB);
		         addCell(2, tmFirstNames.get(x) + " " + tmLastNames.get(x), m_StyleTxtLB);
		         addCell(3, sumTmPrevShowTotals, m_StyleDoubleB);
		         addCell(4, sumTmCurrShowTotals, m_StyleDoubleB);
		         addCell(5, sumTmCarts, m_StyleDoubleB);
		         addCell(6, sumTmUnshipped, m_StyleDoubleB);
		         addCell(7, sumTmShipped, m_StyleDoubleB);
		         addCell(8, division, m_StyleTxtLB);
		         // if this is a TM only report (not division or all), then add it to summary at top too
			     if (!isDirector) {
			         addRow(3);
			         addCell(0, "Totals", m_StyleTxtLB);
			         addCell(1, tmFirstNames.get(x) + " " + tmLastNames.get(x), m_StyleTxtLB);
			         addCell(3, sumTmPrevShowTotals, m_StyleDoubleB);
			         addCell(4, sumTmCurrShowTotals, m_StyleDoubleB);
			         addCell(5, sumTmCarts, m_StyleDoubleB);
			         addCell(6, sumTmUnshipped, m_StyleDoubleB);
			         addCell(7, sumTmShipped, m_StyleDoubleB);
		         }
	
		         // add TM totals to division totals
		         sumDivPrevShowTotals += sumTmPrevShowTotals;
		         sumDivCurrShowTotals += sumTmCurrShowTotals;
		         sumDivCarts += sumTmCarts;
		         sumDivUnshipped += sumTmUnshipped;
		         sumDivShipped += sumTmShipped;
		         
		         curRow += 2; // leaves a row of space between TMs
	    	 }
	    	 
	    	 if (isDirector) {
	    		// add Division total row
		         addRow(curRow);
		         addCell(0, "Totals", m_StyleTxtLB);
		         addCell(1, division + " Division", m_StyleTxtLB);
		         addCell(3, sumDivPrevShowTotals, m_StyleDoubleB);
		         addCell(4, sumDivCurrShowTotals, m_StyleDoubleB);
		         addCell(5, sumDivCarts, m_StyleDoubleB);
		         addCell(6, sumDivUnshipped, m_StyleDoubleB);
		         addCell(7, sumDivShipped, m_StyleDoubleB);
		         // also add the summary line to the top
		         addRow(currentSummaryLine++);
		         addCell(0, "Totals", m_StyleTxtLB);
		         addCell(1, division + " Division", m_StyleTxtLB);
		         addCell(3, sumDivPrevShowTotals, m_StyleDoubleB);
		         addCell(4, sumDivCurrShowTotals, m_StyleDoubleB);
		         addCell(5, sumDivCarts, m_StyleDoubleB);
		         addCell(6, sumDivUnshipped, m_StyleDoubleB);
		         addCell(7, sumDivShipped, m_StyleDoubleB);
	
		         // add Division totals to All-Totals
		         sumAllPrevShowTotals += sumDivPrevShowTotals;
		         sumAllCurrShowTotals += sumDivCurrShowTotals;
		         sumAllCarts += sumDivCarts;
		         sumAllUnshipped += sumDivUnshipped;
		         sumAllShipped += sumDivShipped;
		         
		         curRow += 2;
	    	 }
   	     }
   	     
   	     if (isAll) {
	    	 // add All total row
	    	 addRow(currentSummaryLine++);
	    	 addCell(0, "Totals", m_StyleTxtLB);
		     addCell(1, "All Divisions", m_StyleTxtLB);
		     addCell(3, sumAllPrevShowTotals, m_StyleDoubleB);
		     addCell(4, sumAllCurrShowTotals, m_StyleDoubleB);
		     addCell(5, sumAllCarts, m_StyleDoubleB);
		     addCell(6, sumAllUnshipped, m_StyleDoubleB);
		     addCell(7, sumAllShipped, m_StyleDoubleB);
	    	 ++curRow;
	     }
         
         m_Wrkbk.write(outFile);         
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[CustShowOrderSummary]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, m_OrderData, orderData);
      }
      
      return result;
   }
   
   /**
    * Closes prepared statements and cleans up member variables
    */
   protected void cleanup()
   {
      m_FName = null;
      m_LName = null;
      
      m_OrderData = null;

      m_Sheet = null;
      m_Wrkbk = null;
      
      m_FontNorm = null;
      m_FontBold = null;
      
      m_StyleHdrLeft = null;
      m_StyleTxtC = null;
      m_StyleTxtL = null;      
      m_StyleInt = null;
      m_StyleDouble = null;      
   }
   
   /** 
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport() 
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;

      try {
         m_EdbConn = m_RptProc.getEdbConn();
         
         if (prepareStatements())
            created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[CustShowOrderSummary]", ex);
      }

      finally {
        cleanup();

        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * defines Excel fonts and styles
    */
   private void defineStyles()
   {
      short fontHeight = 11;
      XSSFDataFormat format = m_Wrkbk.createDataFormat();
      
      //
      // Normal Font
      m_FontNorm = m_Wrkbk.createFont();
      m_FontNorm.setFontName("Arial");
      m_FontNorm.setFontHeightInPoints(fontHeight);
      
      //
      // defines bold font
      m_FontBold = m_Wrkbk.createFont();
      m_FontBold.setFontName("Arial");
      m_FontBold.setFontHeightInPoints(fontHeight);
      m_FontBold.setBoldweight(Font.BOLDWEIGHT_BOLD);
      
      //
      // defines style column header, left-justified
      m_StyleHdrLeft = m_Wrkbk.createCellStyle();
      m_StyleHdrLeft.setFont(m_FontBold);
      m_StyleHdrLeft.setAlignment(XSSFCellStyle.ALIGN_LEFT);
      m_StyleHdrLeft.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
      
      m_StyleTxtL = m_Wrkbk.createCellStyle();
      m_StyleTxtL.setAlignment(XSSFCellStyle.ALIGN_LEFT);
      
      m_StyleTxtLB = m_Wrkbk.createCellStyle();
      m_StyleTxtLB.setAlignment(XSSFCellStyle.ALIGN_LEFT);
      m_StyleTxtLB.setFont(m_FontBold);
      
      m_StyleTxtC = m_Wrkbk.createCellStyle();
      m_StyleTxtC.setAlignment(XSSFCellStyle.ALIGN_CENTER);
      
      m_StyleInt = m_Wrkbk.createCellStyle();
      m_StyleInt.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
      m_StyleInt.setDataFormat((short)3);
      
      m_StyleDouble = m_Wrkbk.createCellStyle();
      m_StyleDouble.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
      m_StyleDouble.setDataFormat(format.getFormat("$#,##0"));

      m_StyleDoubleB = m_Wrkbk.createCellStyle();
      m_StyleDoubleB.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
      m_StyleDoubleB.setDataFormat(format.getFormat("$#,##0"));
      m_StyleDoubleB.setFont(m_FontBold);
   }
   
   /**
    * 
    * @param custId
    * @param showYear
    * @return
    * @throws SQLException 
    */
   private double getShipped(String custId, int showId) throws SQLException
   {
      double total = 0.00;
      m_CurrShowShipped.setInt(1, showId);
      m_CurrShowShipped.setInt(2, showId);
      m_CurrShowShipped.setInt(3, showId);
      m_CurrShowShipped.setString(4, custId);
      ResultSet rs = m_CurrShowShipped.executeQuery();
      if (rs.next())
         total = rs.getDouble(1);
      return total;
   }
   
   /**
    * 
    * @param custId
    * @param showYear
    * @return
    * @throws SQLException 
    */
   private double getUnshipped(String custId, int showId) throws SQLException
   {
      double total = 0.00;
      m_CurrShowUnshipped.setInt(1, showId);
      m_CurrShowUnshipped.setInt(2, showId);
      m_CurrShowUnshipped.setInt(3, showId);
      m_CurrShowUnshipped.setString(4, custId);
      ResultSet rs = m_CurrShowUnshipped.executeQuery();
      if (rs.next())
         total = rs.getDouble(1);
      return total;
   }
   
   /**
    * 
    * @param custId
    * @param showYear
    * @return
    * @throws SQLException 
    */
   private double getShowCart(String custId, int showId) throws SQLException
   {
      double total = 0.00;
      m_CurrShowCart.setInt(1, showId);
      m_CurrShowCart.setInt(2, showId);
      m_CurrShowCart.setInt(3, showId);
      m_CurrShowCart.setString(4, custId);
      ResultSet rs = m_CurrShowCart.executeQuery();
      if (rs.next())
         total = rs.getDouble(1);
      return total;
   }
   
   /**
    * 
    * @param custId
    * @param showYear
    * @return
    * @throws Exception 
    */
   private double getShowTotal(String custId, String name, int showYear, int showId) throws Exception
   {
      double total = 0.00;
      
      // How they want the reporting seems to change every year.
      // Additionally, there are a number of packets / promos that need to be included in the reporting,
      // but are not in the show_packet table, meaning they need to be hardcoded or stored elsewhere / in a new table.
      
      if (showYear == 2015) {
         m_PrevShowTotal.setString(1, name);
         m_PrevShowTotal.setString(2, custId);
         ResultSet rs = m_PrevShowTotal.executeQuery();
         if (rs.next())
            total = rs.getDouble(1);
      } else if (showYear == 2016) {
         m_CurrShowTotal.setInt(1, showId);
         m_CurrShowTotal.setInt(2, showId);
         m_CurrShowTotal.setInt(3, showId);
         m_CurrShowTotal.setInt(4, showId);
         m_CurrShowTotal.setInt(5, showId);
         m_CurrShowTotal.setInt(6, showId);
         m_CurrShowTotal.setString(7, custId);
         ResultSet rs = m_CurrShowTotal.executeQuery();
         if (rs.next())
            total = rs.getDouble(1);
      } else {
         throw new Exception("Invalid Show year supplied: " + showYear);
      }
      
      return total;
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(256);      
      boolean isPrepared = false;
      
      if ( m_EdbConn != null ) {
         try {
            sql.append("select ");
            sql.append("customer.customer_id, customer.name cust_name, division ");
            sql.append("from customer ");
            sql.append("join show_cust on show_cust.customer_id = customer.customer_id ");
            sql.append("join show on show.show_id = show_cust.show_id and show.name = '2016 Marketplace' ");            
            sql.append("join cust_rep on cust_rep.customer_id = customer.customer_id ");
            sql.append("join emery_rep on cust_rep.er_id = emery_rep.er_id and emery_rep.first = ? and emery_rep.last = ? ");
            sql.append("join emery_rep_type on emery_rep_type.rep_type_id = cust_rep.rep_type_id  and emery_rep_type.description = 'SALES REP' ");
            sql.append("join sales_division on sales_division.division_id = emery_rep.division_id ");
            sql.append("order by customer.customer_id ");
            m_OrderData = m_EdbConn.prepareStatement(sql.toString());
            
            // 2015 show total (2015 invoices)
            sql.setLength(0);
            sql.append("select round(sum(inv_dtl.ext_sell),2) total ");
            sql.append("   from inv_hdr, inv_dtl, cust_view, promotion, packet, promo_view, vendor_buyers ");
            sql.append("where ");
            sql.append("   inv_hdr.cust_nbr = cust_view.customer_id and ");
            sql.append("   inv_hdr.invoice_nbr = inv_dtl.invoice_nbr and ");
            sql.append("   inv_dtl.promo_nbr = promotion.promo_id and ");
            sql.append("   inv_dtl.promo_nbr = promo_view.promo_id and ");
            sql.append("   inv_dtl.vendor_nbr = vendor_buyers.vendor_id(+) and ");
            sql.append("   promotion.packet_id = packet.packet_id and ");
            sql.append("   promotion.promo_id in ('1173','1174','1153','1154','1155','1158','1181','1182','1183','1184','1185','1186','1196','1188','1189','1191','1192','1197','1235','1237','1239','1244','1198','1201','1202','1203','1204','1205','1206','1207','1208','1209','1210','1211','1212','1213','1214','1215','1216','1217','1218','1219','1220','1221','1222','1223','1224','1225','1226','1227','1228','1229','1230','1231','1232','1236','1238','1240','1241','1242','1243','1399','1195','7004','7008','7009','7016','7018','7017','7108','7109','9967','9968','9969','9971','9972','9973','9974','9965','9975','9976','9977','9978','9964','9979','9981','9982','9983','9984','9985','9986','9987','9989','9994','9995','9997','9998','9999','9966') ");
            sql.append("   and inv_dtl.qty_shipped > 0 ");
            sql.append("   and inv_hdr.invoice_date between ('01-OCT-2014') and ('01-AUG-2015') ");
            sql.append("   and sales_rep = ? ");
            sql.append("   and inv_hdr.cust_nbr = ? ");
            sql.append("group by cust_view.sales_rep, inv_hdr.cust_nbr ");
            m_PrevShowTotal = m_EdbConn.prepareStatement(sql.toString());
            
            // 2016 show total
            sql.setLength(0);
            sql.append("SELECT NVL(Show_Cart.total_dollars, 0) + NVL(Reg_Orders.total_dollars, 0) total_dollars ");
            sql.append("FROM customer ");
            sql.append("LEFT OUTER JOIN ");
            sql.append("  (SELECT customer_id, COUNT(DISTINCT order_header.order_id) show_orders, COUNT(DISTINCT ol_id) total_lines, ");
            sql.append("     SUM(qty_ordered * sell_price)         total_dollars ");
            sql.append("   FROM order_header ");
            sql.append("   JOIN order_line ON order_line.order_id = order_header.order_id ");
            sql.append("   JOIN order_status line_status ON line_status.order_status_id = order_line.order_status_id ");
            sql.append("   WHERE ((packet_id IN (437, 486, 487) OR packet_id IN ");
            sql.append("                                           (SELECT packet_id ");
            sql.append("                                            FROM show_packet ");
            sql.append("                                            WHERE show_id = ?  ");
            sql.append("                                           )) ");
            sql.append("          OR promo_id IN (SELECT promo_id ");
            sql.append("                          FROM promotion ");
            sql.append("                          WHERE packet_id IN (437, 486, 487) OR packet_id IN (SELECT packet_id ");
            sql.append("                                                                              FROM show_packet ");
            sql.append("                                                                              WHERE show_id = ?)) ");
            sql.append("         ) ");
            sql.append("         AND order_header.ORDER_DATE >= (SELECT min(WEB_DISPLAY_START) web_display_start ");
            sql.append("                                         FROM promotion ");
            sql.append("                                         WHERE packet_id IN (SELECT packet_id ");
            sql.append("                                                             FROM show_packet ");
            sql.append("                                                             WHERE show_id = ?) OR packet_id IN (437, 486, 487))  ");
            sql.append("         AND line_status.description IN ('NEW', 'COMPLETE', 'RELEASED', 'FASCOR RELEASED') ");
            sql.append("   GROUP BY customer_id ");
            sql.append("  ) reg_orders USING (customer_id) ");
            sql.append("LEFT OUTER JOIN ");
            sql.append("  (SELECT customer_id, COUNT(DISTINCT show_cart_id) show_orders, COUNT(DISTINCT show_cart_dtl_id) total_lines, ");
            sql.append("     SUM(quantity * nvl(promo_base, 0)) + SUM(quantity * nvl(sab.promo_cost, 0)) total_dollars ");
            sql.append("   FROM show_cart_header ");
            sql.append("   JOIN show_cart_detail USING (Show_Cart_Id) ");
            sql.append("   LEFT JOIN ");
            sql.append("     (SELECT MIN(promo_base) promo_base, item_id ");
            sql.append("      FROM promo_item ");
            sql.append("      WHERE promo_item.promo_id IN ");
            sql.append("            (SELECT promo_id ");
            sql.append("             FROM promotion ");
            sql.append("             WHERE packet_id IN (437, 486, 487) OR packet_id IN ");
            sql.append("                                                   (SELECT packet_id ");
            sql.append("                                                    FROM show_packet ");
            sql.append("                                                    WHERE show_id = ?  ");
            sql.append("                                                   ) ");
            sql.append("            ) ");
            sql.append("      GROUP BY item_id ");
            sql.append("     ) promo_item ON show_cart_detail.item = promo_item.item_id ");
            sql.append("     LEFT JOIN ");
            sql.append("     (SELECT SUM(min_qty * discount_value) promo_cost, asst_promo ");
            sql.append("      FROM show_asst_barcode ");
            sql.append("      JOIN quantity_buy_item USING (qty_buy_id) ");
            sql.append("      GROUP BY asst_promo ");
            sql.append("     ) sab ON sab.asst_promo = show_cart_detail.item AND show_id = ?  ");
            sql.append("   WHERE show_cart_header.active = 1 AND show_cart_header.show_id = ?  ");
            sql.append("   GROUP BY customer_id ");
            sql.append("  )   show_cart USING (customer_id) ");
            sql.append("WHERE customer_id = ? ");
            m_CurrShowTotal = m_EdbConn.prepareStatement(sql.toString());
            
            // 2016 show shipped
            sql.setLength(0);
            sql.append("SELECT NVL(Reg_Orders.total_dollars, 0) total_dollars ");
            sql.append("FROM customer ");
            sql.append("LEFT OUTER JOIN ");
            sql.append("  (SELECT customer_id, COUNT(DISTINCT order_header.order_id) show_orders, COUNT(DISTINCT ol_id) total_lines, ");
            sql.append("     SUM(qty_ordered * sell_price)         total_dollars ");
            sql.append("   FROM order_header ");
            sql.append("   JOIN order_line ON order_line.order_id = order_header.order_id ");
            sql.append("   JOIN order_status line_status ON line_status.order_status_id = order_line.order_status_id ");
            sql.append("   WHERE ((packet_id IN (437, 486, 487) OR packet_id IN ");
            sql.append("                                           (SELECT packet_id ");
            sql.append("                                            FROM show_packet ");
            sql.append("                                            WHERE show_id = ?  ");
            sql.append("                                           )) ");
            sql.append("          OR promo_id IN (SELECT promo_id ");
            sql.append("                          FROM promotion ");
            sql.append("                          WHERE packet_id IN (437, 486, 487) OR packet_id IN (SELECT packet_id ");
            sql.append("                                                                              FROM show_packet ");
            sql.append("                                                                              WHERE show_id = ?)) ");
            sql.append("         ) ");
            sql.append("         AND order_header.ORDER_DATE >= (SELECT min(WEB_DISPLAY_START) web_display_start ");
            sql.append("                                         FROM promotion ");
            sql.append("                                         WHERE packet_id IN (SELECT packet_id ");
            sql.append("                                                             FROM show_packet ");
            sql.append("                                                             WHERE show_id = ?) OR packet_id IN (437, 486, 487))  ");
            sql.append("         AND line_status.description IN ('COMPLETE') ");
            sql.append("   GROUP BY customer_id ");
            sql.append("  ) reg_orders USING (customer_id) ");
            sql.append("WHERE customer_id = ? ");
            m_CurrShowShipped = m_EdbConn.prepareStatement(sql.toString());
            
            // 2016 show unshipped
            sql.setLength(0);
            sql.append("SELECT NVL(Reg_Orders.total_dollars, 0) total_dollars ");
            sql.append("FROM customer ");
            sql.append("LEFT OUTER JOIN ");
            sql.append("  (SELECT customer_id, COUNT(DISTINCT order_header.order_id) show_orders, COUNT(DISTINCT ol_id) total_lines, ");
            sql.append("     SUM(qty_ordered * sell_price)         total_dollars ");
            sql.append("   FROM order_header ");
            sql.append("   JOIN order_line ON order_line.order_id = order_header.order_id ");
            sql.append("   JOIN order_status line_status ON line_status.order_status_id = order_line.order_status_id ");
            sql.append("   WHERE ((packet_id IN (437, 486, 487) OR packet_id IN ");
            sql.append("                                           (SELECT packet_id ");
            sql.append("                                            FROM show_packet ");
            sql.append("                                            WHERE show_id = ?  ");
            sql.append("                                           )) ");
            sql.append("          OR promo_id IN (SELECT promo_id ");
            sql.append("                          FROM promotion ");
            sql.append("                          WHERE packet_id IN (437, 486, 487) OR packet_id IN (SELECT packet_id ");
            sql.append("                                                                              FROM show_packet ");
            sql.append("                                                                              WHERE show_id = ?)) ");
            sql.append("         ) ");
            sql.append("         AND order_header.ORDER_DATE >= (SELECT min(WEB_DISPLAY_START) web_display_start ");
            sql.append("                                         FROM promotion ");
            sql.append("                                         WHERE packet_id IN (SELECT packet_id ");
            sql.append("                                                             FROM show_packet ");
            sql.append("                                                             WHERE show_id = ?) OR packet_id IN (437, 486, 487))  ");
            sql.append("         AND line_status.description IN ('NEW', 'FASCOR RELEASED', 'RELEASED') ");
            sql.append("   GROUP BY customer_id ");
            sql.append("  ) reg_orders USING (customer_id) ");
            sql.append("WHERE customer_id = ? ");
            m_CurrShowUnshipped = m_EdbConn.prepareStatement(sql.toString());
            
            // get show cart
            sql.setLength(0);
            sql.append("SELECT SUM(quantity * nvl(promo_base, 0)) + SUM(quantity * nvl(sab.promo_cost, 0)) total_dollars ");
            sql.append("FROM show_cart_header ");
            sql.append("JOIN show_cart_detail USING (Show_Cart_Id) ");
            sql.append("LEFT JOIN ");
            sql.append("     (SELECT MIN(promo_base) promo_base, item_id ");
            sql.append("      FROM promo_item ");
            sql.append("      WHERE promo_item.promo_id IN ");
            sql.append("            (SELECT promo_id ");
            sql.append("             FROM promotion ");
            sql.append("             WHERE packet_id IN (437, 486, 487) OR packet_id IN ");
            sql.append("                                                   (SELECT packet_id ");
            sql.append("                                                    FROM show_packet ");
            sql.append("                                                    WHERE show_id = ?  ");
            sql.append("                                                   )  ");
            sql.append("            ) ");
            sql.append("      GROUP BY item_id ");
            sql.append(") promo_item ON show_cart_detail.item = promo_item.item_id ");
            sql.append("LEFT JOIN ");
            sql.append("     (SELECT SUM(min_qty * discount_value) promo_cost, asst_promo ");
            sql.append("      FROM show_asst_barcode ");
            sql.append("      JOIN quantity_buy_item USING (qty_buy_id) ");
            sql.append("      GROUP BY asst_promo ");
            sql.append(") sab ON sab.asst_promo = show_cart_detail.item AND show_id = ?  ");
            sql.append("WHERE show_cart_header.active = 1 ");
            sql.append("      and show_cart_header.show_id = ?  ");
            sql.append("      and customer_id = ? ");
            sql.append("GROUP BY customer_id ");
            m_CurrShowCart = m_EdbConn.prepareStatement(sql.toString());

            // get tm names
            sql.setLength(0);
            sql.append("SELECT DISTINCT er.first, er.last, division ");
            sql.append("FROM cust_rep_div_view ");
            sql.append("JOIN web_user USING (er_id) ");
            sql.append("JOIN (SELECT active, er_id, first, last ");
            sql.append("        FROM emery_rep ");
            sql.append("       ) er USING (er_id) ");
            sql.append("JOIN (SELECT customer_id, cust_status_id ");
            sql.append("        FROM customer ");
            sql.append("        JOIN cust_market USING (customer_id) ");
            sql.append("        WHERE mkt_class_id <> 26 ");
            sql.append("        AND market_id = 4 ");
            sql.append("       ) customer USING (customer_id) ");
            sql.append("WHERE rep_type = 'SALES REP' ");
            sql.append("      AND division_id IN (SELECT division ");
            sql.append("                          FROM sa.sales_directors ");
            sql.append("                          WHERE name = ? ");
            sql.append("                         ) ");
            sql.append("      AND er.active = 1 ");
            sql.append("      AND customer.CUST_STATUS_ID <> 3 ");
            sql.append("ORDER BY er.first, er.last ");
            m_GetTMNames = m_PgConn.prepareStatement(sql.toString());
            
            // get director names
            sql.setLength(0);
            sql.append("select name from sa.sales_directors where test_user = 0 order by name ");
            m_GetDirectorNames = m_PgConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[CustShowOrderSummary]", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("[CustShowOrderSummary] prepareStatements - null oracle connection");
      
      return isPrepared;
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {      
      int pcount = params.size();
      Param param = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                  
         if ( param.name.equals("fname") )
            m_FName = param.value;
         
         if ( param.name.equals("lname") )
            m_LName = param.value;
         
         // TODO are we getting a show id or name here?
         
      }
      
      m_FileNames.add(String.format("%s_%s_showorders.xlsx", m_FName, m_LName));
   }
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {
      int col = 0;
      int m_CharWidth = 295;
      
      defineStyles();
      
      //
      // creates Excel title
      addRow(0);
      addCell(col, "Current Customer Show Order Summary", m_StyleHdrLeft);
      
      // add the run time
      Calendar cal = Calendar.getInstance();
      SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MMM-dd HH:mm");
      addCell(col+2, ("Report run at: " + sdf.format(cal.getTime())), m_StyleTxtLB);
      
      //
      // Add the captions for the summary line
      addRow(2);
      m_Sheet.setColumnWidth(col, (8 * m_CharWidth));
      addCell(col, "Summary", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (50 * m_CharWidth));
      addCell(col, "TM / Division", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (22 * m_CharWidth));
      addCell(col, "", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "2015 Total", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "2016 Total", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "2016 Show Cart", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (21 * m_CharWidth));
      addCell(col, "2016 Unshipped Orders", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (19 * m_CharWidth));
      addCell(col, "2016 Shipped Orders", m_StyleTxtC);
      
      //
      // Add the captions for the detail lines
      addRow(9);
      col = 0;
      m_Sheet.setColumnWidth(col, (8 * m_CharWidth));
      addCell(col, "Cust #", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (50 * m_CharWidth));
      addCell(col, "Customer Name", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (22 * m_CharWidth));
      addCell(col, "TM Name", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "2015 Total", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "2016 Total", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "2016 Show Cart", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (21 * m_CharWidth));
      addCell(col, "2016 Unshipped Orders", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (19 * m_CharWidth));
      addCell(col, "2016 Shipped Orders", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (14 * m_CharWidth));
      addCell(col, "Division", m_StyleTxtC);
   }
}
