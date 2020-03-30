/**
 * Title:			ShowSpiffs
 * Description:	    Reports for Show by Vendor or Customer for SPIFFs,
 *                  Extra Entries, etc.
 * Company:			Emery-Waterhouse
 * @author			prichter
 * @version			1.0
 * <p>
 * Create Date: Jan 4, 2007
 * Last Update: $Id: ShowSpiffs.java,v 1.10 2014/03/11 18:35:59 jfisher Exp $
 * <p>
 * History:
 *   $Log: ShowSpiffs.java,v $
 *   Revision 1.10  2014/03/11 18:35:59  jfisher
 *   Changes for TM.
 *
 *   Revision 1.9  2013/03/14 18:06:10  prichter
 *   Fine tunning 2013
 *
 *   Revision 1.8  2013/03/14 17:34:52  prichter
 *   2013 crap
 *
 *   Revision 1.7  2012/08/07 14:47:56  npasnur
 *   Added timestamp to the report file name to allow multiple people to run it.
 *
 *   Revision 1.6  2012/08/01 18:19:55  npasnur
 *   Fixed an issue where the spaces in the file name was causing the report to fail.
 *
 *   Revision 1.5  2011/03/22 12:25:20  prichter
 *   Fixed a bug that caused orders to not be recognized under certain circumstances.
 *
 *   Revision 1.4  2011/03/14 12:07:44  prichter
 *   Added orders already invoiced to the sales for each customer/item
 *
 *   Revision 1.3  2009/02/18 17:17:50  jfisher
 *   Fixed depricated methods after poi upgrade
 *
 *   Revision 1.2  2007/01/17 16:30:30  jfisher
 *   Changed the class name in one exception handler so it was this class name.
 *
 *   Revision 1.1  2007/01/09 18:27:14  prichter
 *   Initial add.  Report of dealer market SPIFFs and Extra Entry deals.
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.web.WebReport;
import com.emerywaterhouse.websvc.Param;


public class ShowSpiffs extends Report
{
   private static final short MAX_COLS = 5;

   // parameter fields
   private int m_ShowId;        // The id of the show being reported
   private int m_RuleId;        // The id of the special detail being reported
   private String m_ReportBy;   // by "Customer" or by "Vendor"
   private String m_VendorId;   // a single vendor id, or "All"
   private String m_CustId;     // a single customer id or "All"
   private int m_TmId;          // a specific TM to run a report for.
   private String m_TmName;    // The TM's name if that was selected.

   private PreparedStatement m_CustItems;    // Main report query
   private PreparedStatement m_CustInvoices; // Orders already invoiced by customer
   private PreparedStatement m_CustOrders;   // Open or picking orders
   private PreparedStatement m_TmData;

   //
   // workbook objects
   private HSSFWorkbook m_Wrkbk;
   private HSSFSheet m_Sheet;

   //
   // Fonts
   private HSSFFont m_FontHeader1;
   private HSSFFont m_FontHeader2;
   private HSSFFont m_FontColumnTitle;
   private HSSFFont m_FontDetailText;
   private HSSFFont m_FontDetailQty;
   private HSSFFont m_FontDetailAmt;

   //
   // cell styles
   private HSSFCellStyle m_StyleHeader1Left;
   private HSSFCellStyle m_StyleHeader1Right;
   private HSSFCellStyle m_StyleHeader2Left;
   private HSSFCellStyle m_StyleHeader2Right;
   private HSSFCellStyle m_StyleColumnTitle;
   private HSSFCellStyle m_StyleDetailText;
   private HSSFCellStyle m_StyleDetailQty;
   private HSSFCellStyle m_StyleDetailAmt;

   //
   // web_report crappola
   private WebReport m_WebRpt;
   private String m_OutFormat = null;
   private int m_WebRptId;
   private int m_Cnt = 0;

   public ShowSpiffs()
   {
      super();

      m_TmId = 0;
      m_ReportBy = "";
      m_VendorId = "";
      m_CustId = "";
   }

   /**
    * Executes the queries and builds the customer report
    *
    * @return true if the report was successfully built
    * @throws Exception
    */
   private boolean buildCustomerReport() throws Exception
   {
      HSSFRow row = null;
      HSSFCell cell = null;
      int rowNum = 0;
      int colNum = 0;
      FileOutputStream outFile = null;
      ResultSet data = null;
      ResultSet salesRs = null;
      boolean result = false;
      int lastVendor = -1;
      String lastCust = "-1";
      StringBuffer fileName = new StringBuffer();
      int qty;
      double amt;
      int tot_qty = 0;
      double tot_amt = 0;
      String tmp = null;
      double last_vnd_min = 0;

      //
      // Get the timestamp for the report file name to allow multiple people to run it.
      tmp = Long.toString(System.currentTimeMillis());

      fileName.append(getShowName());
      fileName.append("-");
      fileName.append(getRuleName());

      if ( m_TmId > 0 )
         fileName.append("-tm-");
      else
         fileName.append("-cust-");

      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xls");
      m_FileNames.add(fileName.toString());

      m_Wrkbk = new HSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      m_Sheet.getPrintSetup().setLandscape(true);
      m_Sheet.getHeader().setCenter(getShowName() + " - " + getRuleName());

      init(m_Wrkbk);

      outFile = new FileOutputStream(m_FilePath + fileName.toString(), false);

      try {
         m_CurAction = "Building customer report";

         m_CustItems.setInt(1, m_ShowId);
         m_CustItems.setInt(2, m_RuleId);
         m_CustItems.setInt(3, m_RuleId);
         m_CustItems.setInt(4, m_ShowId);
         m_CustItems.setInt(5, m_RuleId);
         m_CustItems.setInt(6, m_RuleId);
         m_CustItems.setInt(7, m_ShowId);

         if ( m_TmId > 0 ) {
            m_CustItems.setInt(8, m_TmId);
            m_CustItems.setInt(9, m_ShowId);
            m_CustItems.setInt(10, m_RuleId);
         }
         else {
            m_CustItems.setInt(8, m_ShowId);
            m_CustItems.setInt(9, m_RuleId);
         }

         data = m_CustItems.executeQuery();

         while ( data.next() ) {
            // If this is a new customer, print the customer header
            if ( !data.getString("customer_id").equals(lastCust)   ) {

               // If this is not the first customer, insert a page break
               // if the customer# changes
               if ( !lastCust.equals("-1") ) {
                  row = createRow(rowNum++, MAX_COLS);

                  if ( last_vnd_min != 0 && (tot_amt != 0 || tot_qty != 0 )) {
                     cell = row.getCell(3);
                     cell.setCellStyle(m_StyleDetailAmt);
                     cell.setCellValue(tot_amt);
                     cell = row.getCell(4);
                     cell.setCellStyle(m_StyleDetailAmt);
                     cell.setCellValue(last_vnd_min);
                     tot_amt = 0;
                     tot_qty = 0;
                  }
                  else {
                     if ( data.getInt("min_qty") != 0 ) {
                        cell = row.getCell(4);
                        cell.setCellStyle(m_StyleDetailQty);
                        cell.setCellValue(tot_qty);
                        tot_qty = 0;
                     }
                     else {
                        cell = row.getCell(4);
                        cell.setCellStyle(m_StyleDetailAmt);
                        cell.setCellValue(tot_amt);
                        tot_amt = 0;
                     }
                  }

                  row = createRow(rowNum++, MAX_COLS);
                  m_Sheet.setRowBreak(rowNum++);
                  row = createRow(rowNum++, MAX_COLS);
               }

               tot_amt = 0;
               tot_qty = 0;
               rowNum = createCustomerHeader(data, rowNum);
               lastVendor = -1;
            }

            // If this is a new vendor within the customer, print the vendor header
            if ( data.getInt("vendor_id") != lastVendor ) {
               if ( lastVendor != -1 ) {
                  row = createRow(rowNum++, MAX_COLS);

                  if ( last_vnd_min != 0 ) {
                     cell = row.getCell(3);
                     cell.setCellStyle(m_StyleDetailAmt);
                     cell.setCellValue(tot_amt);
                     cell = row.getCell(4);
                     cell.setCellStyle(m_StyleDetailAmt);
                     cell.setCellValue(last_vnd_min);
                     tot_amt = 0;
                     tot_qty = 0;
                  }

                  else {
                     if ( data.getInt("min_qty") != 0 ) {
                        cell = row.getCell(4);
                        cell.setCellStyle(m_StyleDetailQty);
                        cell.setCellValue(tot_qty);
                        tot_qty = 0;
                     }

                     else {
                        cell = row.getCell(4);
                        cell.setCellStyle(m_StyleDetailAmt);
                        cell.setCellValue(tot_amt);
                        tot_amt = 0;
                     }
                  }
               }
               
               rowNum = createCustVendorHeader(data, rowNum);
               tot_amt = 0;
               tot_qty = 0;
            }

            lastCust = data.getString("customer_id");
            lastVendor = data.getInt("vendor_id");
            last_vnd_min = data.getDouble("vnd_min");

            row = createRow(rowNum++, MAX_COLS);

            for ( int i = 0; i < MAX_COLS; i++ ) {
               cell = row.getCell( i );

               // Set the styles for individual columns
               switch ( i ) {
               case 0:
                  cell.setCellStyle(m_StyleDetailText);
                  break;
               case 1:
                  cell.setCellStyle(m_StyleDetailText);
                  break;
               case 2:
                  cell.setCellStyle(m_StyleDetailText);
                  break;
               case 3:
                  cell.setCellStyle(m_StyleDetailQty);
                  break;
               case 4:
                  cell.setCellStyle(m_StyleDetailQty);
                  break;
               }
            }

            colNum = 0;

            row.getCell(colNum++).setCellValue(new HSSFRichTextString(data.getString("vendor_item_num")));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(data.getString("alt_item_desc")));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(data.getString("item_id")));

            qty = 0;
            amt = 0;

            m_CustInvoices.setInt(1, data.getInt("item_ea_id"));
            m_CustInvoices.setString(2, data.getString("customer_id"));
            m_CustInvoices.setInt(3, m_ShowId);
            salesRs = m_CustInvoices.executeQuery();

            while ( salesRs.next() ) {
               qty += salesRs.getInt("qty_ordered");
               amt += salesRs.getDouble("amt_ordered");
            }

            DbUtils.closeDbConn(null, null, salesRs);

            m_CustOrders.setInt(1, data.getInt("item_ea_id"));
            m_CustOrders.setString(2, data.getString("customer_id"));
            m_CustOrders.setInt(3, m_ShowId);
            salesRs = m_CustOrders.executeQuery();

            while ( salesRs.next() ) {
               qty += salesRs.getInt("qty_ordered");
               amt += salesRs.getDouble("amt_ordered");
            }

            DbUtils.closeDbConn(null, null, salesRs);

            if ( data.getInt("min_qty") != 0 ) {
               row.getCell(colNum++).setCellValue(qty);
               row.getCell(colNum++).setCellValue(data.getInt("min_qty"));
               tot_qty = tot_qty + qty;
            }
            else {
               cell = row.getCell(colNum++);
               cell.setCellStyle(m_StyleDetailAmt);
               cell.setCellValue(amt);

               if ( data.getDouble("min_amt") != 0 ) {
                  cell = row.getCell(colNum++);
                  cell.setCellStyle(m_StyleDetailAmt);
                  cell.setCellValue(data.getFloat("min_amt"));
               }

               tot_amt = tot_amt + amt;
            }
            
            m_Cnt++;
         }

         m_Wrkbk.write(outFile);
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {
         m_FontHeader1 = null;
         m_FontHeader2 = null;
         m_FontColumnTitle = null;
         m_FontDetailText = null;
         m_FontDetailQty = null;
         m_FontDetailAmt = null;
         m_StyleHeader1Left = null;
         m_StyleHeader1Right = null;
         m_StyleHeader2Left = null;
         m_StyleHeader2Right = null;
         m_StyleColumnTitle = null;
         m_StyleDetailText = null;
         m_StyleDetailQty = null;
         m_StyleDetailAmt = null;
         m_Sheet = null;
         m_Wrkbk = null;

         closeRSet(data);
         data = null;

         row = null;

         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error(e);
         }

         outFile = null;
         m_WebRpt.setLineCount(m_Cnt);
      }

      return result;
   }

   /**
    * Builds the email message that will be sent to the customer.  This overrides the default message
    * built by the RptProcessor.
    *
    * @param rptName
    * @return EMail message String
    */
   private String buildEmailText(String rptName)
   {
      StringBuffer msg = new StringBuffer();

      msg.append("The following report are ready for you to pick up:\r\n");
      msg.append("\tShow SPIFFs\r\n\r\n");
      msg.append("To view your reports:\r\n");
      msg.append("\thttp://www.emeryonline.com/emerywh/subscriber/my_account/report_list.jsp\r\n\r\n");
      msg.append("If you have any questions or suggestions, please contact help@emeryonline.com\r\n");
      msg.append("or call 800-283-0236 ext. 1.");

      return msg.toString();
   }

   /**
    * Calls the appropriate report based on the parameters
    *
    * @return true if the report was successfully built
    * @throws Exception
    */
   private boolean buildOutputFile() throws Exception
   {
      boolean result = false;
      
      //
      // Report by customer
      if ( m_ReportBy.equalsIgnoreCase("customer") || m_ReportBy.equalsIgnoreCase("both") ) {
         if ( prepareStatements("customer") ) {
            try {
               result = buildCustomerReport();
            }

            finally {
               closeStatements();
            }
         }
      }

      // Report by vendor
      if ( m_ReportBy.equalsIgnoreCase("vendor") || m_ReportBy.equalsIgnoreCase("both") ) {
         if ( prepareStatements("vendor") ) {
            try {
               result = buildVendorReport();
            }

            finally {
               closeStatements();
            }
         }
      }

      //
      // TM Customer
      if ( m_TmId > 0 ) {
         if ( prepareStatements("tm") ) {
            try {
               getTmData();
               result = buildCustomerReport();
            }

            finally {
               closeStatements();
            }
         }
      }

      return result;
   }

   /**
    * Executes the queries and builds the vendor report
    *
    * @return true if the report was successfully built
    * @throws Exception
    */
   private boolean buildVendorReport() throws Exception
   {
      HSSFRow row = null;
      HSSFCell cell = null;
      int rowNum = 0;
      int colNum = 0;
      FileOutputStream outFile = null;
      ResultSet data = null;
      ResultSet salesRs = null;
      boolean result = false;
      int lastVendor = -1;
      String lastCust = "-1";
      StringBuffer fileName = new StringBuffer();
      int qty;
      double amt;
      int tot_qty = 0;
      double tot_amt = 0;
      double last_vnd_min = 0;

      fileName.append(getShowName());
      fileName.append("-");
      fileName.append(getRuleName());
      fileName.append("-vendor");
      fileName.append(".xls");
      m_FileNames.add(fileName.toString());

      m_Wrkbk = new HSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      m_Sheet.getPrintSetup().setLandscape(true);
      m_Sheet.getHeader().setCenter(getShowName() + " - " + getRuleName());

      init(m_Wrkbk);

      outFile = new FileOutputStream(m_FilePath + fileName.toString(), false);

      try {
         m_CustItems.setInt(1, m_ShowId);
         m_CustItems.setInt(2, m_RuleId);
         m_CustItems.setInt(3, m_RuleId);
         m_CustItems.setInt(4, m_ShowId);
         m_CustItems.setInt(5, m_RuleId);
         m_CustItems.setInt(6, m_RuleId);
         m_CustItems.setInt(7, m_ShowId);
         m_CustItems.setInt(8, m_ShowId);
         m_CustItems.setInt(9, m_RuleId);
         data = m_CustItems.executeQuery();

         m_CurAction = "Building vendor report";

         while ( data.next() ) {
            // If the first record for this vendor, create the vendor header
            if ( data.getInt("vendor_id") != lastVendor  ) {
               // If this is not the first vendor, insert a page break
               // if the vendor# changes
               if ( lastVendor != -1 ) {
                  row = createRow(rowNum++, MAX_COLS);

                  if ( last_vnd_min != 0  && (tot_amt != 0 || tot_qty != 0  )) {
                     cell = row.getCell(3);
                     cell.setCellStyle(m_StyleDetailAmt);
                     cell.setCellValue(tot_amt);
                     cell = row.getCell(4);
                     cell.setCellStyle(m_StyleDetailAmt);
                     cell.setCellValue(last_vnd_min);
                     tot_amt = 0;
                     tot_qty = 0;
                  }

                  else {
                     if ( data.getInt("min_qty") != 0 ) {
                        cell = row.getCell(4);
                        cell.setCellStyle(m_StyleDetailQty);
                        cell.setCellValue(tot_qty);
                     }

                     else {
                        cell = row.getCell(4);
                        cell.setCellStyle(m_StyleDetailAmt);
                        cell.setCellValue(tot_amt);
                     }
                  }


                  row = createRow(rowNum++, MAX_COLS);
                  m_Sheet.setRowBreak(rowNum++);
                  row = createRow(rowNum++, MAX_COLS);
               }

               rowNum = createVendorHeader(data, rowNum);
               tot_amt = 0;
               tot_qty = 0;
               lastVendor = -1;
            }

            // If the first record for this customer/vendor, create the customer subheader
            if ( !data.getString("customer_id").equalsIgnoreCase(lastCust) ) {
               if ( !(lastCust.equals("-1") )) {
                  row = createRow(rowNum++, MAX_COLS);

                  if ( last_vnd_min != 0 ) {
                     cell = row.getCell(3);
                     cell.setCellStyle(m_StyleDetailAmt);
                     cell.setCellValue(tot_amt);
                     cell = row.getCell(4);
                     cell.setCellStyle(m_StyleDetailAmt);
                     cell.setCellValue(last_vnd_min);
                     tot_amt = 0;
                     tot_qty = 0;
                  }
                  else {
                     if ( data.getInt("min_qty") != 0 ) {
                        cell = row.getCell(4);
                        cell.setCellStyle(m_StyleDetailQty);
                        cell.setCellValue(tot_qty);
                     }

                     else {
                        cell = row.getCell(4);
                        cell.setCellStyle(m_StyleDetailAmt);
                        cell.setCellValue(tot_amt);
                     }
                  }
               }
               
               tot_amt = 0;
               tot_qty = 0;
               rowNum = createVndCustHeader(data, rowNum);
            }

            lastCust = data.getString("customer_id");
            lastVendor = data.getInt("vendor_id");
            last_vnd_min = data.getDouble("vnd_min");

            row = createRow(rowNum++, MAX_COLS);
            colNum = 0;

            row.getCell(colNum++).setCellValue(new HSSFRichTextString(data.getString("vendor_item_num")));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(data.getString("item_id")));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(data.getString("alt_item_desc")));

            qty = 0;
            amt = 0;

            m_CustInvoices.setInt(1, data.getInt("item_ea_id"));
            m_CustInvoices.setString(2, data.getString("customer_id"));
            m_CustInvoices.setInt(3, m_ShowId);
            salesRs = m_CustInvoices.executeQuery();

            while ( salesRs.next() ) {
               qty += salesRs.getInt("qty_ordered");
               amt += salesRs.getDouble("amt_ordered");
            }

            DbUtils.closeDbConn(null, null, salesRs);

            m_CustOrders.setInt(1, data.getInt("item_ea_id"));
            m_CustOrders.setString(2, data.getString("customer_id"));
            m_CustOrders.setInt(3, m_ShowId);
            salesRs = m_CustOrders.executeQuery();

            while ( salesRs.next() ) {
               qty += salesRs.getInt("qty_ordered");
               amt += salesRs.getDouble("amt_ordered");
            }

            DbUtils.closeDbConn(null, null, salesRs);

            if ( data.getInt("min_qty") != 0 ) {
               cell = row.getCell(colNum++);
               cell.setCellStyle(m_StyleDetailQty);
               cell.setCellValue(qty);

               if ( data.getInt("min_qty") != 0 ) {
                  cell = row.getCell(colNum++);
                  cell.setCellStyle(m_StyleDetailQty);
                  cell.setCellValue(data.getInt("min_qty"));
               }

               tot_qty = tot_qty + qty;
            }
            else {
               cell = row.getCell(colNum++);
               cell.setCellStyle(m_StyleDetailAmt);
               cell.setCellValue(amt);

               if ( data.getDouble("min_amt") != 0 ) {
                  cell = row.getCell(colNum++);
                  cell.setCellStyle(m_StyleDetailAmt);
                  cell.setCellValue(data.getDouble("min_amt"));
               }

               tot_amt = tot_amt + amt;
            }
            
            m_Cnt++;
         }

         m_Wrkbk.write(outFile);
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {
         m_FontHeader1 = null;
         m_FontHeader2 = null;
         m_FontColumnTitle = null;
         m_FontDetailText = null;
         m_FontDetailQty = null;
         m_FontDetailAmt = null;
         m_StyleHeader1Left = null;
         m_StyleHeader1Right = null;
         m_StyleHeader2Left = null;
         m_StyleHeader2Right = null;
         m_StyleColumnTitle = null;
         m_StyleDetailText = null;
         m_StyleDetailQty = null;
         m_StyleDetailAmt = null;
         m_Sheet = null;
         m_Wrkbk = null;

         closeRSet(data);
         data = null;

         row = null;

         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error(e);
         }

         outFile = null;
         m_WebRpt.setLineCount(m_Cnt);
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      DbUtils.closeDbConn(null, m_CustItems, null);
      DbUtils.closeDbConn(null, m_CustInvoices, null);
      DbUtils.closeDbConn(null, m_CustOrders, null);
      DbUtils.closeDbConn(null, m_TmData, null);
   }

   /**
    * Creates the headings for the customer report
    *
    * @param data - ResultSet containing the report data.  Customer_id and
    * name will be included
    * @param rw - the next row of the spreadsheet
    * @return short next available row number after created the heading
    * @throws SQLException
    */
   private int createCustomerHeader(ResultSet data, int rw) throws SQLException
   {
      HSSFRow row = null;
      HSSFCell cell = null;
      int col = 0;

      //
      // Create the TM name header first if there is one
      if ( m_TmId > 0 ) {
         row = createRow(rw++, MAX_COLS);

         cell = row.getCell( 0);
         cell.setCellType(HSSFCell.CELL_TYPE_STRING);
         cell.setCellStyle(m_StyleHeader1Left);
         cell.setCellValue(new HSSFRichTextString("Emery Rep: " + m_TmName));
      }

      //
      row = createRow(rw++, MAX_COLS);

      cell = row.getCell( 0);
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(m_StyleHeader1Left);
      cell.setCellValue(new HSSFRichTextString(data.getString("custname")));

      cell = row.getCell( (MAX_COLS - 1) );
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(m_StyleHeader1Right);
      cell.setCellValue(new HSSFRichTextString("Cust# " + data.getString("customer_id")));
      //
      // Build the column headings
      row = createRow(rw++, MAX_COLS);
      row = createRow(rw++, MAX_COLS);

      if ( row != null ) {
         for ( int i = 0; i < MAX_COLS; i++ ) {
            cell = row.createCell(i);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellStyle(m_StyleColumnTitle);
         }

         col = 0;

         m_Sheet.setColumnWidth(col, 3000);
         row.getCell(col++).setCellValue(new HSSFRichTextString("Vendor SKU"));

         m_Sheet.setColumnWidth(col, 20000);
         row.getCell(col++).setCellValue(new HSSFRichTextString("Product Name"));

         m_Sheet.setColumnWidth(col, 2300);
         row.getCell(col++).setCellValue(new HSSFRichTextString("Item #"));

         m_Sheet.setColumnWidth(col, 2200);
         row.getCell(col++).setCellValue(new HSSFRichTextString("Ordered"));

         m_Sheet.setColumnWidth(col, 2200);
         row.getCell(col++).setCellValue(new HSSFRichTextString("Minimum"));
      }

      return rw;
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;
      String fileName = null;

      try {
         m_EdbConn = m_RptProc.getEdbConn();
         created = buildOutputFile();

         //
         // Only send this if this was requested from the web.
         if ( created && m_WebRptId > 0 ) {
            fileName = m_FileNames.get(0);
            m_WebRpt.setFileName(fileName);
            
            if ( m_RptProc.getZipped() ) {
               fileName = fileName.substring(0, fileName.indexOf('.')+1) + "zip";
               m_WebRpt.setZipped(true);
            }
            else
               m_WebRpt.setZipped(false);

            m_WebRpt.setFileName(fileName);
            m_WebRpt.setStatus("COMPLETE");

            //
            // Save the web_report entry to the database
            m_WebRpt.update();
            m_EdbConn.commit();

            m_RptProc.setEmailMsg(buildEmailText(fileName));
         }
      }

      catch ( Exception ex ) {
         log.fatal("[ShowSpiffs]", ex);
      }

      finally {
         closeStatements();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    * @param colCnt The number of columns in the row.
    *
    * @return The formatted row of the spreadsheet.
    */
   private HSSFRow createRow(int rowNum, int colCnt)
   {
      HSSFRow row = null;

      if ( m_Sheet == null )
         return row;

      row = m_Sheet.createRow(rowNum);

      //
      // set the type and style of the cell.
      if ( row != null ) {
         for ( int i = 0; i < colCnt; i++ ) {
            row.createCell(i);
         }
      }

      return row;
   }

   /**
    * Creates the vendor subheader for a customer report
    *
    * @param data - ResultSet containing the report data.  The current
    * record will contain the customer_id and name
    * @param rw - The current row of the spreadsheet
    * @return short - the next available row
    * @throws SQLException
    */
   private int createCustVendorHeader(ResultSet data, int rw) throws SQLException
   {
      HSSFRow row = null;
      HSSFCell cell = null;

      //
      row = createRow(rw++, MAX_COLS);
      row = createRow(rw++, MAX_COLS);

      cell = row.getCell( 0);
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(m_StyleHeader2Left);
      cell.setCellValue(new HSSFRichTextString(data.getString("vendorname")));

      //Set the style on the other cells so they'll all be gray
      for ( int i = 0; i < MAX_COLS; i++ ) {
         cell = row.getCell( i );
         cell.setCellStyle(m_StyleHeader2Left);
      }

      cell = row.getCell( (MAX_COLS - 1) );
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(m_StyleHeader2Right);
      cell.setCellValue(new HSSFRichTextString("Booth# " + data.getString("booth")));

      return ++rw;
   }

   /**
    * Creates the report and column headings
    *
    * @param data - ResultSet containing the report data.  The vendor_id and name
    * will be included
    * @param rw - the next available row of the spreadsheet
    * @return short next available row after printing the heading
    * @throws SQLException
    */
   private int createVendorHeader(ResultSet data, int rw) throws SQLException
   {
      HSSFRow row = null;
      HSSFCell cell = null;
      int col = 0;

      //
      row = createRow(rw++, MAX_COLS);

      cell = row.getCell( 0);
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(m_StyleHeader1Left);
      cell.setCellValue(new HSSFRichTextString(data.getString("vendorname")));

      cell = row.getCell( (MAX_COLS - 1) );
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(m_StyleHeader1Right);
      cell.setCellValue(new HSSFRichTextString("Booth# " + data.getString("booth")));
      //
      // Build the column headings
      row = createRow(rw++, MAX_COLS);
      row = createRow(rw++, MAX_COLS);

      if ( row != null ) {
         for ( int i = 0; i < MAX_COLS; i++ ) {
            cell = row.createCell(i);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellStyle(m_StyleColumnTitle);
         }

         col = 0;

         m_Sheet.setColumnWidth(col, 3000);
         row.getCell(col++).setCellValue(new HSSFRichTextString("Vendor SKU"));

         m_Sheet.setColumnWidth(col, 2300);
         row.getCell(col++).setCellValue(new HSSFRichTextString("Item #"));

         m_Sheet.setColumnWidth(col, 20000);
         row.getCell(col++).setCellValue(new HSSFRichTextString("Product Name"));

         m_Sheet.setColumnWidth(col, 2200);
         row.getCell(col++).setCellValue(new HSSFRichTextString("Ordered"));

         m_Sheet.setColumnWidth(col, 2200);
         row.getCell(col++).setCellValue(new HSSFRichTextString("Minimum"));
      }
      return rw;
   }

   /**
    * Creates the customer subheader for a vendor report
    *
    * @param data - A result set containing the report data
    * @param rw - the current row
    * @return short - the next available row number
    * @throws SQLException
    */
   private int createVndCustHeader(ResultSet data, int rw) throws SQLException
   {
      HSSFRow row = null;
      HSSFCell cell = null;

      //
      row = createRow(rw++, MAX_COLS);
      row = createRow(rw++, MAX_COLS);

      cell = row.getCell( 0);
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(m_StyleHeader2Left);
      cell.setCellValue(new HSSFRichTextString(data.getString("custname")));

      //Set the style on the other cells so they'll all be gray
      for ( int i = 0; i < MAX_COLS; i++ ) {
         cell = row.getCell( i );
         cell.setCellStyle(m_StyleHeader2Left);
      }

      cell = row.getCell( (MAX_COLS - 1) );
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(m_StyleHeader2Right);
      cell.setCellValue(new HSSFRichTextString("Cust# " + data.getString("customer_id")));

      return ++rw;
   }

   /**
    * Returns the name of the show rule (SPIFF, or EXTRA ENTRY e.g.)
    *
    * @return String - the show rule name
    * @throws Exception
    */
   private String getRuleName() throws Exception
   {
      Statement stmt = null;
      ResultSet rs = null;

      try {
         stmt = m_EdbConn.createStatement();

         rs = stmt.executeQuery("select name from show_rule where rule_id = " + m_RuleId);
         if ( rs.next() )
            return rs.getString("name") != null ? rs.getString("name").replaceAll(" ", "_") : "";
      }

      finally {
         closeRSet(rs);
         closeStmt(stmt);
         rs = null;
         stmt = null;
      }

      return "rule";
   }

   /**
    *
    * @throws Exception
    */
   private void getTmData() throws Exception
   {
      ResultSet rs = null;

      try {
         m_TmData.setInt(1, m_TmId);
         rs = m_TmData.executeQuery();

         if ( rs.next() )
            m_TmName = rs.getString("name");
      }

      finally {
         closeRSet(rs);
         rs = null;
      }
   }

   /**
    * Returns the name of the show
    *
    * @return String - the show name
    * @throws Exception
    */
   private String getShowName() throws Exception
   {
      Statement stmt = null;
      ResultSet rs = null;

      try {
         stmt = m_EdbConn.createStatement();

         rs = stmt.executeQuery("select name from show where show_id = " + m_ShowId);
         if ( rs.next() )
            return rs.getString("name") != null ? rs.getString("name").trim().replaceAll(" ", "_") : "";
      }

      finally {
         closeRSet(rs);
         closeStmt(stmt);
         rs = null;
         stmt = null;
      }

      return "show";
   }

   /**
    * Retrieve the web_report record from the database.  The web_report_id is
    * stored in parameter 15.  The record should have been created by the
    * PurchHistRptSrv servlet, but if not found, will be created here.
    *
    * @return boolean
    */
   private boolean getWebReport()
   {    
      boolean result = true;

      //
      // set the connection in the WebReport bean
      try {
         m_WebRpt = new WebReport();
         m_WebRpt.setConnection(m_RptProc.getEdbConn());
         m_WebRpt.setReportName("Show SPIFF");
      }

      catch ( Exception ex ) {
         log.error("[ShowSpiffs] Unable to set connection in web report ", ex);
         result = false;
      }

      //
      // Load the web_report id from the 15th parameter.  If an id was passed,
      // load that web_report record into the WebReport bean.  Otherwise, create
      // a new web_report record.
      try {
         if ( m_WebRptId >= 0 )
            m_WebRpt.load(m_WebRptId);
         else
            m_WebRpt.insert();
      }

      catch ( Exception ex ) {
         log.error("[ShowSpiffs] unable to create a web_report record ", ex);
         result = false;
      }

      return result;
   }

   /**
    * Initializes the fonts and cell styles for the given workbook
    *
    * @param wkbk - the HSSFWorkbook object
    */
   private void init(HSSFWorkbook wkbk)
   {
      //
      // Create a fonts and styles for top level header
      m_FontHeader1 = wkbk.createFont();
      m_FontHeader1.setFontHeightInPoints((short)12);
      m_FontHeader1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

      m_StyleHeader1Left = wkbk.createCellStyle();
      m_StyleHeader1Left.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      m_StyleHeader1Left.setFont(m_FontHeader1);

      m_StyleHeader1Right = wkbk.createCellStyle();
      m_StyleHeader1Right.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      m_StyleHeader1Right.setFont(m_FontHeader1);

      //
      // Create a fonts and styles for the subheader
      m_FontHeader2 = wkbk.createFont();
      m_FontHeader2.setFontHeightInPoints((short)11);
      m_FontHeader2.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

      m_StyleHeader2Left = wkbk.createCellStyle();
      m_StyleHeader2Left.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      m_StyleHeader2Left.setFont(m_FontHeader2);
      m_StyleHeader2Left.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
      m_StyleHeader2Left.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

      m_StyleHeader2Right = wkbk.createCellStyle();
      m_StyleHeader2Right.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      m_StyleHeader2Right.setFont(m_FontHeader2);
      m_StyleHeader2Right.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
      m_StyleHeader2Right.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

      //
      // Create a fonts and styles for the column headings
      m_FontColumnTitle = wkbk.createFont();
      m_FontColumnTitle.setFontHeightInPoints((short)10);
      m_FontColumnTitle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

      m_StyleColumnTitle = wkbk.createCellStyle();
      m_StyleColumnTitle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
      m_StyleColumnTitle.setFont(m_FontColumnTitle);

      //
      // Create a fonts and styles for the detail
      m_FontDetailText = wkbk.createFont();
      m_FontDetailText.setFontHeightInPoints((short)10);
      m_FontDetailText.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);

      m_FontDetailQty = wkbk.createFont();
      m_FontDetailQty.setFontHeightInPoints((short)10);
      m_FontDetailQty.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);

      m_FontDetailAmt = wkbk.createFont();
      m_FontDetailAmt.setFontHeightInPoints((short)10);
      m_FontDetailAmt.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);

      m_StyleDetailText = wkbk.createCellStyle();
      m_StyleDetailText.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      m_StyleDetailText.setFont(m_FontDetailText);

      m_StyleDetailQty = wkbk.createCellStyle();
      m_StyleDetailQty.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      m_StyleDetailQty.setFont(m_FontDetailQty);

      m_StyleDetailAmt = wkbk.createCellStyle();
      m_StyleDetailAmt.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      m_StyleDetailAmt.setFont(m_FontDetailAmt);
      m_StyleDetailAmt.setDataFormat((short)4);
   }

   /**
    * Prepares the sql queries for execution.
    *
    * @param reportBy
    * @return true if the statements were successfully prepared
    */
   private boolean prepareStatements(String reportBy)
   {
      StringBuffer sql = new StringBuffer(25);
      boolean isPrepared = false;

      if ( m_EdbConn != null ) {
         try {
            sql.setLength(0);
            sql.append("select sum(qty_shipped) qty_ordered, sum(ext_sell) amt_ordered ");
            sql.append("from show ");
            sql.append("join show_packet on show_packet.show_id = show.show_id ");
            sql.append("join promotion on promotion.packet_id = show_packet.packet_id ");
            sql.append("join promo_item on promo_item.promo_id = promotion.promo_id and promo_item.item_ea_id = ? ");
            sql.append("join inv_dtl on inv_dtl.invoice_date >= current_date - 365 and ");
            sql.append("                inv_dtl.promo_nbr = promo_item.promo_id and ");
            sql.append("                inv_dtl.item_ea_id = promo_item.item_ea_id and ");
            sql.append("                inv_dtl.cust_nbr = ? and ");
            sql.append("                inv_dtl.qty_shipped <> 0 ");
            sql.append("where show.show_id = ? ");
            m_CustInvoices = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select sum(qty_ordered) as qty_ordered, sum(qty_ordered * sell_price) as amt_ordered ");
            sql.append("from show ");
            sql.append("join show_packet on show_packet.show_id = show.show_id ");
            sql.append("join promotion on promotion.packet_id = show_packet.packet_id ");
            sql.append("join promo_item on promo_item.promo_id = promotion.promo_id and promo_item.item_ea_id = ? ");
            sql.append("join order_line on order_line.promo_id = promo_item.promo_id and order_line.item_ea_id = promo_item.item_ea_id ");
            sql.append("join order_header on order_header.order_id = order_line.order_id and order_header.customer_id = ? ");
            sql.append("join order_status header_status on header_status.order_status_id = order_header.order_status_id and ");
            sql.append("     header_status.description in ('NEW','HOLD','WAITING CREDIT APPROVAL','WAITING FOR INVENTORY','RELEASED','FASCOR RELEASED') ");
            sql.append("join order_status line_status on line_status.order_status_id = order_line.order_status_id and  ");
            sql.append("     line_status.description in ('NEW','RELEASE','FASCOR RELEASED')  ");
            sql.append("where show.show_id = ? ");
            m_CustOrders = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select ");
            sql.append("   customer.customer_id, ");
            sql.append("   customer.name custname, ");
            sql.append("   item_entity_attr.item_id, ");
            sql.append("   item_entity_attr.item_ea_id, ");
            sql.append("   preprint_item.alt_item_desc, ");
            sql.append("   vendor.vendor_id, ");
            sql.append("   vendor.name vendorname,   ");
            sql.append("   vendor_item_ea_cross.vendor_item_num, ");
            sql.append("   show_vendor.booth, ");
            sql.append("   show_rule_item.min_qty, ");
            sql.append("   show_rule_item.min_amt, ");
            sql.append("   show_rule_vendor.min_amt vnd_min ");
            sql.append("from ( ");
            sql.append("   select distinct order_header.customer_id, order_line.item_ea_id ");
            sql.append("   from order_line ");
            sql.append("   join order_header on order_header.order_id = order_line.order_id ");
            sql.append("   join order_status on order_status.order_status_id = order_header.order_status_id and ");
            sql.append("      order_status.description in ('NEW','RELEASED','FASCOR RELEASED','WAITING CREDIT APPROVAL','WAITING FOR INVENTORY','HOLD') ");
            sql.append("   join order_status line_status on line_status.order_status_id = order_line.order_status_id and ");
            sql.append("      line_status.description in ('NEW','RELEASED','FASCOR RELEASED') ");
            sql.append("   join promo_item on promo_item.promo_id = order_line.promo_id and promo_item.item_ea_id = order_line.item_ea_id ");
            sql.append("   join promotion on promotion.promo_id = promo_item.promo_id ");
            sql.append("   join show_packet on show_packet.packet_id = promotion.packet_id and show_packet.show_id = ? ");
            sql.append("   join item_entity_attr on item_entity_attr.item_ea_id = order_line.item_ea_id and ");
            sql.append("      ( "); 
            sql.append("          item_entity_attr.item_ea_id in ( ");
            sql.append("             select item_ea_id ");
            sql.append("             from show_rule_item ");
            sql.append("             where show_rule_item.rule_id = ?");
            sql.append("          ) or ");
            sql.append("          item_entity_attr.vendor_id in(select vendor_id from show_rule_vendor where show_rule_vendor.rule_id = ?) ");
            sql.append("      ) ");
            sql.append("   union ");
            sql.append("   select distinct inv_dtl.cust_nbr, inv_dtl.item_ea_id ");
            sql.append("   from inv_dtl ");
            sql.append("   join promo_item on promo_item.promo_id = inv_dtl.promo_nbr and promo_item.item_ea_id = inv_dtl.item_ea_id   ");
            sql.append("   join promotion on promotion.promo_id = promo_item.promo_id ");
            sql.append("   join show_packet on show_packet.packet_id = promotion.packet_id and show_packet.show_id = ?   ");
            sql.append("   join item_entity_attr on item_entity_attr.item_ea_id = inv_dtl.item_ea_id and   ");
            sql.append("      (item_entity_attr.item_ea_id in (select item_ea_id from show_rule_item where show_rule_item.rule_id  = ?) or ");
            sql.append("       item_entity_attr.vendor_id in (select vendor_id from show_rule_vendor where show_rule_vendor.rule_id  = ?) ");
            sql.append("       ) ");
            sql.append("    where inv_dtl.qty_shipped <> 0 and inv_dtl.invoice_date >= (   ");
            sql.append("       select min(po_begin) po_begin ");
            sql.append("       from show_packet   ");
            sql.append("       join  promotion on promotion.packet_id = show_packet.packet_id ");
            sql.append("       where show_packet.show_id = ?  ");
            sql.append("    ) ");
            sql.append(") cust_item ");
            sql.append("join customer on customer.customer_id = cust_item.customer_id ");

            if ( reportBy.equalsIgnoreCase("tm") ) {
               sql.append("join cust_rep_div_view on cust_rep_div_view.customer_id = customer.customer_id and cust_rep_div_view.er_id = ? ");
            }

            if ( reportBy.equalsIgnoreCase("Customer") && !m_CustId.equalsIgnoreCase("all") )
               sql.append(" and customer.customer_id = '" + m_CustId + "' ");

            sql.append("join item_entity_attr on item_entity_attr.item_ea_id = cust_item.item_ea_id ");
            sql.append("join show_packet on show_packet.show_id = ? ");
            sql.append("join promotion on promotion.packet_id = show_packet.packet_id ");
            sql.append("join promo_item on promo_item.promo_id = promotion.promo_id and promo_item.item_ea_id = cust_item.item_ea_id  ");
            sql.append("join preprint_item on preprint_item.promo_item_id = promo_item.promo_item_id ");
            sql.append("join vendor_item_ea_cross on vendor_item_ea_cross.item_ea_id = item_entity_attr.item_ea_id and  ");
            sql.append("   vendor_item_ea_cross.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id    ");

            if ( reportBy.equalsIgnoreCase("Vendor") && !m_VendorId.equalsIgnoreCase("all")) {
               sql.append(" and vendor.vendor_id = " + m_VendorId + " ");
            }

            sql.append("join show_rule on show_rule.show_id = show_packet.show_id and show_rule.rule_id = ? ");
            sql.append("left outer join show_rule_item on show_rule_item.rule_id = show_rule.rule_id and show_rule_item.item_ea_id = item_entity_attr.item_ea_id ");
            sql.append("join show_vendor on show_vendor.show_id = show_packet.show_id and show_vendor.vendor_id = vendor.vendor_id ");
            sql.append("left outer join show_rule_vendor on show_rule_vendor.rule_id = show_rule.rule_id and show_rule_vendor.vendor_id = vendor.vendor_id ");
            sql.append("group by ");
            sql.append("   customer.customer_id, customer.name, item_entity_attr.item_id, item_entity_attr.item_ea_id, ");
            sql.append("   preprint_item.alt_item_desc, vendor.vendor_id, vendor.name, vendor_item_ea_cross.vendor_item_num, ");
            sql.append("   show_vendor.booth, show_rule_item.min_qty, show_rule_item.min_amt, show_rule_vendor.min_amt ");

            if (reportBy.equalsIgnoreCase("Vendor") )
               sql.append("order by vendor.name, customer.name, preprint_item.alt_item_desc ");

            if ( reportBy.equalsIgnoreCase("Customer") || reportBy.equalsIgnoreCase("tm") )
               sql.append("order by customer.name, vendor.name, preprint_item.alt_item_desc ");

            m_CustItems = m_EdbConn.prepareStatement(sql.toString());

            //
            // Emery rep information
            sql.setLength(0);
            sql.append("select first || ' ' || last as name from emery_rep where er_id = ?");
            m_TmData = m_EdbConn.prepareStatement(sql.toString());

            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("[ShowSpiffs]", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("ShowSpiffs.prepareStatements - null oracle connection");

      return isPrepared;
   }

   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      int pcount = params.size();
      Param param = null;
      String email = null;

      m_OutFormat = "EXCEL";
      m_ReportBy = "customer";

      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);

         if ( param.name.equalsIgnoreCase("showid") )
            m_ShowId = Integer.parseInt(param.value);

         if ( param.name.equalsIgnoreCase("ruleid") )
            m_RuleId = Integer.parseInt(param.value);

         if ( param.name.equalsIgnoreCase("reportby") )
            m_ReportBy = param.value;

         if ( param.name.equalsIgnoreCase("custid") )
            m_CustId = param.value;

         if ( param.name.equalsIgnoreCase("vendorid") )
            m_VendorId = param.value;

         if ( param.name.equalsIgnoreCase("tmid") )
            m_TmId = Integer.parseInt(param.value);

         if ( param.name.equalsIgnoreCase("webreportid") ) {
            try {
               m_WebRptId = Integer.parseInt(param.value);
            }

            catch ( Exception ex ) {
               m_WebRptId = -1;
            }
         }

         if ( param.name.equalsIgnoreCase("email") )
            email = param.value;
      }

      if ( getWebReport() ) {
         try {
            m_WebRpt.setEMail(email);
            m_WebRpt.setDocFormat(m_OutFormat);
            m_WebRpt.setLineCount(0);
            m_WebRpt.setZipped(m_RptProc.getZipped());
            m_WebRpt.setStatus("RUNNING");
            m_WebRpt.update();
            m_WebRpt.getConnection().commit();
         }

         catch ( Exception e ) {
            m_WebRpt.setComments("Unable to set web_report parameters " + e.getMessage());

            try {
               m_WebRpt.update();
               m_EdbConn.commit();
            }

            catch ( Exception ex ) {
            }

            log.error("[ShowSpiffs]", e);
         }
      }      
   }

}
