/**
 * File: PurchaseHistory.java
 * Description: Customer Item Purchase History
 *    Rewrite of the report so it works with the new report server.
 *    The original author was Peggy Richter.
 *
 * @author Peggy Richter
 * @author Jeffrey Fisher
 *
 * Create Date: 05/18/2005
 * Last Update: $Id: PurchaseHistory.java,v 1.11 2015/01/30 15:10:34 ebrownewell Exp $
 *
 * History
 *    $Log: PurchaseHistory.java,v $
 *    Revision 1.11  2015/01/30 15:10:34  ebrownewell
 *    *** empty log message ***
 *
 *    Revision 1.10  2013/01/16 19:47:40  jfisher
 *    Removed oracle specific data type
 *
 *    Revision 1.9  2009/02/18 16:13:18  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    09/07/2005 - Additional fixes needed for new report server.  pjr
 *
 *    08/02/2004 - Fixed a bug in the select by RMS function.  pjr
 *
 *    05/03/2004 - Removed the usage of the m_DistList member variable.  This variable gets cleaned up before it can be
 *       used in the email webservice. - jcf
 *
 *    04/07/2004 - Applied Email class changes. - jcf
 *
 *    01/27/2004 - Added more info to server status messages.  pjr
 *
 *    01/14/2004 - Catch broken pipe errors, stop report, and sent email with error.  pjr
 *
 *    01/12/2004 - Added customer id of user requesting report to header
 *
 *    12/12/2002 - Updated pkg name for new POI 1.5.1.  Removed deprecated method
 *       createCell(short column, int type). PD
 *
 *    07/30/2002 - pjr Add customer number to top line of output files.
 *                     Show customer number in ftadmin current action
 *                     Include credits in sale quantities
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.web.WebReport;
import com.emerywaterhouse.websvc.Param;


public class PurchaseHistory extends Report
{
   private final short COL_CNT = 22;
   private final short MAX_LINES = 6000;

   private PreparedStatement m_Invoice;
   private PreparedStatement m_GetRMS;
   private PreparedStatement m_CurSell;
   private PreparedStatement m_CurRetail;

   //parameters
   private Date m_BegDate;
   private Date m_EndDate;
   private String m_InvoiceNum = null;
   private String m_ItemId = null;
   private String m_PO = null;
   private String m_Packet = null;
   private String m_Promo = null;
   private String m_CustId = null;
   private String m_Vendor = null;
   private String m_NRHA = null;
   private String m_MDC = null;
   private String m_FLC = null;
   private String m_RMS = null;
   private String m_UserCustId = null;
   private int m_WebRptId;
   private String m_WebEmail;
   private int m_Cnt = 0;

   //web_report bean
   private WebReport m_WebRpt;

   private String m_OutFormat = null;
   private boolean m_DBError = false;

   /**
    * default constructor
    */
   public PurchaseHistory()
   {
      super();

      m_FileNames.add("purchhist");
      m_WebRptId = -1;
      m_WebEmail = "";
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
         msg.append("\tPURCHASE HISTORY DETAIL REPORT\r\n\r\n");
         msg.append("To view your reports:\r\n");
         msg.append("\thttp://www.emeryonline.com/emerywh/subscriber/my_account/report_list.jsp\r\n\r\n");
         msg.append("If you have any questions or suggestions, please contact help@emeryonline.com\r\n");
         msg.append("or call 800-283-0236 ext. 1.");

         return msg.toString();
      }

   /**
    * Builds the output file
    * @return true if successful, false if not
    */
   private boolean buildOutputFile()
   {
      boolean res = true;

      setCurAction("building output file " + m_CustId);
      try {
         if ( m_OutFormat.equals("EXCEL") )
            res =  buildSpreadsheet();
         else
            res =  buildTextFile();
      }

      catch ( Exception ex ) {
         log.error("exception", ex);
         m_WebRpt.setComments("Unable to build report " + ex.getMessage());
         m_WebRpt.setStatus("ERROR");
      }

      m_WebRpt.setLineCount(m_Cnt);

      return res;
   }

   /**
    * Executes the queries and builds the output file in spreadsheet format
    *
    * @throws FileNotFoundException
    * @return boolean
    */
   private boolean buildSpreadsheet() throws FileNotFoundException
   {
      XSSFWorkbook WrkBk = null;
      XSSFSheet Sheet = null;
      XSSFRow Row = null;
      FileOutputStream OutFile = null;
      ResultSet InvoiceData = null;
      int RowNum;
      boolean Result = false;
      String itemId = null;
      int itemEaId;

      m_FileNames.set(0, m_FileNames.get(0) + m_WebRpt.getWebReportId() + ".xlsx" );
      OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      WrkBk = new XSSFWorkbook();
      Sheet = WrkBk.createSheet();

      RowNum = createHeading(Sheet);
      createCaptions(Sheet, RowNum++);

      try {
         m_Invoice.setString(1, m_CustId);
         m_Invoice.setDate(2, m_BegDate);
         m_Invoice.setDate(3, m_EndDate);

         InvoiceData = m_Invoice.executeQuery();

         setCurAction("generating spreadsheet " + m_CustId);

         while ( InvoiceData.next() && m_Status == RptServer.RUNNING ) {
            try {
               Row = createRow(Sheet, null, RowNum);
               itemId = InvoiceData.getString("item_nbr");
               itemEaId = InvoiceData.getInt("item_ea_id");
               
               Row.getCell(0).setCellValue(new XSSFRichTextString(itemId));
               Row.getCell(1).setCellValue(new XSSFRichTextString(InvoiceData.getString("cust_sku")));
               Row.getCell(2).setCellValue(new XSSFRichTextString(InvoiceData.getString("upc_code")));
               Row.getCell(3).setCellValue(new XSSFRichTextString(InvoiceData.getString("item_descr")));
               Row.getCell(4).setCellValue(new XSSFRichTextString(InvoiceData.getString("invoice_nbr")));
               Row.getCell(5).setCellValue(new XSSFRichTextString(InvoiceData.getString("invoice_date")));
               Row.getCell(6).setCellValue(new XSSFRichTextString(InvoiceData.getString("vendor_name")));
               Row.getCell(7).setCellValue(new XSSFRichTextString(InvoiceData.getString("NRHA")));
               Row.getCell(8).setCellValue(new XSSFRichTextString(InvoiceData.getString("MDC_id")));
               Row.getCell(9).setCellValue(new XSSFRichTextString(InvoiceData.getString("FLC")));
               Row.getCell(10).setCellValue(new XSSFRichTextString(getRMS(itemEaId)));
               Row.getCell(11).setCellValue(new XSSFRichTextString(InvoiceData.getString("ship_unit")));
               Row.getCell(12).setCellValue(InvoiceData.getInt("retail_pack"));
               Row.getCell(13).setCellValue(InvoiceData.getInt("qty_shipped"));
               Row.getCell(14).setCellValue(InvoiceData.getFloat("unit_sell"));
               Row.getCell(15).setCellValue(InvoiceData.getFloat("unit_retail"));
               Row.getCell(16).setCellValue(InvoiceData.getFloat("ext_sell"));
               Row.getCell(17).setCellValue(InvoiceData.getFloat("ext_retail"));
               Row.getCell(18).setCellValue(getCurSell(m_CustId, itemEaId));
               Row.getCell(19).setCellValue(getCurRetail(m_CustId, itemEaId));
               Row.getCell(20).setCellValue(new XSSFRichTextString(InvoiceData.getString("promo_nbr")));
               Row.getCell(21).setCellValue(new XSSFRichTextString(InvoiceData.getString("sell_source")));
               RowNum++;
               m_Cnt++;

               if ( RowNum >= MAX_LINES ) {
                  Sheet = WrkBk.createSheet();

                  createCaptions(Sheet, 0);
                  RowNum = 1;
               }
            }
            //Check for broken pipe errors.  Stop the process and send email.
            catch ( Exception e ) {
               log.error("exception", e);
               m_Status = RptServer.STOPPED;
               m_DBError = true;

               m_ErrMsg.append("Your Purchase History report had the following errors: \r\n");
               m_ErrMsg.append(e.getClass().getName() + "\r\n");
               m_ErrMsg.append(e.getMessage());
            }
         }

         WrkBk.write(OutFile);
         WrkBk.close();
         Result = true;
      }

      catch ( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }

      finally {
         DbUtils.closeDbConn(null, null, InvoiceData);
         InvoiceData = null;

         try {
            OutFile.close();
         }

         catch( Exception e ) {
            log.error(e);
         }

         OutFile = null;
      }

      return Result;
   }

   /**
    * Executes the queries and builds the output file in tab delimeted format
    *
    * @return boolean
    * @throws FileNotFoundException
    */
   private boolean buildTextFile() throws FileNotFoundException
   {
      FileOutputStream OutFile = null;
      ResultSet InvoiceData = null;
      boolean Result = false;
      String itemId = null;
      int itemEaId;
      StringBuffer line = new StringBuffer(1024);
      String tab = "\t";

      m_FileNames.set(0, m_FileNames.get(0) + m_WebRpt.getWebReportId() + ".prn");
      OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      line.append("Customer: " + m_CustId + "\r\n");  //pjr 7/30/2002 - Add customer id
      line.append("Date Range:\t" + m_BegDate + " thru " + m_EndDate + "\r\n");

      if ( m_ItemId.length() != 0 )
         line.append("Item Id:\t" + m_ItemId + "\r\n");

      if ( m_InvoiceNum.length() != 0 )
         line.append("Invoice:\t" + m_InvoiceNum + "\r\n");

      if ( m_PO.length() != 0 )
         line.append("PO:\t" + m_PO + "\r\n");

      if ( m_Promo.length() != 0 )
         line.append("Promo:\t" + m_Promo + "\r\n");

      if ( m_Vendor.length() != 0 )
         line.append("Vendor:\t" + m_Vendor + "\r\n");

      if ( m_NRHA.length() != 0 )
         line.append("NRHA:\t" + m_NRHA + "\r\n");

      if ( m_MDC.length() != 0 )
         line.append("MDC:\t" + m_MDC + "\r\n");

      if ( m_FLC.length() != 0 )
         line.append("FLC:\t" + m_FLC + "\r\n");

      if ( m_RMS.length() != 0 )
         line.append("RMS:\t" + m_RMS + "\r\n\r\n");

      line.append("Item Nbr\tSKU\tUPC\tItem Description\tInvoice Nbr\t");
      line.append("Invoice Date\tVendor\tNRHA\tMDC\tFLC\tRMS\tShip Unit\t");
      line.append("Retail Pack\tQty Shipped\tUnit Cost\tUnit Retail\tExt Cost\tExt Retail\t");
      line.append("Todays Cost\tTodays Retail\tPromo\tPrice Method\r\n");

      try {
         m_Invoice.setString(1, m_CustId);
         m_Invoice.setDate(2, m_BegDate);
         m_Invoice.setDate(3, m_EndDate);

         InvoiceData = m_Invoice.executeQuery();

         setCurAction("generating text file" + m_CustId);

         while ( InvoiceData.next() && m_Status == RptServer.RUNNING ) {
            try {
               itemId = InvoiceData.getString("item_nbr");
               itemEaId = InvoiceData.getInt("item_ea_id");
               
               line.append(itemId + tab);
               line.append(InvoiceData.getString("cust_sku") + tab);
               line.append(InvoiceData.getString("upc_code") + tab);
               line.append(InvoiceData.getString("item_descr") + tab);
               line.append(InvoiceData.getString("invoice_nbr") + tab);
               line.append(InvoiceData.getString("invoice_date") + tab);
               line.append(InvoiceData.getString("vendor_name") + tab);
               line.append(InvoiceData.getString("NRHA") + tab);
               line.append(InvoiceData.getString("MDC_id") + tab);
               line.append(InvoiceData.getString("FLC") + tab);
               line.append(getRMS(itemEaId) + tab);
               line.append(InvoiceData.getString("ship_unit") + tab);
               line.append(InvoiceData.getInt("retail_pack") + tab);
               line.append(InvoiceData.getInt("qty_shipped") + tab);
               line.append(InvoiceData.getFloat("unit_sell") + tab);
               line.append(InvoiceData.getFloat("unit_retail") + tab);
               line.append(InvoiceData.getFloat("ext_sell") + tab);
               line.append(InvoiceData.getFloat("ext_retail") + tab);
               line.append(getCurSell(m_CustId, itemEaId) + tab);
               line.append(getCurRetail(m_CustId, itemEaId) + tab);
               line.append(InvoiceData.getString("promo_nbr") + tab);
               line.append(InvoiceData.getString("sell_source") + "\r\n");
               m_Cnt++;
            }
            //Check for broken pipe errors.  Stop the process and send email.
            catch ( Exception e ) {
               log.error("exception", e);
               m_Status = RptServer.STOPPED;
               m_DBError = true;

               m_ErrMsg.append("Your Purchase History report had the following errors: \r\n");
               m_ErrMsg.append(e.getClass().getName() + "\r\n");
               m_ErrMsg.append(e.getMessage());
            }
         }

         OutFile.write(line.toString().getBytes());
         line.delete(0, line.length());
         Result = true;
      }

      catch ( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }

      finally {
         InvoiceData = null;
         DbUtils.closeDbConn(null, null, InvoiceData);

         try {
            OutFile.close();
         }

         catch( Exception e ) {
            log.error(e);
         }

         OutFile = null;
      }

      return Result;
   }

   /**
    * Perform cleanup on the objects and close db connections ets.  Overrides the base class
    * method.  The base class method will call closeStatements for us.
    */
   protected void cleanup()
   {
      closeStatements();

      m_Invoice = null;
      m_CurSell = null;
      m_CurRetail = null;
      m_GetRMS = null;
      m_WebRpt = null;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      try {
         if ( m_Invoice != null )
            m_Invoice.close();
      }

      catch ( Exception ex ) {
         //log.error(ex);
      }

      try {
         if ( m_GetRMS != null )
            m_GetRMS.close();
      }

      catch ( Exception ex ) {
         //log.error(ex);
      }

      try {
         if ( m_CurRetail != null )
            m_CurRetail.close();
      }

      catch ( Exception ex ) {

      }

      try {
         if ( m_CurRetail != null )
            m_CurSell.close();
      }

      catch ( Exception ex ) {

      }
   }

   /**
    * Builds the captions on the worksheet.
    *
    * @param sheet HSSFSheet
    */
   private void createCaptions(XSSFSheet sheet, int RowNum)
   {
      XSSFRow Row = null;
      XSSFCell Cell = null;

      if ( sheet == null )
         return;

      Row = sheet.createRow(RowNum);

      if ( Row != null ) {
         for ( int i = 0; i < COL_CNT; i++ ) {
            Cell = Row.createCell(i);
            Cell.setCellType(CellType.STRING);
         }

         Row.getCell(0).setCellValue(new XSSFRichTextString("Item Nbr"));
         Row.getCell(1).setCellValue(new XSSFRichTextString("SKU"));
         Row.getCell(2).setCellValue(new XSSFRichTextString("UPC"));
         Row.getCell(3).setCellValue(new XSSFRichTextString("Item Description"));
         Row.getCell(4).setCellValue(new XSSFRichTextString("Invoice Nbr"));
         Row.getCell(5).setCellValue(new XSSFRichTextString("Invoice Date"));
         Row.getCell(6).setCellValue(new XSSFRichTextString("Vendor"));
         Row.getCell(7).setCellValue(new XSSFRichTextString("NRHA"));
         Row.getCell(8).setCellValue(new XSSFRichTextString("MDC"));
         Row.getCell(9).setCellValue(new XSSFRichTextString("FLC"));
         Row.getCell(10).setCellValue(new XSSFRichTextString("RMS"));
         Row.getCell(11).setCellValue(new XSSFRichTextString("Ship Unit"));
         Row.getCell(12).setCellValue(new XSSFRichTextString("Retail Pack"));
         Row.getCell(13).setCellValue(new XSSFRichTextString("Qty Shipped"));
         Row.getCell(14).setCellValue(new XSSFRichTextString("Unit Cost"));
         Row.getCell(15).setCellValue(new XSSFRichTextString("Unit Retail"));
         Row.getCell(16).setCellValue(new XSSFRichTextString("Ext Cost"));
         Row.getCell(17).setCellValue(new XSSFRichTextString("Ext Retail"));
         Row.getCell(18).setCellValue(new XSSFRichTextString("Todays Cost"));
         Row.getCell(19).setCellValue(new XSSFRichTextString("Todays Retail"));
         Row.getCell(20).setCellValue(new XSSFRichTextString("Promo"));
         Row.getCell(21).setCellValue(new XSSFRichTextString("Price Method"));

      }
   }

   /**
    * Creates the heading information that identified the customer and
    * the parameter list  pjr 7/30/2002
    *
    * @param sheet HSSFSheet
    */
   private int createHeading(XSSFSheet sheet)
   {
      XSSFRow Row = null;
      XSSFCell Cell = null;
      short rownum = 0;

      if ( sheet == null )
         return -1;

      Row = sheet.createRow(rownum++);

      if ( Row != null ) {
         Cell = Row.createCell(2);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Customer Id:"));

         Cell = Row.createCell(3);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString(m_CustId));
      }

      Row = sheet.createRow(rownum++);

      if ( Row != null ) {
         Cell = Row.createCell(2);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Requested by cust:"));

         Cell = Row.createCell(3);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString(m_UserCustId));
      }

      Row = sheet.createRow(rownum++);

      if ( Row != null ) {
         Cell = Row.createCell(2);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Date Range:"));

         Cell = Row.createCell(3);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString(m_BegDate + " thru " + m_EndDate));
      }

      if ( m_ItemId.length() != 0 ) {
         Row = sheet.createRow(rownum++);

         if ( Row != null ) {
            Cell = Row.createCell(2);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Item:"));

            Cell = Row.createCell(3);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString(m_ItemId));
         }
      }

      if ( m_Promo.length() != 0 ) {
         Row = sheet.createRow(rownum++);

         if ( Row != null ) {
            Cell = Row.createCell(2);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Promo:"));

            Cell = Row.createCell(3);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString(m_Promo));
         }
      }

      if ( m_Vendor.length() != 0 ) {
         Row = sheet.createRow(rownum++);

         if ( Row != null ) {
            Cell = Row.createCell(2);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Vendor:"));

            Cell = Row.createCell(3);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString(m_Vendor));
         }
      }

      if ( m_NRHA.length() != 0 ) {
         Row = sheet.createRow(rownum++);

         if ( Row != null ) {
            Cell = Row.createCell(2);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("NRHA:"));

            Cell = Row.createCell(3);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString(m_NRHA));
         }
      }

      if ( m_MDC.length() != 0 ) {
         Row = sheet.createRow(rownum++);

         if ( Row != null ) {
            Cell = Row.createCell(2);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("MDC:"));

            Cell = Row.createCell(3);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString(m_MDC));
         }
      }

      if ( m_FLC.length() != 0 ) {
         Row = sheet.createRow(rownum++);

         if ( Row != null ) {
            Cell = Row.createCell(2);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("FLC:"));

            Cell = Row.createCell(3);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString(m_FLC));
         }
      }

      if ( m_RMS.length() != 0 ) {
         Row = sheet.createRow(rownum++);

         if ( Row != null ) {
            Cell = Row.createCell(2);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("RMS:"));

            Cell = Row.createCell(3);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString(m_RMS));
         }
      }

      return ++rownum;
   }

   /**
    * Runs the report and creates any output that is needed.
    *
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {
      boolean created = false;
      String fileName = null;
      m_Status = RptServer.RUNNING;

      try {
         m_EdbConn = m_RptProc.getEdbConn();

         if ( prepareStatements() )
            setCurAction("Retrieving WebReport Record");

            if ( getWebReport() ) {
               try {
                  m_WebRpt.setFileName(m_FileNames.get(0));
                  m_WebRpt.setEMail(m_WebEmail);
                  m_WebRpt.setDocFormat(m_OutFormat);
                  m_WebRpt.setLineCount(0);
                  m_WebRpt.setZipped(m_RptProc.getZipped());
                  m_WebRpt.setStatus("RUNNING");
                  m_WebRpt.getConnection().commit();
               }
   
               catch ( Exception e ) {
                  m_WebRpt.setComments("Unable to set web_report parameters " + e.getMessage());
               }
            }
            created = buildOutputFile() && !m_DBError;

         if ( created ) {
            fileName = m_FileNames.get(0);

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
         log.fatal("[Purchase History]", ex);
      }

      finally {
         cleanup();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Creates a row in the worksheet.
    *
    * @param sheet
    * @param style
    * @param rowNum
    * @return HSSFRow
    */
    private XSSFRow createRow(XSSFSheet sheet, XSSFCellStyle style, int rowNum)
    {
      XSSFRow Row = null;
      XSSFCell Cell = null;

      if ( sheet == null )
         return Row;

      Row = sheet.createRow(rowNum);

      if ( Row != null ) {
         for ( int i = 0; i < COL_CNT; i++ ) {
            if ( ( i >= 12 ) && ( i <= 19 ) ) {
               Cell = Row.createCell(i);
               Cell.setCellType(CellType.NUMERIC);
               Cell.setCellValue(0.0);
            }
            else {
               Cell = Row.createCell(i);
               Cell.setCellType(CellType.STRING);
               Cell.setCellValue(new XSSFRichTextString(""));
            }
         }
      }

      return Row;
    }

    /**
     * Attempt to find the current retail for an item
     *
     * @param custid
     * @param itemid
     * @return double 
     */
    private double getCurRetail(String custId, int itemEaId)
    {
       ResultSet rs = null;
       double retail = 0;
       
       try {
          m_CurRetail.setString(1, custId);
          m_CurRetail.setInt(2, itemEaId);
                    
          rs = m_CurRetail.executeQuery();
          
          if ( rs.next() )
             retail = rs.getDouble(1);
       }
             
       catch (Exception e) {
          log.error("[PurchaseHistory]", e);
       }
       
       finally {
          DbUtils.closeDbConn(null, null, rs);
       }
       
       return retail;
    }

    /**
     * Attempt to find the current sell for an item
     *
     * @param custid
     * @param itemid
     * @return float
     *
     * 5/11/09 - Now uses dia_date instead of dsb_date PD
     * 06/02/2004 - Pass promo's dsb_date to pricing routine when calculating current sell
     */
    private double getCurSell(String custId, int itemEaId)
    {
       ResultSet rs = null;
       double sell = 0;
       
       try {
          m_CurSell.setString(1, custId);
          m_CurSell.setInt(2, itemEaId);
          
          rs = m_CurSell.executeQuery();
          
          if ( rs.next() )
             sell = rs.getDouble("price");
       }
             
       catch (Exception e) {
          log.error("[PurchaseHistory]", e);
       }
       
       finally {
          DbUtils.closeDbConn(null, null, rs);
       }
       
       return sell;
    }

    /**
     * overrides base class method for logging.
     * @return The id of the customer from the params passed to the report.
     * @see com.emerywaterhouse.rpt.server.Report#getCustId()
     */
    public String getCustId()
    {
       return m_UserCustId;
    }

   /**
    * Find an RMS assortment that includes this item
    *
    * @param itemid
    * @return String
    * @throws SQLException
    */
   private String getRMS(int itemeaid) throws SQLException
   {
      ResultSet rs = null;

      try{
         m_GetRMS.setInt(1, itemeaid);
         rs = m_GetRMS.executeQuery();

         if ( rs.next() )
            return rs.getString( "rms_id" );
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return "";
   }

   /**
    * Load the web_report record from the database into the WebReport bean.
    * Note:  The connection must be provided by the report program, because
    *        the WebReport bean is shared by the web
    *
    * @return boolean
    */
   private boolean getWebReport()
   {            
      boolean result = true;
   
      // set the connection in the WebReport bean
      try {         
         m_WebRpt = new WebReport();
         m_WebRpt.setConnection(m_RptProc.getEdbConn());
         m_WebRpt.setReportName("Purchase History Detail Report");
         result = true;
      }

      catch ( Exception e ) {
         result = false;
         log.error("[PurchaseHistory] Unable to set connection in web report", e);         
      }

      //
      //Load the web_report id from the 15th parameter.  If an id was passed,
      //load that web_report record into the WebReport bean.  Otherwise, create
      //a new web_report record.
      if ( result ) {
         try {
            if ( m_WebRptId >= 0 )
               m_WebRpt.load(m_WebRptId);
            else
               m_WebRpt.insert();
         }
   
         catch ( Exception e ) {
            result = false;
            log.error("[PurchaseHistory] unable to create a web_report record", e);         
         }
      }

      return result;
   }

   /**
    * Load parameters into member variables
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {            
      SimpleDateFormat dtFmt = new SimpleDateFormat("MM/dd/yyyy");
                  
      setCurAction("Loading Parameters");

      try {
         m_BegDate = new Date(dtFmt.parse(params.get(0).value.trim()).getTime());
         m_EndDate = new Date(dtFmt.parse(params.get(1).value.trim()).getTime());
      }
      
      catch ( Exception ex ) {
         log.fatal("[PurchaseHistory]", ex);
      }
      
      m_UserCustId = params.get(2).value.trim();
      m_ItemId = params.get(3).value.trim();
      m_Vendor = params.get(4).value.trim();
      m_NRHA = params.get(5).value.trim();
      m_FLC = params.get(6).value.trim();
      m_MDC = params.get(7).value.trim();
      m_Packet = params.get(8).value.trim();
      m_Promo = params.get(9).value.trim();
      m_PO = params.get(10).value.trim();
      m_InvoiceNum = params.get(11).value.trim();
      m_RMS = params.get(12).value.trim();
      m_OutFormat = params.get(13).value.trim();

      try {
         m_WebRptId = Integer.parseInt(params.get(14).value);
      }

      catch ( Exception ex ) {
         m_WebRptId = -1;
      }

      m_CustId = params.get(15).value.trim();
      m_WebEmail = params.get(16).value;
   }

   /**
    * Prepares the sql queries for execution.
    *
    * @return boolean
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer();
      boolean isPrepared = false;

      if ( m_EdbConn != null ) {
         try {
            setCurAction("Preparing Statements");

            sql.append("select inv_hdr.cust_nbr, inv_hdr.cust_name, inv_hdr.invoice_nbr, ");
            sql.append("   to_char(inv_hdr.invoice_date, 'mm/dd/yyyy') as invoice_date, ");
            sql.append("   inv_hdr.cust_po_nbr, inv_dtl.item_nbr, inv_dtl.item_ea_id, inv_dtl.item_descr, inv_dtl.ship_unit, ");
            sql.append("   inv_dtl.qty_shipped, inv_dtl.unit_sell, inv_dtl.unit_retail, ");
            sql.append("   inv_dtl.ext_sell, inv_dtl.ext_retail, nvl(inv_dtl.promo_nbr, ' ') as promo_nbr, ");
            sql.append("   inv_dtl.nrha, inv_dtl.flc, flc.mdc_id, nvl(cust_sku, ' ') as cust_sku, ");
            sql.append("   ejd_item_whs_upc.upc_code, vendor_name, inv_dtl.retail_pack, nvl(sell_source, ' ') as sell_source ");            
            sql.append("from inv_dtl ");
            sql.append("join inv_hdr using (inv_hdr_id) ");
            sql.append("join flc on flc_id = inv_dtl.flc ");
            sql.append("join item_entity_attr on item_entity_attr.item_ea_id = inv_dtl.item_ea_id ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = item_entity_attr.ejd_item_id and primary_upc = 1 and ejd_item_whs_upc.warehouse_id = ");
            sql.append("		(select warehouse_id from warehouse whs where whs.name = inv_dtl.warehouse) ");
            sql.append("where ");
            sql.append("inv_hdr.sale_type in ('WAREHOUSE', 'ACE DIRECT') and ");
            sql.append("inv_hdr.cust_nbr = ? and inv_hdr.invoice_date between ? and ? ");
            
            if ( m_InvoiceNum.length() > 0 )
               sql.append(" and inv_hdr.invoice_nbr = '" + m_InvoiceNum + "'");

            if ( m_ItemId.length() > 0 )
               sql.append(" and item_nbr = '" + m_ItemId + "'");

            if ( m_PO.length() > 0 )
               sql.append(" and cust_po_nbr = '" + m_PO + "'");

            if ( m_Packet.length() > 0 )
               sql.append(" and promo_nbr in (select promo_id from promotion where packet_id = '" + m_Packet + "')");

            if ( m_Promo.length() != 0 )
               sql.append(" and promo_nbr = '" + m_Promo + "'");

            if ( m_Vendor.trim().length() > 0 )
               sql.append(" and vendor_nbr = '" + m_Vendor + "'");

            if ( m_NRHA.length() > 0 )
               sql.append(" and nrha = '" + m_NRHA + "'");

            if ( m_FLC.length() > 0 )
               sql.append(" and flc = '" + m_FLC + "'");

            if ( m_MDC.length() > 0 )
               sql.append(" and mdc_id = '" + m_MDC + "'");

            if ( m_RMS.length() > 0 ) {
               sql.append(" and exists(select * from rms_item where item_ea_id = inv_dtl.item_ea_id " );
               sql.append(" and rms_id = '" + m_RMS + "') ");
            }

            sql.append(" order by nrha, flc, item_nbr");
            m_Invoice = m_EdbConn.prepareStatement(sql.toString());

            m_CurSell = m_EdbConn.prepareStatement("select * from ejd_cust_procs.get_sell_price(?, ?) as sell");
            
            m_CurRetail = m_EdbConn.prepareStatement("select ejd_price_procs.get_retail_price(?, ?) as retil");

            m_GetRMS = m_EdbConn.prepareStatement("select rms_id from rms_item where item_ea_id = ?");

            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.fatal("[PurchaseHistory]", ex);
         }
      }

      return isPrepared;
   }

}
