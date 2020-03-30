/**
 * File: PurchaseHistSummary.java
 * Description: Customer Item Purchase History Summary
 *    Rewrite of the original report so that it works with the new report server.
 *    The original author was Peggy Richter.
 *
 * @author Peggy Richter
 * @author Jeffrey Fisher
 *
 * Create Date: 05/18/2005
 * Last Update: $Id: PurchaseHistSummary.java,v 1.13 2015/01/30 15:10:35 ebrownewell Exp $
 *
 * History
 *    10/05/2005 - Removed isLast() call from writeNeeded.  Throwing exception and not needed.  pjr
 *
 *    09/20/2005 - CR# 646.  Wan't incrementing row count on last row.  pjr
 *
 *    09/07/2005 - Changes needed to work with new report server.  pjr
 *
 *    03/25/2005 - Added log4j logging. jcf
 *
 *    08/02/2004 - Fixed a bug in the select by RMS function.  pjr
 *
 *    05/03/2004 - Removed the usage of the m_DistList member variable.  This variable gets cleaned up before it can be
 *       used in the email webservice. - jcf
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
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
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


public class PurchaseHistSummary extends Report
{
   private final short MAX_LINES = 6000;

   private short m_ColCnt = 15;
   private int m_MaxMnth;

   private PreparedStatement m_Invoice;
   private PreparedStatement m_GetRMS;
   private PreparedStatement m_CurSell;
   private PreparedStatement m_CurRetail;
   private ResultSet m_Data = null;

   //parameters
   private String m_UserCustId;
   private String m_BegDate;
   private String m_EndDate;
   private String m_InvoiceNum;
   private String m_InvoiceDate;
   private String m_ItemId;
   private String m_PO;
   private String m_Packet;
   private String m_Promo;
   private String m_CustId;
   private String m_Vendor;
   private String m_NRHA;
   private String m_MDC;
   private String m_FLC;
   private String m_RMS;
   private String m_Email;
   private int m_Cnt = 0;
   private int m_WebRptId;

   //web_report bean
   private WebReport m_WebRpt;

   private String m_OutFormat = null;

   //report fields
   private String m_ItemFld = null;
   private String m_SkuFld;
   private String m_UpcFld;
   private String m_DescrFld;
   private String m_VendorFld;
   private String m_NrhaFld;
   private String m_MdcFld;
   private String m_FlcFld;
   private String m_RmsFld;
   private String m_ShipUnitFld;
   private int m_RetailPackFld;
   private double m_CostFld;
   private double m_RetailFld;
   private int m_TotQty;
   private int m_PromoQty;
   private String[] m_MnthStr = {"","","","","","","","","","","","","","","","","","","","","","","","","",""};
   private int[] m_MnthInt = {0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0};
   private double[] m_Qtys = {0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0};
   private int m_BegYear;
   private int m_EndYear;
   private int m_BegMnth;
   private int m_EndMnth;
   private boolean m_DBError;


   /**
    * default constructor
    */
   public PurchaseHistSummary()
   {
      super();
      m_FileNames.add("purchhistsumm");
      initFlds();
   }

   /**
    * Load an item into the member variable fields for further processing
    */
   private void addItem()
   {
      String itemId;
      int itemEaId;
      String promoId;
      String moStr;
      String yrStr;

      if ( m_Data != null ) {
         try {
            itemId = m_Data.getString("item_nbr");
            itemEaId = m_Data.getInt("item_ea_id");

            if(itemEaId == 0){
               itemEaId = getItemEaId(m_CustId, itemId);
            }

            m_ItemFld = itemId;
            m_UpcFld = m_Data.getString("upc_code");
            m_DescrFld = m_Data.getString("item_descr");
            m_SkuFld = m_Data.getString("cust_sku");
            m_VendorFld = m_Data.getString("vendor_name");
            m_NrhaFld = m_Data.getString("NRHA");
            m_MdcFld = m_Data.getString("MDC_id");
            m_FlcFld = m_Data.getString("FLC");
            m_RmsFld = getRMS(itemEaId);
            m_ShipUnitFld = m_Data.getString("ship_unit");
            m_RetailPackFld = m_Data.getInt("retail_pack");

            if(itemEaId > 0) {
               m_CostFld = getCurSell(m_CustId, itemEaId);
               m_RetailFld = getCurRetail(m_CustId, itemEaId);
            }
            else{
               m_CostFld = 0.0;
               m_RetailFld = 0.0;
            }

            m_TotQty = m_TotQty + m_Data.getInt("qty_shipped");
            m_InvoiceDate = m_Data.getString("invoice_date");

            promoId = m_Data.getString("promo_nbr").trim();
            if ( (promoId != null) && promoId.length() > 0 )
               m_PromoQty = m_PromoQty + m_Data.getInt("qty_shipped");

            //add the qty shipped to the appropriate month column
            moStr = m_InvoiceDate.substring(0, 2);
            yrStr = m_InvoiceDate.substring(6, 10);
            
            for ( int i = 0; i < m_MaxMnth; i++ ) {
               if ( m_MnthInt[i] == ((Integer.parseInt(yrStr) * 100) + (Integer.parseInt(moStr))) ) {
                  m_Qtys[i] = m_Qtys[i] + m_Data.getInt("qty_shipped");
                  break;
               }
            }
         }

         catch ( Exception e ) {
            log.error("[PurchaseHistSummuary] addItem", e);
         }
      }
   }

   private int getItemEaId(String custId, String itemId) {
      int res = 0;

      String sql = "select * from ejd_item_procs.get_item_ea_id(?, ?)";

      try(PreparedStatement stmt = m_EdbConn.prepareStatement(sql)) {
         stmt.setString(1, itemId);
         stmt.setString(2, custId);

         try(ResultSet rs = stmt.executeQuery()){
            if(rs.next())
               res = rs.getInt("code");
         }
      } 
      
      catch (SQLException e) {
         log.error("[PurchaseHistSummary] Failed to get item ea id after item ea id from inv_dtl is null.", e);
      }

      return res;
   }

   /**
    * Add a row to the report if building an Excel spreadsheet
    *
    * @param sheet
    * @param rowNum
    */
   private void addSpreadsheetRow(XSSFSheet sheet, int rowNum)
   {
      XSSFRow Row;

      Row = createRow(sheet, null, rowNum);
      Row.getCell(0).setCellValue(new XSSFRichTextString(m_ItemFld));
      Row.getCell(1).setCellValue(new XSSFRichTextString(m_SkuFld));
      Row.getCell(2).setCellValue(new XSSFRichTextString(m_UpcFld));
      Row.getCell(3).setCellValue(new XSSFRichTextString(m_DescrFld));
      Row.getCell(4).setCellValue(new XSSFRichTextString(m_VendorFld));
      Row.getCell(5).setCellValue(new XSSFRichTextString(m_NrhaFld));
      Row.getCell(6).setCellValue(new XSSFRichTextString(m_MdcFld));
      Row.getCell(7).setCellValue(new XSSFRichTextString(m_FlcFld));
      Row.getCell(8).setCellValue(new XSSFRichTextString(m_RmsFld));
      Row.getCell(9).setCellValue(new XSSFRichTextString(m_ShipUnitFld));
      Row.getCell(10).setCellValue(m_RetailPackFld);
      Row.getCell(11).setCellValue(m_CostFld);
      Row.getCell(12).setCellValue(m_RetailFld);
      Row.getCell(13).setCellValue(m_TotQty);
      Row.getCell(14).setCellValue(m_PromoQty);

      for ( int i = 0; i < m_MaxMnth; i++ )
         Row.getCell((i + 15)).setCellValue(m_Qtys[i]);
   }

   /**
    * Add a row to the report if building a tab delimited text file
    *
    * @param line
    */
   private void addTextRow(StringBuffer line)
   {
      String tab = "\t";

      line.append(m_ItemFld).append(tab);
      line.append(m_SkuFld).append(tab);
      line.append(m_UpcFld).append(tab);
      line.append(m_DescrFld).append(tab);
      line.append(m_VendorFld).append(tab);
      line.append(m_NrhaFld).append(tab);
      line.append(m_MdcFld).append(tab);
      line.append(m_FlcFld).append(tab);
      line.append(m_RmsFld).append(tab);
      line.append(m_ShipUnitFld).append(tab);
      line.append(m_RetailPackFld).append(tab);
      line.append(m_CostFld).append(tab);
      line.append(m_RetailFld).append(tab);
      line.append(m_TotQty).append(tab);
      line.append(m_PromoQty).append(tab);

      for ( int i = 0; i < m_MaxMnth; i++ )
         line.append(m_Qtys[i]).append(tab);

      line.append("\r\n");
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
      msg.append("\tPURCHASE HISTORY SUMMARY REPORT\r\n\r\n");
      msg.append("To view your reports:\r\n");
      msg.append("\thttp://www.emeryonline.com/emerywh/subscriber/my_account/report_list.jsp\r\n\r\n");
      msg.append("If you have any questions or suggestions, please contact help@emeryonline.com\r\n");
      msg.append("or call 800-283-0236 ext. 1.");

      return msg.toString();
   }

   /**
    * Create either a spreadsheet or tab delimeted text file
    *
    * @return boolean
    */
   private boolean buildOutputFile()
   {
      boolean res = true;

      try {
         setCurAction("Building Output File");
         if ( m_OutFormat.equals("EXCEL") )
            res = buildSpreadsheet();
         else
            res = buildTextFile();

         addItem();
      }
      catch ( Exception ex ) {
         log.error("[PurchaseHistSummary]", ex);
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
    * @throws FileNotFoundException
    */
   private boolean buildSpreadsheet() throws FileNotFoundException
   {
      XSSFWorkbook WrkBk = null;
      XSSFSheet Sheet = null;
      FileOutputStream OutFile = null;
      int RowNum;
      boolean Result = false;
      String itemId = null;

      m_FileNames.set(0, m_FileNames.get(0) + m_WebRpt.getWebReportId() + ".xlsx" );
      OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      WrkBk = new XSSFWorkbook();
      Sheet = WrkBk.createSheet();

      RowNum = createHeading(Sheet);
      createCaptions(Sheet, RowNum++);

      try {
         m_Invoice.setString(1, m_CustId);
         m_Invoice.setString(2, m_BegDate);
         m_Invoice.setString(3, m_EndDate);

         m_Data = m_Invoice.executeQuery();

         setCurAction("generating spreadsheet " + m_CustId);

         while ( m_Data.next() && m_Status == RptServer.RUNNING ) {
            try {
               itemId = m_Data.getString("item_nbr");

               if ( writeNeeded(itemId) ) {
                  addSpreadsheetRow(Sheet, RowNum);
                  RowNum++;
                  m_Cnt++;
                  initFlds();
               }

               addItem();

               if ( RowNum >= MAX_LINES ) {
                  Sheet = WrkBk.createSheet();

                  createCaptions(Sheet, 0);
                  RowNum = 1;
               }
            }
            //Check for broken pipe errors.  Stop the process and send email.
            catch ( Exception e ) {
               log.error("[PurchaseHistSummary]", e);
               m_Status = RptServer.STOPPED;
               m_DBError = true;

               m_ErrMsg.append("Your Purchase History Summary report had the following errors: \r\n");
               m_ErrMsg.append(e.getClass().getName()).append("\r\n");
               m_ErrMsg.append(e.getMessage());
            }
         }

         //pjr 09/20/2005 wasn't doing the final row increment
         if ( m_ItemFld != null ) {
            addSpreadsheetRow(Sheet, RowNum);
            m_Cnt++;
         }

         WrkBk.write(OutFile);
         WrkBk.close();

         if ( m_Data != null ) {
            try {
               m_Data.close();
            }

            catch ( Exception ex ) {
               log.error("[PurchaseHistSummary] ", ex);
            }
         }

         Result = true;
      }

      catch ( Exception ex ) {
         log.error("[PurchHistSummary]", ex);
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName()).append("\r\n");
         m_ErrMsg.append(ex.getMessage());
      }

      finally {
         m_Data = null;

         try {
            OutFile.close();
         }

         catch( Exception e ) {
            log.error("[PurchHistSummary]", e);
         }
      }

      return Result;
   }

   /**
    * Executes the queries and builds the output file in tab delimeted format
    *
    * @throws FileNotFoundException
    * @return boolean
    */
   private boolean buildTextFile() throws FileNotFoundException
   {
      FileOutputStream OutFile;
      boolean Result = false;
      String itemId;
      StringBuffer line = new StringBuffer(1024);

      m_FileNames.set(0, m_FileNames.get(0) + m_WebRpt.getWebReportId() + ".prn");
      OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      line.append("Customer: ").append(m_CustId).append("\r\n");  //pjr 7/30/2002 - Add customer id
      line.append("Request by cust: ").append(m_UserCustId).append("\r\n");  //pjr 1/12/2004
      line.append("Date Range:\t").append(m_BegDate).append(" thru ").append(m_EndDate).append("\r\n");

      if ( m_ItemId.length() != 0 )
         line.append("Item Id:\t").append(m_ItemId).append("\r\n");

      if ( m_InvoiceNum.length() != 0 )
         line.append("Invoice:\t").append(m_InvoiceNum).append("\r\n");

      if ( m_PO.length() != 0 )
         line.append("PO:\t").append(m_PO).append("\r\n");

      if ( m_Promo.length() != 0 )
         line.append("Promo:\t").append(m_Promo).append("\r\n");

      if ( m_Vendor.length() != 0 )
         line.append("Vendor:\t").append(m_Vendor).append("\r\n");

      if ( m_NRHA.length() != 0 )
         line.append("NRHA:\t").append(m_NRHA).append("\r\n");

      if ( m_MDC.length() != 0 )
         line.append("MDC:\t").append(m_MDC).append("\r\n");

      if ( m_FLC.length() != 0 )
         line.append("FLC:\t").append(m_FLC).append("\r\n");

      if ( m_RMS.length() != 0 )
         line.append("RMS:\t").append(m_RMS).append("\r\n\r\n");

      line.append("Item Nbr\tSKU\tUPC\tItem Description\t");
      line.append("Vendor\tNRHA\tMDC\tFLC\tRMS\tShip Unit\t");
      line.append("Retail Pack\tTodays Cost\tTodays Retail\t");
      line.append("Qty Purchased\tPromo Qty\t");
      
      for (int i=0; i < m_MaxMnth; i++)
         line.append(m_MnthStr[i]).append("\t");
      
      line.append("\r\n");

      try {
         m_Invoice.setString(1, m_CustId);
         m_Invoice.setString(2, m_BegDate);
         m_Invoice.setString(3, m_EndDate);

         m_Data = m_Invoice.executeQuery();

         setCurAction("generating text file " + m_CustId);

         while ( m_Data.next() && m_Status == RptServer.RUNNING ) {
            try {
               itemId = m_Data.getString("item_nbr");

               if ( writeNeeded(itemId) ) {
                  addTextRow(line);
                  m_Cnt++;
                  initFlds();
               }

               addItem();
            }
            //Check for broken pipe errors.  Stop the process and send email.
            catch ( Exception e ) {
               log.error("[PurchHistSummary]", e);
               m_Status = RptServer.STOPPED;
               m_DBError = true;

               m_ErrMsg.append("Your Purchase History Summary report had the following errors: \r\n");
               m_ErrMsg.append(e.getClass().getName()).append("\r\n");
               m_ErrMsg.append(e.getMessage());
            }
         }

         //pjr 09/20/2005 wasn't doing the final row increment
         if ( m_ItemFld != null ) {
            addTextRow(line);
            m_Cnt++;
         }

         OutFile.write(line.toString().getBytes());
         line.delete(0, line.length());

         if ( m_Data != null ) {
            try {
               m_Data.close();
            }

            catch ( Exception ex ) {
               log.error("[PurchaseHistSummary] ", ex);
            }
         }

         Result = true;
      }

      catch ( Exception ex ) {
         log.error("[PurchHistSummary]", ex);
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName()).append("\r\n");
         m_ErrMsg.append(ex.getMessage());
      }

      finally {
         m_Data = null;

         try {
            OutFile.close();
         }

         catch( Exception e ) {
            log.error("[PurchHistSummary]", e);
         }
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
      DbUtils.closeDbConn(null, m_Invoice, null);
      DbUtils.closeDbConn(null, m_GetRMS, null);
      DbUtils.closeDbConn(null, m_CurRetail, null);
      DbUtils.closeDbConn(null, m_CurRetail, null);      
   }

   /**
    * Builds the captions on the worksheet.
    *
    * @param sheet
    */
   private void createCaptions(XSSFSheet sheet, int rownum)
   {
      XSSFRow Row;
      XSSFCell Cell;

      if ( sheet == null )
         return;

      Row = sheet.createRow(rownum);

      if ( Row != null ) {
         for ( int i = 0; i < m_ColCnt; i++ ) {
            Cell = Row.createCell(i);
            Cell.setCellType(CellType.STRING);
         }

         Row.getCell(0).setCellValue(new XSSFRichTextString("Item Nbr"));
         Row.getCell(1).setCellValue(new XSSFRichTextString("SKU"));
         Row.getCell(2).setCellValue(new XSSFRichTextString("UPC"));
         Row.getCell(3).setCellValue(new XSSFRichTextString("Item Description"));
         Row.getCell(4).setCellValue(new XSSFRichTextString("Vendor"));
         Row.getCell(5).setCellValue(new XSSFRichTextString("NRHA"));
         Row.getCell(6).setCellValue(new XSSFRichTextString("MDC"));
         Row.getCell(7).setCellValue(new XSSFRichTextString("FLC"));
         Row.getCell(8).setCellValue(new XSSFRichTextString("RMS"));
         Row.getCell(9).setCellValue(new XSSFRichTextString("Ship Unit"));
         Row.getCell(10).setCellValue(new XSSFRichTextString("Retail Pack"));
         Row.getCell(11).setCellValue(new XSSFRichTextString("Todays Cost"));
         Row.getCell(12).setCellValue(new XSSFRichTextString("Todays Retail"));
         Row.getCell(13).setCellValue(new XSSFRichTextString("Qty Purchased"));
         Row.getCell(14).setCellValue(new XSSFRichTextString("Promo Qty"));

         //load the month column headings
         for ( int i = 15; i < m_ColCnt; i++ ) {
            Row.getCell(i).setCellValue(new XSSFRichTextString(m_MnthStr[i - 15]));
         }
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
      XSSFRow Row;
      XSSFCell Cell;
      int rownum = 0;

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
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {
      boolean created = false;
      String fileName;
      m_Status = RptServer.RUNNING;

      try {
         m_EdbConn = m_RptProc.getEdbConn();

         setCurAction("Retrieving Web Report Record");

         if ( getWebReport() ) {
            try {
               m_WebRpt.setFileName(m_FileNames.get(0));
               m_WebRpt.setEMail(m_Email);
               m_WebRpt.setDocFormat(m_OutFormat);
               m_WebRpt.setLineCount(1);
               m_WebRpt.setZipped(m_RptProc.getZipped());
               m_WebRpt.setStatus("RUNNING");
               m_WebRpt.update();
               m_WebRpt.getConnection().commit();               
            }
            
            catch ( Exception e ) {
               m_WebRpt.setComments("Unable to set web_report parameters " + e.getMessage());

               try {
                  m_WebRpt.update();
               }
               catch ( Exception ex ) {
                  log.error("[PurchaseHistSummary] Error when updating web report.", ex);
               }

               log.error("[PurchaseHistSummary]", e);
            }
         }
         else
            log.error("[PurchaseHistSummary] Unable to get web report object");

         if ( prepareStatements() )
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
         log.fatal("[PurchaseHistSummary]", ex);
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
      XSSFCell Cell;

      if ( sheet == null )
         return Row;

      Row = sheet.createRow(rowNum);

      if ( Row != null ) {
         for ( int i = 0; i < m_ColCnt; i++ ) {
            if ( i > 9 ) {
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
    private double getCurRetail(String custid, int itemEaId)
    {
       ResultSet rs = null;
       double retail = 0;
       
       try {
          m_CurRetail.setString(1, custid);
          m_CurRetail.setInt(2, itemEaId);
                    
          rs = m_CurRetail.executeQuery();
          
          if ( rs.next() )
             retail = rs.getDouble(1);
       }
             
       catch (Exception e) {
          log.error("[PurchaseHistSummary]", e);
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
     * @param itemEaId
     * @return float
     *
     * 5/11/09 - Now uses dia_date instead of dsb_date PD
     * 06/02/2004 - Pass promo's dsb_date to pricing routine when calculating current sell
     */
    private double getCurSell(String custid, int itemEaId)
    {
       ResultSet rs = null;
       double sell = 0;
       
       try {
          m_CurSell.setString(1, custid);
          m_CurSell.setInt(2, itemEaId);
          
          rs = m_CurSell.executeQuery();
          
          if ( rs.next() )
             sell = rs.getDouble("price");
       }
             
       catch (Exception e) {
          log.error(String.format("[PurchaseHistory] Error getting current sell for customer %s and item ea id %d", custid, itemEaId), e);
          
          try {
             m_EdbConn.commit();
          } 
          
          catch (SQLException e1) {
             log.error("[PurchaseHistSummary] Failed to commit edbconn", e1);
          }
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
      String rms = "";

      try{
         m_GetRMS.setInt(1, itemeaid);
         rs = m_GetRMS.executeQuery();

         if ( rs.next() )
            rms =  rs.getString( "rms_id" );
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
      }

      return rms;
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
      // set the connection in the WebReport bean
      try {
         m_WebRpt = new WebReport();
         m_WebRpt.setConnection(m_EdbConn);
         m_WebRpt.setReportName("Purchase History Summary Report");
      }

      catch ( Exception e ) {
         log.error("[PurchaseHistSummary] Unable to set connection in web report", e);
         return false;
      }

      //Load the web_report id from the 15th parameter.  If an id was passed,
      //load that web_report record into the WebReport bean.  Otherwise, create
      //a new web_report record.
      try {
         if ( m_WebRptId >= 0 )
            m_WebRpt.load(m_WebRptId);         
         else
            m_WebRpt.insert();

         m_EdbConn.commit();
      } 
      
      catch ( Exception e ) {
         log.error("[PurchaseHistSummary] Summary was unable to create a web_report record", e);
         return false;
      }

      return true;
   }

   /**
    * Initialize all member variables
    */
   private void initFlds()
   {
      m_SkuFld = "";
      m_UpcFld = "";
      m_DescrFld = "";
      m_VendorFld = "";
      m_NrhaFld = "";
      m_MdcFld = "";
      m_FlcFld = "";
      m_RmsFld = "";
      m_ShipUnitFld = "";
      m_RetailPackFld = 0;
      m_CostFld = 0;
      m_RetailFld = 0;
      m_TotQty = 0;
      m_PromoQty = 0;
      m_WebRptId = -1;

      for ( int i = 0; i < 25; i++ )
         m_Qtys[i] = 0;
   }

   /**
    * Prepares the sql queries for execution.
    *
    * @return boolean
    */
   private boolean prepareStatements()
   {
      StringBuilder sql = new StringBuilder();
      boolean isPrepared = false;

      if ( m_EdbConn != null ) {
         try {
            setCurAction("Preparing statements");

            sql.setLength(0);
            sql.append("select inv_hdr.cust_nbr, inv_hdr.cust_name, inv_hdr.invoice_nbr, ");
            sql.append("   to_char(inv_hdr.invoice_date, 'mm/dd/yyyy') invoice_date, ");
            sql.append("   inv_hdr.cust_po_nbr, inv_dtl.item_nbr, inv_dtl.item_ea_id, inv_dtl.item_descr, inv_dtl.ship_unit, ");
            sql.append("   inv_dtl.qty_shipped, nvl(inv_dtl.promo_nbr, ' ') promo_nbr, inv_dtl.nrha, ");
            sql.append("   inv_dtl.flc, flc.mdc_id, nvl(cust_sku, ' ') cust_sku, ejd_item_whs_upc.upc_code, vendor_name, inv_dtl.retail_pack ");
            sql.append("from inv_dtl ");
            sql.append("join inv_hdr using(inv_hdr_id) ");
            sql.append("join flc on flc_id = inv_dtl.flc ");
            sql.append("join item_entity_attr on item_entity_attr.item_ea_id = inv_dtl.item_ea_id ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = item_entity_attr.ejd_item_id and primary_upc = 1 and ejd_item_whs_upc.warehouse_id = ");
            sql.append("		(select warehouse_id from warehouse whs where whs.name = inv_dtl.warehouse) ");
            sql.append("where ");
            sql.append("   inv_hdr.sale_type in ('WAREHOUSE', 'ACE DIRECT') and ");
            sql.append("   inv_hdr.cust_nbr = ? and ");
            sql.append("   inv_hdr.invoice_date >= to_date(?, 'mm/dd/yyyy') and ");
            sql.append("   inv_hdr.invoice_date <= to_date(?, 'mm/dd/yyyy') ");
            
            if ( m_InvoiceNum.length() != 0 )
               sql.append(" and inv_hdr.invoice_nbr = '").append(m_InvoiceNum).append("'");

            if ( m_ItemId.length() != 0 )
               sql.append(" and item_nbr = '").append(m_ItemId).append("'");

            if ( m_PO.length() != 0 )
               sql.append(" and cust_po_nbr = '").append(m_PO).append("'");

            if ( m_Packet.length() != 0 )
               sql.append(" and promo_nbr in (select promo_id from promotion where packet_id = '").append(m_Packet).append("')");

            if ( m_Promo.length() != 0 )
               sql.append(" and promo_nbr = '").append(m_Promo).append("'");

            if ( m_Vendor.length() != 0 )
               sql.append(" and vendor_nbr = '").append(m_Vendor.trim()).append("'");

            if ( m_NRHA.length() != 0 )
               sql.append(" and nrha = '").append(m_NRHA).append("'");

            if ( m_FLC.length() != 0 )
               sql.append(" and flc = '").append(m_FLC).append("'");

            if ( m_MDC.length() != 0 )
               sql.append(" and mdc_id = '").append(m_MDC).append("'");
            
            if ( m_RMS.length() > 0 ) {
               sql.append(" and exists(select * from rms_item where item_ea_id = inv_dtl.item_ea_id " );
               sql.append(" and rms_id = '" + m_RMS + "') ");
            }
            
            sql.append(" order by nrha, flc, item_nbr");

            m_Invoice = m_EdbConn.prepareStatement(sql.toString());

            m_CurSell = m_EdbConn.prepareStatement("select * from ejd_cust_procs.get_sell_price(?, ?)");
            m_CurRetail = m_EdbConn.prepareStatement("select ejd_price_procs.get_retail_price(?, ?) as retil");
            m_GetRMS = m_EdbConn.prepareStatement("select rms_id from rms_item where item_ea_id = ?");

            isPrepared = true;
         } 
         
         catch ( SQLException ex ) {
            log.fatal("[PurchaseHistSummary]", ex);
         }
      }

      return isPrepared;
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      short i = 1;
      int mnth;
      int year;
      int tmpi;
      int max;

      setCurAction("Loading report parameters");
      m_BegDate = params.get(0).value.trim();
      m_EndDate = params.get(1).value.trim();
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
      m_WebRptId = Integer.parseInt(params.get(14).value);
      m_CustId = params.get(15).value.trim();

      //Initialize the months based on the begin and end dates
      m_BegMnth = Integer.parseInt(m_BegDate.substring(0, 2));
      m_EndMnth = Integer.parseInt(m_EndDate.substring(0, 2));
      m_BegYear = Integer.parseInt(m_BegDate.substring(6, 10));
      m_EndYear = Integer.parseInt(m_EndDate.substring(6, 10));
      m_Email = params.get(16).value;

      mnth = m_BegMnth;
      year = m_BegYear;
      tmpi = (year * 100) + mnth;
      max = (m_EndYear * 100) + m_EndMnth;
      i = 0;

      while ( (tmpi <= max) && (i < 25) ) {
         m_MnthInt[i] = tmpi;
         m_MnthStr[i] = String.valueOf(mnth) + "/" + String.valueOf(year);

         if ( mnth == 12 ) {
            mnth = 1;
            year = year + 1;
         }
         else
            mnth = mnth + 1;

         tmpi = (year * 100) + mnth;
         i++;
         m_ColCnt++;
      }

      m_MaxMnth = i;
   }

   /**
    * A new line must be written every time the item number changes.  Return true
    * if this condition exists
    *
    * @param id
    * @return boolean
    */
   private boolean writeNeeded(String id)
   {
      return ( m_ItemFld != null && !(m_ItemFld.equals(id)));
   }
}
