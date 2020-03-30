/**
 * File: PurchaseHistByAcct.java
 * Description: Customer Item Purchase History Summary by Account<p>
 *    Rewritten to work with the new report server.
 *    Original author was Peggy Richter.
 *
 * @author Peggy Richter
 * @author Jeffrey Fisher
 *
 * Create Date: 05/17/2005
 * Last Update: $Id: PurchaseHistByAcct.java,v 1.13 2013/01/16 19:47:40 jfisher Exp $
 *
 * History:
 *    $Log: PurchaseHistByAcct.java,v $
 *    Revision 1.13  2013/01/16 19:47:40  jfisher
 *    Removed oracle specific data type
 *
 *    Revision 1.12  2009/02/18 16:13:18  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    10/05/2005 - Remove the isLast() test in writeNeeded.  Throwing exception and not needed. pjr
 *
 *    09/20/2005 - CR# 646.  Wan't incrementing row count on last row.  pjr
 *
 *    09/07/2005 - Additional changes needed to work with new report server.  pjr
 *
 *    03/25/2005 - Added log4j logging. jcf
 *
 *    05/03/2004 - Removed the usage of the m_DistList member variable.  This variable gets cleaned up before it can be
 *       used in the email webservice. Also fixed some errors with the email subject and wrong class name in the
 *       exception - jcf
 *
 *    01/27/2004 - Added more info to server status messages.  pjr
 *
 *    01/14/2004 - Catch broken pipe errors, stop report, and sent email with error.  pjr
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.StringTokenizer;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.web.WebReport;
import com.emerywaterhouse.websvc.Param;


public class PurchaseHistByAcct extends Report
{
   private final short MAX_LINES = 6000;

   private short m_ColCnt = 15;
   private PreparedStatement m_Invoice;
   private PreparedStatement m_GetRMS;
   private PreparedStatement m_CurSell;
   private PreparedStatement m_CurRetail;
   private ResultSet m_Data = null;

   //parameters
   private String m_BegDate = null;
   private String m_EndDate = null;
   private String m_ItemId = null;
   private String m_Packet = null;
   private String m_Promo = null;
   private String m_CustId = null;
   private String[] m_CustList;
   private String m_Vendor = null;
   private String m_NRHA = null;
   private String m_MDC = null;
   private String m_FLC = null;
   private String m_RMS = null;
   private int m_WebRptId;
   private int m_Cnt = 0;

   //web_report bean
   private WebReport m_WebRpt;

   private String m_OutFormat = null;

   //report fields
   private int m_CustCnt;
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
   private String m_SqlCusts;
   private int m_RetailPackFld;
   private double m_CostFld;
   private double m_RetailFld;
   private int m_TotQty;
   private int m_PromoQty;
   private float[] m_Qtys;

   private boolean m_DBError;

   /**
    * default constructor
    */
   public PurchaseHistByAcct()
   {
      super();

      m_FileNames.add("purchhistbyacct");
      initFlds();
   }

   /**
    * Load an item into the member variable fields for further processing
    */
   private void addItem()
   {
      String itemId;
      String promoId;
      int itemEaId;

      if ( m_Data != null ) {
         try {
            itemId = m_Data.getString("item_nbr");
            itemEaId = m_Data.getInt("item_ea_id");
            
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
            m_CostFld = getCurSell(m_CustId, itemEaId);
            m_RetailFld = getCurRetail(m_CustId, itemEaId);
            m_TotQty = m_TotQty + m_Data.getInt("qty_shipped");

            promoId = m_Data.getString("promo_nbr");
            if ( (promoId == null) || promoId.trim().equals("") ) {
            }
            else
               m_PromoQty = m_PromoQty + m_Data.getInt("qty_shipped");

            //add the qty shipped to the appropriate month column
            for ( int i = 0; i < m_CustCnt; i++ ) {
               if ( m_CustList[i].equals(m_Data.getString("cust_nbr"))) {
                  m_Qtys[i] = m_Qtys[i] + m_Data.getInt("qty_shipped");
                  break;
               }
            }
         }
         catch ( Exception e ) {
            log.error("[PurchaseHistoryByAcct]", e);
         }
      }
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

      for ( int i = 0; i < m_CustCnt; i++ )  {
         Row.getCell((i + 15)).setCellValue(m_Qtys[i]);
         m_Qtys[i] = 0;
      }
   }

   /**
    * Add a row to the report if building a tab delimited text file
    *
    * @param line
    */
   private void addTextRow(StringBuffer line)
   {
      String tab = "\t";

      line.append(m_ItemFld + tab);
      line.append(m_SkuFld + tab);
      line.append(m_UpcFld + tab);
      line.append(m_DescrFld + tab);
      line.append(m_VendorFld + tab);
      line.append(m_NrhaFld + tab);
      line.append(m_MdcFld + tab);
      line.append(m_FlcFld + tab);
      line.append(m_RmsFld + tab);
      line.append(m_ShipUnitFld + tab);
      line.append(m_RetailPackFld + tab);
      line.append(m_CostFld + tab);
      line.append(m_RetailFld + tab);
      line.append(m_TotQty + tab);
      line.append(m_PromoQty + tab);

      for ( int i = 0; i < m_CustCnt; i++ )  {
         line.append(m_Qtys[i] + tab);
         m_Qtys[i] = 0;
      }

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
         msg.append("\tACCOUNT PURCHASE SUMMARY\r\n\r\n");
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
         setCurAction("Build Output File");
         if ( m_OutFormat.equals("EXCEL") )
            res = buildSpreadsheet();
         else
            res = buildTextFile();

         addItem();
      }
      
      catch ( Exception ex ) {
         log.error("[PurchasHistByAcct]", ex);
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
      short RowNum;
      boolean Result = false;
      String itemId = null;

      m_FileNames.set(0, m_FileNames.get(0) + m_WebRpt.getWebReportId() + ".xlsx" );
      OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      WrkBk = new XSSFWorkbook();
      Sheet = WrkBk.createSheet();

      RowNum = createHeading(Sheet);
      createCaptions(Sheet, RowNum);
      RowNum++;

      try {
         m_Invoice.setString(1, m_BegDate);
         m_Invoice.setString(2, m_EndDate);

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
               log.error(e);
               m_Status = RptServer.STOPPED;
               m_DBError = true;

               m_ErrMsg.append("Your Purchase History Summary report had the following errors: \r\n");
               m_ErrMsg.append(e.getClass().getName() + "\r\n");
               m_ErrMsg.append(e.getMessage());
            }

         } //while

         if ( m_ItemFld != null ) {
            addSpreadsheetRow(Sheet, RowNum);
            m_Cnt++; //pjr 09/20/2005 wasn't incrementing line count on last row
         }

         WrkBk.write(OutFile);
         WrkBk.close();

         if ( m_Data != null ) {
            try {
               m_Data.close();
            }

            catch ( Exception ex ) {

            }
         }

         Result = true;
      }  //try

      catch ( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }

      finally {
         m_Data = null;

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
    * @throws FileNotFoundException
    * @return boolean
    */
   private boolean buildTextFile() throws FileNotFoundException
   {
      FileOutputStream OutFile = null;
      boolean Result = false;
      String itemId = null;
      StringBuffer line = new StringBuffer(1024);

      m_FileNames.set(0, m_FileNames.get(0) + m_WebRpt.getWebReportId() + ".prn");
      OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      line.append("Customer: " + m_CustId + "\r\n");  //pjr 7/30/2002 - Add customer id
      line.append("Date Range:\t" + m_BegDate + " thru " + m_EndDate + "\r\n");

      if ( m_ItemId.length() != 0 )
         line.append("Item Id:\t" + m_ItemId + "\r\n");

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

      line.append("Item Nbr\tSKU\tUPC\tItem Description\t");
      line.append("Vendor\tNRHA\tMDC\tFLC\tRMS\tShip Unit\t");
      line.append("Retail Pack\tTodays Cost\tTodays Retail\t");
      line.append("Qty Purchased\tPromo Qty\t");
      for (int i=0; i < m_CustCnt - 1; i++)
         line.append(m_CustList[i] + "\t");
      line.append("\r\n");

      try {
         m_Invoice.setString(1, m_BegDate);
         m_Invoice.setString(2, m_EndDate);

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
               log.error(e);
               m_Status = RptServer.STOPPED;
               m_DBError = true;

               m_ErrMsg.append("Your Purchase History Summary report had the following errors: \r\n");
               m_ErrMsg.append(e.getClass().getName() + "\r\n");
               m_ErrMsg.append(e.getMessage());
            }
         }

         if ( m_ItemFld != null ) {
            addTextRow(line);
            m_Cnt++;  //pjr 09/20/2005 wasn't incrementing count on last row
         }

         OutFile.write(line.toString().getBytes());
         line.delete(0, line.length());

         if ( m_Data != null ) {
            try {
               m_Data.close();
            }

            catch ( Exception ex ) {

            }
         }

         Result = true;
      }

      catch ( Exception ex ) {
         log.error(ex);
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }

      finally {
         m_Data = null;

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
         log.error(ex);
      }

      try {
         if ( m_GetRMS != null )
            m_GetRMS.close();
      }

      catch ( Exception ex ) {
         log.error(ex);
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
    * @param sheet
    */
   private void createCaptions(XSSFSheet sheet, int rownum)
   {
      XSSFRow Row = null;
      XSSFCell Cell = null;

      if ( sheet == null )
         return;

      Row = sheet.createRow(rownum);

      if ( Row != null ) {
         for ( int i = 0; i < (m_ColCnt + m_CustCnt); i++ ) {
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

         for ( int i = 0; i < m_CustCnt; i++ ) {
            int j = (i + 15);
            Row.getCell(j).setCellValue(new XSSFRichTextString(m_CustList[i]));
         }
      }
   }

   /**
    * Creates the heading information that identified the customer and
    * the parameter list  pjr 7/30/2002
    *
    * @param sheet HSSFSheet
    */
   private short createHeading(XSSFSheet sheet)
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
         Cell.setCellValue(new XSSFRichTextString("Requested by cust:"));

         Cell = Row.createCell(3);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString(m_CustId));
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
                  
         if ( prepareStatements() ) {
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
      }

      catch ( Exception ex ) {
         log.fatal("[PurchHistByAcct]", ex);
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
         for ( int i = 0; i < m_ColCnt + m_CustCnt; i++ ) {
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
          log.error("[PurchaseHistoryByAccount]", e);
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
          log.error("[PurchaseHistoryByAccount]", e);
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
       return m_CustId;
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
    * Returns the list of customer in a string that can be used in a query
    * @return The sql
    */
   private String getSqlCust()
   {
      String list = "'";

      for ( int i = 0; i < m_CustCnt; i++ ) {
         list = list + m_CustList[i] + "'";

         if ( i < m_CustCnt - 1 )
            list = list + ",'";
      }

      return list;
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
         m_WebRpt.setConnection(m_RptProc.getEdbConn());
         m_WebRpt.setReportName("Purchase History Summary Report");
      }

      catch ( Exception e ) {
         log.error("[PurchasHistByAccount] Unable to set connection in web report ", e);
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
      }

      catch ( Exception e ) {
         log.error("[PurchasHistByAccount] Purchase History Summary was unable to create a web_report record ", e);
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
   }

   /**
    * Parses the customer list parameter into an array of customer id's
    *
    * @param lst
    * @return String[] containing customer id's
    */
   private String[] loadCustList(String lst)
   {
      String[] custArray;
      StringTokenizer tokens = new StringTokenizer(lst);
      m_CustCnt = tokens.countTokens();
      custArray = new String[m_CustCnt];
      int i = 0;

      setCurAction("Loading Customers");

      while ( tokens.hasMoreTokens() ) {
         custArray[i] = tokens.nextToken();
         setCurAction("Load Customer " + custArray[i]);
         i++;
      }

      return custArray;
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

      if ( m_EdbConn != null) {
         try {
            setCurAction("Preparing Statements");

            sql.append("select inv_hdr.cust_nbr, inv_hdr.cust_name, inv_hdr.invoice_nbr, ");
            sql.append("       to_char(inv_hdr.invoice_date, 'mm/dd/yyyy') as invoice_date, ");
            sql.append("       inv_hdr.cust_po_nbr, ");
            sql.append("       inv_dtl.item_nbr, inv_dtl.item_ea_id, ");
            sql.append("       inv_dtl.item_descr, inv_dtl.ship_unit, inv_dtl.qty_shipped, ");
            sql.append("       nvl(inv_dtl.promo_nbr, ' ') as promo_nbr, ");
            sql.append("       inv_dtl.nrha, inv_dtl.flc, flc.mdc_id, ");
            sql.append("       nvl(cust_sku, ' ') as cust_sku, ");
            sql.append("       upc_code, vendor_name, inv_dtl.retail_pack, ");
            sql.append("       ejd_item_whs_upc.upc_code ");
            sql.append("from inv_dtl ");
            sql.append("join inv_hdr on inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id ");
            sql.append("left join flc on flc.flc_id = inv_dtl.flc ");
            sql.append("join item_entity_attr on item_entity_attr.item_ea_id = inv_dtl.item_ea_id ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = item_entity_attr.ejd_item_id and primary_upc = 1 and ejd_item_whs_upc.warehouse_id = ");
            sql.append("		(select warehouse_id from warehouse whs where whs.name = inv_dtl.warehouse) ");
            sql.append("where inv_hdr.sale_type in ('WAREHOUSE', 'ACE DIRECT') and ");
            sql.append("      inv_hdr.cust_nbr in (" + m_SqlCusts + ") and ");
            sql.append("      inv_hdr.invoice_date >= to_date(?, 'mm/dd/yyyy') and ");
            sql.append("      inv_hdr.invoice_date <= to_date(?, 'mm/dd/yyyy') ");

            if ( m_ItemId.length() > 0 )
               sql.append(" and item_nbr = '" + m_ItemId + "' ");

            if ( m_Packet.length() > 0 )
               sql.append(" and promo_nbr in (select promo_id from promotion where packet_id = '" + m_Packet + "') ");

            if ( m_Promo.length() > 0 )
               sql.append(" and promo_nbr = '" + m_Promo + "' ");

            if ( m_Vendor.length() > 0 )
               sql.append(" and vendor_nbr = '" + m_Vendor + "' ");

            if ( m_NRHA.length() > 0 )
               sql.append(" and nrha = '" + m_NRHA + "' ");

            if ( m_FLC.length() > 0 )
               sql.append(" and flc = '" + m_FLC + "' ");

            if ( m_MDC.length() > 0 )
               sql.append(" and mdc_id = '" + m_MDC + "' ");

            if ( m_RMS.length() > 0 )
               sql.append(" and exists(select * from rms_item where item_ea_id = inv_dtl.item_ea_id and rms_id = '" + m_RMS + "') ");

            sql.append(" order by nrha, flc, item_nbr ");

            m_Invoice = m_EdbConn.prepareStatement(sql.toString());

            m_CurSell = m_EdbConn.prepareStatement("select * from ejd_cust_procs.get_sell_price(?, ?) as sell");
            
            m_CurRetail = m_EdbConn.prepareStatement("select ejd_price_procs.get_retail_price(?, ?) as retil");

            m_GetRMS = m_EdbConn.prepareStatement("select rms_id from rms_item where item_ea_id = ?");

            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.fatal("[PurchHistByAcct]", ex);
         }

         finally {
            sql = null;
         }
      }

      return isPrepared;
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      setCurAction("Loading Report Parameters");

      m_BegDate = params.get(0).value;
      m_EndDate = params.get(1).value;
      m_CustId = params.get(2).value;
      m_ItemId = params.get(3).value.trim();
      m_Vendor = params.get(4).value;
      m_NRHA = params.get(5).value;
      m_FLC = params.get(6).value;
      m_MDC = params.get(7).value;
      m_Packet = params.get(8).value;
      m_Promo = params.get(9).value;
      m_RMS = params.get(12).value;
      m_OutFormat = params.get(13).value;

      try {
         m_WebRptId = Integer.parseInt(params.get(14).value);
      }

      catch ( Exception ex ) {
         m_WebRptId = -1;
      }

      m_CustList = loadCustList(params.get(15).value);
      m_SqlCusts = getSqlCust();

      m_Qtys = new float[m_CustCnt];

      for ( int i = 0; i < m_CustCnt; i++ )
         m_Qtys[i] = 0;

      setCurAction("Retrieving Web Report Record");
      if ( getWebReport() ) {
         try {
            m_WebRpt.setFileName(m_FileNames.get(0));
            m_WebRpt.setEMail(params.get(16).value);
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

            log.error("[PurchHistByAcct]", e);
         }
      }
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
      boolean newLine = false;

      //pjr 10/05/2005 Remove the isLast test
      //newLine = (m_ItemFld != null && !m_ItemFld.equals(id)) || m_Data.isLast();
      newLine = (m_ItemFld != null && !m_ItemFld.equals(id));

      return newLine;
   }

}