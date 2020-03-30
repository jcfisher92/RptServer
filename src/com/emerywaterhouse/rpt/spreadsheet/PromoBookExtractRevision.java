/**
 * File: PromoBookExtractRevision.java
 * Description: Promotion Customer Order Book Extract Revision(This is a rewrite of the original promo book extract report)
 *
 * @author Naresh Pasnur
 *
 * Create Date: 01/16/2012
 *
 * History
 *    $Log: PromoBookExtractRevision.java,v $
 *    Revision 1.17  2014/02/03 20:10:02  jfisher
 *    Removed the exclusive promo check per Anne.
 *
 *    Revision 1.16  2013/02/27 22:54:52  npasnur
 *    ambiguous column issue fix
 *
 *    Revision 1.15  2013/01/16 19:45:43  jfisher
 *    added a check for a null or empty history year param
 *
 *    Revision 1.14  2013/01/16 14:36:23  jfisher
 *    heading change.
 *
 *    Revision 1.13  2013/01/16 13:18:32  jfisher
 *    Added purchase history range and removed oracle specific types.
 *
 *    Revision 1.11  2012/02/19 05:48:13  npasnur
 *    Modified code related to base columns.
 *
 *    Revision 1.10  2012/02/15 05:45:08  npasnur
 *    Sorting the QBs and cleaned up some code.
 *
 *    Revision 1.9  2012/02/09 19:14:34  prichter
 *    Exclude all exclusive promotions
 *
 *    Revision 1.8  2012/02/08 10:40:39  npasnur
 *    Added code for Account Summary Report.
 *
 *    Revision 1.7  2012/02/08 03:05:03  npasnur
 *    Development Commit
 *
 *    Revision 1.6  2012/02/02 01:59:17  npasnur
 *    Development Commit
 *
 *    Revision 1.5  2012/02/02 00:49:21  npasnur
 *    Development Commit
 *
 *    Revision 1.4  2012/02/01 00:18:22  npasnur
 *    Development commit
 *
 *    Revision 1.3  2012/01/31 23:14:49  npasnur
 *    Development commit
 *
 *    Revision 1.2  2012/01/31 19:43:30  npasnur
 *    Development commit
 *
 *    Revision 1.1  2012/01/29 00:23:33  npasnur
 *    initial commit
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Types;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Vector;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.utils.StringFormat;
import com.emerywaterhouse.websvc.Param;

public class PromoBookExtractRevision extends Report
{
   private PreparedStatement m_CustName;        //gets the customer name
   private PreparedStatement m_FamilyTree;      //finds all customers connected to a parent id
   private PreparedStatement m_ItemList;        //creates a list of items to report
   private PreparedStatement m_PurchHist;       //units purchased of an item by one or more stores
   
   private PreparedStatement m_PacketInfo;      //returns packet data
   private PreparedStatement m_StmtQtBuys;      //Item QBs
   private PreparedStatement m_CurRetail;       //find the customer's current retail for an item
   private PreparedStatement m_CurSell;         //fine the customer's current sell for an item

   //parameters
   private String m_CustId;
   private boolean m_ByAccount;
   private boolean m_ByCustomer;
   private String m_Format;

   private ArrayList<String> m_Packets;
   private String m_AcctFileName;
   private String m_CustFileName;
   private String m_AsOfDate;  //pjr 06/29/2005 pass the run date as a parameter for calculating sales history
   private boolean m_UsePacketDate = true;
   private int m_PurchHistYears;
   //
   //report fields
   private String m_PacketId;
   private String m_Title;
   private String m_Vendor;
   private String m_Message;
   private String m_ItemDescr;
   private String m_ItemId;
   private String m_Upc;
   private int m_StockPack;
   private String m_Nbc;
   private String m_Unit;
   private double m_CustSell;
   private double m_PromoSell;
   private double m_CustRetail;
   private double m_RetailC;
   private String m_Terms;
   private String m_Deadline;
   private int m_UnitsPurch;

   //report objects
   private XSSFWorkbook m_WrkBk;
   private XSSFSheet m_Sheet;
   private int m_rowNum = 0;

   // miscellaneous member variables
   private boolean m_Error = false;
   private ArrayList<Integer> m_StoreUnits;
   private ArrayList<String> m_StoreId;

   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;

   private static short BASE_COLS = 14;

   //
   // Column widths
   private static final int CW_VENDOR      = 4000;
   private static final int CW_MESSAGE     = 2100;
   private static final int CW_ITEM_DESC   = 6400;
   private static final int CW_REG_BASE    = 2300;
   private static final int CW_PROMO_COST  = 2300;
   private static final int CW_PCT_SAVED   = 1600;

   private static final int CW_STOCK_PACK  = 1600;
   private static final int CW_CUST_RETAIL = 2300;
   private static final int CW_RETAIL_C    = 2300;

   private static final int CW_PKG         = 1200;
   private static final int CW_ITEM_NO     = 2000;
   private static final int CW_ITEM_UPC    = 2800;
   private static final int CW_12MPH       = 1600;
   private static final int CW_ORD_QTY     = 1600;



   /**
    * default constructor
    */
   public PromoBookExtractRevision()
   {
      super();

      m_Packets = new ArrayList<>();
      m_StoreUnits = new ArrayList<>();
      m_StoreId = new ArrayList<>();
      m_PurchHistYears = 1;
   }

   /**
    * Perform cleanup on the objects and close db connections ets.  Overrides the base class
    * method.  The base class method will call closeStatements for us.
    */
   protected void cleanup()
   {
      m_Packets.clear();
      m_StoreUnits.clear();
      m_StoreId.clear();

      DbUtils.closeDbConn(null, m_FamilyTree, null);
      DbUtils.closeDbConn(null, m_CurRetail, null);
      DbUtils.closeDbConn(null, m_CurSell, null);
      DbUtils.closeDbConn(null, m_ItemList, null);      
      DbUtils.closeDbConn(null, m_PurchHist, null);
      DbUtils.closeDbConn(null, m_StmtQtBuys, null);

      m_FamilyTree = null;
      m_CurRetail = null;
      m_CurSell = null;
      m_ItemList = null;      
      m_PurchHist = null;
      m_Sheet = null;
      m_WrkBk = null;
      m_CustName = null;
      m_StmtQtBuys = null;
   }

   /**
    * Runs the report and creates any output that is needed.
    */
   private void closeReport(String reportType, String custid)
   {
      String FileName;
      FileOutputStream OutFile;
      String tmp;

      try {
         if ( reportType.equals("ACCOUNT") )
            FileName = m_AcctFileName;
         else
            FileName = m_CustFileName.replaceFirst("999999", custid);

         //
         //Append the timestamp to the filename to make it unique.
         tmp = Long.toString(System.currentTimeMillis());
         tmp = tmp.substring(tmp.length() - 5, tmp.length());

         FileName = FileName.replaceFirst("ppp", "("+m_PacketId+")"+"-"+tmp);
         if ( m_Format.equals("EXCEL") )
            FileName += ".xlsx";

         OutFile = new FileOutputStream(m_FilePath + FileName, false);

         if ( m_Format.equals("EXCEL") ) {
            m_WrkBk.write(OutFile);
         }

         try {
            OutFile.close();
         }
         catch ( Exception e ) {
            log.error("[PromoBookExtractRevision]", e );
         }

         //
         // Add the file name to the list of files that will be attached or ftp'd
         m_FileNames.add(FileName);
      }
      catch( Exception ex ) {
         log.error("[PromoBookExtractRevision] report for customer: " + custid, ex);
         m_ErrMsg.append("The report had the following Error: \r\n");
         m_ErrMsg.append(ex.getClass().getName()).append("\r\n").append(ex.getMessage());
      }

      finally {
         m_CellStyles = null;
         m_Sheet = null;
         m_WrkBk = null;
      }
   }


   /**
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    * @param colCnt The number of columns in the row.
    *
    * @return The formatted row of the spreadsheet.
    */
   private XSSFRow createRow(int rowNum, int colCnt)
   {
      XSSFRow row = null;
      XSSFCell cell = null;

      if ( m_Sheet == null )
         return row;

      row = m_Sheet.createRow(rowNum);

      //
      // set the type and style of the cell.
      if ( row != null ) {
         for ( int i = 0; i < colCnt; i++ ) {
            cell = row.createCell(i);
            cell.setCellStyle(m_CellStyles[i]);
         }
      }

      return row;
   }

   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook(String rptType)
   {
      XSSFCellStyle styleVendor;    // Text with left alignment
      XSSFCellStyle styleText;      // Text centered
      XSSFCellStyle styleItemDesc;  // Text with left alignment
      XSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00)
      XSSFCellStyle stylePromoCost; // Promo Cost
      XSSFCellStyle styleStockPack; // Stock Pack
      XSSFCellStyle stylePctSaved; // %Saved

      XSSFFont font = m_WrkBk.createFont();
      font.setFontHeightInPoints((short)8);
      font.setFontName("Arial");
      font.setBold(true);

      styleVendor = m_WrkBk.createCellStyle();
      styleVendor.setFont(font);
      styleVendor.setWrapText(true);
      styleVendor.setAlignment(HorizontalAlignment.LEFT);

      //
      //Assign border for each cell of the row
      styleVendor.setBorderTop(BorderStyle.THIN);
      styleVendor.setBorderBottom(BorderStyle.THIN);
      styleVendor.setBorderLeft(BorderStyle.THIN);
      styleVendor.setBorderRight(BorderStyle.THIN);

      styleText = m_WrkBk.createCellStyle();
      styleText.setFont(font);
      styleText.setAlignment(HorizontalAlignment.CENTER);

      //
      //Assign border for each cell of the row
      styleText.setBorderTop(BorderStyle.THIN);
      styleText.setBorderBottom(BorderStyle.THIN);
      styleText.setBorderLeft(BorderStyle.THIN);
      styleText.setBorderRight(BorderStyle.THIN);

      //
      //Style for item desc
      styleItemDesc = m_WrkBk.createCellStyle();
      styleItemDesc.setFont(font);
      styleItemDesc.setWrapText(true);
      styleItemDesc.setAlignment(HorizontalAlignment.LEFT);

      //
      //Assign border for each cell of the row
      styleItemDesc.setBorderTop(BorderStyle.THIN);
      styleItemDesc.setBorderBottom(BorderStyle.THIN);
      styleItemDesc.setBorderLeft(BorderStyle.THIN);
      styleItemDesc.setBorderRight(BorderStyle.THIN);

      //
      //Style for Promo Cost
      stylePromoCost = m_WrkBk.createCellStyle();
      stylePromoCost.setFont(font);
      stylePromoCost.setWrapText(true);
      stylePromoCost.setAlignment(HorizontalAlignment.RIGHT);

      //
      //Assign border for each cell of the row
      stylePromoCost.setBorderTop(BorderStyle.THIN);
      stylePromoCost.setBorderBottom(BorderStyle.THIN);
      stylePromoCost.setBorderLeft(BorderStyle.THIN);
      stylePromoCost.setBorderRight(BorderStyle.THIN);

      //
      //Style for Stock Pack
      styleStockPack = m_WrkBk.createCellStyle();
      styleStockPack.setFont(font);
      styleStockPack.setWrapText(true);
      styleStockPack.setAlignment(HorizontalAlignment.CENTER);

      //
      //Assign border for each cell of the row
      styleStockPack.setBorderTop(BorderStyle.THIN);
      styleStockPack.setBorderBottom(BorderStyle.THIN);
      styleStockPack.setBorderLeft(BorderStyle.THIN);
      styleStockPack.setBorderRight(BorderStyle.THIN);

      //
      //Style for %saved
      stylePctSaved = m_WrkBk.createCellStyle();
      stylePctSaved.setFont(font);
      stylePctSaved.setWrapText(true);
      stylePctSaved.setAlignment(HorizontalAlignment.CENTER);

      //
      //Assign border for each cell of the row
      stylePctSaved.setBorderTop(BorderStyle.THIN);
      stylePctSaved.setBorderBottom(BorderStyle.THIN);
      stylePctSaved.setBorderLeft(BorderStyle.THIN);
      stylePctSaved.setBorderRight(BorderStyle.THIN);

      styleMoney = m_WrkBk.createCellStyle();
      styleMoney.setFont(font);
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short)8);

      //
      //Assign border for each cell of the row
      styleMoney.setBorderTop(BorderStyle.THIN);// This is working
      styleMoney.setBorderBottom(BorderStyle.THIN);
      styleMoney.setBorderLeft(BorderStyle.THIN);
      styleMoney.setBorderRight(BorderStyle.THIN);

      if ( rptType.equals("ACCOUNT") ) {
         //System.out.println("Stored ID size "+m_StoreId.size());
         int arraySize = m_StoreId.size()+14;
         //System.out.println("array size "+arraySize);

         m_CellStyles = new XSSFCellStyle[arraySize];

         m_CellStyles[0] = styleVendor;
         m_CellStyles[1] = styleText;
         m_CellStyles[2] = styleItemDesc;
         m_CellStyles[3] = styleText;
         m_CellStyles[4] = styleMoney;
         m_CellStyles[5] = stylePromoCost;
         m_CellStyles[6] = styleStockPack;
         m_CellStyles[7] = stylePctSaved;
         m_CellStyles[8] = styleMoney;
         m_CellStyles[9] = styleMoney;
         m_CellStyles[10] = styleText;
         m_CellStyles[11] = styleText;
         m_CellStyles[12] = styleText;
         int colCnt = 12;

         for ( int i = 0; i < m_StoreId.size(); i++ ) {
            colCnt = colCnt + 1;
            m_CellStyles[colCnt] = styleText;
         }

         m_CellStyles[colCnt+1] = styleText;
      }
      else{
         m_CellStyles = new XSSFCellStyle[] {
               styleVendor,    // col 0 Vendor
               styleText,      // col 1 Message
               styleItemDesc,  // col 2 Item Desc
               styleText,      // col 3 UPC
               styleMoney,     // col 4 Cust Cost
               stylePromoCost, // col 5 Promo Cost
               styleStockPack, // col 6 Stock Pack
               stylePctSaved,  // col 7 % Saved
               styleMoney,     // col 8 Cust Retail
               styleMoney,     // col 9 C Mkt Retail
               styleText,      // col 10 Package
               styleText,      // col 11 Item #
               styleText,      // col 12 12MPH
               styleText,      // col 13 Ord qty
         };
      }
   }


   /**
    * Creates the captions for the vendor filter.
    *
    * @see SubRpt#createCaptions(int rowNum)
    */
   public int createRowCaptions(String rptType, int rowNum)
   {
      XSSFRow row;
      XSSFCellStyle styleCaptionsRow;
      XSSFFont fontCaptionsRow;
      int col;

      fontCaptionsRow = m_WrkBk.createFont();
      fontCaptionsRow.setFontHeightInPoints((short)8);
      fontCaptionsRow.setFontName("Arial");
      fontCaptionsRow.setBold(true);

      styleCaptionsRow = m_WrkBk.createCellStyle();
      styleCaptionsRow.setFont(fontCaptionsRow);
      styleCaptionsRow.setAlignment(HorizontalAlignment.CENTER);
      //
      //Shading
      styleCaptionsRow.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
      styleCaptionsRow.setFillPattern(FillPatternType.SOLID_FOREGROUND);

      //
      //Border
      styleCaptionsRow.setBorderTop(BorderStyle.THIN);// This is working
      styleCaptionsRow.setBorderBottom(BorderStyle.THIN);
      styleCaptionsRow.setBorderLeft(BorderStyle.THIN);
      styleCaptionsRow.setBorderRight(BorderStyle.THIN);

      //
      //Additional row for QB1
      rowNum++;
      row = m_Sheet.createRow(rowNum);
      //row.setHeightInPoints((3*m_Sheet.getDefaultRowHeightInPoints()));
      row.setRowStyle(styleCaptionsRow);
      col = 0;
      createCaptionCell(row, col, "Vendor",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_VENDOR);
      createCaptionCell(row, col, "Message",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_MESSAGE);
      createCaptionCell(row, col, "Item Description",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_ITEM_DESC);
      createCaptionCell(row, col, "UPC",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_ITEM_UPC);
      createCaptionCell(row, col, "Cust Cost",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_REG_BASE);
      createCaptionCell(row, col, "Promo Cost",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_PROMO_COST);
      createCaptionCell(row, col, "Stk Pk",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_STOCK_PACK);
      createCaptionCell(row, col, "% Saved",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_PCT_SAVED);
      createCaptionCell(row, col, "Cust Retail",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_CUST_RETAIL);
      createCaptionCell(row, col, "Retail C",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_RETAIL_C);
      createCaptionCell(row, col, "Pkg",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_PKG);
      createCaptionCell(row, col, "Item #",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_ITEM_NO);
      createCaptionCell(row, col, "PH",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_12MPH);

      if ( rptType.equals("ACCOUNT") ) {
          for ( int i = 0; i < m_StoreId.size(); i++ ) {
           BASE_COLS = (short)(BASE_COLS  + 1);
              createCaptionCell(row, col, m_StoreId.get(i),styleCaptionsRow);
              m_Sheet.setColumnWidth(col++, CW_12MPH);
          }
      }

      createCaptionCell(row, col, "Ord Qty",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_ORD_QTY);

      //
      //Additional row for QB1
      rowNum++;
      row = m_Sheet.createRow(rowNum);
      row.setRowStyle(styleCaptionsRow);
      col = 5;
      createCaptionCell(row, col, "QB1",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_PROMO_COST);
      createCaptionCell(row, col, "QB1",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_STOCK_PACK);
      createCaptionCell(row, col, "QB1",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_PCT_SAVED);
      //
      //Additional row for QB2
      rowNum++;
      row = m_Sheet.createRow(rowNum);
      row.setRowStyle(styleCaptionsRow);
      col = 5;
      createCaptionCell(row, col, "QB2",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_PROMO_COST);
      createCaptionCell(row, col, "QB2",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_STOCK_PACK);
      createCaptionCell(row, col, "QB2",styleCaptionsRow);
      m_Sheet.setColumnWidth(col++, CW_PCT_SAVED);

      return ++rowNum;
   }

   protected XSSFCell createCaptionCell(XSSFRow row, int col, String caption, XSSFCellStyle stylCaptions)
   {
      XSSFCell cell = null;
      XSSFCellStyle m_CSCaption;
      XSSFFont font;

      if ( row != null ) {
         font = m_WrkBk.createFont();
         font.setFontHeightInPoints((short)8);
         font.setFontName("Arial");
         font.setBold(true);

         m_CSCaption = m_WrkBk.createCellStyle();
         m_CSCaption.setFont(font);
         m_CSCaption.setAlignment(HorizontalAlignment.CENTER);

         //
         //Shading
         m_CSCaption.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
         m_CSCaption.setFillPattern(FillPatternType.SOLID_FOREGROUND);

         cell = row.createCell(col);
         cell.setCellType(CellType.STRING);
         cell.setCellStyle(stylCaptions);
         cell.setCellValue(new XSSFRichTextString(caption != null ? caption : ""));
      }

      return cell;
   }

   private void createPromoReport(String reportType, String cust)
   {
      ResultSet items = null;
      int count;
      int itemEaId = 0;
      ///BASE_COLS = 14;
      
      
      for (String m_Packet : m_Packets) {
         if (m_Status != RptServer.RUNNING)
            break;

         try {
            m_PacketId = m_Packet;
            //initReport("CUSTOMER", cust);

            m_AsOfDate = getAsOfDate();  //pjr 06/29/2005 Used for sales history date range
            m_ItemList.setString(1, cust);
            m_ItemList.setString(2, m_PacketId);            
            items = m_ItemList.executeQuery();
            setCurAction("Cust:" + cust + " packet:" + m_PacketId);
            count = 0;
            
            if ( m_Error )
               log.warn("[PromoBookExtractRevision] Error condition no report run for " + m_CustId );
            
            while ( items.next() && !m_Error && (m_Status == RptServer.RUNNING) ) {
               itemEaId = items.getInt("item_ea_id");
               
               m_Title = items.getString("title");
               m_Vendor = items.getString("vendor");
               m_Message = items.getString("message");
               m_ItemDescr = items.getString("itemdescr");
               m_ItemId = items.getString("item_id");
               m_Upc = items.getString("upc");
               m_StockPack = items.getInt("stock_pack");
               m_Nbc = items.getString("nbc");
               m_Unit = items.getString("unit");
               m_CustSell = getCurSell(cust, itemEaId, items.getDate("dia_date"));   // 06/02/2004 future date base cost.  pjr **PD 5/11/09 changed to dia_date**
               m_PromoSell = getCurSell(cust, itemEaId, items.getString("promo_id"));
               m_CustRetail = getCurRetail(cust, itemEaId);
               m_RetailC = items.getDouble("retailc");
               m_Terms = items.getString("terms");
               m_Deadline = items.getString("deadline");
               
               setCurAction("Cust:" + cust + " packet:" + m_PacketId + " vendor:" + m_Vendor + " item:" + m_ItemId);

               if ( count == 0 )
                  initReport(reportType, cust);

               count++;

               if (reportType != null && reportType.equals("ACCOUNT")) {
                  int qty;
                  m_UnitsPurch = 0;
                  
                  for (int j = 0; j < m_StoreId.size(); j++) {
                     qty = unitsSold(m_StoreId.get(j), itemEaId);
                     m_StoreUnits.set(j, qty);
                     m_UnitsPurch = m_UnitsPurch + qty;
                  }
               }                
               else {
                  m_UnitsPurch = unitsSold(cust, itemEaId);
               }

               createReportLine(reportType, itemEaId);
            }
            
            //
            //Make sure there is at least one line in the report.
            if (count > 0) {
               createFooter(m_Title);
               closeReport(reportType, cust);
            }
         } 
         
         catch (Exception e) {
            log.error("[PromoBookExtractRevision] report for customer: " + cust, e);
            m_Error = true;
         } 
         
         finally {
            DbUtils.closeDbConn(null, null, items);
            items = null;
            System.gc();
         }
      }
   }

   private void createReportLine(String rptType, int itemEaId)
   {
      XSSFRow row;
      String[] itemDescs;
      String desc1="";
      String desc2="";
      int colCnt = BASE_COLS;
      ResultSet rsetQtyBuys = null;
      int qbCnt = 0;
      double[] qbPrice = new double[2];
      int[] qbQty = new int[2];

      row = createRow(m_rowNum, colCnt);

      //
      //Qty Buy prices
      try{
         qbCnt = 0;

         m_StmtQtBuys.setString(1, m_PacketId != null ? m_PacketId : "");
         m_StmtQtBuys.setInt(2, itemEaId);
         rsetQtyBuys = m_StmtQtBuys.executeQuery();
         
         while(rsetQtyBuys.next()){
            //
            //We need qty buys upto second level only
            if(qbCnt==2)
               break;

            qbPrice[qbCnt] = rsetQtyBuys.getDouble("price");
            qbQty[qbCnt] = rsetQtyBuys.getInt("qty");
            qbCnt++;
         }
      }
      
      catch(Exception e){
         log.fatal("[PromoBookExtractRevision] item: " + m_ItemId, e);
      }

      finally{
         closeRSet(rsetQtyBuys);
      }

      //
      //wrap the item desc to two lines
      itemDescs = wrapText(m_ItemDescr, 80);

      for (int i = 0; i < itemDescs.length; i++) {
         switch(i) {
            case 0:
               desc1 =  itemDescs[i];
            case 1:
               desc2 =  itemDescs[i];
         }
      }

      //
      //When the line has 2 levels of QB Pricing, then set the row height to accomodate 3 lines
      if(qbCnt==2){
         row.setHeightInPoints((3*m_Sheet.getDefaultRowHeightInPoints()));

         if(itemDescs.length > 1)
            row.getCell(2).setCellValue(new XSSFRichTextString(desc1+"\n"+desc2));
         else
            row.getCell(2).setCellValue(new XSSFRichTextString(m_ItemDescr));

         row.getCell(5).setCellValue(new XSSFRichTextString("$"+m_PromoSell+"\n"+"$"+qbPrice[0]+"\n"+"$"+qbPrice[1]));
         row.getCell(6).setCellValue(new XSSFRichTextString(m_StockPack +""+m_Nbc+"\n"+qbQty[0]+"\n"+qbQty[1]));
         //
         //% saved for promo cost,QB1 and QB2.
         String pctSavedPromo = Math.round((m_CustSell - m_PromoSell)*100 / m_CustSell)+"%";
         String pctSavedQB1 = Math.round((m_CustSell - qbPrice[0])*100 / m_CustSell)+"%";
         String pctSavedQB2 = Math.round((m_CustSell - qbPrice[1])*100 / m_CustSell)+"%";
         row.getCell(7).setCellValue(new XSSFRichTextString(pctSavedPromo+"\n"+pctSavedQB1+"\n"+pctSavedQB2));
      }
      else if(qbCnt==1 || itemDescs.length > 1){
         row.setHeightInPoints((2*m_Sheet.getDefaultRowHeightInPoints()));

         if(itemDescs.length > 1)
            row.getCell(2).setCellValue(new XSSFRichTextString(desc1+"\n"+desc2));
         else
            row.getCell(2).setCellValue(new XSSFRichTextString(m_ItemDescr));

         if(qbCnt==1){
            row.getCell(5).setCellValue(new XSSFRichTextString("$"+m_PromoSell+"\n"+"$"+qbPrice[0]));
            row.getCell(6).setCellValue(new XSSFRichTextString(m_StockPack +""+m_Nbc+"\n"+qbQty[0]));
            String pctSavedPromo = Math.round((m_CustSell - m_PromoSell)*100 / m_CustSell)+"%";
            String pctSavedQB1 = Math.round((m_CustSell - qbPrice[0])*100 / m_CustSell)+"%";
            row.getCell(7).setCellValue(new XSSFRichTextString(pctSavedPromo+"\n"+pctSavedQB1));
         }
         else{
            row.getCell(5).setCellValue("$"+m_PromoSell);
            row.getCell(6).setCellValue(new XSSFRichTextString(m_StockPack +""+m_Nbc));
            String pctSavedPromo = Math.round((m_CustSell - m_PromoSell)*100 / m_CustSell)+"%";
            row.getCell(7).setCellValue(new XSSFRichTextString(pctSavedPromo));
         }
      }
      else{
         row.getCell(2).setCellValue(new XSSFRichTextString(m_ItemDescr));
         row.getCell(5).setCellValue("$"+m_PromoSell);
         row.getCell(6).setCellValue(new XSSFRichTextString(m_StockPack +""+m_Nbc));
         String pctSavedPromo = Math.round((m_CustSell - m_PromoSell)*100 / m_CustSell)+"%";
         row.getCell(7).setCellValue(new XSSFRichTextString(pctSavedPromo));
      }

      row.getCell(0).setCellValue(new XSSFRichTextString(m_Vendor));
      row.getCell(1).setCellValue(new XSSFRichTextString(m_Message));
      row.getCell(3).setCellValue(new XSSFRichTextString(m_Upc));
      row.getCell(4).setCellValue(m_CustSell);

      row.getCell(8).setCellValue(m_CustRetail);
      row.getCell(9).setCellValue(m_RetailC);
      row.getCell(10).setCellValue(new XSSFRichTextString(m_Unit));
      row.getCell(11).setCellValue(new XSSFRichTextString(m_ItemId));
      row.getCell(12).setCellValue(m_UnitsPurch);

      if ( rptType != null && rptType.equals("ACCOUNT") ) {
         int col = 12;
         Integer qty;
         for ( int i = 0; i < m_StoreId.size(); i++ ) {
            qty = m_StoreUnits.get(i);
            col = col + 1;
            row.getCell(col).setCellValue(qty);
         }
      }

      m_rowNum++;
   }

   /**
    */
   private void createFooter(String packetTitle)
   {
      Footer footer = m_Sheet.getFooter();

      //
      //Packet Title
      footer.setLeft(packetTitle);

      //
      //Page numbers
      footer.setCenter("Page "+ HeaderFooter.page() + " of " + HeaderFooter.numPages() );

      //
      //Packet ID
      footer.setRight("Packet: "+m_PacketId);
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {
      boolean created = false;
      String custId;
      ResultSet rs = null;
      ResultSet rsStores = null;
      m_Status = RptServer.RUNNING;

      try {
         m_EdbConn = m_RptProc.getEdbConn();

         if ( prepareStatements() ) {
            //
            // If the consolidated account option was chosen, produce that report
            if ( m_ByAccount ) {
               setCurAction("Create account report for " + m_CustId);
               log.info("[PromoBookExtractRevision] Create account report for " + m_CustId);
               
               //
               //find all stores related to this account and load them into vector m_StoreId
               try {
                  m_FamilyTree.setString(1, m_CustId);
                  rsStores = m_FamilyTree.executeQuery();
                  
                  while ( rsStores.next() ) {
                     m_StoreId.add( rsStores.getString("customer_id") );
                     m_StoreUnits.add(0);
                  }
               }
               
               catch ( Exception e ) {
                  log.error("[PromoBookExtractRevision]#account list", e);
                  m_Error = true;
               }
               
               finally {
                  DbUtils.closeDbConn(null, null, rsStores);
               }

               createPromoReport("ACCOUNT", m_CustId);

               //
               // If the individual reports by customer were requested, produce a
               // report for each customer within this account
               if ( m_ByCustomer ) {
                  m_FamilyTree.setString(1, m_CustId);
                  rs = m_FamilyTree.executeQuery();

                  while ( rs.next() && m_Status == RptServer.RUNNING) {
                     custId = rs.getString("customer_id");
                     setCurAction("Create customer report for " + custId);
                     createPromoReport("CUSTOMER",custId);
                  }
               }
            }
            else {
               //
               // if only a single customer report was requested, produce that report
               if ( m_ByCustomer ) {
                  setCurAction("Create customer report for " + m_CustId);
                  createPromoReport("CUSTOMER", m_CustId);
               }
            }

            created = true;
         }
      }

      catch ( Exception ex ) {
         log.fatal("[PromoBookExtractRevision] for customer: " + m_CustId, ex);
         m_ErrMsg.append("The remaining reports experienced errors and were not generated\r\n\r\n");
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         cleanup();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Returns the correct sales history end date for the current packet
    */
   private String getAsOfDate()
   {
      String dateStr = null;
      ResultSet rs = null;

      if ( !m_UsePacketDate )
         return m_AsOfDate;

      try {
         m_PacketInfo.setString(1, m_PacketId);
         rs = m_PacketInfo.executeQuery();

         if ( rs.next() )
            dateStr = rs.getString("RepDate");
         else {
            log.error("[PromoBookExtractRevision] Unable to retrieve date for packet " + m_PacketId);
            m_Error = true;
         }
      }

      catch ( Exception e ) {
         log.error("[PromoBookExtractRevision]", e );
         m_Error = true;
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
      }

      return dateStr;
   }

   /**
    * overrides base class method for logging.
    * @return The id of the customer from the params passed to the report.
    * @see com.emerywaterhouse.rpt.server.Report#getCustId()
    */
   @Override
   public String getCustId()
   {
      return m_CustId;
   }

   /**
    * Gets the name of the customer.
    * @param custid
    * @return the customer name
    */
   private String getCustName(String custid)
   {
      ResultSet rs = null;
      String name = null;

      try {
         m_CustName.setString(1, custid);
         rs = m_CustName.executeQuery();

         while ( rs.next() )
            name = rs.getString("name");
      }

      catch ( Exception e ) {
         log.error("[PromoBookExtractRevision]", e );
         m_Error = true;
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return name;
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
         m_CurRetail.setNull(3, Types.VARCHAR);
         
         rs = m_CurRetail.executeQuery();
         
         if ( rs.next() )
            retail = rs.getDouble(1);
      }
            
      catch (Exception e) {
         log.error("[PromoBookExtractRevision]", e);
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
   private double getCurSell(String custid, int itemEaId, java.sql.Date asOf)
   {
      ResultSet rs = null;
      double sell = 0;
      
      try {
         m_CurSell.setString(1, custid);
         m_CurSell.setInt(2, itemEaId);
         m_CurSell.setNull(3, Types.VARCHAR);
         m_CurSell.setDate(4, asOf);

         rs = m_CurSell.executeQuery();
         
         if ( rs.next() )
            sell = rs.getDouble(1);
      }
            
      catch (Exception e) {
         String msg = String.format("[PromoBookExtractRevision]#getCurSell cust %s, iea %d, date %s", custid, itemEaId, asOf.toString());
         log.error(msg, e);         
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
      }
      
      return sell;
   }

   /**
    * Returns the promotional sell price of an item
    *
    * @param custid
    * @param itemEaId
    * @param promoid
    * @return the current sell price
    */
   private double getCurSell(String custid, int itemEaId, String promoid)
   {
      ResultSet rs = null;
      double sell = 0;
      
      
      try {
         m_CurSell.setString(1, custid);
         m_CurSell.setInt(2, itemEaId);
         m_CurSell.setString(3, promoid);
         m_CurSell.setNull(4, Types.DATE);
         m_CurSell.execute();
         
         rs = m_CurSell.executeQuery();
         
         if ( rs.next() )
            sell = rs.getDouble(1);
      }
            
      catch (Exception e) {
         String msg = String.format("[PromoBookExtractRevision]#getCurSell cust %s, iea %d, promo %s", custid, itemEaId, promoid);
         log.error(msg, e);
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
      }
      
      return sell;
   }

   /**
    * Creates the report headings
    *
    * @param reportType
    * @param custid
    */
   private void initReport(String reportType, String custid)
   {
      try {
         if ( m_Format.equals("EXCEL") ) {
            m_WrkBk = new XSSFWorkbook();
            m_Sheet = m_WrkBk.createSheet();
            
            // This was the original, not sure the exact cols.
            //m_WrkBk.setRepeatingRowsAndColumns(0,-1,-1,6,9);         
            m_Sheet.setRepeatingRows(CellRangeAddress.valueOf("7:10"));
            m_Sheet.setRepeatingColumns(CellRangeAddress.valueOf("A:Z"));
            
            m_Sheet.setMargin(HSSFSheet.LeftMargin,0.15);
            m_Sheet.setMargin(HSSFSheet.RightMargin,0.15);
            m_Sheet.setMargin(HSSFSheet.TopMargin,0.25);
            m_Sheet.setMargin(HSSFSheet.BottomMargin,0.75);
            m_Sheet.getPrintSetup().setLandscape(true);
            setupWorkbook(reportType);
            m_rowNum = 1;
            BASE_COLS = 14;
            m_rowNum = createReportHeader(reportType,custid);
            m_rowNum = createRowCaptions(reportType,m_rowNum);
         }
      }

      catch ( Exception e ) {
         log.error("[PromoBookExtractRevision]", e);
         m_Error = true;
      }

   }

   /**
    */
   public int createReportHeader(String reportType, String custid)
   {
      XSSFRow row;
      XSSFCell cell;
      XSSFFont fontCustomerHdr;
      XSSFFont fontCustomerDet;
      XSSFFont fontPacketTitle;
      XSSFFont fontPacketHeader;
      XSSFFont fontPacketHeaderDet;
      XSSFCellStyle styleCustomerHdr;
      XSSFCellStyle styleCustomerDet;
      XSSFCellStyle stylePacketTitle;
      XSSFCellStyle stylePacketHeader;
      XSSFCellStyle stylePacketHeaderDet;

      //
      //Style for Packet Title
      fontPacketTitle = m_WrkBk.createFont();
      fontPacketTitle.setFontHeightInPoints((short)14);
      fontPacketTitle.setFontName("Arial");
      fontPacketTitle.setBold(true);
      fontPacketTitle.setItalic(true);

      stylePacketTitle = m_WrkBk.createCellStyle();
      stylePacketTitle.setFont(fontPacketTitle);
      stylePacketTitle.setAlignment(HorizontalAlignment.LEFT);

      //
      //Style for Customer
      fontCustomerHdr = m_WrkBk.createFont();
      fontCustomerHdr.setFontHeightInPoints((short)9);
      fontCustomerHdr.setFontName("Arial");
      fontCustomerHdr.setBold(true);

      //
      //Style for Customer
      fontCustomerDet = m_WrkBk.createFont();
      fontCustomerDet.setFontHeightInPoints((short)8);
      fontCustomerDet.setFontName("Arial");
      fontCustomerDet.setBold(true);

      styleCustomerHdr = m_WrkBk.createCellStyle();
      styleCustomerHdr.setFont(fontCustomerHdr);
      styleCustomerHdr.setAlignment(HorizontalAlignment.LEFT);

      styleCustomerDet = m_WrkBk.createCellStyle();
      styleCustomerDet.setFont(fontCustomerDet);
      styleCustomerDet.setAlignment(HorizontalAlignment.LEFT);

      //
      //Style for packet header
      fontPacketHeader = m_WrkBk.createFont();
      fontPacketHeader.setFontHeightInPoints((short)9);
      fontPacketHeader.setFontName("Arial");
      fontPacketHeader.setBold(true);

      //
      //Style for packet header info
      fontPacketHeaderDet = m_WrkBk.createFont();
      fontPacketHeaderDet.setFontHeightInPoints((short)8);
      fontPacketHeaderDet.setFontName("Arial");
      fontPacketHeaderDet.setBold(true);

      stylePacketHeader = m_WrkBk.createCellStyle();
      stylePacketHeader.setFont(fontPacketHeader);
      stylePacketHeader.setAlignment(HorizontalAlignment.LEFT);

      stylePacketHeaderDet = m_WrkBk.createCellStyle();
      stylePacketHeaderDet.setFont(fontPacketHeaderDet);
      stylePacketHeaderDet.setAlignment(HorizontalAlignment.LEFT);

      //Packet Title
      //m_rowNum = m_rowNum;
      row = m_Sheet.createRow(m_rowNum);
      cell = row.createCell(1);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(stylePacketTitle);
      cell.setCellValue(new XSSFRichTextString(m_Title));

      //
      //Packet ID
      m_rowNum = m_rowNum + 1;
      row = m_Sheet.createRow(m_rowNum);
      cell = row.createCell(6);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(stylePacketHeader);
      cell.setCellValue(new XSSFRichTextString("Packet: "));

      cell = row.createCell(8);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(stylePacketHeaderDet);
      cell.setCellValue(new XSSFRichTextString(m_PacketId));

      //
      //Customer#
      m_rowNum = m_rowNum+1;
      row = m_Sheet.createRow(m_rowNum);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleCustomerHdr);

      if ( reportType != null && reportType.equals("ACCOUNT"))
        cell.setCellValue(new XSSFRichTextString("Account #: "));
      else
          cell.setCellValue(new XSSFRichTextString("Customer #: "));

      cell = row.createCell(1);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleCustomerDet);
      cell.setCellValue(new XSSFRichTextString(custid));

      //
      //Order Deadline
      cell = row.createCell(6);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(stylePacketHeader);
      cell.setCellValue(new XSSFRichTextString("Order Deadline: "));

      cell = row.createCell(8);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(stylePacketHeaderDet);
      cell.setCellValue(new XSSFRichTextString(m_Deadline));

      //Customer Name
      m_rowNum = m_rowNum + 1;
      row = m_Sheet.createRow(m_rowNum);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleCustomerHdr);
      cell.setCellValue(new XSSFRichTextString("Name: "));

      cell = row.createCell(1);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleCustomerDet);
      cell.setCellValue(new XSSFRichTextString(getCustName(custid)));

      //
      //Terms
      cell = row.createCell(6);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(stylePacketHeader);
      cell.setCellValue(new XSSFRichTextString("Due: "));

      cell = row.createCell(8);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(stylePacketHeaderDet);
      cell.setCellValue(new XSSFRichTextString(m_Terms));

      return ++m_rowNum;
   }

   /**
    * Prepares the sql queries for execution.
    *
    * @return boolean
    * @throws Exception
    */
   private boolean prepareStatements() throws Exception
   {
      boolean isPrepared = false;
      StringBuilder sql = new StringBuilder();

      if ( m_EdbConn != null ) {
         try {
            m_CustName = m_EdbConn.prepareStatement("select name from customer where customer_id = ?");

            // 06/02/2004 - optionally pass dsb date to pricing routine
            m_CurSell = m_EdbConn.prepareStatement("select round(price, 2) as sell from ejd_cust_procs.get_sell_price(?, ?, ?, ?)");

            m_CurRetail = m_EdbConn.prepareStatement("select ejd_price_procs.get_retail_price(?, ?, ?) as retail");

            m_FamilyTree = m_EdbConn.prepareStatement(
               "select name, customer_id from customer,customer_status " +
               "where customer.cust_status_id = customer_status.cust_status_id and description != 'INACTIVE' " +
               "start with customer_id = ? connect by prior customer_id = parent_id"
            );

            sql.setLength(0);
            sql.append("select ");
            sql.append("flc.description as flc_desc, ");
            sql.append("packet.title, ");
            sql.append("iea.item_id, ");
            sql.append("iea.description as itemdescr, ");
            sql.append("iea.item_ea_id, ");
            sql.append("decode(vendor_shortname.name, NULL, vendor.name, vendor_shortname.name) as vendor, ");
            sql.append("nvl(preprint_item.message, ' ') as message, ");
            sql.append("decode(broken_case.description, 'ALLOW BROKEN CASES', ' ', 'N') as nbc, ");
            sql.append("eiw.stock_pack, ");
            sql.append("ship_unit.unit, ");
            sql.append("upc_code as upc, ");
            sql.append("promotion.promo_id, ");
            sql.append("terms.name as terms, ");
            sql.append("to_char(dia_date, 'mm/dd/yyyy') as deadline, ");
            sql.append("ejd_item_price.retail_c as retailc, ");
            sql.append("dia_date, ");
            sql.append("iea.vendor_id ");
            sql.append("from packet ");
            sql.append("join customer on customer.customer_id = ? ");            
            sql.append("join cust_warehouse on cust_warehouse.customer_id = customer.customer_id and whs_priority = 1 ");
            sql.append("join promotion on promotion.packet_id = packet.packet_id ");
            sql.append("join promo_item on promo_item.promo_id = promotion.promo_id ");
            sql.append("join item_entity_attr iea on iea.item_ea_id = promo_item.item_ea_id ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = iea.ejd_item_id ");
            sql.append("join ejd_item_warehouse eiw on ejd_item.ejd_item_id = eiw.ejd_item_id and eiw.warehouse_id = cust_warehouse.warehouse_id ");
            sql.append("left outer join ejd_item_whs_upc eiwu on eiwu.ejd_item_id = ejd_item.ejd_item_id and eiwu.warehouse_id = eiw.warehouse_id and primary_upc = 1 ");
            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id and ejd_item_price.warehouse_id = eiw.warehouse_id ");
            sql.append("join vendor on vendor.vendor_id = iea.vendor_id ");
            sql.append("left outer join preprint_item on preprint_item.promo_item_id = promo_item.promo_item_id ");
            sql.append("join broken_case on  broken_case.broken_case_id = ejd_item.broken_case_id ");
            sql.append("join ship_unit on ship_unit.unit_id = iea.ship_unit_id ");
            sql.append("join terms on terms.term_id = promotion.term_id ");
            sql.append("left outer join vendor_shortname on vendor_shortname.vendor_id = iea.vendor_id ");
            sql.append("join flc on flc.flc_id = ejd_item.flc_id ");
            sql.append("where packet.packet_id = ? and oe_procs.promo_order_ok(promotion.promo_id, customer.customer_id) = 1 ");
            sql.append("group by ");
            sql.append("   flc.description, promotion.promo_id, iea.item_id, iea.description, vendor_shortname.name, vendor.name, upc_code, retail_c, ");
            sql.append("   preprint_item.message, broken_case.description, eiw.stock_pack, ship_unit.unit, dia_date, iea.vendor_id, terms.name, ");
            sql.append("   dia_date, packet.title, iea.item_ea_id ");
            sql.append("order by flc.description, promotion.promo_id, iea.item_id ");

            m_ItemList = m_EdbConn.prepareStatement(sql.toString());            
            
            sql.setLength(0);
            m_PacketInfo = m_EdbConn.prepareStatement(
               "select to_char(nvl(report_begin_date, current_date), 'mm/dd/yyyy') as repdate from packet where packet_id = ?"
            );

            // pjr 06/29/2005 - pass the end of the date range
            sql.setLength(0);
            sql.append("select sum(qty_shipped) as qty ");
            sql.append("from inv_dtl ");
            sql.append("where cust_nbr = ? and item_ea_id = ? and ");
            sql.append("invoice_date between (to_date(?, 'mm/dd/yyyy') - (365 * ?)) and to_date(?, 'mm/dd/yyyy') ");
            m_PurchHist = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select distinct ");
            sql.append("quantity_buy_item.min_qty qty, item_entity_attr.item_id, item_entity_attr.description, ");
            sql.append("round(quantity_buy_item.discount_value,2) price,promotion.promo_id ");
            sql.append("from packet ");
            sql.append("join promotion on promotion.packet_id = packet.packet_id and packet.packet_id = ? ");
            sql.append("join promo_item on promo_item.promo_id = promotion.promo_id and promo_item.item_ea_id = ? ");
            sql.append("join item_entity_attr on item_entity_attr.item_ea_id = promo_item.item_ea_id ");
            sql.append("join quantity_buy on quantity_buy.packet_id = packet.packet_id ");
            sql.append("join discount on discount.discount_id = quantity_buy.discount_id ");
            sql.append("join quantity_buy_item on quantity_buy_item.qty_buy_id = quantity_buy.qty_buy_id and ");
            sql.append("      quantity_buy_item.item_ea_id = promo_item.item_ea_id ");
            sql.append("order by qty ");
          
            m_StmtQtBuys = m_EdbConn.prepareStatement(sql.toString());

            isPrepared = true;
         }

         catch ( Exception ex ) {
            log.fatal("[PromoBookExtractRevision]", ex);
         }
      }

      return isPrepared;
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    *
    * Note - m_Email, and m_Zipped have been removed from the params the report gets.
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      Param param;

      for (Param param1 : params) {
         param = param1;

         if (param.name.equals("CustomerId"))
            m_CustId = param.value;

         if (param.name.equals("CustSummary"))
            m_ByCustomer = Boolean.parseBoolean(param.value);

         if (param.name.equals("AcctSummary"))
            m_ByAccount = Boolean.parseBoolean(param.value);

         if (param.name.equals("Format"))
            m_Format = param.value;

         if (param.name.equals("AcctFile"))
            m_AcctFileName = param.value;

         if (param.name.equals("CustFile"))
            m_CustFileName = param.value;

         if (param.name.equals("PacketList"))
            m_Packets = StringFormat.parseString(param.value, ';');

         if (param.name.equals("phyears")) {
            if (param.value == null || param.value.length() == 0)
               m_PurchHistYears = 1;
            else
               m_PurchHistYears = Integer.parseInt(param.value);
         }

         if (param.name.equals("DateOption")) {
            m_AsOfDate = param.value;
            m_UsePacketDate = (m_AsOfDate.equals("default"));
         }
      }
   }

   /**
    *
    * @param custid
    * @param itemid
    * @return the number of units sold
    */
   private int unitsSold(String custId, int itemEaId)
   {
      ResultSet rs = null;
      int qty = 0;

      try{
         m_PurchHist.setString(1, custId);
         m_PurchHist.setInt(2, itemEaId);
         m_PurchHist.setString(3, m_AsOfDate);
         m_PurchHist.setInt(4, m_PurchHistYears);
         m_PurchHist.setString(5, m_AsOfDate);

         rs = m_PurchHist.executeQuery();

         if ( rs.next() )
            qty = rs.getInt("qty");
      }

      catch ( Exception e ) {
         log.error("[PromoBookExtractRevision]",  e );
         m_Error = true;
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
      }

      return qty;
   }

   private String [] wrapText (String text, int len)
   {
     // return empty array for null text
     if (text == null)
     return new String [] {};

     // return text if len is zero or less
     if (len <= 0)
     return new String [] {text};

     // return text if less than length
     if (text.length() <= len)
     return new String [] {text};

     char [] chars = text.toCharArray();
     Vector<String> lines = new Vector<>();
     StringBuilder line = new StringBuilder();
     StringBuilder word = new StringBuilder();

      for (char aChar : chars) {
         word.append(aChar);

         if (aChar == ' ') {
            if ((line.length() + word.length()) > len) {
               lines.add(line.toString());
               line.delete(0, line.length());
            }

            line.append(word);
            word.delete(0, word.length());
         }
      }

     // handle any extra chars in current word
     if (word.length() > 0) {
       if ((line.length() + word.length()) > len) {
         lines.add(line.toString());
         line.delete(0, line.length());
       }
       line.append(word);
     }

     // handle extra line
     if (line.length() > 0) {
       lines.add(line.toString());
     }

     String [] ret = new String[lines.size()];
     int c = 0; // counter
     for (Enumeration<String> e = lines.elements(); e.hasMoreElements(); c++) {
       ret[c] = e.nextElement();
     }

     return ret;
   }

}