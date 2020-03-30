/**
 * File: DealerNewPricingSummary.java
 * Description: Report to assist territory/account managers in viewing new pricing summary information for
 *    active customers that are not on contract pricing.  All data is for a specific nrha department.
 *    If a valid emery rep id is passed through, then data is retrieved for customers of that
 *    particular sales rep only.
 *
 *    Rewritten so the report will work with the new report server.
 *    The original author was Paul Davidson.
 *
 * @author Paul Davidson
 * @author Jeffrey Fisher
 *
 * Create Date: 05/11/2005
 * Last Update: $Id: DealerNewPricingSummary.java,v 1.8 2013/01/16 19:50:59 jfisher Exp $
 *
 * History
 *    $Log: DealerNewPricingSummary.java,v $
 *    Revision 1.8  2013/01/16 19:50:59  jfisher
 *    Removed oracle specific data type
 *
 *    Revision 1.7  2009/02/18 15:04:40  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    03/25/2005 - Added log4j logging. jcf
 *
 *    04/07/2004 - Applied Email class changes. - jcf
 *
 *    12/23/2003 - Modified the code to handle the new xml request format - jcf
 *
 *    01/29/2003 - Upped the max runtime to 14 hours. New pricing now retrieved from ts01. PD
 *
 *    01/23/2003 - Modified retail and margin calculations to include retail_pack. PD
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Types;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.Files;
import com.emerywaterhouse.websvc.Param;


public class DealerNewPricingSummary extends Report
{
   private static String FILE_NAME = "dlrnewprcsumm_%d.xlsx";
   private String m_NrhaId;
   private int m_RepId;       // emery rep id
   private boolean m_Zipped;  // zipped flag
   private PreparedStatement m_Customers;
   private PreparedStatement m_Items;
   private PreparedStatement m_QtyShip;
   private PreparedStatement m_SalesRep;
   private CallableStatement m_OldCustSellPrice;
   private CallableStatement m_NewCustSellPrice;
   private CallableStatement m_OldCustRetlPrice;
   private CallableStatement m_NewCustRetlPrice;
   private CallableStatement m_OldBaseSellPrice;
   private CallableStatement m_NewBaseSellPrice;
   private CallableStatement m_EmeryCost;
   private XSSFWorkbook m_WorkBook;
   private XSSFSheet m_Sheet;
   private int m_StartRow;

   private final short ROW_HEIGHT = 14;

   /**
    * default constructor
    */
   public DealerNewPricingSummary()
   {
      super();

      //
      // This report could take a long time to run, as it repeatedly needs to get pricing
      // data from ts01.  Will take even longer when running it for all active customers.
      m_MaxRunTime = RptServer.HOUR * 14;

      //
      // Intialize member variables
      m_NrhaId = null;
      m_Customers = null;
      m_Items = null;
      m_QtyShip = null;
      m_SalesRep = null;
      m_OldCustSellPrice = null;
      m_NewCustSellPrice = null;
      m_OldCustRetlPrice = null;
      m_NewCustRetlPrice = null;
      m_OldBaseSellPrice = null;
      m_NewBaseSellPrice = null;
      m_EmeryCost = null;
      m_RepId = -1;
      m_Zipped = false;

      //
      // Create a new workbook and sheet
      m_WorkBook = new XSSFWorkbook();
      m_Sheet = m_WorkBook.createSheet();
   }

   /**
    * Adds a new row customized specifically for this report.
    *
    * @param rowNum short - the row index.
    * @return XSSFRow - the row object added, or a reference to the existing one.
    */
   private XSSFRow addRow(int rowNum)
   {
      XSSFRow row = m_Sheet.createRow( rowNum);
      XSSFCell cell;
      int col = -1;

      //
      // All rows will have this height
      row.setHeightInPoints(ROW_HEIGHT);

      //
      // Add the cells
      if ( m_RepId == -1 ) {
         cell = row.createCell(++col);  // sales rep name
         cell.setCellType(CellType.STRING);
      }

      cell = row.createCell(++col);  // Customer#
      cell.setCellType(CellType.STRING);

      cell = row.createCell(++col);  // Customer name
      cell.setCellType(CellType.STRING);

      cell = row.createCell(++col);  // Sku count
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // # Items decreased
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // # Items increased
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // # Items equal
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Emery cost - old
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Emery sell - old
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Emery GM$ - old
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Emery GM% - old
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Emery sell - new
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Emery GM$ - new
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Emery GM% - new
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Emery GM$ diff
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Emery GM% diff
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // CRP - old
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Cust GM$ - old
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Cust GM% - old
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // CRP - new
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Cust GM$ - new
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Cust GM% - new
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Cust GM$ diff
      cell.setCellType(CellType.NUMERIC);

      cell = row.createCell(++col);  // Cust GM% diff
      cell.setCellType(CellType.NUMERIC);

      cell = null;

      return row;
   }

   /**
    * Builds the xls file for this report.
    *
    * @return boolean - true if the file was built sucessfully, false otherwise.
    */
   private boolean buildOutputFile() throws Exception
   {
      ResultSet customers;
      ResultSet items;
      ResultSet qty;
      String custId;
      String itemId;
      String repName;
      int skuCount;      // total item count
      int prcDecrCount;  // item count where new sell price is lower than the old one.
      int prcIncrCount;  // item count where new sell price is higher than the old one.
      int prcEqualCount; // item count where the prices are equal
      int retlPck; // item retail pack
      long qtyShip;
      double oldSell;
      double newSell;
      double oldRetl;
      double newRetl;
      double emCost;
      double totBuyOld;            // total emery buy - old
      double totSellOld;           // total cust sell - old
      double totSellNew;           // total cust sell - new
      double totRetlOld;           // total cust retail - old
      double totRetlNew;           // total cust retail - new
      double totEmMarginOld;       // total emery margin - old
      double totEmMarginNew;       // total emery margin - new
      double totEmMarginOldPerc;   // total emery margin percentage - old
      double totEmMarginNewPerc;   // total emery margin percentage - new
      double emMarginDiff;         // difference between old and new emery margins
      double emMarginPercDiff;     // difference between old and new emery margin percs
      double totCustMarginOld;     // total customer margin - old
      double totCustMarginNew;     // total customer margin - new
      double totCustMarginOldPerc; // total customer margin percentage - old
      double totCustMarginNewPerc; // total customer margin percentage - new
      double custMarginDiff;       // difference between old and new cust margins
      double custMarginPercDiff;   // difference between old and new cust margin percs
      XSSFRow row = null;
      int rowIndex = 1;
      int rowTotal;
      int col;
      FileOutputStream outFile = null;
      StringBuffer tmpBuf = new StringBuffer();

      //
      // Build the report headings
      createCaptions();

      //
      // Get the list of active customers that don't have a price contract
      customers = m_Customers.executeQuery();

      //
      // Get the total # rows in customer result set. This is possible since the result
      // set is scrollable.
      customers.last();
      rowTotal = customers.getRow();
      customers.beforeFirst();

      //
      // Loop through each customer in the list
      while ( customers.next() && m_Status != RptServer.STOPPED ) {
         custId = customers.getString("customer_id");
         repName = customers.getString("rep");

         if ( repName == null )
            repName = "";

         //
         // Build progress message
         tmpBuf.setLength(0);
         tmpBuf.append("Building data for ");
         tmpBuf.append(m_RptProc.getUid());
         tmpBuf.append(" - Cust: ");
         tmpBuf.append(custId);
         tmpBuf.append(" - Row ");
         tmpBuf.append(Integer.toString(rowIndex));
         tmpBuf.append(" of ");
         tmpBuf.append(Integer.toString(rowTotal));

         setCurAction(tmpBuf.toString());

         //
         // Initialize counts
         skuCount = 0;
         prcDecrCount = 0;
         prcIncrCount = 0;
         prcEqualCount = 0;

         //
         // Initialize totals
         totBuyOld = 0.0;
         totSellOld = 0.0;
         totSellNew = 0.0;
         totRetlOld = 0.0;
         totRetlNew = 0.0;
         totEmMarginOld = 0.0;
         totEmMarginNew = 0.0;
         totEmMarginOldPerc = 0.0;
         totEmMarginNewPerc = 0.0;
         totCustMarginOld = 0.0;
         totCustMarginNew = 0.0;
         totCustMarginOldPerc = 0.0;
         totCustMarginNewPerc = 0.0;

         m_Items.setString(1, custId);
         m_Items.setString(2, m_NrhaId);
         items = m_Items.executeQuery();

         //
         // Loop through each item that this customer purchased.  Items retrieved are
         // based on a 2 year purchase history and a specific nrha department. Use this
         // loop to determine the sku- and price difference counts.
         while ( items.next() && m_Status != RptServer.STOPPED ) {
            itemId = items.getString("item_nbr");
            retlPck = items.getInt("retail_pack");
            qtyShip = 0;
            skuCount++;

            //
            // Get the old and new sell prices.  Determine if there's an increase,
            // decrease or if they are equal, and increment the appropriate count.
            oldSell = getOldCustSellPrice(custId, itemId);
            newSell = getNewCustSellPrice(custId, itemId);

            if ( newSell < oldSell )
               prcDecrCount++;
            else {
               if ( newSell > oldSell )
                  prcIncrCount++;
               else
                  prcEqualCount++;
            }

            //
            // Get the old and new retail prices.  Make sure this is multiplied by
            // the retail pack, as some items are sold in bulk.
            oldRetl = getOldCustRetlPrice(custId, itemId) * retlPck;
            newRetl = getNewCustRetlPrice(custId, itemId) * retlPck;

            //
            // Get the total qty shipped for this item and customer
            m_QtyShip.setString(1, itemId);
            m_QtyShip.setString(2, custId);
            m_QtyShip.setString(3, m_NrhaId);
            qty = m_QtyShip.executeQuery();

            if ( qty.next() )
               qtyShip = qty.getLong("qty_ship");

            //
            // Close the qty result set
            try {
               qty.close();
            }
            catch ( SQLException e ) {
            }

            //
            // Get the emery cost for the current item
            emCost = getEmeryCost(itemId);

            //
            // Update total cost and sell values
            totBuyOld = totBuyOld + (qtyShip * emCost);
            totSellOld = totSellOld + (qtyShip * oldSell);
            totSellNew = totSellNew + (qtyShip * newSell);

            //
            // Update total retail values
            totRetlOld = totRetlOld + (qtyShip * oldRetl);
            totRetlNew = totRetlNew + (qtyShip * newRetl);
         }

         //
         // Close the items result set
         try {
            items.close();
         }
         catch ( SQLException e ) {
         }

         //
         // Get emery margin totals.  Make sure percentages are rounded to 2 decimals.
         totEmMarginOld = totSellOld - totBuyOld;
         totEmMarginOld = Math.floor(totEmMarginOld * 100 + .5d) / 100;

         totEmMarginNew = totSellNew - totBuyOld;
         totEmMarginNew = Math.floor(totEmMarginNew * 100 + .5d) / 100;

         if ( totSellOld > 0.0 ) {
            totEmMarginOldPerc = (totEmMarginOld/totSellOld) * 100;
            totEmMarginOldPerc = Math.floor(totEmMarginOldPerc * 100 + .5d) / 100;
         }

         if ( totSellNew > 0.0 ) {
            totEmMarginNewPerc = (totEmMarginNew/totSellNew) * 100;
            totEmMarginNewPerc = Math.floor(totEmMarginNewPerc * 100 + .5d) / 100;
         }

         emMarginDiff = totEmMarginNew - totEmMarginOld;
         emMarginDiff = Math.floor(emMarginDiff * 100 + .5d) / 100;

         emMarginPercDiff = totEmMarginNewPerc - totEmMarginOldPerc;
         emMarginPercDiff = Math.floor(emMarginPercDiff * 100 + .5d) / 100;

         //
         // Get customer margin totals
         totCustMarginOld = totRetlOld - totSellOld;
         totCustMarginOld = Math.floor(totCustMarginOld * 100 + .5d) / 100;

         totCustMarginNew = totRetlNew - totSellNew;
         totCustMarginNew = Math.floor(totCustMarginNew * 100 + .5d) / 100;

         if ( totRetlOld > 0.0 ) {
            totCustMarginOldPerc = (totCustMarginOld/totRetlOld) * 100;
            totCustMarginOldPerc = Math.floor(totCustMarginOldPerc * 100 + .5d) / 100;
         }

         if ( totRetlNew > 0.0 ) {
            totCustMarginNewPerc = (totCustMarginNew/totRetlNew) * 100;
            totCustMarginNewPerc = Math.floor(totCustMarginNewPerc * 100 + .5d) / 100;
         }

         custMarginDiff = totCustMarginNew - totCustMarginOld;
         custMarginDiff = Math.floor(custMarginDiff * 100 + .5d) / 100;

         custMarginPercDiff = totCustMarginNewPerc - totCustMarginOldPerc;
         custMarginPercDiff = Math.floor(custMarginPercDiff * 100 + .5d) / 100;

         //
         // Add a new row
         row = addRow((m_StartRow - 1) + rowIndex);
         rowIndex++;
         col = -1;

         //
         // Set the cell values
         if ( m_RepId == -1 )
            row.getCell(++col).setCellValue(new XSSFRichTextString(repName)); // sales rep name

         row.getCell(++col).setCellValue(new XSSFRichTextString(custId)); // customer#
         row.getCell(++col).setCellValue(new XSSFRichTextString(customers.getString("name"))); // customer name
         row.getCell(++col).setCellValue(skuCount); // sku count
         row.getCell(++col).setCellValue(prcDecrCount); // # items decreased
         row.getCell(++col).setCellValue(prcIncrCount); // # items increased
         row.getCell(++col).setCellValue(prcEqualCount);// # items equal
         row.getCell(++col).setCellValue(totBuyOld); // emery cost - old
         row.getCell(++col).setCellValue(totSellOld); // emery sell - old
         row.getCell(++col).setCellValue(totEmMarginOld); // emery GM$ - old
         row.getCell(++col).setCellValue(totEmMarginOldPerc); // emery GM% - old
         row.getCell(++col).setCellValue(totSellNew); // emery sell - new
         row.getCell(++col).setCellValue(totEmMarginNew); // emery GM$ - new
         row.getCell(++col).setCellValue(totEmMarginNewPerc); // emery GM% - new
         row.getCell(++col).setCellValue(emMarginDiff); // emery GM$ diff
         row.getCell(++col).setCellValue(emMarginPercDiff); // emery GM% diff
         row.getCell(++col).setCellValue(totRetlOld); // CRP - old
         row.getCell(++col).setCellValue(totCustMarginOld); // cust GM$ - old
         row.getCell(++col).setCellValue(totCustMarginOldPerc); // cust GM% - old
         row.getCell(++col).setCellValue(totRetlNew); // CRP - new
         row.getCell(++col).setCellValue(totCustMarginNew); // cust GM$ - new
         row.getCell(++col).setCellValue(totCustMarginNewPerc); // cust GM% - new
         row.getCell(++col).setCellValue(custMarginDiff); // cust GM$ diff
         row.getCell(++col).setCellValue(custMarginPercDiff); // cust GM% diff
      }

      //
      // Close the customers result set
      try {
         customers.close();
      }
      catch ( SQLException e ) {
         log.error("exception", e);
      }

      if ( m_Status != RptServer.STOPPED ) {
         //
         // Set the time stamp for this file build.  The time stamp will be appended to
         // the end of the file name and will prevent concurrent writes to the same file
         m_FileNames.add(String.format(FILE_NAME, System.currentTimeMillis()));
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

         setCurAction("Writing final xls file: " + m_FileNames.get(0));

         //
         // Write the test output to disk
         try {
            m_WorkBook.write(outFile);
         }

         catch (IOException e) {
            log.error(e);
         }

         try {
            outFile.close();
         }
         catch ( Exception e ) {
         }
      }

      custId = null;
      repName = null;
      itemId = null;
      outFile = null;
      tmpBuf = null;

      return true;
   }

   /**
    * Handles any cleanup.  Closes any open statements and the connection wrapper object.
    */
   protected void cleanup()
   {
      m_NrhaId = null;
      m_WorkBook = null;
      m_Sheet = null;

      closeStatements();
   }

   /**
    * Closes any open statements.
    */
   private void closeStatements()
   {
      if ( m_Customers != null ) {
         try {
            m_Customers.close();
         }
         catch ( SQLException e ) {

         }

         m_Customers = null;
      }

      if ( m_Items != null ) {
         try {
            m_Items.close();
         }
         catch ( SQLException e) {

         }

         m_Items = null;
      }

      if ( m_QtyShip != null ) {
         try {
            m_QtyShip.close();
         }
         catch ( SQLException e ) {

         }

         m_QtyShip = null;
      }

      if ( m_SalesRep != null ) {
         try {
            m_SalesRep.close();
         }
         catch ( SQLException e ) {

         }

         m_SalesRep = null;
      }

      if ( m_OldCustSellPrice != null ) {
         try {
            m_OldCustSellPrice.close();
         }
         catch ( SQLException e ) {

         }

         m_OldCustSellPrice = null;
      }

      if ( m_NewCustSellPrice != null ) {
         try {
            m_NewCustSellPrice.close();
         }
         catch ( SQLException e ) {

         }

         m_NewCustSellPrice = null;
      }

      if ( m_OldCustRetlPrice != null ) {
         try {
            m_OldCustRetlPrice.close();
         }
         catch ( SQLException e ) {

         }

         m_OldCustRetlPrice = null;
      }

      if ( m_NewCustRetlPrice != null ) {
         try {
            m_NewCustRetlPrice.close();
         }
         catch ( SQLException e ) {

         }

         m_NewCustRetlPrice = null;
      }

      if ( m_OldBaseSellPrice != null ) {
         try {
            m_OldBaseSellPrice.close();
         }
         catch ( SQLException e ) {

         }

         m_OldBaseSellPrice = null;
      }

      if ( m_NewBaseSellPrice != null ) {
         try {
            m_NewBaseSellPrice.close();
         }
         catch ( SQLException e ) {

         }

         m_NewBaseSellPrice = null;
      }

      if ( m_EmeryCost != null ) {
         try {
            m_EmeryCost.close();
         }
         catch ( SQLException e ) {

         }

         m_EmeryCost = null;
      }
   }

   /**
    * Builds the captions on the worksheet.
    */
   private void createCaptions() throws SQLException
   {
      XSSFRow row = null;
      XSSFCell cell = null;
      XSSFFont font;
      XSSFCellStyle style;
      CellRangeAddress region;
      ResultSet salesRep = null;
      String repName = null;
      int colCnt = 23;
      int col = -1;
      int rw = -1;

      if ( m_RepId == -1 )
         colCnt++;

      font = m_WorkBook.createFont();
      font.setFontHeightInPoints((short) 15);
      font.setFontName("Arial");
      style = m_WorkBook.createCellStyle();
      style.setFont(font);

      //
      // Show the report title in larger font
      row = m_Sheet.createRow(++rw);
      row.setHeightInPoints(ROW_HEIGHT + 5);
      cell = row.createCell( 0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(style);
      cell.setCellValue(new XSSFRichTextString("Dealer Report Summary"));
      cell = row.createCell( 1);
      cell.setCellType(CellType.STRING);
      cell = row.createCell( 2);
      cell.setCellType(CellType.STRING);

      region = new CellRangeAddress(0,  0, 0,  2);
      m_Sheet.addMergedRegion(region);

      //
      // Show the current date
      row = m_Sheet.createRow(++rw);
      row.setHeightInPoints(ROW_HEIGHT);
      cell = row.createCell( 0);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(
            new SimpleDateFormat("MM/dd/yyyy").format(new java.util.Date())
         )
      );

      //
      // If filtering by sales rep, show his/her name once before the column headings
      if ( m_RepId != -1 ) {
         salesRep = m_SalesRep.executeQuery();

         if ( salesRep.next() )
            repName = salesRep.getString("rep");

         if ( repName == null )
            repName = "";

         row = m_Sheet.createRow(++rw);
         row.setHeightInPoints(ROW_HEIGHT);
         cell = row.createCell( 0);
         cell.setCellType(CellType.STRING);
         cell.setCellValue(new XSSFRichTextString("TM: " + repName));

         try {
            salesRep.close();
         }
         catch ( SQLException e ) {
            log.error(e);
         }
      }

      //
      // Add two more rows
      row = m_Sheet.createRow(++rw);
      row.setHeightInPoints(ROW_HEIGHT);
      row = m_Sheet.createRow(++rw);
      row.setHeightInPoints(ROW_HEIGHT);

      m_StartRow = ++rw;

      //
      // Build the column headings

      for ( int i = 0; i < colCnt; i++ ) {
         cell = row.createCell(i);
         cell.setCellType(CellType.STRING);
      }

      if ( m_RepId == -1 ) {
         row.getCell(++col).setCellValue(new XSSFRichTextString("TM"));
         m_Sheet.setColumnWidth(col, 4266);
      }

      row.getCell(++col).setCellValue(new XSSFRichTextString("Customer #"));
      m_Sheet.setColumnWidth(col, 3000);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Customer Name"));
      m_Sheet.setColumnWidth(col, 8533);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Sku Count"));

      row.getCell(++col).setCellValue(new XSSFRichTextString("# Items Decreased"));
      m_Sheet.setColumnWidth(col, 4266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("# Items Increased"));
      m_Sheet.setColumnWidth(col, 4266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("# Items Equal"));
      m_Sheet.setColumnWidth(col, 3500);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Emery Cost - Old"));
      m_Sheet.setColumnWidth(col, 4266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Emery Sell - Old"));
      m_Sheet.setColumnWidth(col, 4266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Emery GM$ - Old"));
      m_Sheet.setColumnWidth(col, 4266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Emery GM% - Old"));
      m_Sheet.setColumnWidth(col, 4266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Emery Sell - New"));
      m_Sheet.setColumnWidth(col, 4266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Emery GM$ - New"));
      m_Sheet.setColumnWidth(col, 4266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Emery GM% - New"));
      m_Sheet.setColumnWidth(col, 4266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Emery GM$ Difference"));
      m_Sheet.setColumnWidth(col, 5266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Emery GM% Difference"));
      m_Sheet.setColumnWidth(col, 5266);

      row.getCell(++col).setCellValue(new XSSFRichTextString("CRP - Old"));
      m_Sheet.setColumnWidth(col, 3500);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Customer GM$ - Old"));
      m_Sheet.setColumnWidth(col, 4866);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Customer GM% - Old"));
      m_Sheet.setColumnWidth(col, 4866);

      row.getCell(++col).setCellValue(new XSSFRichTextString("CRP - New"));
      m_Sheet.setColumnWidth(col, 3500);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Customer GM$ - New"));
      m_Sheet.setColumnWidth(col, 4866);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Customer GM% - New"));
      m_Sheet.setColumnWidth(col, 4866);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Customer GM$ Difference"));
      m_Sheet.setColumnWidth(col, 5866);

      row.getCell(++col).setCellValue(new XSSFRichTextString("Customer GM% Difference"));
      m_Sheet.setColumnWidth(col, 5866);
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {
      boolean created = false;
      String fileName = null;
      String zipFile = null;

      try {
         m_OraConn = m_RptProc.getOraConn();

         if ( prepareStatements() ) {
            created = buildOutputFile();

            if ( created && m_Status != RptServer.STOPPED ) {
               //
               // Remove the report files from the application server, as these have been
               // moved to the ftp server.
               try {
                  fileName = m_FilePath + m_FileNames.get(0);
                  Files.delete(fileName);

                  if ( m_Zipped ) {
                     zipFile = fileName.substring(0, fileName.lastIndexOf('.')) + ".zip";
                     Files.delete(zipFile);
                  }
               }

               catch ( Exception ex ) {
                  log.error("exception: " + ex);
               }
            }
         }
      }

      catch ( Exception ex ) {
         log.fatal("exception:", ex);
      }

      finally {
        closeStatements();
      }

      return created;
   }

   /**
    * Gets the current customer specific sell price from pr03, else gets the current base
    * price if an exception occurred.  If an exception occurs getting the base, it just
    * returns zero.
    *
    * @param custId String - the input customer identifier.
    * @param itemId String - the input item identifier.
    * @return double - the current customer specific sell price.
    */
   public double getOldCustSellPrice(String custId, String itemId)
   {
      double sell = 0.0;

      try {
         m_OldCustSellPrice.setString(2, custId);
         m_OldCustSellPrice.setString(3, itemId);
         m_OldCustSellPrice.execute();

         sell = m_OldCustSellPrice.getDouble(1);
      }
      catch ( SQLException e1 ) {
         try {
            //
            // Use the base price if an exception occurred
            m_OldBaseSellPrice.setString(2, itemId);
            m_OldBaseSellPrice.execute();

            sell = m_OldBaseSellPrice.getDouble(1);
         }
         catch ( SQLException e2 ) {
            sell = 0.0;
         }
      }

      return sell;
   }

   /**
    * Gets the current customer specific retail price from pr03, else returns zero if some
    * exception occurred.
    *
    * @param custId String - the input customer identifier.
    * @param itemId String - the input item identifier.
    * @return double - the current customer specific retail price.
    */
   public double getOldCustRetlPrice(String custId, String itemId)
   {
      double retl;

      try {
         m_OldCustRetlPrice.setString(2, custId);
         m_OldCustRetlPrice.setString(3, itemId);
         m_OldCustRetlPrice.execute();

         retl = m_OldCustRetlPrice.getDouble(1);
      }
      catch ( SQLException e ) {
         retl = 0.0;
      }

      return retl;
   }

   /**
    * Gets the new customer specific sell price from ts01, else gets the new base price if an
    * exception occurred.  If an exception occurs getting the base, it just returns zero.
    *
    * @param custId String - the input customer identifier.
    * @param itemId String - the input item identifier.
    * @return double - the new customer specific sell price.
    */
   public double getNewCustSellPrice(String custId, String itemId)
   {
      double sell = 0.0;

      try {
         m_NewCustSellPrice.setString(2, custId);
         m_NewCustSellPrice.setString(3, itemId);
         m_NewCustSellPrice.execute();

         sell = m_NewCustSellPrice.getDouble(1);
      }
      catch ( SQLException e1 ) {
         try {
            //
            // Use the base price if an exception occurred
            m_NewBaseSellPrice.setString(2, itemId);
            m_NewBaseSellPrice.execute();

            sell = m_NewBaseSellPrice.getDouble(1);
         }
         catch ( SQLException e2 ) {
            sell = 0.0;
         }
      }

      return sell;
   }

   /**
    * Gets the input item's current emery cost from pr03.
    *
    * @param itemId String - the input item identifier.
    * @return double - the emery cost of the input item.
    */
   public double getEmeryCost(String itemId)
   {
      double cost = 0.0;

      try {
         m_EmeryCost.setString(2, itemId);
         m_EmeryCost.execute();

         cost = m_EmeryCost.getDouble(1);
      }
      catch ( SQLException e ) {
         log.error("exception: " + e.getMessage());
      }

      return cost;
   }

   /**
    * Gets the new customer specific retail price from ts01, else returns zero if some
    * exception occurred.
    *
    * @param custId String - the input customer identifier.
    * @param itemId String - the input item identifier.
    * @return double - the new customer specific retail price.
    */
   public double getNewCustRetlPrice(String custId, String itemId)
   {
      double retl;

      try {
         m_NewCustRetlPrice.setString(2, custId);
         m_NewCustRetlPrice.setString(3, itemId);
         m_NewCustRetlPrice.execute();

         retl = m_NewCustRetlPrice.getDouble(1);
      }
      catch ( SQLException e ) {
         retl = 0.0;
      }

      return retl;
   }

   /**
    * Gets the repid of the territory manager for this report.
    *
    * @return long - the emery rep id of the territory manager.
    */
   public int getRepId()
   {
      return m_RepId;
   }

   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {
      boolean isPrepared = false;
      StringBuffer sql = new StringBuffer();

      if ( m_OraConn != null ) {
         try {
            //
            // Gets list of active customers that are not using contract pricing.  Optionally
            // filtered by emery rep id.
            sql.append("select distinct customer.customer_id, name, first||' '||last as rep ");
            sql.append("from customer, customer_status, cust_rep, emery_rep, cust_market_view ");
            sql.append("where customer.cust_status_id = customer_status.cust_status_id and ");
            sql.append("customer_status.description = 'ACTIVE' and ");

            if ( m_RepId != -1 )
               sql.append("cust_rep.er_id = " + m_RepId + " and ");

            sql.append("customer.customer_id = cust_rep.customer_id(+) and ");
            sql.append("cust_rep.er_id = emery_rep.er_id(+) and ");
            sql.append("customer.customer_id = cust_market_view.customer_id and ");
            sql.append("cust_market_view.market = 'CUSTOMER TYPE' and ");
            sql.append("cust_market_view.class not in ('BACKHAULS', 'EMPLOYEE', 'NAT ACCT', 'EMERY') and ");
            sql.append("customer.customer_id not in ( ");
            sql.append("select cpm.customer_id ");
            sql.append("from cust_price_method cpm, price_method pm ");
            sql.append("where cpm.price_method_id = pm.price_method_id and ");
            sql.append("pm.description = 'CONTRACT') ");
            sql.append("order by customer_id");

            //
            // Make customer resultset scrollable so can get total # rows
            m_Customers = m_OraConn.prepareStatement(sql.toString(), ResultSet.TYPE_SCROLL_INSENSITIVE,
               ResultSet.CONCUR_READ_ONLY);

            sql.setLength(0);

            //
            // Gets list of distinct items based on a 2 year purchase history for a particular
            // customer and nrha department.  This will be used in getting the sku- and price
            // difference counts.
            sql.append("select distinct item_nbr, item.retail_pack ");
            sql.append("from inv_dtl, item, flc, mdc ");
            sql.append("where cust_nbr = ? and ");
            sql.append("sale_type = 'WAREHOUSE' and ");
            sql.append("mdc.nrha_id = ? and ");
            sql.append("inv_dtl.item_nbr = item.item_id and ");
            sql.append("invoice_date > add_months(sysdate, -24) and ");
            sql.append("invoice_date <= sysdate and ");
            sql.append("item.flc_id = flc.flc_id and ");
            sql.append("flc.mdc_id = mdc.mdc_id");

            m_Items = m_OraConn.prepareStatement(sql.toString());

            sql.setLength(0);

            //
            // Gets the item total qty purchased for the time period, customer and nrha dept.
            sql.append("select sum(qty_shipped) as qty_ship ");
            sql.append("from inv_dtl, item, flc, mdc ");
            sql.append("where item_id = ? and ");
            sql.append("cust_nbr = ? and ");
            sql.append("sale_type = 'WAREHOUSE' and ");
            sql.append("mdc.nrha_id = ? and ");
            sql.append("inv_dtl.item_nbr = item.item_id and ");
            sql.append("invoice_date > add_months(sysdate, -24) and ");
            sql.append("invoice_date <= sysdate and ");
            sql.append("item.flc_id = flc.flc_id and ");
            sql.append("flc.mdc_id = mdc.mdc_id");

            m_QtyShip = m_OraConn.prepareStatement(sql.toString());

            sql.setLength(0);

            sql.append("select first||' '||last as rep ");
            sql.append("from emery_rep ");
            sql.append("where er_id = " + m_RepId);

            m_SalesRep = m_OraConn.prepareStatement(sql.toString());

            //
            // Gets the old customer specific sell price.  Pulled from pr03.
            m_OldCustSellPrice = m_OraConn.prepareCall(
               "begin ? := cust_procs.getsellprice(?, ?); end;"
            );
            m_OldCustSellPrice.registerOutParameter(1, Types.DOUBLE);

            //
            // Gets the new customer specific sell price.  Pulled from ts01.
            m_NewCustSellPrice = m_OraConn.prepareCall(
               "begin ? := cust_procs.getsellprice@ts01(?, ?); end;"
            );
            m_NewCustSellPrice.registerOutParameter(1, Types.DOUBLE);

            //
            // Gets the old customer specific retail price.  Pulled from pr03.
            m_OldCustRetlPrice = m_OraConn.prepareCall(
               "begin ? := round(cust_procs.getretailprice(?, ?), 2); end;"
            );
            m_OldCustRetlPrice.registerOutParameter(1, Types.DOUBLE);

            //
            // Gets the new customer specific retail price.  Pulled from ts01.
            m_NewCustRetlPrice = m_OraConn.prepareCall(
               "begin ? := round(cust_procs.getretailprice@ts01(?, ?), 2); end;"
            );
            m_NewCustRetlPrice.registerOutParameter(1, Types.DOUBLE);

            //
            // Gets the old base sell price.  This is pulled from pr03.
            m_OldBaseSellPrice = m_OraConn.prepareCall(
               "begin ? := item_price_procs.todays_sell(?); end;"
            );
            m_OldBaseSellPrice.registerOutParameter(1, Types.DOUBLE);

            //
            // Gets the new base sell price.  This is pulled from ts01.
            m_NewBaseSellPrice = m_OraConn.prepareCall(
               "begin ? := item_price_procs.todays_sell@ts01(?); end;"
            );
            m_NewBaseSellPrice.registerOutParameter(1, Types.DOUBLE);

            //
            // Gets the current emery cost for the input item.  This is pulled from pr03.
            m_EmeryCost = m_OraConn.prepareCall(
               "begin ? := item_price_procs.todays_buy(?); end;"
            );
            m_EmeryCost.registerOutParameter(1, Types.DOUBLE);

            isPrepared = true;
         }

         catch ( Exception ex ) {

         }
      }

      sql = null;
      return isPrepared;
   }

   /**
    * Process the report parameters.
    *
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      //
      // Initialize parameters needed by the report
      m_NrhaId = params.get(0).value;
      m_RepId = Integer.parseInt(params.get(1).value);
      m_Zipped = Boolean.parseBoolean(params.get(2).value);
   }
}
