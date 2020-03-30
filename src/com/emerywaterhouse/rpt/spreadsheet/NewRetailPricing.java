/**
 * File: NewRetailPricing.java
 * Description: Report to assist territory managers in viewing default and CRP pricing for a specific customer.
 *    Based on a specific purchase history period that defaults to 2 years.  Totals by item, flc,
 *    nrha, and store.  Mainly used for comparing old versus new pricing data.
 *    <p>
 *    NOTE: New pricing data currently retrieved from ts01.<p>
 *     
 *    The original author is Paul Davidson.
 *
 * @author Paul Davidson
 * @author Jeffrey Fisher
 *
 * Create Data: 05/13/2005
 * Last Update: $Id: NewRetailPricing.java,v 1.9 2014/03/17 18:37:50 epearson Exp $
 * 
 * History
 *    $Log: NewRetailPricing.java,v $
 *    Revision 1.9  2014/03/17 18:37:50  epearson
 *    updated characters to UTF-8
 *
 *    Revision 1.8  2011/04/25 06:56:30  npasnur
 *    Changed the params order for method CellRangeAddress after POI upgrade
 *
 *    Revision 1.7  2009/02/18 17:17:50  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.6  2006/02/23 16:02:39  jfisher
 *    removed reference to logger and used the static logger in the report object.
 *
 *    Revision 1.5  2006/01/03 15:34:03  jfisher
 *    Fixed the type safety warning by adding a suppress tag
 *
 *    03/25/2005 - Added log4j logging. jcf
 * 
 *    09/30/2004 - removed uread variables. - jcf
 * 
 *    04/07/2004 - Applied Email class changes. - jcf
 *
 *    12/23/2003 - Modified the code to handle the new xml request format. - jcf
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Types;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.StringFormat;
import com.emerywaterhouse.websvc.Param;


public class NewRetailPricing extends Report
{
   private String m_CustId;
   private String m_NrhaId;   
   private int m_NumCust; // Number of customers to report on
   private int m_FileIndex; // Index indicating which file is being built
   private ArrayList<String> m_CustList;

   private PreparedStatement m_Items;
   private PreparedStatement m_CustAddr;
   private CallableStatement m_OldCustSellPrice;
   private CallableStatement m_NewCustSellPrice;
   private CallableStatement m_OldCustRetlPrice;
   private CallableStatement m_NewCustRetlPrice;
   private CallableStatement m_OldBaseRetails;
   private CallableStatement m_NewBaseRetails;
   private PreparedStatement m_CrpItem;
   private PreparedStatement m_CrpMkts;

   private HSSFWorkbook m_WorkBook;
   private HSSFSheet m_Sheet;
   private HSSFFont m_Font;

   private HSSFCellStyle m_StyleLeft;  // Horizontal left alignment
   private HSSFCellStyle m_StyleRght;  // Horizontal right alignment
   private HSSFCellStyle m_StyleCent;  // Horizontally centered

   private final int ROW_HEIGHT = 12;
   private final int REGION_BORDERED = 1;
   private final int REGION_BORDER_BOTTOM = 2;
   
   /**
    * default constructor
    */
   public NewRetailPricing()
   {
      super();
      
      m_MaxRunTime = 8;

      //
      // Create a new workbook. A new sheet will be created and removed for each report
      // since a separate report needs to be generated per customer.
      m_WorkBook = new HSSFWorkbook();

      //
      // Create the font for this workbook
      m_Font = m_WorkBook.createFont();
      m_Font.setFontHeightInPoints((short) 7);
      m_Font.setFontName("Arial");

      //
      // Setup the cell styles used in this report
      m_StyleLeft = m_WorkBook.createCellStyle();
      m_StyleLeft.setFont(m_Font);
      m_StyleLeft.setAlignment(HSSFCellStyle.ALIGN_LEFT);

      m_StyleRght = m_WorkBook.createCellStyle();
      m_StyleRght.setFont(m_Font);
      m_StyleRght.setAlignment(HSSFCellStyle.ALIGN_RIGHT);

      m_StyleCent = m_WorkBook.createCellStyle();
      m_StyleCent.setFont(m_Font);
      m_StyleCent.setAlignment(HSSFCellStyle.ALIGN_CENTER);
   }
   
   
   /**
    * Convenience method that adds a new String type cell with no borders and the specified alignment.
    *
    * @param rowNum int - the row index.
    * @param colNum int - the column index.
    * @param val String - the cell value.
    * @param align int - the horizontal cell alignment.
    *
    * @return HSSFCell - the newly added String type cell, or a reference to the existing one.
    */
   private HSSFCell addCell(int rowNum, int colNum, String val, int align)
   {
      HSSFCell cell = addCell(rowNum, colNum, HSSFCellStyle.BORDER_NONE, HSSFCellStyle.BORDER_NONE,
         HSSFCellStyle.BORDER_NONE, HSSFCellStyle.BORDER_NONE, align);

      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellValue(new HSSFRichTextString(val));

      return cell;
   }

   /**
    * Convenience method that adds a new numeric type cell with no borders and the specified alignment.
    *
    * @param rowNum int - the row index.
    * @param colNum int - the column index.
    * @param val double - the cell value.
    * @param align int - the horizontal cell alignment.
    *
    * @return HSSFCell - the newly added numeric type cell, or a reference to the existing one.
    */
   private HSSFCell addCell(int rowNum, int colNum, double val, int align)
   {
      HSSFCell cell = addCell(rowNum, colNum, HSSFCellStyle.BORDER_NONE, HSSFCellStyle.BORDER_NONE,
         HSSFCellStyle.BORDER_NONE, HSSFCellStyle.BORDER_NONE, align);

      cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
      cell.setCellValue(val);

      return cell;
   }

   /**
    * Adds a new cell with the specified borders and horizontal alignment.
    *
    * @param rowNum int - the row index.
    * @param colNum int - the column index.
    * @param bordTop int - top border constant.
    * @param bordLeft int - left border constant.
    * @param bordBottom int - bottom border constant.
    * @param bordRight int - right border constant.
    * @param align int - the horizontal cell alignment.
    *
    * @return HSSFCell - the newly added cell, or a reference to the existing one.
    */
   private HSSFCell addCell(int rowNum, int colNum, int bordTop, int bordLeft, int bordBottom, int bordRight, int align)
   {
      HSSFRow row = addRow(rowNum);
      HSSFCell cell = row.getCell(colNum);

      if ( cell == null )
         cell = row.createCell(colNum);

      switch ( align ) {
         case HSSFCellStyle.ALIGN_LEFT:
            cell.setCellStyle(m_StyleLeft);
            break;

         case HSSFCellStyle.ALIGN_RIGHT:
            cell.setCellStyle(m_StyleRght);
            break;

         case HSSFCellStyle.ALIGN_CENTER:
            cell.setCellStyle(m_StyleCent);
            break;

         default:
            cell.setCellStyle(m_StyleLeft);
            break;
      }

      //
      // Set up the cell borders
      /*
      style.setBorderTop(bordTop);
      style.setBorderLeft(bordLeft);
      style.setBorderBottom(bordBottom);
      style.setBorderRight(bordRight);
      */

      row = null;

      return cell;
   }

   /**
    * Convenience method that adds a completely bordered cell with the specified alignment.
    *
    * @param rowNum int - the row index.
    * @param colNum int - the column index.
    * @param align int - the horizontal cell alignment.
    * @return HSSFCell - the newly added cell, or a reference to the existing one.
    */
   private HSSFCell addBordCell(int rowNum, int colNum, int align)
   {
      return addCell(rowNum, colNum, HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_THIN,
         HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_THIN, align);
   }

   /**
    * Convenience method that adds a merged region with the given value and border style.  This method assumes
    * that the region will only have 1 row, and 1 or more cells.  Also assumes a string type value.
    *
    * @param rowNum int - start from this row.
    * @param colFrom int - start from this column.
    * @param colTo int - to this column.
    * @param value String - region cell value.
    * @param borderStyle int - region border style.  See region border constants.
    * @param merge boolean - if true then don't merge it.  This is here because of POI bug#16362. See bugzilla.
    */
   private void addRegion(int rowNum, int colFrom, int colTo, String value, int borderStyle, boolean merge)
   {
      HSSFCell cell = null;

      addRow(rowNum);

      cell = addBordCell(rowNum, colFrom, HSSFCellStyle.ALIGN_LEFT);
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellValue(new HSSFRichTextString(value));

      if ( merge ) {
         CellRangeAddress region = new CellRangeAddress(rowNum, rowNum, colFrom, colTo);
         m_Sheet.addMergedRegion(region);
      }

      //
      // Have to construct borders around the merged region manually, since this is not yet
      // supported by the POI api.  Commented out, but may add this later.
      /*
      for ( int i = colFrom; i <= colTo; i++ ) {
         cell = addCell(rowNum, i, HSSFCellStyle.ALIGN_LEFT);

         //
         // Set up the region border
         switch ( borderStyle ) {
            case REGION_BORDERED:
               if ( i == colFrom ) {
                  setCellBorders(rowNum, i, HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_THIN,
                     HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_NONE);
               }
               else {
                  if ( i == colTo ) {
                     setCellBorders(rowNum, i, HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_NONE,
                        HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_THIN);
                  }
                  else {
                     setCellBorders(rowNum, i, HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_NONE,
                        HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_NONE);
                  }
               }
               break;

            case REGION_BORDER_HORZ:
               setCellBorders(rowNum, i, HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_NONE,
                  HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_NONE);
               break;

            case REGION_BORDER_BOTTOM:
               setCellBorders(rowNum, i, HSSFCellStyle.BORDER_NONE, HSSFCellStyle.BORDER_NONE,
                  HSSFCellStyle.BORDER_THIN, HSSFCellStyle.BORDER_NONE);
               break;

            case REGION_NOBORDER:
               break;
         }
      }
      */
   }

   /**
    * Adds a new row or returns the existing one.
    *
    * @param rowNum int - the row index.
    * @return HSSFRow - the row object added, or a reference to the existing one.
    */
   private HSSFRow addRow(int rowNum)
   {
      HSSFRow row = m_Sheet.getRow(rowNum);

      if ( row == null )
         row = m_Sheet.createRow( rowNum);

      //
      // All rows will have this height
      row.setHeightInPoints(ROW_HEIGHT);

      return row;
   }
   
   /**
    * Builds the xls file for this report.  This is a long method that could be improved upon,
    * however this is also partly due to the complex nature of the report.
    *
    * @return boolean - true if the file was built sucessfully, false otherwise.
    */
   private boolean buildOutputFile() throws Exception
   {
      String name = null;
      String addr1 = null;
      String addr2 = null;
      String city = null;
      String state = null;
      String zip = null;
      String itm = null;
      ResultSet custAddr;
      ResultSet items;
      ResultSet crpMkts;
      ResultSet crpOpt;
      String storeCrpMkt = "C"; // store crp market
      String imageCrpMkt = "C"; // image crp market
      String crpLvl = "";  // crp level: S, D, F, I, M
      String crpMkt = "";  // crp market: A, B, C, D
      //String crpTyp = "";  // crp type: market, margin, variance, price
      FileOutputStream outFile = null;
      long qtyShip;
      int retlPck;

      double oldSell;  // old customer cost
      double newSell;  // new customer cost
      double oldRetl;  // old customer retail
      double newRetl;  // new customer retail

      double oldRetlA; // old A-mkt retail
      double oldRetlB; // old B-mkt retail
      double oldRetlC; // old C-mkt retail
      double oldRetlD; // old D-mkt retail
      double newRetlA; // new A-mkt retail
      double newRetlB; // new B-mkt retail
      double newRetlC; // new C-mkt retail
      double newRetlD; // new D-mkt retail

      double oldRetlMgn;  // old retail margin
      double oldRetlAMgn; // old A-mkt margin
      double oldRetlBMgn; // old B-mkt margin
      double oldRetlCMgn; // old C-mkt margin
      double oldRetlDMgn; // old D-mkt margin
      double oldRetlMgnPerc;  // old retail margin percentage
      double oldRetlAMgnPerc; // old A-mkt margin percentage
      double oldRetlBMgnPerc; // old B-mkt margin percentage
      double oldRetlCMgnPerc; // old C-mkt margin percentage
      double oldRetlDMgnPerc; // old D-mkt margin percentage

      double newRetlMgn;  // new retail margin
      double newRetlAMgn; // new A-mkt margin
      double newRetlBMgn; // new B-mkt margin
      double newRetlCMgn; // new C-mkt margin
      double newRetlDMgn; // new D-mkt margin
      double newRetlMgnPerc;  // new retail margin percentage
      double newRetlAMgnPerc; // new A-mkt margin percentage
      double newRetlBMgnPerc; // new B-mkt margin percentage
      double newRetlCMgnPerc; // new C-mkt margin percentage
      double newRetlDMgnPerc; // new D-mkt margin percentage

      double oldFlcSellTot = 0.0; // flc old ext sell total
      double oldFlcRetlTot = 0.0; // flc old ext retail total
      double oldFlcAMktTot = 0.0; // flc old A-mkt ext retail total
      double oldFlcBMktTot = 0.0; // flc old B-mkt ext retail total
      double oldFlcCMktTot = 0.0; // flc old C-mkt ext retail total
      double oldFlcDMktTot = 0.0; // flc old D-mkt ext retail total

      double newFlcSellTot = 0.0; // flc new ext sell total
      double newFlcRetlTot = 0.0; // flc new ext retail total
      double newFlcAMktTot = 0.0; // flc new A-mkt ext retail total
      double newFlcBMktTot = 0.0; // flc new B-mkt ext retail total
      double newFlcCMktTot = 0.0; // flc new C-mkt ext retail total
      double newFlcDMktTot = 0.0; // flc new D-mkt ext retail total

      double oldSellTot = 0.0; // old ext sell total
      double oldRetlTot = 0.0; // old ext retail total
      double oldAMktTot = 0.0; // old A-mkt ext retail total
      double oldBMktTot = 0.0; // old B-mkt ext retail total
      double oldCMktTot = 0.0; // old C-mkt ext retail total
      double oldDMktTot = 0.0; // old D-mkt ext retail total

      double newSellTot = 0.0; // new ext sell total
      double newRetlTot = 0.0; // new ext retail total
      double newAMktTot = 0.0; // new A-mkt ext retail total
      double newBMktTot = 0.0; // new B-mkt ext retail total
      double newCMktTot = 0.0; // new C-mkt ext retail total
      double newDMktTot = 0.0; // new D-mkt ext retail total

      double gmPerc = 0.0;
      double gm = 0.0;

      int row = -1;
      int rowIndex;
      int rowTotal;
      StringBuffer prgBuf = new StringBuffer();
      String prevFlc = null; // previous flc id
      String prevFlcDesc = null; // previous flc description
      String nrhaDesc = "";
      StringBuffer fileName = new StringBuffer();

      //
      // Execute customer address query
      m_CustAddr.setString(1, m_CustId);
      custAddr = m_CustAddr.executeQuery();

      //
      // Get customer address info
      if ( custAddr.next() ) {
         name = custAddr.getString(1);
         addr1 = custAddr.getString(2);
         addr2 = custAddr.getString(3);
         city = custAddr.getString(4);
         state = custAddr.getString(5);
         zip = custAddr.getString(6);
      }

      //
      // Close customer address result set
      try {
         custAddr.close();
      }
      catch ( SQLException e ) {
      }

      //
      // Write customer information in top-left corner
      addRegion(++row,  0,  3, name, REGION_BORDER_BOTTOM, true);
      addRegion(++row,  0,  3, addr2, REGION_BORDER_BOTTOM, true);
      addRegion(++row,  0,  3, addr1, REGION_BORDER_BOTTOM, true);
      addRegion(++row,  0,  3, city, REGION_BORDER_BOTTOM, true);
      m_Sheet.setColumnWidth( 0,  1200);
      m_Sheet.setColumnWidth( 1,  1000);
      m_Sheet.setColumnWidth( 2,  1000);
      m_Sheet.setColumnWidth( 3,  2000);

      //
      // Write customer state
      addCell(row,  4, state, HSSFCellStyle.ALIGN_LEFT);
      m_Sheet.setColumnWidth( 4,  1000);

      //
      // Write customer zip
      addCell(row,  5, zip, HSSFCellStyle.ALIGN_LEFT);
      m_Sheet.setColumnWidth( 5,  1500);

      //
      // Write customer# caption and value
      addCell(0,  6, "CUST. NO", HSSFCellStyle.ALIGN_LEFT);
      m_Sheet.setColumnWidth( 6,  2000);
      addCell(0,  7, m_CustId, HSSFCellStyle.ALIGN_LEFT);
      m_Sheet.setColumnWidth( 7,  2000);

      //
      // Execute crp markets statement
      m_CrpMkts.setString(1, m_CustId);
      crpMkts = m_CrpMkts.executeQuery();

      //
      // Get the store and image level crp market defaults
      if ( crpMkts.next() ) {
         if ( crpMkts.getString("row_type").equals("STORE") )
            storeCrpMkt = crpMkts.getString("mkt");
         else
            imageCrpMkt = crpMkts.getString("mkt");
      }

      //
      // Close crp markets resultset
      try {
         crpMkts.close();
      }
      catch ( SQLException e ) {
      }

      //
      // Write market default caption and value
      addCell(row - 1,  9, "CURRENT MARKET DEFAULT:", HSSFCellStyle.ALIGN_LEFT);
      addCell(row - 1,  10, storeCrpMkt, HSSFCellStyle.ALIGN_LEFT);
      m_Sheet.setColumnWidth( 9,  5000);
      m_Sheet.setColumnWidth( 10,  1000);

      //
      // Write image default caption and value
      addCell(row,  9, "CURRENT IMAGE DEFAULT:", HSSFCellStyle.ALIGN_LEFT);
      addCell(row,  10, imageCrpMkt, HSSFCellStyle.ALIGN_LEFT);

      //
      // Write global CRP options
      addRegion(0,  12,  16, "CRP OPTIONS", REGION_BORDER_BOTTOM, true);
      addCell(1,  12, "STORE", HSSFCellStyle.ALIGN_LEFT);
      addCell(1,  13, "S", HSSFCellStyle.ALIGN_LEFT);
      addCell(1,  15, "ITEM", HSSFCellStyle.ALIGN_LEFT);
      addCell(1,  16, "I", HSSFCellStyle.ALIGN_LEFT);
      addCell(2,  12, "DEPT", HSSFCellStyle.ALIGN_LEFT);
      addCell(2,  13, "D", HSSFCellStyle.ALIGN_LEFT);
      addCell(2,  15, "RPM", HSSFCellStyle.ALIGN_LEFT);
      addCell(2,  16, "R", HSSFCellStyle.ALIGN_LEFT);
      addCell(3,  12, "FLC", HSSFCellStyle.ALIGN_LEFT);
      addCell(3,  13, "F", HSSFCellStyle.ALIGN_LEFT);
      addCell(3,  15, "IMAGE", HSSFCellStyle.ALIGN_LEFT);
      addCell(3,  16, "M", HSSFCellStyle.ALIGN_LEFT);

      for ( int i = 12; i <= 17; i++)
         m_Sheet.setColumnWidth(i,  1700);

      m_Sheet.setColumnWidth( 18,  500);

      //
      // Write the instructions block
      addRegion(0,  19,  24, "INSTRUCTIONS:", REGION_BORDER_BOTTOM, true);
      addRegion(1,  19,  24, "", REGION_BORDER_BOTTOM, true);
      addRegion(2,  19,  24, "", REGION_BORDER_BOTTOM, true);
      addRegion(3,  19,  24, "", REGION_BORDER_BOTTOM, true);

      for ( int i = 19; i <= 24; i++)
         m_Sheet.setColumnWidth(i,  1700);

      //
      // Write current pricing caption
      addRegion(5,  12,  17, "CURRENT PRICING", REGION_BORDERED, true);

      //
      // Write new pricing caption
      addRegion(5,  19,  24, "NEW PRICING", REGION_BORDERED, true);

      //
      // Execute the main items query
      m_Items.setString(1, m_CustId);
      m_Items.setString(2, m_NrhaId);
      items = m_Items.executeQuery();

      //
      // Get the total # rows in the items result set. This is possible since the result
      // set is scrollable.
      items.last();
      rowTotal = items.getRow();
      items.beforeFirst();
      rowIndex = 1;

      //
      // Add some row spaces
      addRow(++row);
      addRow(++row);
      addRow(++row);

      createCaptions(row);

      //
      // Loop through each item and write the info
      while ( items.next() && m_Status != RptServer.STOPPED ) {
         row++;

         //
         // Build progress message
         prgBuf.setLength(0);
         prgBuf.append("Building file ");
         prgBuf.append(Integer.toString(m_FileIndex));
         prgBuf.append(" of ");
         prgBuf.append(Integer.toString(m_NumCust));
         prgBuf.append(" - Row ");
         prgBuf.append(Integer.toString(rowIndex));
         prgBuf.append(" of ");
         prgBuf.append(Integer.toString(rowTotal));
         prgBuf.append(" - for ");
         prgBuf.append(m_RptProc.getUid());

         setCurAction(prgBuf.toString());

         itm = items.getString("item_nbr");
         qtyShip = items.getLong("qty_ship");
         retlPck = items.getInt("retail_pack");
         nrhaDesc = items.getString("nrha_desc");
         oldSell = getOldCustSellPrice(m_CustId, itm);
         newSell = getNewCustSellPrice(m_CustId, itm);
         oldRetl = getOldCustRetlPrice(m_CustId, itm) * retlPck;
         newRetl = getNewCustRetlPrice(m_CustId, itm) * retlPck;

         //
         // Execute the crp item statement
         m_CrpItem.setString(1, m_CustId);
         m_CrpItem.setString(2, itm);
         crpOpt = m_CrpItem.executeQuery();

         if ( crpOpt.next() ) {
            crpLvl = crpOpt.getString("crp");
            crpMkt = crpOpt.getString("market_id");

            if ( crpMkt == null )
               crpMkt = "";            
         }
         else {
            crpLvl = "";
            crpMkt = "";            
         }

         try {
            crpOpt.close();
         }
         catch ( SQLException e) {
         }

         //
         // If there's a change in flc, show GM$ and GM% totals for the previous flc
         if ( prevFlc != null && !items.getString("flc_id").equals(prevFlc) ) {
            addRow(++row);

            addRegion(row,  4,  10, "TOTAL FOR FLC " + prevFlc + " " + prevFlcDesc, REGION_BORDERED, false);
            addRegion(row,  12,  13, "CURRENT COST", REGION_BORDERED, false);
            addCell(row,  14, "GM%", HSSFCellStyle.ALIGN_CENTER);
            addCell(row,  15, "GM$", HSSFCellStyle.ALIGN_CENTER);
            addRegion(row,  19,  20, "NEW COST", REGION_BORDERED, false);
            addCell(row,  21, "GM%", HSSFCellStyle.ALIGN_CENTER);
            addCell(row,  22, "GM$", HSSFCellStyle.ALIGN_CENTER);
            addRegion(row,  23,  24, "FLC CRP", REGION_BORDERED, false);

            addRow(++row);

            //
            // Write old retail caption and GM totals
            addRegion(row,  12,  13, "CURRENT RETAIL", REGION_BORDERED, false);
            if ( oldFlcRetlTot != 0 )
               gmPerc = (oldFlcRetlTot - oldFlcSellTot)/oldFlcRetlTot * 100;
            else
               gmPerc = 0;
            addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER); // GM%
            gm = oldFlcRetlTot - oldFlcSellTot;
            addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER); // GM$

            //
            // Write new retail caption and GM totals
            addRegion(row,  19,  20, "NEW RETAIL", REGION_BORDERED, false);
            if ( newFlcRetlTot != 0 )
               gmPerc = (newFlcRetlTot - newFlcSellTot)/newFlcRetlTot * 100;
            else
               gmPerc = 0;
            addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
            gm = newFlcRetlTot - newFlcSellTot;
            addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

            addRow(++row);

            //
            // Write old A-mkt retail caption and GM totals
            addRegion(row,  12,  13, "CURRENT A MKT", REGION_BORDERED, false);
            if ( oldFlcAMktTot != 0 )
               gmPerc = (oldFlcAMktTot - oldFlcSellTot)/oldFlcAMktTot * 100;
            else
               gmPerc = 0;
            addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
            gm = oldFlcAMktTot - oldFlcSellTot;
            addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

            //
            // Write new A-mkt retail caption and GM totals
            addRegion(row,  19,  20, "NEW A MKT", REGION_BORDERED, false);
            if ( newFlcAMktTot != 0 )
               gmPerc = (newFlcAMktTot - newFlcSellTot)/newFlcAMktTot * 100;
            else
               gmPerc = 0;
            addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
            gm = newFlcAMktTot - newFlcSellTot;
            addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

            addRow(++row);

            //
            // Write old B-mkt retail caption and GM totals
            addRegion(row,  12,  13, "CURRENT B MKT", REGION_BORDERED, false);
            if ( oldFlcBMktTot != 0 )
               gmPerc = (oldFlcBMktTot - oldFlcSellTot)/oldFlcBMktTot * 100;
            else
               gmPerc = 0;
            addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
            gm = oldFlcBMktTot - oldFlcSellTot;
            addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

            //
            // Write new B-mkt retail caption and GM totals
            addRegion(row,  19,  20, "NEW B MKT", REGION_BORDERED, false);
            if ( newFlcBMktTot != 0 )
               gmPerc = (newFlcBMktTot - newFlcSellTot)/newFlcBMktTot * 100;
            else
               gmPerc = 0;
            addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
            gm = newFlcBMktTot - newFlcSellTot;
            addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

            addRow(++row);

            //
            // Write old C-mkt retail caption and GM totals
            addRegion(row,  12,  13, "CURRENT C MKT", REGION_BORDERED, false);
            if ( oldFlcCMktTot != 0 )
               gmPerc = (oldFlcCMktTot - oldFlcSellTot)/oldFlcCMktTot * 100;
            else
               gmPerc = 0;
            addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
            gm = oldFlcCMktTot - oldFlcSellTot;
            addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

            //
            // Write new C-mkt retail caption and GM totals
            addRegion(row,  19,  20, "NEW C MKT", REGION_BORDERED, false);
            if ( newFlcCMktTot != 0)
               gmPerc = (newFlcCMktTot - newFlcSellTot)/newFlcCMktTot * 100;
            else
               gmPerc = 0;
            addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
            gm = newFlcCMktTot - newFlcSellTot;
            addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

            addRow(++row);

            //
            // Write old D-mkt retail caption and GM totals
            addRegion(row,  12,  13, "CURRENT D MKT", REGION_BORDERED, false);
            if ( oldFlcDMktTot != 0 )
               gmPerc = (oldFlcDMktTot - oldFlcSellTot)/oldFlcDMktTot * 100;
            else
               gmPerc = 0;
            addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
            gm = oldFlcDMktTot - oldFlcSellTot;
            addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

            //
            // Write new D-mkt retail caption and GM totals
            addRegion(row,  19,  20, "NEW D MKT", REGION_BORDERED, false);
            if ( newFlcDMktTot != 0 )
               gmPerc = (newFlcDMktTot - newFlcSellTot)/newFlcDMktTot * 100;
            else
               gmPerc = 0;
            addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
            gm = newFlcDMktTot - newFlcSellTot;
            addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

            //
            // Add some row spaces
            addRow(++row);
            addRow(++row);
            addRow(++row);

            //
            // Re-initialize sell and retail totals for the next flc
            oldFlcSellTot = 0.0;
            oldFlcRetlTot = 0.0;
            oldFlcAMktTot = 0.0;
            oldFlcBMktTot = 0.0;
            oldFlcCMktTot = 0.0;
            oldFlcDMktTot = 0.0;
            newFlcSellTot = 0.0;
            newFlcRetlTot = 0.0;
            newFlcAMktTot = 0.0;
            newFlcBMktTot = 0.0;
            newFlcCMktTot = 0.0;
            newFlcDMktTot = 0.0;
         }

         //
         // Initialize percentages
         oldRetlMgnPerc = 0.0;
         oldRetlAMgnPerc = 0.0;
         oldRetlBMgnPerc = 0.0;
         oldRetlCMgnPerc = 0.0;
         oldRetlDMgnPerc = 0.0;
         newRetlMgnPerc = 0.0;
         newRetlAMgnPerc = 0.0;
         newRetlBMgnPerc = 0.0;
         newRetlCMgnPerc = 0.0;
         newRetlDMgnPerc = 0.0;

         //
         // Write nrha value
         addCell(row,  0, items.getString("nrha_id"), HSSFCellStyle.ALIGN_CENTER);

         //
         // Write mdc value
         addCell(row,  1, items.getString("mdc_id"), HSSFCellStyle.ALIGN_CENTER);

         //
         // Write flc value
         addCell(row,  2, items.getString("flc_id"), HSSFCellStyle.ALIGN_CENTER);

         //
         // Write item#
         addCell(row,  3, itm, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write item description
         addRegion(row,  4,  10, items.getString("description"), REGION_BORDERED, false);

         //
         // Write qty shipped
         addCell(row,  11, qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old cost caption and value
         addRegion(row,  12,  13, "CURRENT COST", REGION_BORDERED, false);
         addCell(row,  14, oldSell, HSSFCellStyle.ALIGN_RIGHT);

         //
         // CRP@ caption (current pricing)
         addCell(row,  15, "CRP@", HSSFCellStyle.ALIGN_CENTER);

         //
         // GM% caption (current pricing)
         addCell(row,  16, "GM%", HSSFCellStyle.ALIGN_CENTER);

         //
         // GM$ caption (current pricing)
         addCell(row,  17, "GM$", HSSFCellStyle.ALIGN_CENTER);

         //
         // New cost caption and value
         addRegion(row,  19,  20, "NEW COST", REGION_BORDERED, false);
         addCell(row,  21, newSell, HSSFCellStyle.ALIGN_RIGHT);

         //
         // CRP@ caption (new pricing)
         addCell(row,  22, "CRP@", HSSFCellStyle.ALIGN_CENTER);

         //
         // GM% caption (new pricing)
         addCell(row,  23, "GM%", HSSFCellStyle.ALIGN_CENTER);

         //
         // GM$ caption (new pricing)
         addCell(row,  24, "GM$", HSSFCellStyle.ALIGN_CENTER);

         addRow(++row);

         //
         // Old customer retail caption and value
         addRegion(row,  12,  13, "CURRENT RETAIL", REGION_BORDERED, false);
         addCell(row,  14, oldRetl, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write crp level indicator.  If its market, write the level in the
         // appropriate retail market row.
         if ( crpMkt.equals("") )
            addCell(row,  15, crpLvl, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write old retail margin$ - make sure its rounded to 2 decimals
         oldRetlMgn = (oldRetl * qtyShip) - (oldSell *qtyShip);
         oldRetlMgn = Math.floor(oldRetlMgn * 100 + .5d) / 100;
         addCell(row,  17, oldRetlMgn, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old retail margin% - make sure its rounded to 2 decimals
         if ( oldRetl != 0 ) {
            oldRetlMgnPerc = (oldRetlMgn/(oldRetl * qtyShip)) * 100;
            oldRetlMgnPerc = Math.floor(oldRetlMgnPerc * 100 + .5d) / 100;
         }
         addCell(row,  16, oldRetlMgnPerc, HSSFCellStyle.ALIGN_RIGHT);

         //
         // New customer retail caption and value
         addRegion(row,  19,  20, "NEW RETAIL", REGION_BORDERED, false);
         addCell(row,  21, newRetl, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write crp level indicator.  If its market, write the level in the
         // appropriate retail market row.
         if ( crpMkt.equals("") )
            addCell(row,  22, crpLvl, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write new retail margin$ - make sure its rounded to 2 decimals
         newRetlMgn = (newRetl * qtyShip) - (newSell *qtyShip);
         newRetlMgn = Math.floor(newRetlMgn * 100 + .5d) / 100;
         addCell(row,  24, newRetlMgn, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new retail margin% - make sure its rounded to 2 decimals
         if ( newRetl != 0 ) {
            newRetlMgnPerc = (newRetlMgn/(newRetl * qtyShip)) * 100;
            newRetlMgnPerc = Math.floor(newRetlMgnPerc * 100 + .5d) / 100;
         }
         addCell(row,  23, newRetlMgnPerc, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old extended cost value (COGS)
         addCell(row,  27, oldSell * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old extended retail value
         addCell(row,  28, oldRetl * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new extended retail value
         addCell(row,  29, newRetl * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Get old base retails
         try {
            m_OldBaseRetails.setString(1, itm);
            m_OldBaseRetails.execute();

            oldRetlA = m_OldBaseRetails.getDouble(2) * retlPck;
            oldRetlB = m_OldBaseRetails.getDouble(3) * retlPck;
            oldRetlC = m_OldBaseRetails.getDouble(4) * retlPck;
            oldRetlD = m_OldBaseRetails.getDouble(5) * retlPck;
         }
         catch ( SQLException e ) {
            oldRetlA = 0.0;
            oldRetlB = 0.0;
            oldRetlC = 0.0;
            oldRetlD = 0.0;
         }

         //
         // Get NEW base retails
         try {
            m_NewBaseRetails.setString(1, itm);
            m_NewBaseRetails.execute();

            newRetlA = m_NewBaseRetails.getDouble(2) * retlPck;
            newRetlB = m_NewBaseRetails.getDouble(3) * retlPck;
            newRetlC = m_NewBaseRetails.getDouble(4) * retlPck;
            newRetlD = m_NewBaseRetails.getDouble(5) * retlPck;
         }
         catch ( SQLException e ) {
            newRetlA = 0.0;
            newRetlB = 0.0;
            newRetlC = 0.0;
            newRetlD = 0.0;
         }

         //
         // Accumulate global totals.  This is equivalent to the dept and store totals
         // since the report is for a specific nrha department.
         oldSellTot = oldSellTot + (oldSell * qtyShip);
         oldRetlTot = oldRetlTot + (oldRetl * qtyShip);
         oldAMktTot = oldAMktTot + (oldRetlA * qtyShip);
         oldBMktTot = oldBMktTot + (oldRetlB * qtyShip);
         oldCMktTot = oldCMktTot + (oldRetlC * qtyShip);
         oldDMktTot = oldDMktTot + (oldRetlD * qtyShip);
         newSellTot = newSellTot + (newSell * qtyShip);
         newRetlTot = newRetlTot + (newRetl * qtyShip);
         newAMktTot = newAMktTot + (newRetlA * qtyShip);
         newBMktTot = newBMktTot + (newRetlB * qtyShip);
         newCMktTot = newCMktTot + (newRetlC * qtyShip);
         newDMktTot = newDMktTot + (newRetlD * qtyShip);

         //
         // Accumulate flc sell and retail totals.  This has to be separate since
         // these totals are re-initialized for each unique flc.
         oldFlcSellTot = oldFlcSellTot + (oldSell * qtyShip);
         oldFlcRetlTot = oldFlcRetlTot + (oldRetl * qtyShip);
         oldFlcAMktTot = oldFlcAMktTot + (oldRetlA * qtyShip);
         oldFlcBMktTot = oldFlcBMktTot + (oldRetlB * qtyShip);
         oldFlcCMktTot = oldFlcCMktTot + (oldRetlC * qtyShip);
         oldFlcDMktTot = oldFlcDMktTot + (oldRetlD * qtyShip);
         newFlcSellTot = newFlcSellTot + (newSell * qtyShip);
         newFlcRetlTot = newFlcRetlTot + (newRetl * qtyShip);
         newFlcAMktTot = newFlcAMktTot + (newRetlA * qtyShip);
         newFlcBMktTot = newFlcBMktTot + (newRetlB * qtyShip);
         newFlcCMktTot = newFlcCMktTot + (newRetlC * qtyShip);
         newFlcDMktTot = newFlcDMktTot + (newRetlD * qtyShip);

         //
         // Add a new row for the A-mkt values
         addRow(++row);

         //
         // Write old A-mkt retail caption and value
         addRegion(row,  12,  13, "CURRENT A MKT", REGION_BORDERED, false);
         addCell(row,  14, oldRetlA, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write crp level indicator for A-mkt.
         if ( crpMkt.equals("A") )
            addCell(row,  15, crpLvl, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write old A-mkt retail margin$
         oldRetlAMgn = (oldRetlA * qtyShip) - (oldSell * qtyShip);
         oldRetlAMgn = Math.floor(oldRetlAMgn * 100 + .5d) / 100;
         addCell(row,  17, oldRetlAMgn, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old A-mkt retail margin%
         if ( oldRetlA != 0 ) {
            oldRetlAMgnPerc = (oldRetlAMgn/(oldRetlA * qtyShip)) * 100;
            oldRetlAMgnPerc = Math.floor(oldRetlAMgnPerc * 100 + .5d) / 100;
         }
         addCell(row,  16, oldRetlAMgnPerc, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new A-mkt retail caption and value
         addRegion(row,  19,  20, "NEW A MKT", REGION_BORDERED, false);
         addCell(row,  21, newRetlA, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write crp level indicator for A-mkt.
         if ( crpMkt.equals("A") )
            addCell(row,  22, crpLvl, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write new A-mkt retail margin$
         newRetlAMgn = (newRetlA * qtyShip) - (newSell * qtyShip);
         newRetlAMgn = Math.floor(newRetlAMgn * 100 + .5d) / 100;
         addCell(row,  24, newRetlAMgn, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new A-mkt retail margin%
         if ( newRetlA != 0 ) {
            newRetlAMgnPerc = (newRetlAMgn/(newRetlA * qtyShip)) * 100;
            newRetlAMgnPerc = Math.floor(newRetlAMgnPerc * 100 + .5d) / 100;
         }
         addCell(row,  23, newRetlAMgnPerc, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old extended A-mkt retail
         addCell(row,  28, oldRetlA * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new extended A-mkt retail
         addCell(row,  29, newRetlA * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Add a new row for the B-mkt values
         addRow(++row);

         //
         // Write old B-mkt retail caption and value
         addRegion(row,  12,  13, "CURRENT B MKT", REGION_BORDERED, false);
         addCell(row,  14, oldRetlB, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write crp level indicator for B-mkt.
         if ( crpMkt.equals("B") )
            addCell(row,  15, crpLvl, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write old B-mkt retail margin$
         oldRetlBMgn = (oldRetlB * qtyShip) - (oldSell * qtyShip);
         oldRetlBMgn = Math.floor(oldRetlBMgn * 100 + .5d) / 100;
         addCell(row,  17, oldRetlBMgn, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old B-mkt retail margin%
         if ( oldRetlB != 0 ) {
            oldRetlBMgnPerc = (oldRetlBMgn/(oldRetlB * qtyShip)) * 100;
            oldRetlBMgnPerc = Math.floor(oldRetlBMgnPerc * 100 + .5d) / 100;
         }
         addCell(row,  16, oldRetlBMgnPerc, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new B-mkt retail caption and value
         addRegion(row,  19,  20, "NEW B MKT", REGION_BORDERED, false);
         addCell(row,  21, newRetlB, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write crp level indicator for B-mkt.
         if ( crpMkt.equals("B") )
            addCell(row,  22, crpLvl, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write new B-mkt retail margin$
         newRetlBMgn = (newRetlB * qtyShip) - (newSell * qtyShip);
         newRetlBMgn = Math.floor(newRetlBMgn * 100 + .5d) / 100;
         addCell(row,  24, newRetlBMgn, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new B-mkt retail margin%
         if ( newRetlB != 0 ) {
            newRetlBMgnPerc = (newRetlBMgn/(newRetlB * qtyShip)) * 100;
            newRetlBMgnPerc = Math.floor(newRetlBMgnPerc * 100 + .5d) / 100;
         }
         addCell(row,  23, newRetlBMgnPerc, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old extended B-mkt retail
         addCell(row,  28, oldRetlB * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new extended B-mkt retail
         addCell(row,  29, newRetlB * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Add a new row for the C-mkt values
         addRow(++row);

         //
         // Write old C-mkt retail caption and value
         addRegion(row,  12,  13, "CURRENT C MKT", REGION_BORDERED, false);
         addCell(row,  14, oldRetlC, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write crp level indicator for C-mkt.
         if ( crpMkt.equals("C") )
            addCell(row,  15, crpLvl, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write old C-mkt retail margin$
         oldRetlCMgn = (oldRetlC * qtyShip) - (oldSell * qtyShip);
         oldRetlCMgn = Math.floor(oldRetlCMgn * 100 + .5d) / 100;
         addCell(row,  17, oldRetlCMgn, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old C-mkt retail margin%
         if ( oldRetlC != 0 ) {
            oldRetlCMgnPerc = (oldRetlCMgn/(oldRetlC * qtyShip)) * 100;
            oldRetlCMgnPerc = Math.floor(oldRetlCMgnPerc * 100 + .5d) / 100;
         }
         addCell(row,  16, oldRetlCMgnPerc, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new C-mkt retail caption and value
         addRegion(row,  19,  20, "NEW C MKT", REGION_BORDERED, false);
         addCell(row,  21, newRetlC, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write crp level indicator for C-mkt.
         if ( crpMkt.equals("C") )
            addCell(row,  22, crpLvl, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write new C-mkt retail margin$
         newRetlCMgn = (newRetlC * qtyShip) - (newSell * qtyShip);
         newRetlCMgn = Math.floor(newRetlCMgn * 100 + .5d) / 100;
         addCell(row,  24, newRetlCMgn, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new C-mkt retail margin%
         if ( newRetlC != 0 ) {
            newRetlCMgnPerc = (newRetlCMgn/(newRetlC * qtyShip)) * 100;
            newRetlCMgnPerc = Math.floor(newRetlCMgnPerc * 100 + .5d) / 100;
         }
         addCell(row,  23, newRetlCMgnPerc, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old extended C-mkt retail
         addCell(row,  28, oldRetlC * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new extended C-mkt retail
         addCell(row,  29, newRetlC * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Add a new row for the D-mkt values
         addRow(++row);

         //
         // Write old D-mkt retail caption and value
         addRegion(row,  12,  13, "CURRENT D MKT", REGION_BORDERED, false);
         addCell(row,  14, oldRetlD, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write crp level indicator for D-mkt.
         if ( crpMkt.equals("D") )
            addCell(row,  15, crpLvl, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write old D-mkt retail margin$
         oldRetlDMgn = (oldRetlD * qtyShip) - (oldSell * qtyShip);
         oldRetlDMgn = Math.floor(oldRetlDMgn * 100 + .5d) / 100;
         addCell(row,  17, oldRetlDMgn, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old D-mkt retail margin%
         if ( oldRetlD != 0 ) {
            oldRetlDMgnPerc = (oldRetlDMgn/(oldRetlD * qtyShip)) * 100;
            oldRetlDMgnPerc = Math.floor(oldRetlDMgnPerc * 100 + .5d) / 100;
         }
         addCell(row,  16, oldRetlDMgnPerc, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new D-mkt retail caption and value
         addRegion(row,  19,  20, "NEW D MKT", REGION_BORDERED, false);
         addCell(row,  21, newRetlD, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write crp level indicator for D-mkt.
         if ( crpMkt.equals("D") )
            addCell(row,  22, crpLvl, HSSFCellStyle.ALIGN_CENTER);

         //
         // Write new D-mkt retail margin$
         newRetlDMgn = (newRetlD * qtyShip) - (newSell * qtyShip);
         newRetlDMgn = Math.floor(newRetlDMgn * 100 + .5d) / 100;
         addCell(row,  24, newRetlDMgn, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new D-mkt retail margin%
         if ( newRetlD != 0 ) {
            newRetlDMgnPerc = (newRetlDMgn/(newRetlD * qtyShip)) * 100;
            newRetlDMgnPerc = Math.floor(newRetlDMgnPerc * 100 + .5d) / 100;
         }
         addCell(row,  23, newRetlDMgnPerc, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write old extended D-mkt retail
         addCell(row,  28, oldRetlD * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         //
         // Write new extended D-mkt retail
         addCell(row,  29, newRetlD * qtyShip, HSSFCellStyle.ALIGN_RIGHT);

         addRow(++row);

         //
         // Store the current flc at loop's end.  This will become the "previous"
         // flc in the next iteration of the loop.
         prevFlc = items.getString("flc_id");
         prevFlcDesc = items.getString("flc_desc");

         //
         // Increment the row index
         rowIndex++;
      }

      try {
         items.close();
      }
      catch ( SQLException e ) {
      }

      items = null;
      crpOpt = null;
      prgBuf.setLength(0);
      prgBuf = null;

      addRow(++row);
      addRow(++row);

      //
      // Add last flc total section
      if ( prevFlc != null) {
         //
         // Write flc total captions
         addRegion(row,  4,  10, "TOTAL FOR FLC " + prevFlc + " " + prevFlcDesc, REGION_BORDERED, false);
         addRegion(row,  12,  13, "CURRENT COST", REGION_BORDERED, false);
         addCell(row,  14, "GM%", HSSFCellStyle.ALIGN_CENTER);
         addCell(row,  15, "GM$", HSSFCellStyle.ALIGN_CENTER);
         addRegion(row,  19,  20, "NEW COST", REGION_BORDERED, false);
         addCell(row,  21, "GM%", HSSFCellStyle.ALIGN_CENTER);
         addCell(row,  22, "GM$", HSSFCellStyle.ALIGN_CENTER);
         addRegion(row,  23,  24, "FLC CRP", REGION_BORDERED, false);

         addRow(++row);

         //
         // Write old retail caption and GM totals (last flc)
         addRegion(row,  12,  13, "CURRENT RETAIL", REGION_BORDERED, false);
         if ( oldFlcRetlTot != 0 )
            gmPerc = (oldFlcRetlTot - oldFlcSellTot)/oldFlcRetlTot * 100;
         else
            gmPerc = 0;
         addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER); // GM%
         gm = oldFlcRetlTot - oldFlcSellTot;
         addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER); // GM$

         //
         // Write new retail caption and GM totals (last flc)
         addRegion(row,  19,  20, "NEW RETAIL", REGION_BORDERED, false);
         if ( newFlcRetlTot != 0 )
            gmPerc = (newFlcRetlTot - newFlcSellTot)/newFlcRetlTot * 100;
         else
            gmPerc = 0;
         addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
         gm = newFlcRetlTot - newFlcSellTot;
         addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

         addRow(++row);

         //
         // Write old A-mkt retail caption and GM totals (last flc)
         addRegion(row,  12,  13, "CURRENT A MKT", REGION_BORDERED, false);
         if ( oldFlcAMktTot != 0 )
            gmPerc = (oldFlcAMktTot - oldFlcSellTot)/oldFlcAMktTot * 100;
         else
            gmPerc = 0;
         addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
         gm = oldFlcAMktTot - oldFlcSellTot;
         addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

         //
         // Write new A-mkt retail caption and GM totals (last flc)
         addRegion(row,  19,  20, "NEW A MKT", REGION_BORDERED, false);
         if ( newFlcAMktTot != 0 )
            gmPerc = (newFlcAMktTot - newFlcSellTot)/newFlcAMktTot * 100;
         else
            gmPerc = 0;
         addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
         gm = newFlcAMktTot - newFlcSellTot;
         addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

         addRow(++row);

         //
         // Write old B-mkt retail caption and GM totals (last flc)
         addRegion(row,  12,  13, "CURRENT B MKT", REGION_BORDERED, false);
         if ( oldFlcBMktTot != 0 )
            gmPerc = (oldFlcBMktTot - oldFlcSellTot)/oldFlcBMktTot * 100;
         else
            gmPerc = 0;
         addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
         gm = oldFlcBMktTot - oldFlcSellTot;
         addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

         //
         // Write new B-mkt retail caption and GM totals (last flc)
         addRegion(row,  19,  20, "NEW B MKT", REGION_BORDERED, false);
         if ( newFlcBMktTot != 0 )
            gmPerc = (newFlcBMktTot - newFlcSellTot)/newFlcBMktTot * 100;
         else
            gmPerc = 0;
         addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
         gm = newFlcBMktTot - newFlcSellTot;
         addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

         addRow(++row);

         //
         // Write old C-mkt retail caption and GM totals (last flc)
         addRegion(row,  12,  13, "CURRENT C MKT", REGION_BORDERED, false);
         if ( oldFlcCMktTot != 0 )
            gmPerc = (oldFlcCMktTot - oldFlcSellTot)/oldFlcCMktTot * 100;
         else
            gmPerc = 0;
         addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
         gm = oldFlcCMktTot - oldFlcSellTot;
         addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

         //
         // Write new C-mkt retail caption and GM totals (last flc)
         addRegion(row,  19,  20, "NEW C MKT", REGION_BORDERED, false);
         if ( newFlcCMktTot != 0)
            gmPerc = (newFlcCMktTot - newFlcSellTot)/newFlcCMktTot * 100;
         else
            gmPerc = 0;
         addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
         gm = newFlcCMktTot - newFlcSellTot;
         addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

         addRow(++row);

         //
         // Write old D-mkt retail caption and GM totals (last flc)
         addRegion(row,  12,  13, "CURRENT D MKT", REGION_BORDERED, false);
         if ( oldFlcDMktTot != 0 )
            gmPerc = (oldFlcDMktTot - oldFlcSellTot)/oldFlcDMktTot * 100;
         else
            gmPerc = 0;
         addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
         gm = oldFlcDMktTot - oldFlcSellTot;
         addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

         //
         // Write new D-mkt retail caption and GM totals (last flc)
         addRegion(row,  19,  20, "NEW D MKT", REGION_BORDERED, false);
         if ( newFlcDMktTot != 0 )
            gmPerc = (newFlcDMktTot - newFlcSellTot)/newFlcDMktTot * 100;
         else
            gmPerc = 0;
         addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
         gm = newFlcDMktTot - newFlcSellTot;
         addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

         //
         // Add some row spaces
         addRow(++row);
         addRow(++row);
         addRow(++row);
      }

      //
      // Add the dept total captions
      addRegion(row,  4,  10, "TOTAL FOR DEPARTMENT " + m_NrhaId + " " + nrhaDesc, REGION_BORDERED, false);
      addRegion(row,  12,  13, "CURRENT COST", REGION_BORDERED, false);
      addCell(row,  14, "GM%", HSSFCellStyle.ALIGN_CENTER);
      addCell(row,  15, "GM$", HSSFCellStyle.ALIGN_CENTER);
      addRegion(row,  19,  20, "NEW COST", REGION_BORDERED, false);
      addCell(row,  21, "GM%", HSSFCellStyle.ALIGN_CENTER);
      addCell(row,  22, "GM$", HSSFCellStyle.ALIGN_CENTER);
      addRegion(row,  23,  24, "DEPT CRP", REGION_BORDERED, false);

      addRow(++row);

      //
      // Write old retail caption and GM totals (dept)
      addRegion(row,  12,  13, "CURRENT RETAIL", REGION_BORDERED, false);
      if ( oldRetlTot != 0 )
         gmPerc = (oldRetlTot - oldSellTot)/oldRetlTot * 100;
      else
         gmPerc = 0;
      addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER); // GM%
      gm = oldRetlTot - oldSellTot;
      addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER); // GM$

      //
      // Write new retail caption and GM totals (dept)
      addRegion(row,  19,  20, "NEW RETAIL", REGION_BORDERED, false);
      if ( newRetlTot != 0 )
         gmPerc = (newRetlTot - newSellTot)/newRetlTot * 100;
      else
         gmPerc = 0;
      addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
      gm = newRetlTot - newSellTot;
      addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

      addRow(++row);

      //
      // Write old A-mkt retail caption and GM totals (dept)
      addRegion(row,  12,  13, "CURRENT A MKT", REGION_BORDERED, false);
      if ( oldAMktTot != 0 )
         gmPerc = (oldAMktTot - oldSellTot)/oldAMktTot * 100;
      else
         gmPerc = 0;
      addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
      gm = oldAMktTot - oldSellTot;
      addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

      //
      // Write new A-mkt retail caption and GM totals (dept)
      addRegion(row,  19,  20, "NEW A MKT", REGION_BORDERED, false);
      if ( newAMktTot != 0 )
         gmPerc = (newAMktTot - newSellTot)/newAMktTot * 100;
      else
         gmPerc = 0;
      addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
      gm = newAMktTot - newSellTot;
      addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

      addRow(++row);

      //
      // Write old B-mkt retail caption and GM totals (dept)
      addRegion(row,  12,  13, "CURRENT B MKT", REGION_BORDERED, false);
      if ( oldBMktTot != 0 )
         gmPerc = (oldBMktTot - oldSellTot)/oldBMktTot * 100;
      else
         gmPerc = 0;
      addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
      gm = oldBMktTot - oldSellTot;
      addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

      //
      // Write new B-mkt retail caption and GM totals (dept)
      addRegion(row,  19,  20, "NEW B MKT", REGION_BORDERED, false);
      if ( newBMktTot != 0 )
         gmPerc = (newBMktTot - newSellTot)/newBMktTot * 100;
      else
         gmPerc = 0;
      addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
      gm = newBMktTot - newSellTot;
      addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

      addRow(++row);

      //
      // Write old C-mkt retail caption and GM totals (dept)
      addRegion(row,  12,  13, "CURRENT C MKT", REGION_BORDERED, false);
      if ( oldCMktTot != 0 )
         gmPerc = (oldCMktTot - oldSellTot)/oldCMktTot * 100;
      else
         gmPerc = 0;
      addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
      gm = oldCMktTot - oldSellTot;
      addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

      //
      // Write new C-mkt retail caption and GM totals (dept)
      addRegion(row,  19,  20, "NEW C MKT", REGION_BORDERED, false);
      if ( newCMktTot != 0 )
         gmPerc = (newCMktTot - newSellTot)/newCMktTot * 100;
      else
         gmPerc = 0;
      addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
      gm = newCMktTot - newSellTot;
      addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

      addRow(++row);

      //
      // Write old D-mkt retail caption and GM totals (dept)
      addRegion(row,  12,  13, "CURRENT D MKT", REGION_BORDERED, false);
      if ( oldDMktTot != 0)
         gmPerc = (oldDMktTot - oldSellTot)/oldDMktTot * 100;
      else
         gmPerc = 0;
      addCell(row,  14, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
      gm = oldDMktTot - oldSellTot;
      addCell(row,  15, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

      //
      // Write new D-mkt retail caption and GM totals (dept)
      addRegion(row,  19,  20, "NEW D MKT", REGION_BORDERED, false);
      if ( newDMktTot != 0 )
         gmPerc = (newDMktTot - newSellTot)/newDMktTot * 100;
      else
         gmPerc = 0;
      addCell(row,  21, round(gmPerc, 2), HSSFCellStyle.ALIGN_CENTER);
      gm = newDMktTot - newSellTot;
      addCell(row,  22, round(gm, 2), HSSFCellStyle.ALIGN_CENTER);

      if ( m_Status != RptServer.STOPPED ) {
         //
         // Build the report file name         
         fileName.append("newrps");
         fileName.append("_");
         fileName.append(m_CustId);
         fileName.append("_");
         fileName.append(Long.toString(System.currentTimeMillis()));
         fileName.append(".xls");

         m_FileNames.add(fileName.toString());
         outFile = new FileOutputStream(m_FilePath + fileName.toString(), false);
         
         //
         // Write the test output to disk
         try {
            m_WorkBook.write(outFile);
         }
         catch (IOException e) {
         }

         try {
            outFile.close();
         }
         catch ( Exception e ) {
         }
      }

      //
      // Nullify refs
      fileName = null;
      custAddr = null;
      crpMkts = null;
      name = null;
      addr1 = null;
      addr2 = null;
      city = null;
      state = null;
      zip = null;
      itm = null;
      storeCrpMkt = null;
      imageCrpMkt = null;
      crpLvl = null;
      crpMkt = null;      
      prevFlc = null;
      prevFlcDesc = null;
      nrhaDesc = null;

      return true;
   }

   /**
    * Handles any cleanup.  Closes any open statements and the connection wrapper object.
    */
   protected void cleanup()
   {
      setCurAction("Cleaning up...");

      try {
         //
         // If there's a sheet left over, make sure its cleaned up
         if ( m_WorkBook.getNumberOfSheets() > 0 )
            if ( m_WorkBook.getSheetAt(0) != null )
               m_WorkBook.removeSheetAt(0);
      }

      catch ( Exception e ) {
         log.error(e);
      }

      finally {
         m_WorkBook = null;
         m_Sheet = null;
         m_StyleLeft = null;
         m_StyleRght = null;
         m_StyleCent = null;
         m_Font = null;
         m_CustId = null;
         m_NrhaId = null;
        
         closeStatements();         
      }
   }

   /**
    * Closes any open statements.
    */
   private void closeStatements()
   {
      //
      // Close customer address statement
      if ( m_CustAddr != null ) {
         try {
            m_CustAddr.close();
         }
         catch ( SQLException e) {
         }

         m_CustAddr = null;
      }

      //
      // Close main items prepared statement
      if ( m_Items != null ) {
         try {
            m_Items.close();
         }
         catch ( SQLException e) {
         }

         m_Items = null;
      }

      //
      // Close old customer sell price statement
      if ( m_OldCustSellPrice != null ) {
         try {
            m_OldCustSellPrice.close();
         }
         catch ( SQLException e ) {
            
         }

         m_OldCustSellPrice = null;
      }

      //
      // Close new customer sell price statement
      if ( m_NewCustSellPrice != null ) {
         try {
            m_NewCustSellPrice.close();
         }
         catch ( SQLException e ) {
            
         }

         m_NewCustSellPrice = null;
      }

      //
      // Close old customer retail price statement
      if ( m_OldCustRetlPrice != null ) {
         try {
            m_OldCustRetlPrice.close();
         }
         catch ( SQLException e ) {
            
         }

         m_OldCustRetlPrice = null;
      }

      //
      // Close new customer retail price statement
      if ( m_NewCustRetlPrice != null ) {
         try {
            m_NewCustRetlPrice.close();
         }
         catch ( SQLException e ) {
            
         }

         m_NewCustRetlPrice = null;
      }

      //
      // Close old base retails statement
      if ( m_OldBaseRetails != null ) {
         try {
            m_OldBaseRetails.close();
         }
         catch ( SQLException e ) {
            
         }

         m_OldBaseRetails = null;
      }

      //
      // Close new base retails statement
      if ( m_NewBaseRetails != null ) {
         try {
            m_NewBaseRetails.close();
         }
         catch ( SQLException e ) {
            
         }

         m_NewBaseRetails = null;
      }

      //
      // Close crp item statement
      if ( m_CrpItem != null ) {
         try {
            m_CrpItem.close();
         }
         catch ( SQLException e ) {
            
         }

         m_CrpItem = null;
      }

      //
      // Close crp markets statement
      if ( m_CrpMkts != null ) {
         try {
            m_CrpMkts.close();
         }
         catch ( SQLException e ) {
            
         }

         m_CrpMkts = null;
      }
   }

   /**
    * Creates captions for the item listing section.
    */
   private void createCaptions(int row)
   {
      //
      // Nrha caption
      addCell(row,  0, "NRHA", HSSFCellStyle.ALIGN_CENTER);

      //
      // Mdc caption
      addCell(row,  1, "MDC", HSSFCellStyle.ALIGN_CENTER);

      //
      // Flc caption
      addCell(row,  2, "FLC", HSSFCellStyle.ALIGN_CENTER);

      //
      // Item# caption
      addCell(row,  3, "ITEM NO.", HSSFCellStyle.ALIGN_CENTER);

      //
      // Item description caption
      addRegion(row,  4,  10, "DESCRIPTION", REGION_BORDERED, true);

      //
      // Qty shipped caption
      addCell(row,  11, "QTY RY", HSSFCellStyle.ALIGN_CENTER);

      //
      // Cost/retail caption  (Current Pricing)
      addRegion(row,  12,  14, "COST/RETAIL", REGION_BORDERED, true);

      //
      // Crp option caption  (Current Pricing)
      addRegion(row,  15,  17, "CRP OPTION/ GROSS MGN", REGION_BORDERED, true);

      //
      // Cost/retail caption  (New Pricing)
      addRegion(row,  19,  21, "COST/RETAIL", REGION_BORDERED, true);

      //
      // Crp option caption  (Current Pricing)
      addRegion(row,  22,  24, "CRP OPTION/ GROSS MGN", REGION_BORDERED, true);

      //
      // New crp caption
      addCell(row,  25, "NEW CRP", HSSFCellStyle.ALIGN_CENTER);

      //
      // COGS caption
      addCell(row,  27, "COGS", HSSFCellStyle.ALIGN_CENTER);

      //
      // Ext retail caption
      addCell(row,  28, "EXT RET", HSSFCellStyle.ALIGN_CENTER);

      //
      // New ext retail caption
      addCell(row,  29, "NEW EXT", HSSFCellStyle.ALIGN_CENTER);
   }
   
   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {
      boolean created = false;
      Iterator<String> iterCust = null;
      
      try {         
         m_OraConn = m_RptProc.getOraConn();
         
         if ( prepareStatements() ) {
            iterCust = m_CustList.iterator();
            
            //
            // Iterate over the customer list and build a report for each customer
            while ( iterCust.hasNext() && m_Status != RptServer.STOPPED ) {
               m_CustId = iterCust.next();

               //
               // Add the sheet for the current report.  This will be removed after the
               // current report file has been built.
               m_Sheet = m_WorkBook.createSheet();
               created = buildOutputFile();
               
               if ( created && m_Status != RptServer.STOPPED ) {            
                  try {
                     //TODO add the rest of the code in the old report.  It's really screwed up.
                  }
                  
                  catch ( Exception ex ) {
                     log.error("exception: " + ex);
                  }
               }
               
               m_FileIndex++;
            }
         }
      }
      
      catch ( Exception ex ) {
         log.fatal("exception:", ex);
      }
      
      finally {
        cleanup(); 
      }
      
      return created;
   }

   /**
    * Gets the current customer specific sell price from pr03, else returns zero if some
    * exception occurred.
    *
    * @param custId String - the input customer identifier.
    * @param itemId String - the input item identifier.
    * @return double - the current customer specific sell price.
    */
   public double getOldCustSellPrice(String custId, String itemId)
   {
      double sell;

      try {
         m_OldCustSellPrice.setString(2, custId);
         m_OldCustSellPrice.setString(3, itemId);
         m_OldCustSellPrice.execute();

         sell = m_OldCustSellPrice.getDouble(1);
      }
      catch ( SQLException e ) {
         //
         // Set to zero if an exception occurred
         sell = 0.0;
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
         //
         // Set to zero if an exception occurred
         retl = 0.0;
      }

      return retl;
   }

   /**
    * Gets the NEW customer specific sell price from ts01, else returns zero if some
    * exception occurred.
    *
    * @param custId String - the input customer identifier.
    * @param itemId String - the input item identifier.
    * @return double - the new customer specific sell price.
    */
   public double getNewCustSellPrice(String custId, String itemId)
   {
      double sell;

      try {
         m_NewCustSellPrice.setString(2, custId);
         m_NewCustSellPrice.setString(3, itemId);
         m_NewCustSellPrice.execute();

         sell = m_NewCustSellPrice.getDouble(1);
      }
      catch ( SQLException e ) {
         //
         // Set to zero if an exception occurred
         sell = 0.0;
      }

      return sell;
   }

   /**
    * Gets the NEW customer specific retail price from ts01, else returns zero if some
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
    * Prepares the sql queries for execution.
    * 
    * @return true if the statements were prepared, false if not.
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer();
      boolean isPrepared = false;

      if ( m_OraConn != null ) {
         try {            
            //
            // Build customer address sql
            sql.append("select /*+rule*/ name, addr1, addr2, city, state, postal_code ");
            sql.append("from cust_address_view, customer ");
            sql.append("where customer.customer_id = ? and addrtype = 'SHIPPING' and ");
            sql.append("cust_address_view.customer_id = customer.customer_id");

            m_CustAddr = m_OraConn.prepareStatement(sql.toString());

            sql.setLength(0);

            //
            // Builds main item sql.  Gets distinct item information over a 2 year purchase
            // history, including the total quantity shipped over that period.
            sql.append("select item_nbr, item.description, item.retail_pack, ");
            sql.append("   flc.flc_id, flc.description as flc_desc, mdc.mdc_id, ");
            sql.append("   nrha.nrha_id, nrha.description as nrha_desc, ");
            sql.append("   sum(qty_shipped) as qty_ship ");
            sql.append("from inv_dtl, item, flc, mdc, nrha ");
            sql.append("where cust_nbr = ? and ");
            sql.append("   sale_type = 'WAREHOUSE' and ");
            sql.append("   nrha.nrha_id = ? and ");
            sql.append("   inv_dtl.item_nbr = item.item_id and ");
            sql.append("   item.flc_id = flc.flc_id and ");
            sql.append("   flc.mdc_id = mdc.mdc_id and ");
            sql.append("   mdc.nrha_id = nrha.nrha_id and ");
            sql.append("   inv_dtl.invoice_date > add_months(sysdate, -24) and ");
            sql.append("   inv_dtl.invoice_date <= sysdate ");
            sql.append("group by item_nbr, item.description, item.retail_pack, flc.flc_id, ");
            sql.append("   flc.description, mdc.mdc_id, nrha.nrha_id, nrha.description ");
            sql.append("order by nrha_id, mdc_id, flc_id, item_nbr");

            //
            // Make customer resultset scrollable so can get total # rows
            m_Items = m_OraConn.prepareStatement(sql.toString(), ResultSet.TYPE_SCROLL_INSENSITIVE,
               ResultSet.CONCUR_READ_ONLY);

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

            sql.setLength(0);

            //
            // Get all the old base retail values from pr03
            sql.append("declare ");
            sql.append("   itm varchar2(7); ");
            sql.append("begin ");
            sql.append("   itm := ?; ");
            sql.append("   ? := round(item_price_procs.todays_retaila(itm), 2); ");
            sql.append("   ? := round(item_price_procs.todays_retailb(itm), 2); ");
            sql.append("   ? := round(item_price_procs.todays_retailc(itm), 2); ");
            sql.append("   ? := round(item_price_procs.todays_retaild(itm), 2); ");
            sql.append("end;");

            m_OldBaseRetails = m_OraConn.prepareCall(sql.toString());
            m_OldBaseRetails.registerOutParameter(2, Types.DOUBLE);
            m_OldBaseRetails.registerOutParameter(3, Types.DOUBLE);
            m_OldBaseRetails.registerOutParameter(4, Types.DOUBLE);
            m_OldBaseRetails.registerOutParameter(5, Types.DOUBLE);

            sql.setLength(0);

            //
            // Get all the new base retail values from ts01
            sql.append("declare ");
            sql.append("   itm varchar2(7); ");
            sql.append("begin ");
            sql.append("   itm := ?; ");
            sql.append("   ? := round(item_price_procs.todays_retaila@ts01(itm), 2); ");
            sql.append("   ? := round(item_price_procs.todays_retailb@ts01(itm), 2); ");
            sql.append("   ? := round(item_price_procs.todays_retailc@ts01(itm), 2); ");
            sql.append("   ? := round(item_price_procs.todays_retaild@ts01(itm), 2); ");
            sql.append("end;");

            m_NewBaseRetails = m_OraConn.prepareCall(sql.toString());
            m_NewBaseRetails.registerOutParameter(2, Types.DOUBLE);
            m_NewBaseRetails.registerOutParameter(3, Types.DOUBLE);
            m_NewBaseRetails.registerOutParameter(4, Types.DOUBLE);
            m_NewBaseRetails.registerOutParameter(5, Types.DOUBLE);

            sql.setLength(0);

            //
            // Build crp item statement
            sql.append("select decode(row_type, 'NRHA', 'D', 'IMAGE', 'M', substr(row_type, 1, 1)) as crp, ");
            sql.append("crp_type, market_id ");
            sql.append("from cust_crp, crp_option ");
            sql.append("where cust_crp_id = (select cust_procs.item_crp_id(?, ?) from dual) and ");
            sql.append("cust_crp.crp_opt_id = crp_option.crp_opt_id");

            m_CrpItem = m_OraConn.prepareStatement(sql.toString());

            sql.setLength(0);

            //
            // Build crp markets statement
            sql.append("select row_type, nvl(market_id, 'C') as mkt ");
            sql.append("from cust_crp, crp_option ");
            sql.append("where customer_id = ? and ");
            sql.append("   row_type in ('STORE', 'IMAGE') and ");
            sql.append("   sen_code_id = 0 and ");
            sql.append("   cust_crp.crp_opt_id = crp_option.crp_opt_id");
            m_CrpMkts = m_OraConn.prepareStatement(sql.toString());
            
            isPrepared = true;            
         }
         
         catch ( Exception ex ) {
            log.fatal("exception: ", ex);
         }
         
         finally {
            sql = null;
         }
      }
      
      return isPrepared;      
   }

   /**
    * Rounds input value to specified precision.  The precision being the number
    * of decimal places.
    *
    * @param val double - the input value to round.
    * @param precision int - the number of decimal places in the result.
    * @return double - the rounded value.
    */
   private double round(double val, int precision)
   {
      int factor = 1;

      if ( val == 0 )
         return val;

      //
      // Get the multiplication factor
      for ( int i = 1; i <= precision; i++ )
         factor = factor * 10;

      //
      // Round sell price to x decimal places
      return Math.floor(val * factor + .5d) / factor;
   }
   
   /**
    * Process the report parameters.
    *    Note - the suppress warning is set so because there doesn't seem to be a way to remove
    *       the warning when calling parse string.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */  
   public void setParams(ArrayList<Param> params)
   {      
      //
      // Get all the customer#'s and store them.
      m_CustList = StringFormat.parseString(params.get(0).value, ';');
      m_NumCust = m_CustList.size();

      //
      // Initialize additional parameters needed by the report
      m_NrhaId = params.get(1).value;      
      m_FileIndex = 1;
   }
}
