/**
 * File: VndPriceChange.java
 * Description: Reports price changes by Vendor.
 *
 * @author Jacob Heric
 *
 *
 * Create Date: 07/24/2008
 * Last Update: $Id: VndPriceChange.java,v 1.23 2014/12/03 23:16:23 everge Exp $
 *
 * History:
 *    $Log: VndPriceChange.java,v $
 *    Revision 1.23  2014/12/03 23:16:23  everge
 *    Fixed issue where row indices were stored as short integers, limiting row count to 32,768. Switched from HSSF to generic SS usermodel. More explicit logging. Other minor fixes.
 *
 *    Revision 1.22  2013/09/11 14:25:42  tli
 *    Converted the facilityId to facilityName when needed
 *
 *    Revision 1.21  2013/09/09 18:33:38  tli
 *    Replace SkuQty web service call with item_qty_view
 *
 *    Revision 1.20  2012/10/11 14:23:13  jfisher
 *    Changes to deal with the timeout on the sku quantity web service.
 *
 *    Revision 1.19  2012/08/29 19:53:02  jfisher
 *    Switched web service calls from Wasp to Axis2
 *
 *    Revision 1.18  2012/05/05 06:05:52  pberggren
 *    Removed redundant loading of system properties.
 *
 *    Revision 1.17  2012/05/03 07:55:10  prichter
 *    Fix to web service ip address
 *
 *    Revision 1.16  2012/05/03 04:46:54  pberggren
 *    Added server.properties call to force report to .57
 *
 *    Revision 1.15  2009/02/18 16:13:18  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.14  2008/11/19 15:40:25  jfisher
 *    Fixed a bunch of logic errors with the report when data wasn't available.  Fixed the null pointer exception.
 *
 *    Revision 1.13  2008/11/18 12:19:05  jfisher
 *    Changed the order by on the slot pricing query to order by price level
 *
 *    Revision 1.12  2008/11/17 22:09:37  pdavidson
 *    Fix to use facility ID instead of name when calling SkuQty web service
 *
 *    Revision 1.11  2008/09/03 21:57:54  jheric
 *    Check divisor, not dividend in margin calc.
 *
 *    Revision 1.10  2008/08/29 19:53:36  jfisher
 *    Fixed some warnings, production version.
 */

package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import com.emerywaterhouse.fascor.Facility;
import com.emerywaterhouse.rpt.helper.QtyBreakPrice;
import com.emerywaterhouse.rpt.helper.SlotPrice;
import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class VndPriceChange extends Report
{
   private final SpreadsheetVersion SS_VERSION = SpreadsheetVersion.EXCEL2007; //or SpreadsheetVersion.EXCEL2007
   private static final int TOTAL_QTY_BREAKS = 3;  //Total minimum # of qty break prices to attempt to show
   private static final int TOTAL_PRICE_SLOTS = 4; //Total minimum # of price breaks to attempt to show.

   private PreparedStatement m_SqlVPC;     // vendor price change statement
   private PreparedStatement m_QtyBreak;   // Quantity break prices
   private PreparedStatement m_SlotPrice;  // Slot Prices
   private PreparedStatement m_UnitsSold;  // Units sold (by facility).
   private PreparedStatement m_CanSell;    // Item active indicator by facility
   private PreparedStatement m_ItemDCQty;  // Item available in facility
   
   private String m_QtyBreakTotalSQL;      // Quantity break price totals (used for drawing variable length column headers)
   private String m_SlotPriceTotalSQL;     // Slot Price totals (used for drawing variable length column headers)
   private int m_QtyBrkTotal;              // Note maximum number of qty breaks on this report
   private int m_SlotPriceTotal;           // Note maximum number of slot price breaks on this report
   private int m_VendorId;                 // Selected EIS Vendor Number
   private boolean m_Active;               // Active only items?
   private Workbook m_Workbook;
   private Sheet m_Sheet;
   private Row m_Row;
   private Font m_FontNorm;
   private Font m_FontBold;
   private CellStyle m_StyleHdrLeft;
   private CellStyle m_StyleHdrLeftWrap;
   private CellStyle m_StyleHdrCntr;
   private CellStyle m_StyleHdrCntrWrap;
   private CellStyle m_StyleHdrRghtWrap;
   private CellStyle m_StyleDtlLeft;
   private CellStyle m_StyleDtlLeftWrap;
   private CellStyle m_StyleDtlRightWrap;
   private CellStyle m_StyleDtlCntr;
   private CellStyle m_StyleDtlRght;
   private CellStyle m_StyleDtlRght2d;
   private CellStyle m_StyleDtlRght3d;
   private CellStyle m_StyleDtlRght4d;
   private CellStyle m_StyleNewLine;
   private CreationHelper m_CreateHelper;
   private FileOutputStream m_OutputStream;
   private ArrayList<Facility> m_FacilityList;  // List of all whse facilities
   private String m_FacilityCondition;          // SQL facility condition fragment.
   private String m_WhseId;                     // Warehouse input parameter

    /**
    * default constructor
    */
   public VndPriceChange()
   {
      super();
      
      m_Workbook = new XSSFWorkbook();
      String fileName = "VndPriceChange-" + String.valueOf(System.currentTimeMillis()) + ".xlsx";
      m_Sheet = m_Workbook.createSheet("VndPriceChange");
      m_CreateHelper = m_Workbook.getCreationHelper();
      defineStyles();
      m_FileNames.add(fileName);
      m_FacilityList = new ArrayList<Facility>();
      
   }

   /**
    * adds a numeric type cell to current row at col p_Col in current sheet
    *
    * @param col     0-based column number of spreadsheet cell
    * @param value   numeric value to be stored in cell
    * @param style   Excel style to be used to display cell
    */

   private void addCell(int col, double value, CellStyle style)
   {
      Cell cell = m_Row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellStyle(style);

      //
      // Math.round(V * D) / D is required at this (final) stage to store the
      // true decimal value rather than the Java float value
      // (for example, 2.82 rather than 2.81999993324279)
      if ( style == m_StyleDtlRght2d )
         cell.setCellValue(Math.round(value * 100d) / 100d);
      else {
         if ( style == m_StyleDtlRght3d )
            cell.setCellValue(Math.round(value * 1000d) / 1000d);
         else {
            if ( style == m_StyleDtlRght4d)
               cell.setCellValue(Math.round(value * 10000d) / 10000d);
            else
               cell.setCellValue(value);
         }
      }

      cell = null;
   }

   /**
    * adds a text type cell to current row at col p_Col in current sheet
    *
    * @param col     0-based column number of spreadsheet cell
    * @param value   text value to be stored in cell
    * @param style   Excel style to be used to display cell
    */
   private void addCell(int col, String value, CellStyle style)
   {
      Cell cell = m_Row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(m_CreateHelper.createRichTextString(value));
      cell.setCellStyle(style);
      cell = null;
   }

   /**
    * adds a integer type cell to current row at col in current sheet
    *
    * @param col     0-based column number of spreadsheet cell
    * @param value   Integer value to be stored in cell
    * @param style   Excel style to be used to display cell
    */
   private void addCell(int col, Integer value, CellStyle style)
   {
      Cell cell = m_Row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellValue(value);
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
    * Builds a facility sql condition snippet for use in the prep. statements,
    * of the form: ('PORTLAND', 'PITSTON').
    *
    * @return String - the facility sql condition snippet.
    * @throws SQLException
    */
   private String buildFacilityCondition()
   {
      StringBuffer tmp = new StringBuffer();
      int i = 1;

      try {
         tmp.append("(");
         for (Facility f : this.getFacilityList()){
            tmp.append("'");
            tmp.append(f.getName());
            tmp.append("'");
            tmp.append(i != this.getFacilityList().size() ? ", " : "");
            i++;
         }

         tmp.append(")");

      }

      catch(Exception e) {
         log.fatal("[VndPriceChange] buildFacilityCondition ", e);
         m_ErrMsg.append("The report had the following Error(s) in: \r\n");
         m_ErrMsg.append(e.getClass().getName() + "\r\n" + e.getMessage());
      }
      return tmp.toString();
   }

   /**
    * Opens the Excel spreadsheet and creates the title and the column headings
    *
    * @return  row number of first detail row (below header rows)
    */
   private int buildTitle()
   {
      int col = 0;
      int row = 0;
      int charWidth = 295;
      StringBuffer hdr = new StringBuffer();
      ResultSet rsQtyBrk = null;
      ResultSet rsSlotPrice = null;
      Statement stm = null;

      try {
         // creates Excel title
         addRow(row++);

         hdr.append("Vendor Price Change Review Report for ");
         hdr.append(" Vendor ");
         hdr.append(m_VendorId);
         hdr.append(" ");
         hdr.append(getVndName());

         //
         //Indicate if active items only.
         if ( m_Active )
            hdr.append("  (Active items only). ");

         //
         //Add report date.
         DateFormat df = new SimpleDateFormat("MM/dd/yyyy");
         hdr.append(" " + df.format(new java.util.Date(System.currentTimeMillis())));

         addCell(col, hdr.toString(), m_StyleHdrLeft);
         hdr = null;
         addRow(row++);// blank row between title and headings
         // creates Excel column headings
         addRow(row++);

         //
         // Computes approximate HSSF character width based on "Arial" size "10" and
         // HHSF characteristics, to be used as a multiplier to set column widths.
         //

         // Item ID
         m_Sheet.setColumnWidth(col, (8 * charWidth));
         addCell(col, "Item ID", m_StyleHdrLeftWrap);

         // Vendor Item ID
         m_Sheet.setColumnWidth(++col, (12 * charWidth));
         addCell(col, "Vendor Item ID", m_StyleHdrLeftWrap);

         // Item Description
         m_Sheet.setColumnWidth(++col, (30 * charWidth));
         addCell(col, "Description", m_StyleHdrLeftWrap);

         // Dealer Pack
         m_Sheet.setColumnWidth(++col, (6 * charWidth));
         addCell(col, "Dealer Pack", m_StyleHdrCntrWrap);

         // Stock Pack
         m_Sheet.setColumnWidth(++col, (6 * charWidth));
         addCell(col, "Stock Pack", m_StyleHdrCntrWrap);

         // NBC
         m_Sheet.setColumnWidth(++col, (6 * charWidth));
         addCell(col, "NBC", m_StyleHdrCntrWrap);

         // Ship Unit
         m_Sheet.setColumnWidth(++col, (6 * charWidth));
         addCell(col, "Ship Unit", m_StyleHdrCntrWrap);

         // Product Category Code
         m_Sheet.setColumnWidth(++col, (6 * charWidth));
         addCell(col, "PC Code", m_StyleHdrCntrWrap);

         // Product Category Description
         m_Sheet.setColumnWidth(++col, (20 * charWidth));
         addCell(col, "PC Description", m_StyleHdrCntrWrap);

         //Primary UPC
         m_Sheet.setColumnWidth(++col, (12 * charWidth));
         addCell(col, "Primary UPC",   m_StyleHdrLeftWrap);

         // Emery Cost
         m_Sheet.setColumnWidth(++col, (9 * charWidth));
         addCell(col, "Emery Cost", m_StyleHdrCntrWrap);

         // Emery Base
         m_Sheet.setColumnWidth(++col, (9 * charWidth));
         addCell(col, "Emery Base", m_StyleHdrCntrWrap);

         // Margin
         m_Sheet.setColumnWidth(++col, (9 * charWidth));
         addCell(col, "Margin", m_StyleHdrCntrWrap);

         // New Cost
         m_Sheet.setColumnWidth(++col, (9 * charWidth));
         addCell(col, "New Cost", m_StyleHdrCntrWrap);

         // New  Base
         m_Sheet.setColumnWidth(++col, (9 * charWidth));
         addCell(col, "New Base", m_StyleHdrCntrWrap);

         //
         //Slot prices/qty total.  For these we need to peek ahead.
         stm = m_EdbConn.createStatement();
         rsSlotPrice = stm.executeQuery(m_SlotPriceTotalSQL);

         if (rsSlotPrice.next()){
            m_SlotPriceTotal = rsSlotPrice.getInt("total");
         }
         //
         //Cleanup
         closeStatement(stm);

         //If there are no slot prices on this report, set a
         //reasonable minimum so columns for new slots can be drawn.
         if (m_SlotPriceTotal < TOTAL_PRICE_SLOTS)
            m_SlotPriceTotal = TOTAL_PRICE_SLOTS;

         for (int i = 1; i <= m_SlotPriceTotal; i++){
            // Slot i Price Min Qty
            m_Sheet.setColumnWidth(++col, (8 * charWidth));
            addCell(col, "Slot " + i + " Min Qty", m_StyleDtlRightWrap);

            // Slot i Price
            m_Sheet.setColumnWidth(++col, (8 * charWidth));
            addCell(col, "Slot " + i, m_StyleDtlRightWrap);

            // Slot i Margin
            m_Sheet.setColumnWidth(++col, (8 * charWidth));
            addCell(col, "Slot " + i + " Margin", m_StyleDtlRightWrap);

            // New Slot i New Price
            m_Sheet.setColumnWidth(++col, (8 * charWidth));
            addCell(col, "New Slot " + i, m_StyleDtlRightWrap);
         }

         //
         //Qty break total.  For these we need to peek ahead so we draw enough headers.
         stm = m_EdbConn.createStatement();
         rsQtyBrk = stm.executeQuery(m_QtyBreakTotalSQL);

         if (rsQtyBrk.next()){
            m_QtyBrkTotal = rsQtyBrk.getInt("total");
         }

         //If there are no qty breaks on this report, set a
         //reasonable minimum so columns for new qty breaks can be drawn.
         if (m_QtyBrkTotal < TOTAL_QTY_BREAKS)
            m_QtyBrkTotal = TOTAL_QTY_BREAKS;

         for (int i = 1; i <= m_QtyBrkTotal; i++){
            // Qty Break i Min Qty
            m_Sheet.setColumnWidth(++col, (8 * charWidth));
            addCell(col, "Qty Brk" + i, m_StyleDtlRightWrap);

            // Qty  i %
            m_Sheet.setColumnWidth(++col, (8 * charWidth));
            addCell(col, "Qty " + i + " %", m_StyleDtlRightWrap);
         }

         // Retail A
         m_Sheet.setColumnWidth(++col, (8 * charWidth));
         addCell(col, "Retail A", m_StyleHdrCntrWrap);

         // Retail B
         m_Sheet.setColumnWidth(++col, (8 * charWidth));
         addCell(col, "Retail B", m_StyleHdrCntrWrap);

         // Retail C
         m_Sheet.setColumnWidth(++col, (8 * charWidth));
         addCell(col, "Retail C", m_StyleHdrCntrWrap);

         // Retail D
         m_Sheet.setColumnWidth(++col, (8 * charWidth));
         addCell(col, "Retail D", m_StyleHdrCntrWrap);

         // Units Sold
         for (Facility f : this.getFacilityList()){
            m_Sheet.setColumnWidth(++col, (8 * charWidth));
            addCell(col, f.getName() + " Units Sold", m_StyleDtlRightWrap);
         }

         // Sensitivity Code
         m_Sheet.setColumnWidth(++col, (8 * charWidth));
         addCell(col, "Sensitivity Code", m_StyleDtlRightWrap);

         // On Hand
         for(Facility f : this.getFacilityList()){
            m_Sheet.setColumnWidth(++col, (8 * charWidth));
            addCell(col, f.getName() + " On Hand", m_StyleDtlRightWrap);
         }

         // Can Sell (AKA Active)
         for(Facility f : this.getFacilityList()){
            m_Sheet.setColumnWidth(++col, (8 * charWidth));
            addCell(col, f.getName() + " Can Sell", m_StyleDtlRightWrap);
         }

         // FLC
         m_Sheet.setColumnWidth(++col, (8 * charWidth));
         addCell(col, "FLC", m_StyleDtlLeftWrap);

         // Catalog
         m_Sheet.setColumnWidth(++col, (8 * charWidth));
         addCell(col, "Catalog", m_StyleDtlLeftWrap);
      }

      catch ( Exception ex ) {
         log.fatal("VndPriceChange.openWorkbook ", ex);
         m_ErrMsg.append("The report had the following Error(s) in: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n" + ex.getMessage());
      }

      finally {
         closeResultSet(rsQtyBrk);
         closeResultSet(rsSlotPrice);
         closeStatement(stm);
      }

      // returns first data row number
      return row;
   }

   /**
    * Builds Excel workbook.
    * <p>
    * Opens, builds and closes the Excel spreadsheet.
    * @return true if the workbook was built, false if not.
    */
   private boolean buildWorkbook()
   {
      boolean result = false;
      int row = 0;
      int col = 0;
      int book = 2;
      int maxRows = SS_VERSION.getLastRowIndex();
      String itemId = null;
      ResultSet rsVPC = null;
      ResultSet rsQtyBrk = null;
      ResultSet rsSlotPrice = null;
      ResultSet rsUnitsSold = null;
      ResultSet rsCanSell = null;
      Float cost = null;
      BigDecimal bd = null;         // for rounding

      //
      //Some maps and list to facilitate drawing variable length columnar data (e.g. by facility)
      Map<Facility, Integer> unitsSold = new HashMap<Facility, Integer>();
      Map<Facility, Integer> canSell = new HashMap<Facility, Integer>();
      List<QtyBreakPrice> qtyBreaks = new ArrayList<QtyBreakPrice>();
      List<SlotPrice> slotPrices = new ArrayList<SlotPrice>();
      row = buildTitle();

      try {
         setCurAction("Running the vendor price change report");
         m_SqlVPC.setInt(1, m_VendorId);
         rsVPC = m_SqlVPC.executeQuery();

         while ( rsVPC.next() && m_Status == RptServer.RUNNING ) {
            //
            //Clear helper maps & lists.
            unitsSold.clear();
            qtyBreaks.clear();
            slotPrices.clear();
            canSell.clear();

            addRow(row);

            //Item ID (store item number for later lookups)
            itemId = rsVPC.getString("item_id");
            addCell(col++, itemId, m_StyleDtlLeftWrap);
            addCell(col++, rsVPC.getString("vendor_item_num"), m_StyleDtlLeftWrap);
            addCell(col++, rsVPC.getString("description"), m_StyleDtlLeftWrap);
            addCell(col++, rsVPC.getString("retail_pack"), m_StyleDtlRght);
            addCell(col++, rsVPC.getString("stock_pack"), m_StyleDtlRght);
            addCell(col++, rsVPC.getString("broken_case"), m_StyleDtlLeft);
            addCell(col++, rsVPC.getString("unit"), m_StyleDtlLeft);
            addCell(col++, rsVPC.getString("category_code"), m_StyleDtlLeft);
            addCell(col++, rsVPC.getString("product_category"), m_StyleDtlLeft);
            addCell(col++, rsVPC.getString("upc_code"), m_StyleDtlCntr);
            cost = rsVPC.getFloat("cost");
            addCell(col++, cost, m_StyleDtlRght2d);
            addCell(col++, rsVPC.getFloat("base"), m_StyleDtlRght2d);
            addCell(col++, rsVPC.getFloat("margin"), m_StyleDtlRght2d);
            addCell(col++, rsVPC.getFloat("new_cost"), m_StyleDtlRght2d);
            addCell(col++, rsVPC.getFloat("new_base"), m_StyleDtlRght2d);

            //Slot quantities and Percent statement
            m_SlotPrice.setString(1, itemId);
            rsSlotPrice = m_SlotPrice.executeQuery();

            //
            //Store them before adding, this facilitates drawing as many as we have (which may
            //be less than the # of columns) and initializing the rest.
            while (rsSlotPrice.next())
               slotPrices.add(new SlotPrice(rsSlotPrice.getDouble("price"), rsSlotPrice.getInt("min_qty")));

            SlotPrice spTmp = null;

            //
            //Slot quantity, price, margin and new price
            for (int i = 0; i < m_SlotPriceTotal; i++){
               spTmp = null;
               //
               //Only retrieve price if it exists
               if ( slotPrices.size() > i )
                  spTmp = slotPrices.get(i);

               //
               //Enter blanks when not slot prices found
               if ( spTmp == null ) {
                  //
                  //Min qty.
                  addCell(col++, "", m_StyleDtlRght);
                  //
                  //Blank Price
                  addCell(col++, "", m_StyleDtlRght);
               }
               else {
                  // Slot i Price Min Qty
                  addCell(col++, spTmp.getMinQty(), m_StyleDtlRght);
                  //
                  //Slot i Price
                  addCell(col++, spTmp.getPrice(), m_StyleDtlRght);
               }

               //
               // Slot i Margin, calculate only if we have a cost AND slot price.
               // Mind I'm using bigdecimal to do the round.
               if ( cost != null && cost > 0 && spTmp != null && spTmp.getPrice() > 0 ) {
                  bd = new BigDecimal(100 - 100 * cost / spTmp.getPrice());
                  bd = bd.setScale(2, BigDecimal.ROUND_HALF_UP);
                  addCell(col++, bd.floatValue(), m_StyleDtlRght);
               }
               else
                  addCell(col++, "", m_StyleDtlRght);

               //
               // New Slot i New Price, blank for now (leave it up to user)
               addCell(col++, "", m_StyleDtlRght);
            }

            //
            // Qty Break Qty & Percent sql
            m_QtyBreak.setString(1, itemId);
            rsQtyBrk = m_QtyBreak.executeQuery();

            //
            // Store them before adding, this facilitates drawing as many as we have (which may
            // be less than the # of columns) and initializing the rest.
            while ( rsQtyBrk.next() )
               qtyBreaks.add(new QtyBreakPrice(rsQtyBrk.getInt("min_qty"), rsQtyBrk.getDouble("percent")));

            QtyBreakPrice qbTmp = null;

            //
            //Qty Break Pricing
            for ( int i = 0; i < m_QtyBrkTotal; i++ ) {
               qbTmp = null;

               //
               //Only get qty break if they exist
               if ( qtyBreaks.size() > i) {
                  qbTmp = qtyBreaks.get(i);
               }

               //When no breaks found, enter blanks
               if ( qbTmp == null ) {
                  //
                  // Qty Break i Min Qty
                  addCell(col++, "", m_StyleDtlRght);
                  //
                  //Qty Break i %
                  addCell(col++, "", m_StyleDtlRght);
               }
               else {
                  //
                  // Qty Break i Min Qty,
                  addCell(col++, qbTmp.getMinQty(), m_StyleDtlRght);
                  //
                  // Qty Break i Min Qty
                  addCell(col++, qbTmp.getPercent(), m_StyleDtlRght);
               }
            }


            //Retail A
            addCell(col++, rsVPC.getFloat("retail_a"), m_StyleDtlRght2d);

            //Retail B
            addCell(col++, rsVPC.getFloat("retail_b"), m_StyleDtlRght2d);

            //Retail C
            addCell(col++, rsVPC.getFloat("retail_c"), m_StyleDtlRght2d);

            //Retail D
            addCell(col++, rsVPC.getFloat("retail_d"), m_StyleDtlRght2d);

            //Units Sold, by facility
            m_UnitsSold.setString(1, itemId);
            rsUnitsSold = m_UnitsSold.executeQuery();

            //
            //Gather and store units sold by facility so we can put them in the file in the correct order, etc.
            while (rsUnitsSold.next()){
               unitsSold.put(this.getFacilityByName(rsUnitsSold.getString("warehouse")), rsUnitsSold.getInt("units_sold"));
            }

            //
            //Units sold by facility
            for (Facility f : this.getFacilityList()){
               addCell(col++, unitsSold.get(f) == null ? new Integer(0) : unitsSold.get(f).intValue(), m_StyleDtlRght);
            }

            //Sensitivity Code
            addCell(col++, rsVPC.getString("sen_code_id"), m_StyleDtlRght);

            //Units on Hand by facility
            for (Facility f : this.getFacilityList()){
               addCell(col++, this.getOnHand(itemId, f.getFas_facility_id()), m_StyleDtlRght);
            }

            //Item Active, by facility
            m_CanSell.setString(1, itemId);
            rsCanSell = m_CanSell.executeQuery();

            //
            //Gather and store item active information by facility so we can put them in the file in the correct order.
            while (rsCanSell.next()){
               canSell.put(this.getFacilityByName(rsCanSell.getString("name")), rsCanSell.getInt("active"));
            }

            //
            //Item Active indicator by facility.
            for(Facility f : this.getFacilityList()){
               addCell(col++, canSell.get(f) != null && canSell.get(f).intValue() == 1 ? "Yes" : "No", m_StyleDtlLeft);
            }

            //FLC
            addCell(col++, rsVPC.getString("flc_id"), m_StyleDtlLeft);

            //Catalog
            addCell(col++, rsVPC.getString("in_catalog").equals("0") ? "No" : "Yes", m_StyleDtlLeft);

            row++;
            //if rows exceed spreadsheet limit, roll excess over into a new sheet
            if (row > maxRows) {
               m_Sheet = m_Workbook.createSheet(m_Sheet.getSheetName() + " - pg " + book++);
               row = buildTitle();
            }
            col = 0;
         }

         closeWorkbook();
         result = true;
      }
      catch ( Exception ex ) {
         log.fatal("[VndPriceChange] buildWorkbook ", ex);
         m_ErrMsg.append("The report had the following Error(s) in: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n" + ex.getMessage());
      }

      finally {
         closeResultSet(rsVPC);
         closeResultSet(rsQtyBrk);
         closeResultSet(rsSlotPrice);
         closeResultSet(rsUnitsSold);
      }

      return result;
   }

   /*
    * cleans up member variables
    */
   protected void cleanup()
   {
      // closes statements
      closeStatement(m_SqlVPC);
      closeStatement(m_QtyBreak);
      closeStatement(m_SlotPrice);
      closeStatement(m_UnitsSold);
      closeStatement(m_CanSell);
      closeStmt(m_ItemDCQty);
      
      m_StyleDtlRght = null;
      m_StyleDtlCntr = null;
      m_StyleDtlLeft = null;
      m_StyleHdrRghtWrap = null;
      m_StyleHdrCntrWrap = null;
      m_StyleHdrLeftWrap = null;
      m_StyleHdrLeft = null;
      m_FontBold = null;
      m_Sheet = null;
      m_Workbook = null;
      m_OutputStream = null;
      m_Status = RptServer.STOPPED;
      m_ItemDCQty = null;
   }

   /**
    * closes output stream, deletes tempoaray disk file
    */
   private void closeOutputStream()
   {
      try {
         if ( m_OutputStream != null ) {
            m_OutputStream.close();
         }
      }

      catch ( Exception ex ) {
         log.error("[VndPriceChange] closeOutputStream()", ex);
      }

      finally {
         m_OutputStream = null;
      }
   }

   /**
    * Generic ResultSet close procedure.  Logs any exception.
    * @param data the result set to be closed
    */
   public void closeResultSet(ResultSet data)
   {
      try {
         data.close();
      }
      catch ( Exception e ) {
         log.fatal("[VndPriceChange] closeResultSet ", e);
         m_ErrMsg.append("The report had the following Error(s) in: \r\n");
         m_ErrMsg.append(e.getClass().getName() + "\r\n" + e.getMessage());
      }
   }

   /**
    * closes a single statement identified by p_Statement
    *
    * @param statement - name of statement to be closed
    */
   private void closeStatement(Statement statement)
   {
      try {
         if (statement != null)
            statement.close();
      }

      catch ( Exception e ) {
         log.fatal("[VndPriceChange] closeStatement ", e);
         m_ErrMsg.append("The report had the following Error(s) in: \r\n");
         m_ErrMsg.append(e.getClass().getName() + "\r\n" + e.getMessage());
      }
   }

   /**
    * Closes the Excel spreadsheet.
    * <p>
    * Writes the Excel workbook to the output stream, which also creates a disk file
    * (output file) in the default reports folder.
    * FTPs the output disk file to the user.
    * Closes the workbook.
    * @throws IOException
    */
   private void closeWorkbook()
   {
      setCurAction("writing excel spreadsheet");

      if ( m_Workbook != null ) {
         if ( m_OutputStream != null )
            try {
               m_Workbook.write(m_OutputStream);
            } catch (IOException e) {
               e.printStackTrace();
            }
         else
            log.error("[VndPriceChange] null output stream object, unable to write spreadsheet");
      }
      else
         log.error("[VndPriceChange] null workbook, unable to write spreadsheet");

      // closes spreadsheet
      m_StyleDtlRght = null;
      m_StyleDtlCntr = null;
      m_StyleDtlLeft = null;
      m_StyleHdrRghtWrap = null;
      m_StyleHdrCntrWrap = null;
      m_StyleHdrLeftWrap = null;
      m_StyleHdrLeft = null;
      m_FontBold = null;
      m_Sheet = null;
      m_Workbook = null;
   }

   /**
    * Creates the excel spreadsheet.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;

      try {
         m_EdbConn = m_RptProc.getEdbConn();
         openOutputStream();

         //
         //In order to use prepared statements with a dynamic list of facilities,
         //we must know the facilities before preparation.
         loadFacilities();
         this.setFacilityCondition(buildFacilityCondition());

         if ( prepareStatements() ) {
            created = buildWorkbook();
         }

         setCurAction("complete");
      }

      catch ( Exception ex ) {
         log.fatal("[VndPriceChange]", ex);
      }

      finally {
         closeOutputStream();
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
      // m_CustomDataFormat is used to define a non-standard data format when
      // defining a style. For example, "0.00" is a bulit-in format, but "0.000"
      // and "0.0000" are custom formats.
      DataFormat m_CustomDataFormat;
      m_CustomDataFormat = m_Workbook.createDataFormat();
      
      // defines normal font
      m_FontNorm = m_Workbook.createFont();
      m_FontNorm.setFontName("Arial");
      m_FontNorm.setFontHeightInPoints((short) 10);

      // defines bold font
      m_FontBold = m_Workbook.createFont();
      m_FontBold.setFontName("Arial");
      m_FontBold.setFontHeightInPoints((short)10);
      m_FontBold.setBold(true);

      // defines style column header, left-justified
      m_StyleHdrLeft = m_Workbook.createCellStyle();
      m_StyleHdrLeft.setFont(m_FontBold);
      m_StyleHdrLeft.setAlignment(HorizontalAlignment.LEFT);
      m_StyleHdrLeft.setVerticalAlignment(VerticalAlignment.TOP);

      // defines style column header, left-justified, wrap text
      m_StyleHdrLeftWrap = m_Workbook.createCellStyle();
      m_StyleHdrLeftWrap.setFont(m_FontBold);
      m_StyleHdrLeftWrap.setAlignment(HorizontalAlignment.LEFT);
      m_StyleHdrLeftWrap.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleHdrLeftWrap.setWrapText(true);

      // defines style column header, center-justified
      m_StyleHdrCntr = m_Workbook.createCellStyle();
      m_StyleHdrCntr.setFont(m_FontBold);
      m_StyleHdrCntr.setAlignment(HorizontalAlignment.CENTER);
      m_StyleHdrCntr.setVerticalAlignment(VerticalAlignment.TOP);

      // defines style column header, center-justified, wrap text
      m_StyleHdrCntrWrap = m_Workbook.createCellStyle();
      m_StyleHdrCntrWrap.setFont(m_FontBold);
      m_StyleHdrCntrWrap.setAlignment(HorizontalAlignment.CENTER);
      m_StyleHdrCntrWrap.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleHdrCntrWrap.setWrapText(true);

      // defines style column header, right-justified, wrap text
      m_StyleHdrRghtWrap = m_Workbook.createCellStyle();
      m_StyleHdrRghtWrap.setFont(m_FontBold);
      m_StyleHdrRghtWrap.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleHdrRghtWrap.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleHdrRghtWrap.setWrapText(true);

      // defines style detail data cell, left-justified
      m_StyleDtlLeft = m_Workbook.createCellStyle();
      m_StyleDtlLeft.setFont(m_FontNorm);
      m_StyleDtlLeft.setAlignment(HorizontalAlignment.LEFT);
      m_StyleDtlLeft.setVerticalAlignment(VerticalAlignment.TOP);

      // defines style detail data cell, left-justified, wrap text
      m_StyleDtlLeftWrap = m_Workbook.createCellStyle();
      m_StyleDtlLeftWrap.setFont(m_FontNorm);
      m_StyleDtlLeftWrap.setAlignment(HorizontalAlignment.LEFT);
      m_StyleDtlLeftWrap.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleDtlLeftWrap.setWrapText(true);

      // defines style detail data cell, right-justified, wrap text
      m_StyleDtlRightWrap = m_Workbook.createCellStyle();
      m_StyleDtlRightWrap.setFont(m_FontNorm);
      m_StyleDtlRightWrap.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtlRightWrap.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleDtlRightWrap.setWrapText(true);

      // defines style detail data cell, center-justified
      m_StyleDtlCntr = m_Workbook.createCellStyle();
      m_StyleDtlCntr.setFont(m_FontNorm);
      m_StyleDtlCntr.setAlignment(HorizontalAlignment.CENTER);
      m_StyleDtlCntr.setVerticalAlignment(VerticalAlignment.TOP);

      // defines style detail data cell, center-justified
      m_StyleNewLine = m_Workbook.createCellStyle();
      m_StyleNewLine.setFont(m_FontNorm);
      m_StyleNewLine.setWrapText( true );
      m_StyleNewLine.setAlignment(HorizontalAlignment.CENTER);
      m_StyleNewLine.setVerticalAlignment(VerticalAlignment.TOP);


      // defines style detail data cell, right-justified
      m_StyleDtlRght = m_Workbook.createCellStyle();
      m_StyleDtlRght.setFont(m_FontNorm);
      m_StyleDtlRght.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtlRght.setVerticalAlignment(VerticalAlignment.TOP);

      // defines style detail data cell, right-justified with 2 decimal places
      //  (built-in data format)
      m_StyleDtlRght2d = m_Workbook.createCellStyle();
      m_StyleDtlRght2d.setFont(m_FontNorm);
      m_StyleDtlRght2d.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtlRght2d.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleDtlRght2d.setDataFormat(m_CustomDataFormat.getFormat("0.00"));

      // defines style detail data cell, right-justified with 3 decimal places
      //  (custom data format)
      m_StyleDtlRght3d = m_Workbook.createCellStyle();
      m_StyleDtlRght3d.setFont(m_FontNorm);
      m_StyleDtlRght3d.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtlRght3d.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleDtlRght3d.setDataFormat(m_CustomDataFormat.getFormat("0.000"));

      // defines style detail data cell, right-justified with 4 decimal places
      //  (custom data format)
      m_StyleDtlRght4d = m_Workbook.createCellStyle();
      m_StyleDtlRght4d.setFont(m_FontNorm);
      m_StyleDtlRght4d.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtlRght4d.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleDtlRght4d.setDataFormat(m_CustomDataFormat.getFormat("0.0000"));
   }

   /**
    * Returns the on hand quantity of an item at a facility by calling a web service.
    *
    * @param item String - the item id
    * @param facility String - the fascor facility id
    * @return int - the quantity on hand.
    * @throws Exception
    */
   private int getOnHand(String itemId, String facilityId)
   {
	   int qty = 0;
	   ResultSet rset = null;
	      
	   if ( itemId != null && itemId.length() == 7 ) {
	         try {
	        	 m_ItemDCQty.setString(1, itemId);
	        	 if(facilityId.equals("01"))
		        		m_ItemDCQty.setString(2, "PORTLAND");
		         else if(facilityId.equals("04"))
		        		 m_ItemDCQty.setString(2, "PITTSTON");
	            rset = m_ItemDCQty.executeQuery();

	            if ( rset.next() )
	            	qty = rset.getInt("available_qty");
	         }
	         catch (Exception e) {
	            log.fatal("[VndPriceChange] getOnHand ", e);
	            m_ErrMsg.append("The report had the following Error(s) in: \r\n");
	            m_ErrMsg.append(e.getClass().getName() + "\r\n" + e.getMessage());
	         }
	         finally {
	            closeRSet(rset);
	            rset = null;
	         }
	      }

	      return qty;
   }

   /**
    * Gets the name of the vendor based on the internal vendor id.
    * @return The name of the vendor.
    * @throws SQLException
    */
   private String getVndName()
   {
      Statement stmt = null;
      ResultSet rs = null;
      String name = "";

      try {
         stmt = m_EdbConn.createStatement();
         rs = stmt.executeQuery(String.format("select name from vendor where vendor_id = %d", m_VendorId));

         if ( rs.next() )
            name = rs.getString("name");
      }
      catch ( Exception ex ) {
         log.error("[VndPriceChange]", ex);
      }

      finally {
         closeResultSet(rs);
         closeStatement(stmt);

         rs = null;
         stmt = null;
      }

      return name;
   }

   /**
    * Loads list of whse facility ids and names.
    *
    * @throws SQLException
    */
   private void loadFacilities()
   {
      ResultSet rs = null;
      Statement stm = null;
      StringBuffer tmp = new StringBuffer();

      try {
         if (m_WhseId == null || m_WhseId.equals("") || m_WhseId.equals("ALL")){
            tmp.append("select fas_facility_id, name from warehouse where fas_facility_id is not null");
         }
         else{
            tmp.append("select fas_facility_id, name from warehouse where fas_facility_id = '");
            tmp.append(m_WhseId);
            tmp.append("'");
         }

         m_FacilityList.clear();
         stm = m_EdbConn.createStatement();
         rs = stm.executeQuery(tmp.toString());

         while ( rs.next() )
            m_FacilityList.add(new Facility(rs.getString("name"), rs.getString("fas_facility_id")));
      }
      catch(Exception e) {
         log.fatal("[VndPriceChange] loadFacilities ", e);
         m_ErrMsg.append("The report had the following Error(s) in: \r\n");
         m_ErrMsg.append(e.getClass().getName() + "\r\n" + e.getMessage());
      }
      finally {
         closeRSet(rs);
         closeStatement(stm);
         rs = null;
      }
   }


   /**
    * Opens the output stream
    * @throws FileNotFoundException
    */
   private void openOutputStream()
   {
      try {
         m_FileNames.set(0, m_RptProc.getUid() + "-" + m_FileNames.get(0));
         m_OutputStream = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      } 
      catch (Exception e)
      {
         log.fatal("[VndPriceChange] openOutputStream ", e);
         m_ErrMsg.append("The report had the following Error(s) in: \r\n");
         m_ErrMsg.append(e.getClass().getName() + "\r\n" + e.getMessage());
      }
      
   }

   /**
    * Edb query to get sales data based on parameters
    * @return true if the statements are prepared, false if not.
    */
   private boolean prepareStatements()
   {
      boolean isPrepared = false;
      StringBuffer sql = new StringBuffer();

      if (  m_EdbConn != null ) {
         try {
            sql.append("select distinct");
            sql.append("   v.vendor_id, v.name, item_entity_attr.item_id, item_entity_attr.description, ");
            sql.append("   vic.vendor_item_num, item_entity_attr.retail_pack, ejd_item_warehouse.stock_pack, ");
            sql.append("   decode(bc.description, 'ALLOW BROKEN CASES', 'N', 'Y') as broken_case, ");
            sql.append("   su.unit, pc.description as product_category, pc.category_code, upc.upc_code, ");
            sql.append("   ip.sen_code_id, ejd_item.flc_id, ejd.IsInCatalog(item_entity_attr.item_id) in_catalog, ");
            sql.append("   ejd.item_price_procs.todays_buy(item_entity_attr.item_id) as cost, ");
            sql.append("   (SELECT pp.buy ");
            sql.append("      FROM ejd_pending_price pp ");
            sql.append("      WHERE pp.ejd_item_id = item_entity_attr.ejd_item_id AND ");
            sql.append("            pp.buy_date IN ");
            sql.append("            (SELECT MIN(pp2.buy_date) FROM ejd_pending_price pp2 ");
            sql.append("             WHERE pp2.buy_date > TRUNC(now) AND ");
            sql.append("               pp2.ejd_item_id = item_entity_attr.ejd_item_id AND ");
            sql.append("               pp2.approved_by IS NOT NULL) limit 1) AS new_cost, ");
            sql.append("   ip.sell as base, ");
            sql.append("   (SELECT  pp.sell ");
            sql.append("      FROM ejd_pending_price  pp ");
            sql.append("      WHERE  pp.ejd_item_id =  item_entity_attr.ejd_item_id AND ");
            sql.append("         pp.buy_date IN ");
            sql.append("         (SELECT MIN(pp2.buy_date) FROM ejd_pending_price pp2 ");
            sql.append("          WHERE pp2.buy_date > TRUNC(now) AND ");
            sql.append("              pp2.ejd_item_id = item_entity_attr.ejd_item_id AND ");
            sql.append("              pp2.approved_by IS NOT NULL) limit 1) AS new_base, ");
            sql.append("   decode(ip.sell, 0, 0, round(100 - 100 * (ejd.item_price_procs.todays_buy(item_entity_attr.item_id) / ip.sell), 2)) as margin, ");
            sql.append("   ip.retail_a, ip.retail_b, ip.retail_c, ip.retail_d ");
            sql.append("from ");
            sql.append("   item_entity_attr ");
            sql.append("   inner join ejd_item on item_entity_attr.ejd_item_id = ejd_item.ejd_item_id ");
            sql.append("   inner join ejd_item_warehouse on item_entity_attr.ejd_item_id = ejd_item_warehouse.ejd_item_id ");

            //
            //To determine active items,  get all active items only once associated
            //with selected facilities
            if (m_Active) {
               sql.append("   and ejd_item_warehouse.active = 1 ");
               sql.append("   inner join warehouse w on ejd_item_warehouse.warehouse_id = w.warehouse_id and  w.name in " + this.getFacilityCondition() );

            }
            sql.append("   inner join ejd_item_price ip on ip.ejd_item_id = item_entity_attr.ejd_item_id  and ip.warehouse_id = ejd_item_warehouse.warehouse_id ");
            sql.append("   inner join vendor v on item_entity_attr.vendor_id = v.vendor_id ");
            sql.append("   inner join vendor_item_ea_cross vic on v.vendor_id = vic.vendor_id and item_entity_attr.item_ea_id = vic.item_ea_id ");
            sql.append("   inner join broken_case bc on ejd_item.broken_case_id = bc.broken_case_id ");
            sql.append("   inner join ship_unit su on item_entity_attr.ship_unit_id = su.unit_id ");
            sql.append("   left outer join item_ea_product_category ipc on item_entity_attr.item_ea_id = ipc.item_ea_id ");
            sql.append("   left outer join product_category pc on ipc.prod_cat_id = pc.prod_cat_id ");
            sql.append("   left outer join ejd_item_whs_upc upc on item_entity_attr.ejd_item_id = upc.ejd_item_id and primary_upc = 1 and upc.warehouse_id = ejd_item_warehouse.warehouse_id ");
            sql.append("where ");
            sql.append("   v.vendor_id = ?  ");
            sql.append("order by v.name, item_entity_attr.item_id ");
            m_SqlVPC = m_EdbConn.prepareStatement(sql.toString());

            //
            //The qty break pricing statement.
            sql.setLength(0);
            sql.append("select percent, min_qty from item_ea_qty_discount ");
            sql.append(" join item_entity_attr on item_ea_qty_discount.item_ea_id = item_entity_attr.item_ea_id  ");
            sql.append(" where item_entity_attr.item_id = ? and packet_id is null order by min_qty ");
            m_QtyBreak = m_EdbConn.prepareStatement(sql.toString());

            //
            //The slot pricing statement.
            sql.setLength(0);
            sql.append("select price, min_qty from item_commodity_price where price_id = (select ejd.item_price_procs.todays_sell_id(?) ) order by price_level");
            m_SlotPrice = m_EdbConn.prepareStatement(sql.toString());

            //
            //Get maximum number of qty breaks to expect for this vendor (used to draw variable length column headers).
            sql.setLength(0);
            sql.append("select MAX(iqdcount.total) total   ");
            sql.append("  from (select count(iqd.min_qty) total  ");
            sql.append("from ");
            sql.append("   item_entity_attr i ");
            sql.append("   inner join item_ea_qty_discount iqd on i.item_ea_id = iqd.item_ea_id  ");

            //
            //To determine active items, use a subquery with distinct to get all active items only once associated
            //with selected facilities
            if (m_Active){
               sql.append("  inner join (select distinct iw1.item_id, iw1.active ");
               sql.append("              from item_warehouse iw1  ");
               sql.append("                   inner join warehouse w on iw1.warehouse_id = w.warehouse_id   ");
               sql.append("               where iw1.active = 1 and ");
               sql.append("                  w.name in " + this.getFacilityCondition() + ") iw on i.item_id = iw.item_id ");
            }

            sql.append("where ");
            sql.append("   i.vendor_id = '" + m_VendorId + "' ");
            sql.append("group by iqd.item_ea_id) iqdcount ");
            m_QtyBreakTotalSQL = sql.toString();

            //
            //Get maximum number of slot price breaks to expect for this vendor (used to draw variable length column headers).
            sql.setLength(0);
            sql.append("select MAX(icpcount.total) total ");
            sql.append(" from (select count(icp.min_qty) total ");
            sql.append("from ");
            sql.append("   item_entity_attr i ");
            sql.append("   inner join item_commodity_price icp on icp.price_id = (select ejd.item_price_procs.todays_sell_id(i.item_id) )");

            //
            //To determine active items, use a subquery with distinct to get all active items only once associated
            //with selected facilities
            if (m_Active){
               sql.append("   inner join (select distinct iw1.ejd_item_id, iw1.active  ");
               sql.append("               from ejd_item_warehouse iw1 ");
               sql.append("                  inner join warehouse w on iw1.warehouse_id = w.warehouse_id ");
               sql.append("               where iw1.active = 1 and ");
               sql.append("                  w.name in " + this.getFacilityCondition() + ") iw on i.ejd_item_id = iw.ejd_item_id ");
            }

            sql.append("where ");
            sql.append("   i.vendor_id = '" + m_VendorId + "' ");
            sql.append("group by icp.price_id) icpcount ");
            m_SlotPriceTotalSQL = sql.toString();

            //
            // Units sold
            sql.setLength(0);
            sql.append("select warehouse, sum(qty_shipped) as units_sold ");
            sql.append("from inv_dtl ");
            sql.append("where item_nbr = ? and ");
            sql.append("   warehouse in " + this.getFacilityCondition());
            sql.append("   and (invoice_date > (now - interval '1 year')) and invoice_date <= now ");
            sql.append("group by warehouse ");
            m_UnitsSold = m_EdbConn.prepareStatement(sql.toString());

            //
            // Item active by facility
            sql.setLength(0);
            sql.append("select w.name, eiw.active, eiw.in_catalog ");
            sql.append("from ejd_item_warehouse eiw ");
            sql.append("   inner join warehouse w on eiw.warehouse_id = w.warehouse_id  ");
            sql.append("   inner join item_entity_attr on item_entity_attr.ejd_item_id = eiw.ejd_item_id and item_type_id = 1 ");
            sql.append("where item_entity_attr.item_id = ? and ");
            sql.append("   w.name in " + this.getFacilityCondition());
            m_CanSell = m_EdbConn.prepareStatement(sql.toString());

            m_ItemDCQty = m_EdbConn.prepareStatement("select avail_qty as available_qty " +
             		"from ejd_item_warehouse " +
                    "join item_entity_attr on item_entity_attr.ejd_item_id = ejd_item_warehouse.ejd_item_id and item_type_id = 1 " +
             		"where item_entity_attr.item_id = ? and warehouse_id = (select warehouse_id from warehouse where name = ?) ");
            
            isPrepared = true;
            sql = null;
         }

         catch( Exception ex ) {
            log.fatal("exception: " + ex);
            m_ErrMsg.append(ex.getMessage());
         }
      }

      return isPrepared;
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      //
      //processes user parameters from EIS
      for (Param p : params) {
         if ( p.name.equals("vndId") )
            m_VendorId = Integer.parseInt(p.value);

         if ( p.name.equals("active"))
            m_Active = p.value.trim().equalsIgnoreCase("true");

         if ( p.name.equals("dc"))
            m_WhseId = p.value.trim();
      }

   }

   /**
    * Gets a facility, by name, from m_FacilityList (a convenience method).
    *
    * @param name - String facility name
    * @throws Exception
    */
   private Facility getFacilityByName(String name)
   {
      Facility facility = null;

      if (name == null || name.equals(""))
         return facility;

      for (Facility f : this.getFacilityList()){
         if (f.getName().equals(name))
            return f;
      }

      return facility;
   }

   /**
    * @return ArrayList<Facility> - list of facilities to report on.
    */
   public ArrayList<Facility> getFacilityList()
   {
      return m_FacilityList;
   }

   /**
    * @param facilityList - list of facilities to report on.
    */
   public void setFacilityList(ArrayList<Facility> facilityList)
   {
      m_FacilityList = facilityList;
   }

   /**
    * @return String - the facility sql condition snippet
    */
   public String getFacilityCondition()
   {
      return m_FacilityCondition;
   }

   /**
    * @param facilityCondition
    */
   public void setFacilityCondition(String facilityCondition)
   {
      m_FacilityCondition = facilityCondition;
   }
}
