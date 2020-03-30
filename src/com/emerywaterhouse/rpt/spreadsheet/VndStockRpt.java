/**
 * File: VndStockRpt.java
 * Description: Class <code>VndStockRpt</code> creates the 'Vendor Stocking Report' - an Excel
 *    file of sales data for a selected date range and vendor, fine line class and/or item.
 *    
 *    Data is presented at the invoice detail level, including vendor, customer, fine
 *    line class and item identification and sales data including unit cost, sell and
 *    retail pricing, dollars shipped, units shipped and current units on-hand.
 *    
 *    Data is sorted by vendor, customer account, customer and item.
 *    
 *    Rewrite of the class to work with the new report server.
 *    Original author was Peter Peter de Zeeuw
 *
 * @author Peter de Zeeuw
 * @author Jeffrey Fisher
 *
 * Create Date: 05/23/2005
 * Last Update: $Id: VndStockRpt.java,v 1.13 2014/07/31 21:05:48 sgillis Exp $
 * 
 * History
 *    $Log: VndStockRpt.java,v $
 *    Revision 1.13  2014/07/31 21:05:48  sgillis
 *    removed shorts
 *
 *    Revision 1.12  2014/04/16 15:23:36  tli
 *    Output xlsx file
 *
 *    Revision 1.11  2014/04/15 20:28:46  tli
 *    Added ItemVelocityProject and PromoServiceLevel reports
 *
 *    Revision 1.10  2009/03/03 20:04:08  smurdock
 *    tweaked to slect off (partitioned) inv_dtl date instead of inv_hdr
 *
 *    Revision 1.9  2009/02/18 15:44:56  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    03/25/2005 - Added log4j logging. jcf
 * 
 *    09/30/2004 - removed unneeded cast to double. - jcf
 * 
 *    06/17/2004 - Modified to use 'max()' vs 'and rowid = 1' in sub-queries. - pdz
 *
 *    05/03/2004 - Removed the usage of the m_DistList member variable.  This
 *                 variable gets cleaned up before it can be used in the email
 *                 webservice. - jcf
 *
 *    04/07/2004 - Applied Email class changes. - jcf 
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class VndStockRpt extends Report
{  
   private PreparedStatement m_SqlSales;  // sales data based on selection parameters
   private Date m_FromDate;             // Start Date of selected date range  
   private Date m_ThruDate;             // End Date of selected date range
   private String m_VendorId;             // Selected EIS Vendor Number, may be "" to select all vendors   
   private String m_FlcId; 
   private String m_ItemId;
   private XSSFWorkbook m_Workbook;
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
   private CellStyle m_StyleDtlCntr;
   private CellStyle m_StyleDtlRght;
   private CellStyle m_StyleDtlRght2d;
   private CellStyle m_StyleDtlRght3d;
   private CellStyle m_StyleDtlRght4d;
   private FileOutputStream m_OutputStream;
      
   /**
    * default constructor
    */
   public VndStockRpt()
   {
      super();
      
      m_FileNames.add("VndStockRpt-" + String.valueOf(System.currentTimeMillis()) + ".xlsx");
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
      int index = 0;
      ResultSet rsSales = null;
      //ResultSet rsPrice = null;
      float emeryCost = -1;
      float baseCost = -1;
      float retailA = -1;
      float retailB = -1;
      float retailC = -1;
      float retailD = -1;
      String itemNbr = null;
      StringBuffer sb1 = new StringBuffer();
      
      row = openWorkbook();
      
      try {       
         // binds v_RS_Sales variable parameters (first = 1)
         if ( !m_FromDate.equals(m_ThruDate) ) {            
            m_SqlSales.setDate(++index, m_FromDate);
            m_SqlSales.setDate(++index, m_ThruDate);
         }
         else {            
            m_SqlSales.setDate(++index, m_FromDate);
         }
         
         if ( !m_VendorId.equals("") ) {            
            m_SqlSales.setString(++index, m_VendorId);            
         }
         
         if ( !m_FlcId.equals("") ) {            
            m_SqlSales.setString(++index, m_FlcId);
         }
         
         if ( !m_ItemId.equals("") ) {            
            m_SqlSales.setString(++index, m_ItemId);
         }
         
         // opens Oracle query result set
         rsSales = m_SqlSales.executeQuery();
         
         // processes each row returned from Oracle query result set
         while ( rsSales.next() && m_Status == RptServer.RUNNING ) {
            // gets value for item_nbr bind variables
            itemNbr = rsSales.getString("item_nbr");
            setCurAction("processing item: " + itemNbr);
            
            // initializes Price fields
            emeryCost = -1;
            baseCost = -1;
            retailA = -1;
            retailB = -1;
            retailC = -1;
            retailD = -1;
            
            
            
            emeryCost = rsSales.getFloat("emery_cost");
            baseCost = rsSales.getFloat("base_cost");
            retailA = rsSales.getFloat("retail_a");
            retailB = rsSales.getFloat("retail_b");
            retailC = rsSales.getFloat("retail_c");
            retailD = rsSales.getFloat("retail_d");
            
            
            // loads Oracle data into spread sheet cells for new row
            addRow(row);
            
            // column 0 (A), Vendor Name
            addCell(col++, rsSales.getString("vendor_name"), m_StyleDtlLeftWrap);
            
            // column 1 (B), Cust Number
            addCell(col++, rsSales.getString("cust_nbr"), m_StyleDtlLeft);
            
            // column 2 (C), Cust Name
            addCell(col++, rsSales.getString("cust_name"), m_StyleDtlLeftWrap);
            
            // column 3 (D), Ship To Addr 1
            addCell(col++, rsSales.getString("ship_to_addr_1"), m_StyleDtlLeftWrap);
            
            // column 4 (E), Ship To Addr 2
            addCell(col++, rsSales.getString("ship_to_addr_2"), m_StyleDtlLeftWrap);
            
            // column 5 (F), Ship To City
            addCell(col++, rsSales.getString("ship_to_city"), m_StyleDtlLeftWrap);
            
            // column 6 (G), Ship To State
            addCell(col++, rsSales.getString("ship_to_state"), m_StyleDtlCntr);
            
            // column 7 (H), Ship To Zip
            addCell(col++, rsSales.getString("ship_to_zip"), m_StyleDtlLeft);
            
            // column 8 (I), Phone Number
            sb1.append(rsSales.getString("phone_number"));
            
            if ( sb1.length() == 10 ) {
               sb1.insert(6, "-");
               sb1.insert(3, ") ");
               sb1.insert(0, "(");
            }
            
            addCell(col++, sb1.toString(), m_StyleDtlLeft);
            sb1.setLength(0);
            
            // column 9 (J), Fax Number
            sb1.append(rsSales.getString("fax_number"));
            if ( sb1.length() == 10 ) {
               sb1.insert(6, "-");
               sb1.insert(3, ") ");
               sb1.insert(0, "(");
            }
            
            addCell(col++, sb1.toString(), m_StyleDtlLeft);
            sb1.setLength(0);
            
            // column 10 (K), Territory Manager
            addCell(col++, rsSales.getString("territory_manager"), m_StyleDtlLeftWrap);
            
            // column 11 (L), Emery Item Nbr
            addCell(col++, itemNbr, m_StyleDtlCntr);
            
            // column 12 (M), Vendor Part Nbr
            addCell(col++, rsSales.getString("vendor_item_num"), m_StyleDtlLeft);
            
            // column 13 (N), Ship Unit
            addCell(col++, rsSales.getString("ship_unit"), m_StyleDtlCntr);
            
            // column 14 (O), Dealer Pack
            addCell(col++, rsSales.getInt("stock_pack"), m_StyleDtlCntr);
            
            // column 15 (P), Primary UPC
            addCell(col++, rsSales.getString("upc_code"), m_StyleDtlLeft);
            
            // column 16 (Q), Item Description
            addCell(col++, rsSales.getString("item_descr"), m_StyleDtlLeftWrap);
            
            // column 17 (R), Emery Cost
            if ( emeryCost >= 0 ) {
               addCell(col, emeryCost, m_StyleDtlRght4d);
            }
            col++;
            
            // column 18 (S), Base Cost
            if ( baseCost >= 0 ) {
               addCell(col, baseCost, m_StyleDtlRght3d);
            }
            col++;
            
            // column 19 (T), A Mkt Retail
            if ( retailA >= 0 ) {
               addCell(col, retailA, m_StyleDtlRght2d);
            }
            col++;
            
            // column 20 (U), B Mkt Retail
            if ( retailB >= 0 ) {
               addCell(col, retailB, m_StyleDtlRght2d);
            }
            col++;
            
            // column 21 (V), C Mkt Retail
            if ( retailC >= 0 ) {
               addCell(col, retailC, m_StyleDtlRght2d);
            }
            col++;
            
            // column 22 (W), D Mkt Retail
            if ( retailD >= 0 ) {
               addCell(col, retailD, m_StyleDtlRght2d);
            }
            col++;
            
            // column 23 (X), Sens Code
            addCell(col++, rsSales.getString("rtr_sensitivity"), m_StyleDtlCntr);
            
            // column 24 (Y), Units Sold
            addCell(col++, rsSales.getInt("qty_shipped"), m_StyleDtlRght);
            
            // column 25 (Z), Dollars Sold
            addCell(col++, rsSales.getFloat("ext_sell"), m_StyleDtlRght2d);
            
            // column 26 (AA), Units On Hand
            addCell(col++, rsSales.getInt("ptld_on_hand"), m_StyleDtlRght);
            
            // column 27 (AB), Units On Hand
            addCell(col++, rsSales.getInt("pitt_on_hand"), m_StyleDtlRght);
            
            // column 28 (AC), Fine Line Class
            addCell(col++, rsSales.getString("flc"), m_StyleDtlCntr);
            
            row++;
            col = 0;
         }
         
         closeWorkbook();
         result = true;         
      }
      
      catch ( Exception ex ) {
         log.error("[VndStockRpt]", ex);           
         m_ErrMsg.append(ex.getMessage());
      }
      
      finally {
         if ( rsSales != null ) {
            try {               
               rsSales.close();
               rsSales = null;
            }
            
            catch ( Exception ex ) {
               
            }
         }
      }
      
      return result;
   }
   
   /*
    * cleans up member variables
    */
   protected void cleanup()
   {
      // closes prepared statements
      closePreparedStatement(m_SqlSales);
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
         log.error("[VndStockRpt]", ex);
      }
      
      finally {       
         m_OutputStream = null;
      }
   }
   
   /**
    * closes a single prepared statement identified by p_Statement
    *
    * @param p_Statement   name of prepared statement to be closed
    */
   private void closePreparedStatement(PreparedStatement p_Statement)
   {
      try {
         if (p_Statement != null)
            p_Statement.close();
      }
      
      catch ( Exception ex ) {
         
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
   private void closeWorkbook() throws IOException
   {
      //
      // Don't write anything to disk if the report has been stopped.
      if ( m_Status == RptServer.RUNNING ) {
         setCurAction("writing excel spreadsheet");
         m_Workbook.write(m_OutputStream);
      }
      
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
         
         if ( prepareStatements() ) {
            created = buildWorkbook();            
         }
         
         setCurAction("complete");
      }
      
      catch ( Exception ex ) {
         log.fatal("[VndStockRpt]", ex);
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
      
      // defines normal font
      m_FontNorm = m_Workbook.createFont();
      m_FontNorm.setFontName("Arial");
      m_FontNorm.setFontHeightInPoints((short)10);
      
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
      
      // defines style detail data cell, center-justified
      m_StyleDtlCntr = m_Workbook.createCellStyle();
      m_StyleDtlCntr.setFont(m_FontNorm);
      m_StyleDtlCntr.setAlignment(HorizontalAlignment.CENTER);
      m_StyleDtlCntr.setVerticalAlignment(VerticalAlignment.TOP);
      
      // defines style detail data cell, right-justified
      m_StyleDtlRght = m_Workbook.createCellStyle();
      m_StyleDtlRght.setFont(m_FontNorm);
      m_StyleDtlRght.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtlRght.setVerticalAlignment(VerticalAlignment.TOP);
      
   // m_CustomDataFormat is used to define a non-standard data format when
      // defining a style. For example, "0.00" is a bulit-in format, but "0.000"
      // and "0.0000" are custom formats.
      DataFormat m_CustomDataFormat;
      m_CustomDataFormat = m_Workbook.createDataFormat();
      
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
    * opens the output stream
    * @throws FileNotFoundException 
    */
   private void openOutputStream() throws FileNotFoundException
   {
      m_FileNames.set(0, m_RptProc.getUid() + "-" + m_FileNames.get(0));
      m_OutputStream = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
   }
   
   /**
    * opens the Excel spreadsheet and creates the title and the column headings
    *
    * @return  row number of first detail row (below header rows)
    */
   private int openWorkbook()
   {
      int col = 0;
      int m_CharWidth = 295;
      
      // creates workbook
      m_Workbook = new XSSFWorkbook();
      
      // creates sheet 0
      m_Sheet = m_Workbook.createSheet("VndStockRpt");
      
      // defines styles
      defineStyles();
      
      // creates Excel title
      addRow(0);
      StringBuffer hdr = new StringBuffer();
      hdr.append("Vendor Stocking Report for ");
      hdr.append(m_FromDate);
      
      if (!m_ThruDate.equals(m_FromDate)) {
         hdr.append(" - ");
         hdr.append(m_ThruDate);
         
         if (!m_VendorId.equals("")) {
            hdr.append("  Vendor ");
            hdr.append(m_VendorId);
         }
         
         if (!m_FlcId.equals("")) {
            hdr.append("  FLC ");
            hdr.append(m_FlcId);
         }
         
         if (!m_ItemId.equals("")) {
            hdr.append("  Item ");
            hdr.append(m_ItemId);
         }
      }
      
      addCell(col, hdr.toString(), m_StyleHdrLeft);
      hdr = null;
      
      // creates Excel column headings
      addRow(2);
      
      //
      // Computes approximate HSSF character width based on "Arial" size "10" and
      // HHSF characteristics, to be used as a multiplier to set column widths.
      //      
      // column 0 (A), Vendor Name
      m_Sheet.setColumnWidth(col, (30 * m_CharWidth));
      addCell(col, "Vendor Name", m_StyleHdrLeftWrap);
      
      // column 1 (B), Cust Number
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "Cust Number", m_StyleHdrLeftWrap);
      
      // column 2 (C), Cust Name
      m_Sheet.setColumnWidth(++col, (25 * m_CharWidth));
      addCell(col, "Cust Name", m_StyleHdrLeftWrap);
      
      // column 3 (D), Ship To Addr 1
      m_Sheet.setColumnWidth(++col, (25 * m_CharWidth));
      addCell(col, "Ship To Address", m_StyleHdrLeftWrap);
      
      // column 4 (E), Ship To Addr 2
      m_Sheet.setColumnWidth(++col, (20 * m_CharWidth));
      addCell(col, "Address 2", m_StyleHdrLeftWrap);
      
      // column 5 (F), Ship To City
      m_Sheet.setColumnWidth(++col, (16 * m_CharWidth));
      addCell(col, "City", m_StyleHdrLeftWrap);
      
      // column 6 (G), Ship To State
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "State", m_StyleHdrCntr);
      
      // column 7 (H), Ship To Zip
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "Zip", m_StyleHdrLeft);
      
      // column 8 (I), Phone Number
      m_Sheet.setColumnWidth(++col, (11 * m_CharWidth));
      addCell(col, "Phone Number", m_StyleHdrLeftWrap);
      
      // column 9 (J), Fax Number
      m_Sheet.setColumnWidth(++col, (11 * m_CharWidth));
      addCell(col, "Fax Number", m_StyleHdrLeftWrap);
      
      // column 10 (K), Territory Manager
      m_Sheet.setColumnWidth(++col, (13 * m_CharWidth));
      addCell(col, "Territory Manager", m_StyleHdrLeftWrap);
      
      // column 11 (L), Emery Item Nbr
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "Emery Item Nbr", m_StyleHdrCntrWrap);
      
      // column 12 (M), Vendor Part Nbr
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "Vendor Part Nbr", m_StyleHdrLeftWrap);
      
      // column 13 (N), Ship Unit
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "Ship Unit", m_StyleHdrCntrWrap);
      
      // column 14 (O), Dealer Pack
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "Dealer Pack",   m_StyleHdrCntrWrap);
      
      // column 15 (P), Primary UPC
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Primary UPC",     m_StyleHdrLeftWrap);
      
      // column 16 (Q), Item Description
      m_Sheet.setColumnWidth(++col, (45 * m_CharWidth));
      addCell(col, "Item Description", m_StyleHdrLeftWrap);
      
      // column 17 (R), Emery Cost
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "Emery Cost",  m_StyleHdrRghtWrap);
      
      // column 18 (S), Base Cost
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "Base Cost", m_StyleHdrRghtWrap);
      
      // column 19 (T), A-Market Retail
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "A-Market Retail", m_StyleHdrRghtWrap);
      
      // column 20 (U), B-Market Retail
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "B-Market Retail", m_StyleHdrRghtWrap);
      
      // column 21 (V), C-Market Retail
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "C-Market Retail", m_StyleHdrRghtWrap);
      
      // column 22 (W), D-Market Retail
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "D-Market Retail", m_StyleHdrRghtWrap);
      
      // column 23 (X), Sens Code
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "Sens Code", m_StyleHdrCntrWrap);
      
      // column 24 (Y), Units Sold
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "Units Sold", m_StyleHdrRghtWrap);
      
      // column 25 (Z), Dollars Sold
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "Dollars Sold", m_StyleHdrRghtWrap);
      
      // column 26 (AA), On Hand
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "On Hand Ptld", m_StyleHdrRghtWrap);
      
      // column 27 (AB), On Hand
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "On Hand Pitt", m_StyleHdrRghtWrap);
      
      // column 28 (AC), FLC
      m_Sheet.setColumnWidth(++col, (4 * m_CharWidth));
      addCell(col, "FLC", m_StyleHdrCntrWrap);
      
      // returns first data row number
      return 3;
   }
   
   /**
    * Oracle query to get sales data based on parameters
    * @return true if the statements are prepared, false if not.
    */
   private boolean prepareStatements()
   {
      boolean isPrepared = false;
      StringBuffer sql = new StringBuffer();
      
      if ( m_EdbConn != null ) {
         try {
            sql.append("select ");
            sql.append("inv_dtl.vendor_name, ");
            sql.append("inv_hdr.cust_nbr, ");
            sql.append("inv_hdr.cust_name, ");
            sql.append("ship_to_addr_1, ");
            sql.append("coalesce(ship_to_addr_2, ' ') as ship_to_addr_2, ");
            sql.append("ship_to_city, ");
            sql.append("ship_to_state, ");
            sql.append("ship_to_zip, ");
            sql.append("coalesce(( ");
            sql.append("   select phone_number ");
            sql.append("   from cust_contact_phone_view ");
            sql.append("   where ");
            sql.append("      cust_contact_phone_view.customer_id = inv_hdr.cust_nbr and ");
            sql.append("      cust_contact_phone_view.type = 'BUSINESS' and rownum = 1 ");
            sql.append("), ' ') as phone_number, ");
            sql.append("coalesce (( ");
            sql.append("   select phone_number ");
            sql.append("   from cust_contact_phone_view ");
            sql.append("   where ");
            sql.append("      cust_contact_phone_view.customer_id = inv_hdr.cust_nbr and ");
            sql.append("      cust_contact_phone_view.type = 'BUSINESS FAX' and rownum = 1 ");
            sql.append("), ' ') as fax_number, ");
            sql.append("repname as territory_manager, ");
            sql.append("inv_dtl.item_nbr, ");
            sql.append("vendor_item_num, ");
            sql.append("inv_dtl.ship_unit, ");
            sql.append("inv_dtl.stock_pack, ");
            //sql.append("upc_code, ");
            sql.append("(select upc_code from ejd_item_whs_upc where ejd_item_id = iea.ejd_item_id ");
            sql.append("and warehouse_id = war.warehouse_id order by primary_upc desc limit 1) as upc_code, ");
            sql.append("inv_dtl.item_descr, ");
            sql.append("inv_dtl.rtr_sensitivity, ");
            sql.append("qty_shipped, ");
            sql.append("ext_sell, ");
            sql.append("ptld_whs.qoh as ptld_on_hand, ");
            sql.append("pitt_whs.qoh as pitt_on_hand, ");
            //sql.append("inv_dtl.flc ");
            sql.append("inv_dtl.flc, ");
            sql.append("coalesce(eip.buy, -1) emery_cost, ");
            sql.append("coalesce(eip.sell, -1) base_cost, ");
            sql.append("coalesce(eip.retail_a, -1) retail_a, ");
            sql.append("coalesce(eip.retail_b, -1) retail_b, ");
            sql.append("coalesce(eip.retail_c, -1) retail_c, ");
            sql.append("coalesce(eip.retail_d, -1) retail_d ");
            sql.append("from inv_hdr ");
            sql.append("join inv_dtl on inv_dtl.inv_hdr_id = inv_hdr.inv_hdr_id ");
            sql.append("join item_entity_attr iea on iea.item_id = inv_dtl.item_nbr and iea.item_type_id < 8 ");
            //sql.append("left outer join item_upc on item_upc.item_id = inv_dtl.item_nbr and primary_upc = 1 ");
            //sql.append("join item_warehouse ptld_whs on ptld_whs.item_id = inv_dtl.item_nbr and ptld_whs.warehouse_id = 1 ");
            //sql.append("join item_warehouse pitt_whs on pitt_whs.item_id = inv_dtl.item_nbr and pitt_whs.warehouse_id = 2 ");
            //sql.append("join vendor_item_cross on vendor_item_cross.item_id = inv_dtl.item_nbr and vendor_item_cross.vendor_id = inv_dtl.vendor_nbr ");
            sql.append("join ejd_item_warehouse ptld_whs on ptld_whs.ejd_item_id = iea.ejd_item_id and ptld_whs.warehouse_id = 1 ");
            sql.append("join ejd_item_warehouse pitt_whs on pitt_whs.ejd_item_id = iea.ejd_item_id and pitt_whs.warehouse_id = 2 ");
            sql.append("join vendor_item_ea_cross vic on vic.item_ea_id = iea.item_ea_id and vic.vendor_id = inv_dtl.vendor_nbr ");
            sql.append("join warehouse war on war.name = inv_hdr.warehouse and war.warehouse_id in (1, 2) ");
            sql.append("join cust_rep_div_view on cust_rep_div_view.customer_id = inv_hdr.cust_nbr and rep_type = 'SALES REP' ");
            sql.append("left join ejd_item_price eip on iea.ejd_item_id = eip.ejd_item_id and eip.warehouse_id = war.warehouse_id ");
            sql.append("where ");
            sql.append("inv_dtl.qty_shipped > 0 ");
                                    
            if ( !m_FromDate.equals(m_ThruDate) )
               sql.append(" and inv_hdr.invoice_date between ? and ? ");            
            else
               sql.append(" and inv_hdr.invoice_date = ? ");
                        
            if ( m_VendorId.length() > 0 )
               sql.append(" and inv_dtl.vendor_nbr = ? ");
                        
            if ( m_FlcId.length() > 0 )
               sql.append(" and inv_dtl.flc = ? ");
                        
            if ( m_ItemId.length() > 0 )
               sql.append(" and inv_dtl.item_nbr = ? ");
                        
            sql.append("order ");
            sql.append("   by inv_dtl.vendor_name, inv_hdr.cust_acct, inv_hdr.cust_nbr, inv_dtl.item_nbr");
            
            m_SqlSales = m_EdbConn.prepareStatement(sql.toString());
            m_SqlSales.setFetchSize(500);
                        
            //sql.setLength(0);
            //sql.append("select");
            //sql.append(" coalesce(buy, -1) emery_cost,");
            //sql.append(" coalesce(sell, -1) base_cost,");
            //sql.append(" coalesce(retail_a, -1) retail_a,");
            //sql.append(" coalesce(retail_b, -1) retail_b,");
            //sql.append(" coalesce(retail_c, -1) retail_c,");
            //sql.append(" coalesce(retail_d, -1) retail_d");
            //sql.append(" from ejd_item_price eip ");
            //sql.append(" inner join item_entity_attr iea on iea.ejd_item_id = eip.ejd_item_id and iea.item_type_id < 8 ");
            //sql.append(" where eip.warehouse_id in (1,2) ");
            //sql.append(" and iea.item_id = ? ");
            //sql.append(" limit 1 ");

            //sql.append(" from item_price");
            //sql.append(" where item_id = ?");
            //sql.append(" and sell_date = ");
            //sql.append(" (select max(sell_date)");
            //sql.append("  from item_price");
            //sql.append("  where item_id = ?");
            //sql.append("  and merch_initial is not null");
            //sql.append("  and sell_date <= trunc(now))");
            //m_SqlPrice = m_EdbConn.prepareStatement(sql.toString());
            
            isPrepared = true;
            sql = null;            
         }
         
         catch( Exception ex ) {
            log.fatal("[VndStockRpt]", ex);
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
      SimpleDateFormat dtFmt = new SimpleDateFormat("dd-MMM-yyyy");
      
      try {
         m_FromDate = new Date(dtFmt.parse(params.get(0).value.trim()).getTime());
         m_ThruDate = new Date(dtFmt.parse(params.get(1).value.trim()).getTime());
         
         m_VendorId = params.get(2).value;      
         m_FlcId = params.get(3).value;
         m_ItemId = params.get(4).value;
      }
      
      catch ( Exception ex ) {
         log.fatal("[VndStockRpt]", ex);
      }
      
      finally {
         dtFmt = null;
      }           
   }
}
