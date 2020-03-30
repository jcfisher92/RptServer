/**
 * File: DisplayBreakdown.java
 * Description: Export for the display breakdown screen in EIS
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 10/29/2015
 * Last Update: 10/29/2015
 *
 * History:
 */

package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
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
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class DisplayBreakdown extends Report
{
   private PreparedStatement m_BrkData;
   private PreparedStatement m_CustRet;
   private PreparedStatement m_CustSell;

   private XSSFWorkbook m_Workbook;
   private Sheet m_Sheet;
   private Font m_FontNorm;
   private Font m_FontBold;
   private CellStyle m_StyleHdrLeft;
   private CellStyle m_StyleHdrRight;
   private CellStyle m_StyleDtlLeft;
   private CellStyle m_StyleDtlInt;
   private CellStyle m_StyleDtl2d;
   private CellStyle m_StyleDtl3d;

   private Row m_Row;
   private String m_VndId;
   private String m_VndName;
   private String m_CustId;
   private String m_CustName;
   private String m_DispNum;
   private String m_Desc;
   private int m_DispId;
   private int m_WhsId;

   public DisplayBreakdown()
   {
      super();
      
      m_WhsId = 1;
   }

   /**
    * adds a numeric type cell to current row at col p_Col in current sheet
    *
    * @param col     0-based column number of spreadsheet cell
    * @param value   numeric value to be stored in cell
    */
   private void addCell(int col, int value)
   {
      Cell cell = m_Row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellStyle(m_StyleDtlInt);
      cell.setCellValue(value);

      cell = null;
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
      cell.setCellValue(value);

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
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      boolean result = false;
      int rowNum = 0;
      int colNum = 0;
      FileOutputStream outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      ResultSet brkData = null;
      String itemId = "";
      int itemEaId = 0;

      try {
         rowNum = openWorkbook();

         m_BrkData.setInt(1, m_WhsId);
         m_BrkData.setInt(2, m_DispId);
         brkData = m_BrkData.executeQuery();

         while ( brkData.next() && m_Status == RptServer.RUNNING ) {
            addRow(rowNum);

            if ( m_Row != null ) {
               itemId = brkData.getString("item_id");
               itemEaId = brkData.getInt("item_ea_id");

               addCell(colNum++, itemId, m_StyleDtlLeft);
               addCell(colNum++, brkData.getString("description"), m_StyleDtlLeft);
               addCell(colNum++, brkData.getString("upc_code"), m_StyleDtlLeft);
               addCell(colNum++, brkData.getInt("quantity"));
               addCell(colNum++, brkData.getDouble("sell"), m_StyleDtl3d);
               addCell(colNum++, brkData.getDouble("retail_c"), m_StyleDtl2d);

               if ( m_CustId != null && m_CustId.length() == 6 ) {
                  addCell(colNum++, getCustSell(itemEaId), m_StyleDtl3d);
                  addCell(colNum++, getCustRetail(itemId), m_StyleDtl2d);
               }
               else
                  colNum += 2;

               addCell(colNum, brkData.getString("page"), m_StyleDtlLeft);
            }

            rowNum++;
            itemId = "";
            colNum = 0;
         }

         m_Workbook.write(outFile);
         brkData.close();
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[DisplayBreakdown]", ex);
      }

      finally {
         try {
            outFile.close();
         }

         catch( Exception ex ) {
            log.error("[DisplayBreakdown]", ex);
         }

         outFile = null;
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_BrkData);
      closeStmt(m_CustRet);
      closeStmt(m_CustSell);

      m_BrkData = null;
      m_CustRet = null;
      m_CustSell = null;
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
         if ( prepareStatements() )
            created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[DisplayBreakdown]", ex);
      }

      finally {
         closeStatements();

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
      //
      // Used to define a non-standard data format when
      // defining a style. For example, "0.00" is a bulit-in format, but "0.000"
      // and "0.0000" are custom formats.
      DataFormat customDataFormat = m_Workbook.createDataFormat();

      //
      // defines normal font
      m_FontNorm = m_Workbook.createFont();
      m_FontNorm.setFontName("Arial");
      m_FontNorm.setFontHeightInPoints((short)10);

      //
      // defines bold font
      m_FontBold = m_Workbook.createFont();
      m_FontBold.setFontName("Arial");
      m_FontBold.setFontHeightInPoints((short)10);
      m_FontBold.setBold(true);

      //
      // defines style column header, left-justified
      m_StyleHdrLeft = m_Workbook.createCellStyle();
      m_StyleHdrLeft.setFont(m_FontBold);
      m_StyleHdrLeft.setAlignment(HorizontalAlignment.LEFT);
      m_StyleHdrLeft.setVerticalAlignment(VerticalAlignment.TOP);

      m_StyleHdrRight = m_Workbook.createCellStyle();
      m_StyleHdrRight.setFont(m_FontBold);
      m_StyleHdrRight.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleHdrRight.setVerticalAlignment(VerticalAlignment.TOP);

      m_StyleDtlLeft = m_Workbook.createCellStyle();
      m_StyleDtlLeft.setFont(m_FontNorm);
      m_StyleDtlLeft.setAlignment(HorizontalAlignment.LEFT);
      m_StyleDtlLeft.setVerticalAlignment(VerticalAlignment.TOP);

      m_StyleDtlInt = m_Workbook.createCellStyle();
      m_StyleDtlInt.setFont(m_FontNorm);
      m_StyleDtlInt.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtlInt.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleDtlInt.setDataFormat((short)3);

      //
      // defines style detail data cell, right-justified with 2 decimal places
      //  (built-in data format)
      m_StyleDtl2d = m_Workbook.createCellStyle();
      m_StyleDtl2d.setFont(m_FontNorm);
      m_StyleDtl2d.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtl2d.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleDtl2d.setDataFormat(customDataFormat.getFormat("0.00"));

      //
      // defines style detail data cell, right-justified with 3 decimal places
      //  (custom data format)
      m_StyleDtl3d = m_Workbook.createCellStyle();
      m_StyleDtl3d.setFont(m_FontNorm);
      m_StyleDtl3d.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtl3d.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleDtl3d.setDataFormat(customDataFormat.getFormat("0.000"));
   }

   /**
    *
    * @param itemId
    * @return
    * @throws SQLException
    */
   private double getCustRetail(String itemId) throws SQLException
   {
      double retail = 0.0;
      ResultSet rs = null;

      try {
         m_CustRet.setString(1, m_CustId);
         m_CustRet.setString(2, itemId);

         rs = m_CustRet.executeQuery();

         if ( rs.next() )
            retail = rs.getDouble(1);
      }
      
      catch ( SQLException ex ) {
         log.error("[DisplayBreakdown]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return retail;
   }

   /**
    *
    * @param itemId
    * @return
    * @throws SQLException
    */
   private double getCustSell(Integer itemEaId) throws SQLException
   {
      double sell = 0.0;
      ResultSet rs = null;

      try {
         m_CustSell.setString(1, m_CustId);
         m_CustSell.setInt(2, itemEaId);

         rs = m_CustSell.executeQuery();

         if ( rs.next() )
            sell = rs.getDouble(1);
      }
      
      catch ( SQLException ex ) {
         log.error("[DisplayBreakdown]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return sell;
   }

   /**
    * opens the Excel spreadsheet and creates the title and the column headings
    *
    * @return  row number of first detail row (below header rows)
    */
   private int openWorkbook()
   {
      int col = 0;
      int row = 0;
      int charWidth = 295;

      m_Workbook = new XSSFWorkbook();
      m_Sheet = m_Workbook.createSheet();
      defineStyles();

      addRow(row++);
      addCell(col, "Description: " + m_Desc, m_StyleHdrLeft);

      addRow(row++);
      addCell(col, String.format("Vendor: %s  VndName: %s",  m_VndId, m_VndName), m_StyleHdrLeft);

      addRow(row++);
      addCell(col, "Vend Display#: " + m_DispNum, m_StyleHdrLeft);

      addRow(row++);
      addCell(col, String.format("Cust Nbr: %s  Cust Name: %s", m_CustId, m_CustName), m_StyleHdrLeft);

      col = 0;
      row++;
      addRow(row);

      m_Sheet.setColumnWidth(col, (7 * charWidth));
      addCell(col++, "Item Nbr", m_StyleHdrLeft);

      m_Sheet.setColumnWidth(col, (40 * charWidth));
      addCell(col++, "Item Desc", m_StyleHdrLeft);

      m_Sheet.setColumnWidth(col, (13 * charWidth));
      addCell(col++, "UPC", m_StyleHdrLeft);

      m_Sheet.setColumnWidth(col, (5 * charWidth));
      addCell(col++, "Qty", m_StyleHdrRight);

      m_Sheet.setColumnWidth(col, (11 * charWidth));
      addCell(col++, "Base Cost", m_StyleHdrRight);

      m_Sheet.setColumnWidth(col, (11 * charWidth));
      addCell(col++, "Retail C", m_StyleHdrRight);

      m_Sheet.setColumnWidth(col, (11 * charWidth));
      addCell(col++, "Cust Cost", m_StyleHdrRight);

      m_Sheet.setColumnWidth(col, (11 * charWidth));
      addCell(col++, "Cust Retail", m_StyleHdrRight);

      m_Sheet.setColumnWidth(col, (15 * charWidth));
      addCell(col, "Cat Page", m_StyleHdrLeft);

      return ++row;
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
            sql.append("   item_entity_attr.item_ea_id, item_entity_attr.item_id, upc_code, quantity, ");
            sql.append("   item_entity_attr.description, page, sell, retail_c ");
            sql.append("from display_brkdwn_item ");
            sql.append("join item_entity_attr on item_entity_attr.item_ea_id = display_brkdwn_item.item_ea_id ");
            sql.append("join warehouse on warehouse.warehouse_id = ? ");
            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_price.warehouse_id = warehouse.warehouse_id ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = item_entity_attr.ejd_item_id and ");
            sql.append("   ejd_item_whs_upc.warehouse_id = warehouse.warehouse_id and primary_upc = 1 ");
            sql.append("left outer join web_item_ea on web_item_ea.item_ea_id = display_brkdwn_item.item_ea_id ");
            sql.append("where display_brkdwn_item.display_brkdwn_id = ? ");
            sql.append("order by item_entity_attr.item_id");

            m_BrkData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select retail_price_procs.getretailprice(?, ?) as retail");
            m_CustRet = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select price from ejd_cust_procs.get_sell_price(?, ?)");
            m_CustSell = m_EdbConn.prepareStatement(sql.toString());

            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("[DisplayBreakdown]", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("[DisplayBreakdown] prepareStatements - null oracle connection");

      return isPrepared;
   }

   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();
      String tmp = Long.toString(System.currentTimeMillis());
      int pcount = params.size();
      Param param = null;

      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);

         if ( param.name.equals("vndid") )
            m_VndId = param.value;

         if ( param.name.equals("vndname") )
            m_VndName = param.value;

         if ( param.name.equals("custid") )
            m_CustId = param.value;

         if ( param.name.equals("custname") )
            m_CustName = param.value;

         if ( param.name.equals("dispnum") )
            m_DispNum = param.value;
         

         //
         // We'll let this fail with an exception if no ID comes in.  The
         // display ID is needed as the param in the query so might as well
         // stop it here if it doesn't exist.
         if ( param.name.equals("dispid") )
               m_DispId = Integer.parseInt(param.value);

         if (param.name.equals("desc") )
            m_Desc = param.value;
         
         if ( param.name.equals("whsid") ) {
            if ( param.value != null && param.value.length() > 0 )
               m_DispId = Integer.parseInt(param.value);
         }
      }

      fileName.append("display_brkdwn");
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
   }
   
   
   /* Main for debugging
   public static void main(String[] args) {
   	DisplayBreakdown cat = new DisplayBreakdown();

      Param p1 = new Param();
      p1.name = "vndid";
      p1.value = "703535";
      Param p2 = new Param();
      p2.name = "vndname";
      p2.value = "APEX TOOL GROUP,LLC";
      Param p3 = new Param();
      p3.name = "custid";
      p3.value = "059005";
      Param p4 = new Param();
      p4.name = "custname";
      p4.value = "LMC-ALLEN LUMBER CO #100";
      Param p5 = new Param();
      p5.name = "dispnum";
      p5.value = "CMHT38";
      Param p6 = new Param();
      p6.name = "dispid";
      p6.value = "1";
      ArrayList<Param> params = new ArrayList<Param>();
      params.add(p1);
      params.add(p2);
      params.add(p3);
      params.add(p4);
      params.add(p5);
      params.add(p6);
      
      cat.m_FilePath = "C:\\EXP\\";
      
   	java.util.Properties connProps = new java.util.Properties();
   	connProps.put("user", "ejd");
   	connProps.put("password", "boxer");
   	try {
   		cat.m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33:5444/emery_jensen",connProps);
   	} catch (Exception e) {
   		e.printStackTrace();
   	}
      
      cat.setParams(params);
      cat.createReport();
   }*/

}
