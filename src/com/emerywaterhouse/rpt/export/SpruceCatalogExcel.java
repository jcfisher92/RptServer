/**
 * File: SpruceCatalogExcel.java
 * Description: Exports the catalog data in Spruce.net format.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 01/22/2013
 * Last Update: 2016/12/14 jfisher
 *
 * History
 *    Refactored the name 2016/12/14 jfisher
 *    
 *    $Log: SpruceCatalogExcel.java,v $
 *    Revision 1.2  2013/02/19 14:26:48  jfisher
 *    changes for rp johnson
 *
 *    Revision 1.1  2013/02/07 13:46:06  jfisher
 *    Initial Add
 *
 */
package com.emerywaterhouse.rpt.export;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class SpruceCatalogExcel extends Report
{
   private static int maxCols = 18;
   //
   // workbook entries.
   private HSSFWorkbook m_Wrkbk;
   private HSSFSheet m_Sheet;
   private HSSFFont m_FontBold;
   private HSSFFont m_FontNormal;

   //
   // The cell styles for each of the base columns in the spreadsheet.
   private HSSFCellStyle[] m_CellStyles;

   //
   // params
   private String m_CustId;
   private String m_VndCode;

   private PreparedStatement m_CatData;

   /**
    *
    */
   public SpruceCatalogExcel()
   {
      super();

      m_CustId = "";
      m_Wrkbk = new HSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();

      setupWorkbook();
   }


   /**
    * Clean up any resources
    *
    * @see java.lang.Object#finalize()
    */
   @Override
   public void finalize() throws Throwable
   {
      if ( m_CellStyles != null ) {
         for ( int i = 0; i < m_CellStyles.length; i++ )
            m_CellStyles[i] = null;
      }

      m_Sheet = null;
      m_Wrkbk = null;
      m_CellStyles = null;

      super.finalize();
   }

   /**
    * Executes the queries and builds the output file
    *
    * @return true if the report was successfully built
    * @throws FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      HSSFRow row = null;
      int rowNum = 0;
      int colNum = 0;
      FileOutputStream outFile = null;
      boolean result = false;
      ResultSet catData = null;
      String itemId = null;
      String vndSku = null;
      int vndId = 0;
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         m_CatData.setString(1, m_CustId);
         catData = m_CatData.executeQuery();
         m_CurAction = "Building output file";

         while ( catData.next() && getStatus() != RptServer.STOPPED ) {
            row = createRow(rowNum++, maxCols);
            colNum = 0;
            itemId = catData.getString("item_id");
            vndId = catData.getInt("vendor_id");
            vndSku = catData.getString("vendor_item_num");

            m_CurAction = String.format("procesing item %s", itemId);

            //
            // Spruce's model number is a 16 char max.
            if ( vndSku.length() > 16 )
               vndSku = vndSku.substring(0, 16);

            row.getCell(colNum++).setCellValue(new HSSFRichTextString(itemId));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(m_VndCode));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(catData.getString("description")));
            row.getCell(colNum++).setCellValue(catData.getDouble("listprice"));
            row.getCell(colNum++).setCellValue(catData.getDouble("unitprice"));
            row.getCell(colNum++).setCellValue(catData.getDouble("unitcost"));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(catData.getString("vendorum")));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(catData.getString("altum")));
            row.getCell(colNum++).setCellValue(catData.getInt("AltUMQtyConv"));
            row.getCell(colNum++).setCellValue(catData.getInt("AltUMQtyConvTo"));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(catData.getString("upc_code")));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(Integer.toString(vndId)));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(vndSku));
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(catData.getString("nrha_id"))); // department
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(catData.getString("mdc_id")));  // class
            row.getCell(colNum++).setCellValue(new HSSFRichTextString(catData.getString("flc_id")));  // fineline
            row.getCell(colNum++).setCellValue(catData.getDouble("weight"));
            row.getCell(colNum++).setCellValue(catData.getInt("order_multiple"));
         }

         m_Wrkbk.write(outFile);
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append(String.format("The Spruce catalog export for customer %s had the following errors: \r\n", m_CustId));
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[SpruceCatalogExcel]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, m_CatData, catData);
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {

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
         m_OraConn = m_RptProc.getOraConn();
         if ( prepareStatements() )
            created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[SpruceCatalogExcel]", ex);
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
    * @param maxCols The number of columns in the row.
    *
    * @return The formatted row of the spreadsheet.
    * @throws Exception
    */
   private HSSFRow createRow(int rowNum, int maxCols) throws Exception
   {
      HSSFRow row = null;
      HSSFCell cell = null;

      if ( m_Sheet != null ) {
         row = m_Sheet.createRow(rowNum);

         //
         // set the type and style of the cell.
         if ( row != null ) {
            for ( int i = 0; i < maxCols; i++ ) {
               cell = row.createCell(i);
               cell.setCellStyle(m_CellStyles[i]);
            }
         }
      }
      else
         throw new Exception("null worksheet");

      return row;
   }

   /**
    * Prepares the sql queries for execution.
    *
    * @return true if the statements were successfully prepared
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;

      if ( m_OraConn != null ) {
         try {
            sql.setLength(0);
            sql.append("select ");
            sql.append("   item.item_id, ");
            sql.append("   decode(bmi_item.web_descr, null, item.description, bmi_item.web_descr) description, ");
            sql.append("   cust_procs.getretailprice(cust_warehouse.customer_id, item.item_id) listprice, ");
            sql.append("   cust_procs.getretailprice(cust_warehouse.customer_id, item.item_id) unitprice, ");
            sql.append("   cust_procs.getsellprice(cust_warehouse.customer_id, item.item_id) unitcost, ");
            sql.append("   ship_unit.unit VendorUM, ship_unit.unit altUM, 1 AltUMQtyConv, 1 AltUMQtyConvTo, ");
            sql.append("   item_upc.upc_code, item.vendor_id, vendor_item_cross.vendor_item_num, ");
            sql.append("   mdc.nrha_id, mdc.mdc_id, flc.flc_id, item.weight, ");
            sql.append("   decode(broken_case.description, 'ALLOW BROKEN CASES', 1, item.stock_pack) order_multiple ");
            sql.append("from cust_warehouse ");
            sql.append("join item_warehouse on item_warehouse.warehouse_id = cust_warehouse.warehouse_id and item_warehouse.in_catalog = 1 ");
            sql.append("join item on item.item_id = item_warehouse.item_id ");
            sql.append("join vendor_item_cross on vendor_item_cross.vendor_id = item.vendor_id and vendor_item_cross.item_id = item.item_id ");
            sql.append("left outer join item_upc on item_upc.item_id = item.item_id and item_upc.primary_upc = 1 ");
            sql.append("join flc on flc.flc_id = item.flc_id ");
            sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
            sql.append("join ship_unit on ship_unit.unit_id = item.ship_unit_id ");
            sql.append("join broken_case on broken_case.broken_case_id = item.broken_case_id ");
            sql.append("left outer join bmi_item on bmi_item.item_id = item.item_id ");
            sql.append("where customer_id = ? ");
            sql.append("order by item.item_id");

            m_CatData = m_OraConn.prepareStatement(sql.toString());

            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("[SpruceCatalogExcel]", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("[SpruceCatalogExcel] - null oracle connection");

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
      int pcount = params.size();
      Param param = null;

      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);

         if ( param.name.equalsIgnoreCase("custid") )
            m_CustId = param.value;

         if ( param.name.equalsIgnoreCase("vndcode") )
            m_VndCode = param.value;
      }

      fileName.append(".xls");
      m_FileNames.add(String.format("%d-%s-emery-catalog.xls",System.currentTimeMillis(), m_CustId));
   }

   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {
      HSSFCellStyle styleText;      // Text left justified
      HSSFCellStyle styleInt;       // Style with 0 decimals
      HSSFCellStyle styleDec;       // 2 decimal positions
      HSSFCellStyle styleDec3;      // 3 decimal positions
      HSSFDataFormat fmt;
      //
      // Create a font that is normal size & bold
      m_FontBold = m_Wrkbk.createFont();
      m_FontBold.setFontHeightInPoints((short)8);
      m_FontBold.setFontName("Arial");
      m_FontBold.setBold(true);

      //
      // Create a font that is normal size & bold
      m_FontNormal = m_Wrkbk.createFont();
      m_FontNormal.setFontHeightInPoints((short)8);
      m_FontNormal.setFontName("Arial");

      styleText = m_Wrkbk.createCellStyle();
      styleText.setAlignment(HorizontalAlignment.LEFT);
      styleText.setFont(m_FontNormal);

      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);
      styleInt.setFont(m_FontNormal);

      styleDec = m_Wrkbk.createCellStyle();
      styleDec.setAlignment(HorizontalAlignment.RIGHT);
      styleDec.setDataFormat((short)4);
      styleDec.setFont(m_FontNormal);

      fmt = m_Wrkbk.createDataFormat();
      styleDec3 = m_Wrkbk.createCellStyle();
      styleDec3.setAlignment(HorizontalAlignment.RIGHT);
      styleDec3.setDataFormat(fmt.getFormat("#.000"));
      styleDec3.setFont(m_FontNormal);

      m_CellStyles = new HSSFCellStyle[] {
         styleText,     // col 0 vendor sku
         styleText,     // col 1 vendor code
         styleText,     // col 2 description
         styleDec,      // col 3 List price
         styleDec,      // col 4 Unit price
         styleDec3,     // col 5 cost
         styleText,     // col 6 vendor uom
         styleText,     // col 7 alt uom
         styleInt,      // col 8 alt uom conv
         styleInt,      // col 9 alt uom conv to
         styleText,     // col 10 upc
         styleText,     // col 11 vendor id
         styleText,     // col 12 model number (16 char only)
         styleText,     // col 13 dept
         styleText,     // col 14 class
         styleText,     // col 15 fine line
         styleDec,      // col 16 weight
         styleInt       // col 17 order multiple
      };

      styleText = null;
      styleInt = null;
      styleDec = null;
   }
}
