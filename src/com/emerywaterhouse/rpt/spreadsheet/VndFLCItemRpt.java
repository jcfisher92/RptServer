/**
 * File: VndFLCItemRpt.java
 * Description: Class <code>VndFLCItemRpt</code> creates the 'Vendor FLC Item Report' - an Excel
 *    file of items for a select vendor and/or flc class.
 *
 *    Data includes vendor id, vendor, item id, vendor item number, item description, Emery cost,
 *    Emery base, Retails A, B, C, and D, ship unit, retail unit, stock pack, UPC, Sensitivity Code,
 *    FLC, In-Catalog, and NBC.
 *
 *    Data is sorted by vendor and item.
 *
 *    Rewrite of the class to work with the new report server.
 *    Original author was Peter de Zeeuw
 *
 * @author Seth Murdock
 *
 *
 * Create Date: 02/20/2006
 * Last Update: $Id: VndFLCItemRpt.java,v 1.25 2013/09/11 14:25:42 tli Exp $
 *
 * History:
 *    $Log: VndFLCItemRpt.java,v $
 *    Revision 1.25.0000000001 2017/7/26 1:19 sjaguilar
 *    Removed velocity so this works on EDB not Oracle
 *    
 *    Revision 1.25  2013/09/11 14:25:42  tli
 *    Converted the facilityId to facilityName when needed
 *
 *    Revision 1.24  2013/09/09 18:33:38  tli
 *    Replace SkuQty web service call with item_qty_view
 *
 *    Revision 1.23  2012/10/29 14:06:27  jfisher
 *    Updated the default timeout.
 *
 *    Revision 1.22  2012/10/11 14:23:13  jfisher
 *    Changes to deal with the timeout on the sku quantity web service.
 *
 *    Revision 1.21  2012/08/29 19:53:02  jfisher
 *    Switched web service calls from Wasp to Axis2
 *
 *    Revision 1.20  2011/09/20 13:43:43  smurdock
 *    removed  sytem output test line
 *
 *    Revision 1.19  2010/07/25 03:09:08  epearson
 *    Added Item Type column to id vendors as primary or secondary to an item
 *
 *    Revision 1.18  2010/07/13 14:11:12  epearson
 *    Changed style of "Emery Cost" field to show 4 decimal places.
 *
 *    Revision 1.17  2009/02/17 22:30:08  jfisher
 *    Added the accpac average cost and made some other code mods to clean things up.
 *
 *    Revision 1.16  2008/08/11 23:26:56  jfisher
 *    removed unused import
 *
 *    Revision 1.15  2008/08/06 01:02:15  jheric
 *    Fixed bug handling warehouses.
 *
 *    Revision 1.14  2008/08/04 22:20:10  jheric
 *    Move warehouse class to external library(in wsbeans).
 *
 *    Revision 1.13  2008/07/24 17:17:09  jheric
 *    Tested, finished up.
 *
 *    Revision 1.12  2008/07/17 10:37:49  jheric
 *    Cleaned up junky comments.
 *
 *    Revision 1.11  2008/07/11 03:39:28  jheric
 *    Added two additional fields (new cost and new buy).
 *
 *    Revision 1.10  2008/07/10 04:49:12  jheric
 *    Rewrote to accept and handle multiple facilities.
 *
 *    Revision 1.9  2008/05/02 07:07:30  jfisher
 *    Added the velocity code field
 *
 *    Revision 1.8  2006/04/12 11:23:46  smurdock
 *    added call to web service for Fascor locations
 *
 *    Revision 1.7  2006/03/17 14:21:24  smurdock
 *    *** empty log message ***
 *
 *    Revision 1.6  2006/03/07 19:30:57  smurdock
 *    final again -- took out credit invoices in when computing total and average order per Tracy Nantel
 *
 *    Revision 1.5  2006/03/07 17:55:55  smurdock
 *    final again
 *
 *    Revision 1.4  2006/03/06 17:56:35  smurdock
 *    added setup date and qty on hand
 *
 *    Revision 1.3  2006/03/03 15:02:19  jfisher
 *    Fixed bogus formatting and fixed warning with hidden log variable.
 *
 */

package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

import javax.xml.namespace.QName;

import org.apache.axis2.AxisFault;
import org.apache.axis2.addressing.EndpointReference;
import org.apache.axis2.client.Options;
import org.apache.axis2.rpc.client.RPCServiceClient;
import org.apache.axis2.transport.http.HTTPConstants;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.helper.Warehouse;
import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class VndFLCItemRpt extends Report
{
   public enum WsClients { AvgCost, SkuData, SkuQty };   
   private static final String avgCostSvcName   = "ItemCost";
   private static final String avgCostSvcMethod = "getItemAverageCost";
   private static final String avgCostSvcNsUri  = "http://websvc.emerywaterhouse.com";

   private PreparedStatement m_SqlVFI;     // vendor, FLC based on selection parameters, also has velocity.
   private PreparedStatement m_SqlAvgOrd;  // Order Total and Average for past year by item, if requested
   private PreparedStatement m_ItemDCQty;
   private PreparedStatement m_FasLocs;
   
   private String m_VendorId;              // Selected EIS Vendor Number, may be "" to select all vendors
   private String m_FlcId;                 // Selected FLC Id, "" selects all FLC's
   private String m_AvgOrd;                // Are order totals wanted?

   private XSSFWorkbook m_Workbook;
   private XSSFSheet m_Sheet;
   private XSSFRow m_Row;
   private XSSFFont m_FontNorm;
   private XSSFFont m_FontBold;
   private XSSFCellStyle m_StyleHdrLeft;
   private XSSFCellStyle m_StyleHdrLeftWrap;
   private XSSFCellStyle m_StyleHdrCntr;
   private XSSFCellStyle m_StyleHdrCntrWrap;
   private XSSFCellStyle m_StyleHdrRghtWrap;
   private XSSFCellStyle m_StyleDtlLeft;
   private XSSFCellStyle m_StyleDtlLeftWrap;
   private XSSFCellStyle m_StyleDtlCntr;
   private XSSFCellStyle m_StyleDtlRght;
   private XSSFCellStyle m_StyleDtlRght2d;
   private XSSFCellStyle m_StyleDtlRght3d;
   private XSSFCellStyle m_StyleDtlRght4d;
   private XSSFCellStyle m_StyleNewLine;
   private FileOutputStream m_OutputStream;

   private ArrayList<Warehouse> m_WhsList;      // A list of warehouses.
   private String m_WhsId;                     // Warehouse input parameter

   //
   // Web service clients used for each web service
   //
   // Web service classes
   RPCServiceClient m_Client;
   
   private QName m_CostMeth;
   @SuppressWarnings("rawtypes")
   private Class[] m_CostReturnType;
   private EndpointReference m_CostEndPointRef;

   private int m_TryCount;

   /**
    * default constructor
    */
   public VndFLCItemRpt()
   {
      super();

      m_FileNames.add("VndFLCItemRpt-" + String.valueOf(System.currentTimeMillis()) + ".xlsx");
      m_WhsList = new ArrayList<Warehouse>();
   }


   /*
    * cleans up member variables
    */
   protected void cleanup()
   {
      m_Status = RptServer.STOPPED;

      // closes prepared statements
      closePreparedStatement(m_SqlVFI);
      closePreparedStatement(m_SqlAvgOrd);
      closePreparedStatement(m_ItemDCQty);
      closePreparedStatement(m_FasLocs);

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

      m_Client = null;      
      m_CostMeth = null;
      m_CostReturnType = null;

      if ( m_WhsList != null )
         m_WhsList.clear();
   }

   /**
    * adds a numeric type cell to current row at col p_Col in current sheet
    *
    * @param col     0-based column number of spreadsheet cell
    * @param value   numeric value to be stored in cell
    * @param style   Excel style to be used to display cell
    */
   private void addCell(int col, double value, XSSFCellStyle style)
   {
      XSSFCell cell = m_Row.createCell(col);
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
   private void addCell(int col, String value, XSSFCellStyle style)
   {
      XSSFRichTextString tmp = new XSSFRichTextString(value);
      XSSFCell cell = m_Row.createCell(col);

      try {
         cell.setCellType(CellType.STRING);
         cell.setCellValue(tmp);
         cell.setCellStyle(style);
      }

      finally {
         cell = null;
         tmp = null;
      }
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
      short row = 0;
      short col = 0;
      short index = 0;
      int whsId = 0;
      String whsName = null;
      String itemNbr = null;
      String primaryVendorId = null;      
      ResultSet rsVFI = null;
      ResultSet rsAvgOrd = null;

      try {
         setCurAction("Running the vendor flc item report");
         row = openWorkbook();
                  
         if ( !m_VendorId.equals("") )
            m_SqlVFI.setString(++index, m_VendorId);

         if ( !m_FlcId.equals("") )
            m_SqlVFI.setString(++index, m_FlcId);

         // opens Edb query result set
         rsVFI = m_SqlVFI.executeQuery();

         // processes each row returned from Edb query result set
         while ( rsVFI.next() && m_Status == RptServer.RUNNING ) {
            itemNbr = rsVFI.getString("item_id");
            primaryVendorId = rsVFI.getString("primary_vendor");
            whsId = rsVFI.getInt("warehouse_id");
            whsName = rsVFI.getString("whs_name");

            // loads Edb data into spread sheet cells for new row
            addRow(row);

            addCell(col++, whsName, m_StyleDtlLeftWrap);
            addCell(col++, rsVFI.getString("vendor_id"), m_StyleDtlLeftWrap);
            addCell(col++, rsVFI.getString("name"), m_StyleDtlLeftWrap);

            // identifies whether vendor is primary or secondary for this item
            addCell(col++, primaryVendorId.equals(m_VendorId) ? "Primary" : "Secondary", m_StyleDtlLeftWrap);

            addCell(col++, itemNbr, m_StyleDtlLeftWrap);
            addCell(col++, rsVFI.getString("setup_date"), m_StyleDtlCntr);
            addCell(col++, rsVFI.getInt("avail_qty"), m_StyleDtlCntr);

            addCell(col++, rsVFI.getString("vendor_item_num"), m_StyleDtlLeftWrap);
            addCell(col++, rsVFI.getString("description"), m_StyleDtlLeftWrap);
            addCell(col++, rsVFI.getFloat("buy"), m_StyleDtlRght4d);
            addCell(col++, rsVFI.getFloat("new_cost"), m_StyleDtlRght2d);
            addCell(col++, getAvgCost(itemNbr, whsId), m_StyleDtlRght2d);

            addCell(col++, rsVFI.getFloat("sell"), m_StyleDtlRght2d);
            addCell(col++, rsVFI.getFloat("new_base"), m_StyleDtlRght2d);
            addCell(col++, rsVFI.getFloat("retail_a"), m_StyleDtlRght2d);
            addCell(col++, rsVFI.getFloat("retail_b"), m_StyleDtlRght2d);
            addCell(col++, rsVFI.getFloat("retail_c"), m_StyleDtlRght2d);
            addCell(col++, rsVFI.getFloat("retail_d"), m_StyleDtlRght2d);
            addCell(col++, rsVFI.getString("unit"), m_StyleDtlCntr);
            addCell(col++, rsVFI.getString("retail_pack"), m_StyleDtlCntr);
            addCell(col++, rsVFI.getString("stock_pack"), m_StyleDtlCntr);
            addCell(col++, rsVFI.getString("upc_code"), m_StyleDtlLeftWrap);
            addCell(col++, rsVFI.getString("sen_code_id"), m_StyleDtlCntr);
            addCell(col++, rsVFI.getString("flc_id"), m_StyleDtlCntr);
            addCell(col++, rsVFI.getString("in_catalog"), m_StyleDtlCntr);
            addCell(col++, rsVFI.getString("broken_case"), m_StyleDtlCntr);
            addCell(col++, getLocs(itemNbr, whsName), m_StyleDtlLeftWrap);
            addCell(col++, rsVFI.getString("velocity"), m_StyleDtlCntr);

            if ( !m_AvgOrd.equals("") ) {
               m_SqlAvgOrd.setString(1, itemNbr);
               rsAvgOrd = m_SqlAvgOrd.executeQuery();

               //
               // if we have order totals for this item_d, put them in
               if (rsAvgOrd.next()){
                  // column 23 (x), Total Ordered
                  addCell(col++, rsAvgOrd.getString("tot_qty"), m_StyleDtlCntr);
                  // column 24 (y), Average Order
                  addCell(col++, rsAvgOrd.getString("avg_ord_qty"), m_StyleDtlCntr);
               }
               // Result set returned no rows: this item was not ordered in past year, insert zeroes
               else {
                  // column 23 (x), Total Ordered
                  addCell(col++, "0", m_StyleDtlCntr);
                  // column 24 (y), Average Order
                  addCell(col++, "0", m_StyleDtlCntr);
               }
            }

            if ( rsAvgOrd != null ) {
               try {
                  rsAvgOrd.close();
                  rsAvgOrd = null;
               }

               catch ( Exception ex ) {

               }
            }

            row++;
            col = 0;
         }

         closeWorkbook();
         result = true;
      }


      catch ( Exception ex ) {
         log.error("exception:", ex);
         m_ErrMsg.append(ex.getMessage());
      }

      finally {
         if ( rsVFI != null ) {
            try {
               rsVFI.close();
               rsVFI = null;
            }

            catch ( Exception ex ) {

            }
         }
      }

      return result;
   }


   /**
    * closes output stream, deletes temporary disk file
    */
   private void closeOutputStream()
   {
      try {
         if ( m_OutputStream != null ) {
            m_OutputStream.close();
         }
      }

      catch ( Exception ex ) {
         log.error("VndStockRpt.closeOutputStream() " + ex);
      }

      finally {
         m_OutputStream = null;
      }
   }

   /**
    * closes a single prepared statement identified by p_Statement
    *
    * @param p_Statement   name of statement to be closed
    */
   private void closePreparedStatement(Statement p_Statement)
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
      setCurAction("writing excel spreadsheet");
      m_Workbook.write(m_OutputStream);

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
   @Override
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;

      try {         
         m_EdbConn = m_RptProc.getEdbConn();
         openOutputStream();
         init();
         
         if ( prepareStatements() )            
            created = buildWorkbook();

         setCurAction("complete");
      }

      catch ( Exception ex ) {
         log.fatal("exception:", ex);
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
      XSSFDataFormat customDataFormat = m_Workbook.createDataFormat();


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
      m_StyleDtlRght2d.setDataFormat(customDataFormat.getFormat("0.00"));


      // defines style detail data cell, right-justified with 3 decimal places
      //  (custom data format)
      m_StyleDtlRght3d = m_Workbook.createCellStyle();
      m_StyleDtlRght3d.setFont(m_FontNorm);
      m_StyleDtlRght3d.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtlRght3d.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleDtlRght3d.setDataFormat(customDataFormat.getFormat("0.000"));

      // defines style detail data cell, right-justified with 4 decimal places
      //  (custom data format)
      m_StyleDtlRght4d = m_Workbook.createCellStyle();
      m_StyleDtlRght4d.setFont(m_FontNorm);
      m_StyleDtlRght4d.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDtlRght4d.setVerticalAlignment(VerticalAlignment.TOP);
      m_StyleDtlRght4d.setDataFormat(customDataFormat.getFormat("0.0000"));
   }

   /**
    * Gets the average cost from the accounting system for the specified item.
    *
    * @param itemId The item to get the cost for.
    * @param facId The emery warehouse id
    *
    * @return The average cost for the item in the facility.
    *
    * @throws Exception
    */
   private double getAvgCost(String itemId, int facId) throws Exception
   {
      double cost = 0.00;
      RPCServiceClient client = getWsClient(WsClients.AvgCost);
      Object[] wsArgs = new Object[] { itemId, facId };
      Object[] response = null;
      long startTime = System.currentTimeMillis();
      long totalTime = 0;

      try {
         m_TryCount++;
         response = client.invokeBlocking(m_CostMeth, wsArgs, m_CostReturnType);
         cost = (Double)response[0];
      }

      catch ( AxisFault ex ) {
         if ( ex.getMessage().equals("Read timed out") ) {
            totalTime = (System.currentTimeMillis() - startTime);
            log.error(
                  String.format("[VndFLCItemRpt] actual time: %d; client options timeout: ",
                        totalTime,
                        client.getOptions().getTimeOutInMilliSeconds()
                        )
                  );

            if ( m_TryCount < 2 )
               cost = getAvgCost(itemId, facId);
            else {
               log.error(String.format("[VndFLCItemRpt] Unable to get avareage cost after %d attemtpts; returning -1", m_TryCount));
               cost = -1;
            }
         }
      }

      m_TryCount = 0;

      return cost;
   }

   /**
    * Creates the web service client and instantiate anything specific to a web service call.
    * Sets the correct EPR based on what the client type is.
    *
    * @return A reference to the ServiceClient object.
    */
   private RPCServiceClient getWsClient(WsClients wsc) throws Exception
   {
      String url = System.getProperty("soap.service.url");
      Options options = null;
      Integer soTimeout = new Integer(2 * 60000); // two minutes

      if ( m_Client == null ) {
         if ( url != null && url.length() > 0 ) {
            m_Client = new RPCServiceClient();

            m_CostEndPointRef = new EndpointReference(url + avgCostSvcName);
            m_CostMeth = new QName(avgCostSvcNsUri, avgCostSvcMethod);
            m_CostReturnType = new Class[] { Double.class };
            
            options = m_Client.getOptions();
            options.setProperty(HTTPConstants.REUSE_HTTP_CLIENT, "true");
            options.setProperty(HTTPConstants.SO_TIMEOUT, soTimeout);
            options.setProperty(HTTPConstants.CONNECTION_TIMEOUT, soTimeout);
            m_Client.setOptions(options);
         }
         else
            throw new Exception("Missing soap service url property");
      }

      if ( m_Client != null ) {
         switch ( wsc ) {
         case AvgCost: {
            m_Client.getOptions().setTo(m_CostEndPointRef);
            break;
         }

         default:
            break;
         }
      }

      return m_Client;
   }

   /**
    * @param item - String item identifier
    * @param facility - String facility identifier
    * @return String - concatenated locations for given item, location.
    * @throws Exception
    */
   public String getLocs(String item, String facility) throws Exception
   {
      StringBuffer tmp = new StringBuffer();
      ResultSet rs = null;
      int count = 0;
      
      m_FasLocs.setString(1, item);
      m_FasLocs.setString(2, facility);
            
      rs = m_FasLocs.executeQuery();
      
      try {
         while ( rs.next() ) {
            if ( count > 0 )
               tmp.append(", ");
            
            tmp.append(rs.getString(1));
            count++;   
         }
      }
      
      finally {
         rs.close();         
      }
      
      return tmp.toString();
   }
  
   /**
    * Initialize any internal data structures that are needed for processing.
    */
   private void init()
   {
      Statement stmt = null;
      ResultSet rs = null;
      StringBuffer sql = new StringBuffer();

      try {
         stmt = m_EdbConn.createStatement();
         sql.append("select warehouse_id, name, fas_facility_id, accpac_wh_id from warehouse ");

         if ( m_WhsId.length() > 0 )
            sql.append(String.format("where fas_facility_id = '%s'", m_WhsId));
         else
            sql.append("where fas_facility_id is not null");

         //System.out.println(sql.toString());
         rs = stmt.executeQuery(sql.toString());

         while ( rs.next() ) {
            m_WhsList.add(
               new Warehouse(
                  rs.getInt("warehouse_id"),
                  rs.getString("fas_facility_id"),
                  rs.getString("accpac_wh_id"),
                  rs.getString("name")
               )
            );
         }
      }

      catch ( SQLException ex ) {
         log.error("[VndFlcItem]", ex);
      }

      finally {
         closeRSet(rs);
         closeStmt(stmt);

         rs = null;
         stmt = null;
         sql = null;
      }
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
   private short openWorkbook()
   {
      short col = 0;
      int m_CharWidth = 295;

      // creates workbook
      m_Workbook = new XSSFWorkbook();

      // creates sheet 0
      m_Sheet = m_Workbook.createSheet("VndItemFLC");

      // defines styles
      defineStyles();

      // creates Excel title
      addRow(0);
      StringBuffer hdr = new StringBuffer();
      hdr.append("Vendor Item FLC Report for ");

      if (!m_VendorId.equals("")) {
         hdr.append("  Vendor ");
         hdr.append(m_VendorId);
      }

      if (!m_FlcId.equals("")) {
         hdr.append("  FLC ");
         hdr.append(m_FlcId);
      }

      addCell(col, hdr.toString(), m_StyleHdrLeft);
      hdr = null;

      // creates Excel column headings
      addRow(2);

      //
      // Computes approximate HSSF character width based on "Arial" size "10" and
      // HHSF characteristics, to be used as a multiplier to set column widths.
      
      //
      // column 0 (A), warehouse
      m_Sheet.setColumnWidth(col, (10 * m_CharWidth));
      addCell(col, "Warehouse", m_StyleHdrLeftWrap);

      //
      // column 1 (A), Vendor ID
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Vendor ID", m_StyleHdrLeftWrap);

      // column 2 (B), Vendor Name
      m_Sheet.setColumnWidth(++col, (30 * m_CharWidth));
      addCell(col, "Vendor", m_StyleHdrLeftWrap);

      // column 1B (B), Primary/Secondary Vendor
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Vend Type", m_StyleHdrLeftWrap);

      // column 3 (C), Item_ID
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Item ID", m_StyleHdrLeftWrap);

      // column 4 (D), Setup Date
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell (col, "Setup Date", m_StyleHdrCntrWrap);

      // column Qty On Hand
      m_Sheet.setColumnWidth(++col, (7 * m_CharWidth));
      addCell(col, "On Hand", m_StyleHdrCntrWrap);
      
      // column 5 (F), Vendor Item Number
      m_Sheet.setColumnWidth(++col, (25 * m_CharWidth));
      addCell(col, "Vendor Item Number", m_StyleHdrLeftWrap);

      // column 6 (G), Item Description
      m_Sheet.setColumnWidth(++col, (60 * m_CharWidth));
      addCell(col, "Item", m_StyleHdrLeftWrap);

      // column 7 (H), Emery Cost
      m_Sheet.setColumnWidth(++col, (9 * m_CharWidth));
      addCell(col, "Emery Cost", m_StyleHdrCntrWrap);

      // column 8 (I), New Cost
      m_Sheet.setColumnWidth(++col, (9 * m_CharWidth));
      addCell(col, "New Cost", m_StyleHdrCntrWrap);
     
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell (col, "Avg Cost", m_StyleHdrCntrWrap);

      // column 9 (J), Emery Base
      m_Sheet.setColumnWidth(++col, (9 * m_CharWidth));
      addCell(col, "Emery Base", m_StyleHdrCntrWrap);

      // column 10 (K), New  Base
      m_Sheet.setColumnWidth(++col, (9 * m_CharWidth));
      addCell(col, "New Base", m_StyleHdrCntrWrap);

      // column 11 (L), Retail A
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Retail A", m_StyleHdrCntrWrap);

      // column 12 (M), Retail B
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Retail B", m_StyleHdrCntrWrap);

      // column 13 (N), Retail C
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Retail C", m_StyleHdrCntrWrap);

      // column 14 (O), Retail D
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Retail D", m_StyleHdrCntrWrap);

      // column 15 (P), Ship Unit
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "Ship Unit", m_StyleHdrCntrWrap);

      // column 16 (Q), Retail Pack
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "Retail Pack", m_StyleHdrCntrWrap);

      // column 17 (R), Stock Pack
      m_Sheet.setColumnWidth(++col, (6 * m_CharWidth));
      addCell(col, "Stock Pack", m_StyleHdrCntrWrap);

      // column 18 (S), Primary UPC
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "UPC", m_StyleHdrCntrWrap);

      // column 19 (T), Sensitivity Code
      m_Sheet.setColumnWidth(++col, (4 * m_CharWidth));
      addCell(col, "Sen", m_StyleHdrCntrWrap);

      // column 20 (U), FLC
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "FLC", m_StyleHdrCntrWrap);

      // column 21 (V), In Catalog
      m_Sheet.setColumnWidth(++col, (4 * m_CharWidth));
      addCell(col, "Cat", m_StyleHdrCntrWrap);

      // column 22 (W), Broken Case
      m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
      addCell(col, "NBC", m_StyleHdrCntrWrap);

      m_Sheet.setColumnWidth(++col, (40 * m_CharWidth));
      addCell (col, "Whs Location", m_StyleHdrCntrWrap);

      //
      // column 24 (Y), Location
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Velocity", m_StyleHdrCntrWrap);

      // if user wanted order totals, we need two more columns
      if ( !m_AvgOrd.equals("") ) {
         // column 25 (Z), Total Ordered
         m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
         addCell(col, "Total Ord", m_StyleHdrCntrWrap);

         // column 25 (AA), Average Order
         m_Sheet.setColumnWidth(++col, (5 * m_CharWidth));
         addCell(col, "Avg Ord",  m_StyleHdrCntrWrap);
      }

      // returns first data row number
      return 3;
   }

   /**
    * Edb query to get sales data based on parameters
    * @return true if the statements are prepared, false if not.
    */
   private boolean prepareStatements()
   {
      boolean isPrepared = false;
      StringBuffer sql = new StringBuffer();
      String whsIds = "";
      int i = 0;

      if ( m_EdbConn != null && m_EdbConn != null ) {
         //
         // Can't use a parameter for an in ().  Just build the list from the
         // warehouse list which gets initialized early.
         for ( Warehouse w : m_WhsList ) {
            if ( i > 0 )
               whsIds += ",";
            
            whsIds += Integer.toString(w.emeryId);
            i++;
         }
         
         try {
            sql.append("select ");
            sql.append("vendor.vendor_id, ");
            sql.append("vendor.name, ");
            sql.append("item_entity_attr.item_id, ");
            sql.append("item_entity_attr.vendor_id as primary_vendor, ");
            sql.append("to_char(ejd_item.setup_date,'mm/dd/yyyy') as setup_date, ");
            sql.append("vendor_item_num, ");
            sql.append("warehouse.warehouse_id, ");
            sql.append("warehouse.name as whs_name, ");
            sql.append("item_entity_attr.description, ");
            sql.append("buy, ");
            sql.append("( ");
            sql.append("   select buy ");
            sql.append("   from ejd_pending_price ");
            sql.append("   where ejd_pending_price.ejd_item_id = ejd_item.ejd_item_id and buy_date in ( ");
            sql.append("      select min(buy_date) ");
            sql.append("      from ejd_pending_price epp2 ");
            sql.append("      where epp2.buy_date > current_date and epp2.ejd_item_id = ejd_item.ejd_item_id and epp2.warehouse_id = warehouse.warehouse_id ");
            sql.append("   )and ejd_pending_price.warehouse_id = warehouse.warehouse_id ");
            sql.append(") as new_cost, ");                     
            sql.append("sell, ");
            sql.append("( ");
            sql.append("   select sell ");
            sql.append("   from ejd_pending_price ");
            sql.append("   where ejd_pending_price.ejd_item_id = ejd_item.ejd_item_id and sell_date in ( ");
            sql.append("      select min(sell_date) ");
            sql.append("      from ejd_pending_price epp2 ");
            sql.append("      where epp2.sell_date > current_date and epp2.ejd_item_id = ejd_item.ejd_item_id and epp2.warehouse_id = warehouse.warehouse_id ");
            sql.append("   )and ejd_pending_price.warehouse_id = warehouse.warehouse_id ");
            sql.append(") as new_base, ");
            sql.append("retail_a, ");
            sql.append("retail_b, ");
            sql.append("retail_c, ");
            sql.append("retail_d, ");
            sql.append("unit, ");
            sql.append("retail_pack, ");
            sql.append("stock_pack, ");
            sql.append("( ");
            sql.append("   select upc_code ");
            sql.append("   from ejd_item_whs_upc ");
            sql.append("   where ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_whs_upc.warehouse_id = warehouse.warehouse_id ");
            sql.append("   order by primary_upc desc, upc_code limit 1 ");
            sql.append(") as upc_code, ");
            sql.append("sen_code_id, ");
            sql.append("flc_id, ");
            sql.append("decode(in_catalog, 0, 'N', 'Y') as in_catalog, ");
            sql.append("decode(broken_case.description, 'ALLOW BROKEN CASES', 'N', 'Y') as broken_case, ");
            sql.append("velocity, ");
            sql.append("avail_qty ");
            sql.append("from ejd_item ");
            sql.append(String.format("join warehouse on warehouse.warehouse_id in (%s) ", whsIds));
            sql.append("join item_entity_attr using(ejd_item_id) ");
            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id and ejd_item_price.warehouse_id = warehouse.warehouse_id ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id and ejd_item_warehouse.warehouse_id = warehouse.warehouse_id ");
            sql.append("join item_velocity using(velocity_id) ");
            sql.append("join vendor_item_ea_cross on vendor_item_ea_cross.item_ea_id = item_entity_attr.item_ea_id ");
            sql.append("join vendor on vendor.vendor_id = vendor_item_ea_cross.vendor_id ");
            sql.append("join ship_unit on ship_unit.unit_id = item_entity_attr.ship_unit_id ");
            sql.append("join broken_case using (broken_case_id) ");

            if ( !m_VendorId.equals("") && !m_FlcId.equals("") )
               sql.append("where vendor.vendor_id = ? and flc_id = ? ");
            else {
               if ( !m_VendorId.equals("") )
                  sql.append("where vendor.vendor_id = ? ");
               else
                  sql.append("where flc_id = ? ");
            }

            sql.append("order by name, item_id, warehouse_id ");

            m_SqlVFI = m_EdbConn.prepareStatement(sql.toString());

            //
            // we prepare the next query only if the user requested order totals
            if ( !m_AvgOrd.equals("") ) {
               sql.setLength(0);
               sql.append("select ");
               sql.append("coalesce(sum(qty_ordered), 0) as tot_qty, ");
               sql.append("coalesce(round(sum(qty_ordered)/count(item_nbr)), 0) as avg_ord_qty ");
               sql.append("from inv_dtl ");
               sql.append("where ");
               sql.append("item_nbr = ? and tran_type = 'SALE' and ");
               sql.append("invoice_date between (current_date - interval '12' month) and current_date ");
               m_SqlAvgOrd = m_EdbConn.prepareStatement(sql.toString());
            }

            sql.setLength(0);
            sql.append("select avail_qty ");
            sql.append("from item_entity_attr ");
            sql.append("join warehouse on warehouse.warehouse_id = ? ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ");
            sql.append("   ejd_item_warehouse.warehouse_id = warehouse.warehouse_id ");
            sql.append("where item_id = ? ");
            m_ItemDCQty = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select loc_id ");
            sql.append("from loc_allocation ");
            sql.append("where sku = ? and warehouse = ?");
            m_FasLocs = m_EdbConn.prepareStatement(sql.toString());

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
         if ( p.name.equals("vndId"))
            m_VendorId = p.value.trim();

         if ( p.name.equals("flcId"))
            m_FlcId = p.value.trim();

         if ( p.name.equals("ordqty"))
            m_AvgOrd = p.value.trim();

         if ( p.name.equals("dc"))
            m_WhsId = p.value.trim();
      }

   }
}
