/**
 * File: ItemExtract.java
 * Description: Listing of all items sold by emery.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 02/12/2009
 * Last Update: $Id: ItemExtract.java,v 1.8 2013/09/11 14:25:42 tli Exp $
 *
 * History
 *    $Log: ItemExtract.java,v $
 *    Revision 1.8  2013/09/11 14:25:42  tli
 *    Converted the facilityId to facilityName when needed
 *
 *    Revision 1.7  2013/09/09 18:33:38  tli
 *    Replace SkuQty web service call with item_qty_view
 *
 *    Revision 1.6  2012/10/05 14:16:05  jfisher
 *    removed duplicated line of code
 *
 *    Revision 1.5  2012/10/05 14:15:20  jfisher
 *    Changes to deal with the timeout on the sku quantity web service.
 *
 *    Revision 1.4  2012/08/29 19:53:02  jfisher
 *    Switched web service calls from Wasp to Axis2
 *
 *    Revision 1.3  2009/07/16 20:12:56  smurdock
 *    added nbc and dc columns for Tom Poole
 *
 *    Revision 1.2  2009/02/17 22:31:08  jfisher
 *    Added sheet processing for results > 65K.  Added some code cleanup.
 *
 *    Revision 1.1  2009/02/12 19:46:40  jfisher
 *    Initial add
 *
 */

package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import javax.xml.namespace.QName;

import org.apache.axis2.addressing.EndpointReference;
import org.apache.axis2.client.Options;
import org.apache.axis2.rpc.client.RPCServiceClient;
import org.apache.axis2.transport.http.HTTPConstants;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class ItemExtract extends Report
{
   private static final short maxCols = 15;

   private static final String avgCostSvcName   = "ItemCost";
   private static final String avgCostSvcMethod = "getItemAverageCost";
   private static final String avgCostSvcNsUri  = "http://websvc.emerywaterhouse.com";

   private String m_Dept;                    // A department filter.
   private String m_OrderBy;                 // The sort order for the query.
   private String m_VndId;                   // A vendor id to limit the query listing.
   
   private PreparedStatement m_ItemData;
      
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;

   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;

   //
   // Web service classes
   RPCServiceClient m_Client;
   
   private QName m_CostMeth;
   private EndpointReference m_CostEndPointRef;
   @SuppressWarnings("rawtypes")
   private Class[] m_CostReturnType;
   

   public ItemExtract()
   {
      super();
      
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet("Item List Page 1");

      m_VndId = "";
      m_Dept = "";
      m_OrderBy = "";

      setupWorkbook();
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable
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
      m_Dept = null;
      m_OrderBy = null;
      m_VndId = null;
      m_Client = null;
      m_CostMeth = null;
      m_CostReturnType = null;

      super.finalize();
   }

   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);      
      String msg1 = "processing page %d, row %d, item %s";
      String msg2 = "processing page %d, row %d, vendor %d, item %s";
      int rowNum = 0;
      int colNum = 0;
      int vndId = 0;
      int pageNum = 1;
      boolean result = false;
      XSSFRichTextString itemId = null;
      XSSFRow row = null;
      ResultSet itemData = null;
      
      try {         
         rowNum = createCaptions();
         itemData = m_ItemData.executeQuery();

         while ( itemData.next() && m_Status == RptServer.RUNNING ) {
            vndId = itemData.getInt("vendor_id");
            itemId = new XSSFRichTextString(itemData.getString("item_id"));

            if ( m_OrderBy.equalsIgnoreCase("item_id") )
               setCurAction(String.format(msg1, pageNum, rowNum, itemId));
            else
               setCurAction(String.format(msg2, pageNum, rowNum, vndId, itemId));

            row = createRow(rowNum++, maxCols);
            colNum = 0;

            if ( row != null ) {
               row.getCell(colNum++).setCellValue(vndId);
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(itemData.getString("name")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(itemData.getString("item_id")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(itemData.getString("description")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(itemData.getString("upc_code")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(itemData.getString("unit")));
               row.getCell(colNum++).setCellValue(itemData.getInt("stock_pack"));
               row.getCell(colNum++).setCellValue(itemData.getInt("retail_pack"));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(itemData.getString("nbc")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(itemData.getString("whs_name")));               
               row.getCell(colNum++).setCellValue(itemData.getDouble("avail_qty"));
               row.getCell(colNum++).setCellValue( getAvgCost(itemId.getString(), itemData.getInt("warehouse_id")) );
               row.getCell(colNum++).setCellValue(itemData.getDouble("buy"));
               row.getCell(colNum++).setCellValue(itemData.getDouble("sell"));
               row.getCell(colNum++).setCellValue(itemData.getDouble("retail_c"));
            }

            if ( rowNum > 65000 ) {
               pageNum++;
               m_Sheet = m_Wrkbk.createSheet("Item List Page " + pageNum);
               rowNum = createCaptions();
            }
         }

         m_Wrkbk.write(outFile);
         DbUtils.closeDbConn(null, null, itemData);

         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("The report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {
         row = null;
         
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[ItemExtract]", e);
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
      closeStmt(m_ItemData);
   }

   /**
    * Sets the captions on the report.
    */
   private int createCaptions()
   {
      XSSFRow row = null;
      int rowNum = 0;
      int colNum = 0;

      if ( m_Sheet != null ) {
         //
         // Create the row for the captions.
         row = m_Sheet.createRow(rowNum);

         if ( row != null ) {
            for ( int i = 0; i < maxCols; i++ ) {
               row.createCell(i);
            }

            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Vendor#"));
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Vendor"));
            m_Sheet.setColumnWidth(colNum++, 8000);
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Item#"));
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Description"));
            m_Sheet.setColumnWidth(colNum++, 4000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("UPC"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Unit"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Stock Pack"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Ret Pack"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("NBC"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("DC"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("On Hand"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Avg Cost"));            
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Buy"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Base"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Ret C"));
         }

         rowNum++;
      }

      return rowNum;
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

         if ( prepareStatements() ) {            
            created = buildOutputFile();
         }
      }

      catch ( Exception ex ) {
         log.fatal("exception:", ex);
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
    * @param colCnt The number of columns in the row.
    *
    * @return The formatted row of the spreadsheet.
    */
   private XSSFRow createRow(int rowNum, short colCnt)
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
    * Gets the average cost from the accounting system for the specified item.
    *
    * @param itemId The item to get the cost for.
    * @param facId The emery warehouse id
    *
    * @return The average cost for the item in the facility.
    * @throws Exception
    *
    * @throws LookupException
    */
   private double getAvgCost(String itemId, int facId) throws Exception
   {
      double cost = 0.00;
      RPCServiceClient client = getWsClient();
      Object[] wsArgs = new Object[] { ACCESS_KEY, itemId, facId };
      Object[] response = null;

      response = client.invokeBlocking(m_CostMeth, wsArgs, m_CostReturnType);
      cost = (Double)response[0];

      return cost;
   }

   /**
    * Creates the web service client and instantiate anything specific to a web service call.
    * Sets the correct EPR based on what the client type is.
    *
    * @return A reference to the ServiceClient object.
    */
   private RPCServiceClient getWsClient() throws Exception
   {
      String url = System.getProperty("soap.service.url");
      Options options = null;

      //
      // Setup the client once the first time through with all the default stuff
      if ( m_Client == null ) {
         if ( url != null && url.length() > 0 ) {
            m_Client = new RPCServiceClient();

            m_CostEndPointRef = new EndpointReference(url + avgCostSvcName);
            m_CostMeth = new QName(avgCostSvcNsUri, avgCostSvcMethod);
            m_CostReturnType = new Class[] { Double.class };
        
            options = m_Client.getOptions();
            options.setProperty(HTTPConstants.REUSE_HTTP_CLIENT, "true");
            m_Client.setOptions(options);
         }
         else
            throw new Exception("Missing soap service url property");
      }

      //
      // Set the endpoint based on which web service to call.
      if ( m_Client != null ) {
         m_Client.getOptions().setTo(m_CostEndPointRef);
      }

      return m_Client;
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(512);
      boolean isPrepared = false;
      String orderBy = "order by item_id";

      if ( m_EdbConn != null ) {
         try {
            sql.append("select item_data.* ");
            sql.append("from ( ");
            sql.append("   select ");
            sql.append("   vendor.vendor_id, vendor.name, item_entity_attr.item_id, item_entity_attr.description,  ");
            sql.append("   ship_unit.unit, stock_pack, retail_pack, ");
            sql.append("   decode(broken_case_id,1,'Y','N') as nbc, ");
            sql.append("   warehouse.name as whs_name, warehouse.warehouse_id, ");
            sql.append("   upc_code, buy, sell, retail_c, avail_qty ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join warehouse on warehouse.warehouse_id = 1 ");
            sql.append("join item_type on item_type.item_type_id = item_entity_attr.item_type_id and itemtype = 'STOCK' ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            
            //
            // If the user added a vendor filter, add it here.  Preparing with a
            if ( m_VndId.length() > 0 )
               sql.append(String.format("and vendor.vendor_id = %s ", m_VndId));
            
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.warehouse_id = warehouse.warehouse_id ");
            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_price.warehouse_id = warehouse.warehouse_id ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = item_entity_attr.ejd_item_id and primary_upc = 1 and  ");
            sql.append("   ejd_item_whs_upc.warehouse_id = warehouse.warehouse_id ");
            sql.append("join ship_unit on ship_unit.unit_id = item_entity_attr.ship_unit_id ");
            sql.append("union ");
            sql.append("select ");
            sql.append("   vendor.vendor_id, vendor.name, item_entity_attr.item_id, item_entity_attr.description, ");
            sql.append("   ship_unit.unit, stock_pack, retail_pack, ");
            sql.append("   decode(broken_case_id,1,'Y','N') as nbc, ");
            sql.append("   warehouse.name as whs_name, warehouse.warehouse_id, ");
            sql.append("   upc_code, buy, sell, retail_c, ");
            sql.append("   avail_qty ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join warehouse on warehouse.warehouse_id = 2 ");
            sql.append("join item_type on item_type.item_type_id = item_entity_attr.item_type_id and itemtype = 'STOCK' ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            
            //
            // If the user added a vendor filter, add it here.  Preparing with a
            if ( m_VndId.length() > 0 )
               sql.append(String.format("and vendor.vendor_id = %s ", m_VndId));
            
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.warehouse_id = warehouse.warehouse_id ");
            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_price.warehouse_id = warehouse.warehouse_id ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = item_entity_attr.ejd_item_id and primary_upc = 1 and ");
            sql.append("   ejd_item_whs_upc.warehouse_id = warehouse.warehouse_id ");
            sql.append("join ship_unit on ship_unit.unit_id = item_entity_attr.ship_unit_id ");
            sql.append(") item_data ");
            
            //
            // Add a department filter.  No department number means all departments.
            if ( m_Dept.length() > 0 ) {
               sql.append("join vendor_dept on vendor_dept.vendor_id = item_data.vendor_id ");
               sql.append("join emery_dept on emery_dept.dept_id = vendor_dept.dept_id ");
               sql.append(String.format("and dept_num = '%s' ", m_Dept));
            }
            
            if ( m_OrderBy.length() > 0 )
               orderBy = String.format("order by %s", m_OrderBy);

            sql.append(orderBy);

            m_ItemData = m_EdbConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("[ItemExtract]", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("[ItemExtract] null DB connecion");

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

         if ( param.name.equals("sort") ) {
            m_OrderBy = param.value;

            //
            // Always sort by item id even when a different column is selected as the primary
            // sort order.
            if ( !m_OrderBy.equalsIgnoreCase("item_id") )
               m_OrderBy+= ", item_id";
         }

         if ( param.name.equals("vnd") )
            m_VndId = param.value;

         if ( param.name.equals("dept") )
            m_Dept = param.value;
      }

      fileName.append("item_extract");
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
   }

   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {
      XSSFCellStyle styleText;      // Text left justified
      XSSFCellStyle styleInt;       // Style with 0 decimals
      XSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00)

      styleText = m_Wrkbk.createCellStyle();
      styleText.setAlignment(HorizontalAlignment.LEFT);

      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);

      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short)8);

      m_CellStyles = new XSSFCellStyle[] {
         styleText,     // col 0 vnd id
         styleText,     // col 1 vnd name
         styleText,     // col 2 item id
         styleText,     // col 3 item desc
         styleText,     // col 4 upc
         styleText,     // col 5 unit
         styleInt,      // col 6 retail pack
         styleText,     // col 7 nbc -- N equals No Broken Case per Tom
         styleText,     // col 8 distribution center
         styleInt,      // col 9 stock pack
         styleInt,      // col 10 on hand
         styleMoney,    // col 11 avg cost
         styleMoney,    // col 12 buy
         styleMoney,    // col 13 base
         styleMoney,    // col 14 ret c
      };

      styleText = null;
      styleInt = null;
      styleMoney = null;
   }
}
