/**
 * File: CustPriceQuont.java
 * Description: Customer quote on items.
 *
 * @author ?
 *
 * Create Date: 02/08/2006
 * Last Update: 05/15/2017
 *
 * History
 *    Fixed a bug with the getParent query where it was returning the result even though
 *       the parent was null.  jfisher
 *
 *    Now handles ACE warehouses, specifically Wilton.  Fixed a number of bugs and some
 *    Ridiculous code.  The setup date is now displayed correctly. jfisher
 *
 *    Revision 1.24  2014/11/05 18:14:35  everge
 *    Fixed a bug that caused the 'Customer SKU' field to be empty for customers that share SKUs with a parent.
 *    Will now fill in with the customer's parent's customer SKUs if the customer has a parent.
 *
 *    Revision 1.23  2014/10/20 16:27:27  ebrownewell
 *    Updated to allow this report to show virtual items if a virtual vendor is selected.
 *
 *    Revision 1.22  2013/09/11 14:25:42  tli
 *    Converted the facilityId to facilityName when needed
 *
 *    Revision 1.21  2013/09/09 18:33:38  tli
 *    Replace SkuQty web service call with item_qty_view
 *
 *    Revision 1.20  2013/02/26 22:13:30  prichter
 *    Oracle upgrade.  Fixed ambiguous columns.
 *
 *    Revision 1.19  2012/10/05 13:55:06  jfisher
 *    Changes to deal with the timeout on the sku quantity web service.
 *
 *    Revision 1.18  2012/08/29 19:53:02  jfisher
 *    Switched web service calls from Wasp to Axis2
 *
 *    Revision 1.17  2012/05/08 09:00:04  pberggren
 *    Pointer to SkuQty was commented out, now resolved.
 *
 *    Revision 1.16  2012/05/05 06:07:42  pberggren
 *    Removed redundant loading of system properties.
 *
 *    Revision 1.15  2012/05/03 07:55:10  prichter
 *    Fix to web service ip address
 *
 *    Revision 1.14  2012/05/03 03:21:01  pberggren
 *    Added new code for test server location PBerggren
 *
 *    Revision 1.13  2011/07/27 04:53:53  epearson
 *    changed method used in date formatting compatible with poi 3.2
 *
 *    Revision 1.12  2011/07/27 04:22:59  epearson
 *    Added DC column and fixed setup date format
 *
 *    Revision 1.11  2010/03/22 21:49:45  smurdock
 *    now reports Portland and Pittson on hand in two columns if user does not choose a DC.
 *
 *    Revision 1.10  2010/01/30 02:19:06  smurdock
 *    added call to wsdl getFacilitySkuQty to get Pittston on hnads correct when requested
 *
 *    Revision 1.9  2009/02/17 22:56:08  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.8  2009/02/12 17:13:52  smurdock
 *     report ONLY  primary vendor
 *
 *    (join to vendor_item_cross is now on BOTH item_id and vendor_id)
 *
 *    Revision 1.7  2009/02/05 16:15:08  smurdock
 *    added select by distribution center
 *
 *    updated join sytax to current ANSI
 *
 *    Revision 1.6  2008/10/29 21:10:41  jfisher
 *    Fixed some warnings
 *
 *    Revision 1.5  2006/09/05 18:53:25  jfisher
 *    Fixed problem with value being put into a cell as an int instead of a double.
 *
 *    Revision 1.4  2006/04/17 18:09:22  jfisher
 *    switched the getActQty method to call the new SkuQty web service.
 *
 *    Revision 1.3  2006/03/03 14:16:07  jfisher
 *    Modified the price routines to deal with exceptions.  Ready for production.
 *
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


/**
 * Customer Price Quote
 *
 */
public class CustPriceQuote extends Report
{
   //
   // Report type filter identifiers
   private static final short FT_CUST              = 0;
   private static final short FT_CUST_VND          = 1;
   private static final short FT_CUST_FLC          = 2;
   private static final short FT_CUST_VND_FLC      = 3;
   private static final short FT_CUST_ITEM         = 4;
   private static final short FT_CUST_VND_ITEM     = 5;
   private static final short FT_CUST_FLC_ITEM     = 6;
   private static final short FT_CUST_VND_FLC_ITEM = 7;
   private static final short FT_MSI               = 8;
   private static final short FT_NRHA              = 9;

   private String m_CustId;
   private String m_FlcId;
   private String m_ItemEaId;
   private String m_Msi;
   private String m_MsiDate;
   private String m_NrhaId;
   private String m_RptDate;
   private String m_DC;
   private short m_FilterType;
   private int m_VndId;

   private PreparedStatement m_CustCost;
   private PreparedStatement m_CustData;
   private PreparedStatement m_CustRetail;
   private PreparedStatement m_CustParent;
   private PreparedStatement m_CustWhs;
   private PreparedStatement m_InvDtl;
   private PreparedStatement m_RmsId;
   private PreparedStatement m_ItemPrices;

   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;

   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;

   /**
    * Default constructor.
    */
   public CustPriceQuote()
   {
      super();

      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
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

      super.finalize();
   }

   /**
    * Executes the queries and builds the output file
    *
    * @return true if the file was built, false if not.
    * @throws FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      XSSFRow row = null;
      FileOutputStream outFile = null;
      ResultSet custData = null;
      ResultSet invDtl = null;
      int rowNum = 1;
      int colNum = 0;
      int whsId = 0;
      boolean result = false;
      int itemEaId = 0;
      int ejdItemId = 0;
      String item = "";
      String upc = null;
      String flc = null;
      int qtyShipped = 0;
      double extSell = 0.0;
      double custCost = 0.0;
      double[] priceData = new double[7];

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         setupWorkbook();
         rowNum = createCaptions();

         m_CustData.setString(1, getParentId(m_CustId));
         custData = m_CustData.executeQuery();

         while ( custData.next() && m_Status == RptServer.RUNNING ) {
            row = createRow(rowNum);
            itemEaId = custData.getInt("item_ea_id");
            ejdItemId = custData.getInt("ejd_item_id");
            whsId = custData.getInt("warehouse_id");
            item = custData.getString("item_id");
            upc = custData.getString("upc_code");
            flc = custData.getString("flc_id");
            
            custCost = getCustCost(itemEaId);


            setCurAction("processing item: " + item);

            m_InvDtl.setString(1, item);

            //
            // If there is a vendor filter, then use both the vendor id
            // and the bull id.
            if ( m_VndId > 0 )
               m_InvDtl.setString(2, Integer.toString(m_VndId));

            invDtl = m_InvDtl.executeQuery();

            try {
               if ( invDtl.next() ) {
                  qtyShipped = invDtl.getInt("qty_shipped");
                  extSell = invDtl.getDouble("ext_sell");
               }
            }

            finally {
               closeRSet(invDtl);
            }

            //
            // Get the pricing for each item unless it's in the
            // misc recovery FLC.  Those don't have price records.
            if ( !flc.equals("9998") )
               getPricing(ejdItemId, whsId, priceData);

            row.getCell(colNum++).setCellValue(new XSSFRichTextString(item));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("customer_sku")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(upc != null ? upc : ""));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("in_catalog")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("vendor_item_num")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("description")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("name")));
            row.getCell(colNum++).setCellValue(custData.getInt("stock_pack"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("nbc")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("shp_unit")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("ret_unit")));

            //
            // Some items may not be in customer contracts etc.  Leave it blank if that's the case.
            if ( custCost > -1.0 )
               row.getCell(colNum++).setCellValue(custCost);
            else
               colNum++;

            row.getCell(colNum++).setCellValue(getCustRetail(item));
            row.getCell(colNum++).setCellValue(qtyShipped);
            row.getCell(colNum++).setCellValue(extSell);
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(flc));
            row.getCell(colNum++).setCellValue(custData.getDouble("item_freight"));
            row.getCell(colNum++).setCellValue(priceData[0]);  // buy
            row.getCell(colNum++).setCellValue(priceData[1]);  // sell
            row.getCell(colNum++).setCellValue(priceData[2]);  // retail a
            row.getCell(colNum++).setCellValue(priceData[3]);  // retail b
            row.getCell(colNum++).setCellValue(priceData[4]);  // retail c
            row.getCell(colNum++).setCellValue(priceData[5]);  // retail d

            //
            // A zero value for a sensitivity code id represents null.
            // ACE items won't have a sensitivity code.
            if ( priceData[6] > 0 )
               row.getCell(colNum++).setCellValue((int)priceData[6]);
            else
               colNum++;

            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("disposition")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(getRmsId(item)));
            row.getCell(colNum++).setCellValue(custData.getInt("vendor_id"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("warehouse")));
            row.getCell(colNum++).setCellValue(custData.getInt("qoh"));            
            row.getCell(colNum).setCellValue(custData.getDate("setup_date"));

            colNum = 0;
            rowNum++;
            upc = null;

            for ( int i = 0; i < priceData.length; i++ )
               priceData[i] = 0.0;
         }

         m_Wrkbk.write(outFile);
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         log.error(String.format("[CustPriceQuote] item id %s whs id %d ", item, whsId), ex);
      }

      finally {
         row = null;
         priceData = null;

         closeRSet(custData);

         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[CustPriceQuote]", e);
         }

         outFile = null;
      }

      return result;
   }

   /**
    * Builds the sql based on the type of filter requested by the user.
    * @return A complete sql statement.
    * @throws SQLException
    */
   private String buildSql() throws SQLException
   {
      StringBuffer sql = new StringBuffer();
      StringBuffer join = new StringBuffer();
      StringBuffer where = new StringBuffer();
      
      sql.append("select \r\n");
      sql.append("   item_entity_attr.item_ea_id, item_entity_attr.ejd_item_id, item_entity_attr.item_id, customer_sku, upc_code, \r\n");
      sql.append("   vendor_item_num, item_entity_attr.description, vendor.name, decode(in_catalog, 0, 'N', 'Y') as in_catalog, \r\n");
      sql.append("   stock_pack, decode(broken_case.description, 'ALLOW BROKEN CASES', '', 'NBC') as nbc, \r\n");
      sql.append("   ejd_item.flc_id, frt_cwt, disposition, vendor.vendor_id, ship_unit.unit as shp_unit, \r\n");
      sql.append("   ejd_price_procs.get_item_freight(item_entity_attr.item_ea_id, ejd_item_warehouse.warehouse_id) as item_freight, \r\n");
      sql.append("   retail_unit.unit as ret_unit, ejd_item.setup_date, qoh, warehouse.name as warehouse, warehouse.warehouse_id \r\n");
      sql.append("from item_entity_attr \r\n");
      sql.append("join customer on customer.customer_id = ? \r\n");
      sql.append("join cust_warehouse on cust_warehouse.customer_id = customer.customer_id and whs_priority = 1 \r\n");
      sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id \r\n");
      sql.append("join item_type on item_type.item_type_id = item_entity_attr.item_type_id and itemtype in ('STOCK', 'ACE', 'EXPANDED ASST', 'VIRTUAL', 'ASSORTMENT') \r\n");
      sql.append("left outer join vendor_item_ea_cross on vendor_item_ea_cross.item_ea_id = item_entity_attr.item_ea_id and \r\n");
      sql.append("  item_entity_attr.vendor_id = vendor_item_ea_cross.vendor_id \r\n");
      sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id \r\n");      
      sql.append("join ship_unit on ship_unit.unit_id = item_entity_attr.ship_unit_id \r\n");
      sql.append("join retail_unit on retail_unit.unit_id = item_entity_attr.ret_unit_id \r\n");
      sql.append("join broken_case on broken_case.broken_case_id = ejd_item.broken_case_id \r\n");
      sql.append("left outer join item_ea_cross on item_ea_cross.item_ea_id = item_entity_attr.item_ea_id and item_ea_cross.customer_id = customer.customer_id \r\n");
      sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id and ejd_item_warehouse.warehouse_id = cust_warehouse.warehouse_id \r\n");
      sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = ejd_item.ejd_item_id and \r\n");
      sql.append("  ejd_item_whs_upc.warehouse_id = cust_warehouse.warehouse_id and primary_upc = 1 \r\n");
      sql.append("join warehouse on warehouse.warehouse_id = cust_warehouse.warehouse_id \r\n");
      
      // temporarily off until the warehouse can be figured out.
      /*if ( m_DC != null && m_DC.length() > 0 )         
         sql.append(String.format("and warehouse.name = '%s' ", m_DC));
      else
         sql.append(String.format(" and warehouse.warehouse_id in (%s) ", getCustWarehouse()));*/
      
      sql.append("join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id and disposition in ('BUY-SELL', 'NOBUY', 'NOBUY-NOSELL') \r\n");
            
      switch ( m_FilterType ) {
         case FT_CUST_VND:
            where.append(String.format("where item_entity_attr.vendor_id = %d \r\n", m_VndId));
         break;

         case FT_CUST_FLC:
            where.append(String.format("where ejd_item.flc_id = '%s' \r\n", m_FlcId));
         break;

         case FT_CUST_VND_FLC:
            where.append(String.format("where ejd_item.flc_id = '%s' and ", m_FlcId));
            where.append(String.format("item_entity_attr.vendor_id = %d \r\n", m_VndId));
         break;

         case FT_CUST_ITEM:
            where.append(String.format("where item_entity_attr.item_ea_id in (%s) \r\n", m_ItemEaId));
         break;

         case FT_CUST_VND_ITEM:
            where.append(String.format("where item_entity_attr.item_ea_id in (%s) and ", m_ItemEaId));
            where.append(String.format("item_entity_attr.vendor_id = %d \r\n", m_VndId));
         break;

         case FT_CUST_FLC_ITEM:
            where.append(String.format("where item_entity_attr.item_ea_id in (%s) and  ", m_ItemEaId));
            where.append(String.format("ejd_item.flc_id = '%s' \r\n", m_FlcId));
         break;

         case FT_CUST_VND_FLC_ITEM:
            where.append(String.format("where item_entity_attr.item_ea_id in (%s) and ", m_ItemEaId));
            where.append(String.format("ejd_item.flc_id = '%s' and ", m_FlcId));
            where.append(String.format("item_entity_attr.vendor_id = %d \r\n", m_VndId));
         break;

         case FT_MSI:
            join.append("join order_line_error ole on ole.item_id = item_entity_attr.item_id ");
            join.append("join order_header_error ohe on ");
            join.append("   ohe.ohe_id = ole.ohe_id and ");
            join.append("   instr(ohe.source_ref, '.MSI') > 0 and ");
            join.append("   ohe.errormsg = 'Miscellaneous Transmission' and ");
            join.append(String.format("   substr( ohe.source_ref, 1, 6 ) = '%s' and ", m_MsiDate));
            join.append(String.format("   ohe.source_id_num = '%s' ", m_Msi));
         break;

         case FT_NRHA:
            //
            // Have to add the extra tables here or the query performs full table scans
            // when they aren't used.
            join.append("join flc on flc.flc_id = ejd_item.flc_id  ");
            join.append("join mdc on mdc.mdc_id = flc.mdc_id ");
            join.append("join nrha on nrha.nrha_id = mdc.nrha_id and nrha.nrha_id = ");
            join.append(m_NrhaId);
         break;
      }

      sql.append(join);
      sql.append(where);
      sql.append("order by name, item_id");
      
      return sql.toString();
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_CustData);
      closeStmt(m_CustCost);
      closeStmt(m_CustRetail);
      closeStmt(m_CustParent);
      closeStmt(m_CustWhs);
      closeStmt(m_InvDtl);
      closeStmt(m_RmsId);
      closeStmt(m_ItemPrices);
   }

   /**
    * Creates the report title and the captions.
    */
   private int createCaptions()
   {
      XSSFFont fontTitle;
      XSSFCellStyle styleTitle;     // Bold, left
      XSSFCellStyle styleCaption;   // Bold, centered
      XSSFRow row = null;
      XSSFCell cell = null;
      int rowNum = 0;
      int col = 0;
      StringBuffer caption = new StringBuffer();

      if ( m_Sheet == null )
         return 0;

      fontTitle = m_Wrkbk.createFont();
      fontTitle.setFontHeightInPoints((short)10);
      fontTitle.setFontName("Arial");
      fontTitle.setBold(true);

      styleTitle = m_Wrkbk.createCellStyle();
      styleTitle.setFont(fontTitle);
      styleTitle.setAlignment(HorizontalAlignment.LEFT);

      styleCaption = m_Wrkbk.createCellStyle();
      styleCaption.setFont(fontTitle);
      styleCaption.setAlignment(HorizontalAlignment.CENTER);

      //
      // set the report title
      caption.append("Customer Price Quote: ");
      caption.append(m_CustId);

      row = m_Sheet.createRow(rowNum++);

      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitle);
      cell.setCellValue(new XSSFRichTextString(caption.toString()));

      //
      // Set the report selection criteria title
      caption.setLength(0);
      caption.append("selection: ");

      //
      // Labeled so we don't have a gigantic nested if statement.  When there are 16 possible
      // distribution centers, this could be ugly.
      DISTRO: {
         if ( m_DC != null && m_DC.length() > 0 ) {
            if( m_DC.equalsIgnoreCase("portland") ) {
              caption.append(" Portland ");
              break DISTRO;
            }

            if ( m_DC.equalsIgnoreCase("pittston") ) {
               caption.append(" Pittston ");
               break DISTRO;
            }

            if ( m_DC.equalsIgnoreCase("wilton") ) {
               caption.append(" Wilton ");
               break DISTRO;
            }
         };
      }

      switch ( m_FilterType ) {
         case FT_CUST:
            caption.append("Customer");
         break;

         case FT_CUST_VND:
            caption.append("Customer, Vendor");
         break;

         case FT_CUST_FLC:
            caption.append("Customer, Fine Line Class");
         break;

         case FT_CUST_ITEM:
            caption.append("Customer, Item List");
         break;

         case FT_CUST_VND_ITEM:
            caption.append("Customer, Vendor, Item List");
         break;

         case FT_CUST_FLC_ITEM:
            caption.append("Customer, Fine Line Class, Item List");
         break;

         case FT_CUST_VND_FLC_ITEM:
            caption.append("Customer, Vendor, Fine Line Class, Item List");
         break;

         case FT_MSI:
            caption.append("MSI");
         break;
      }

      caption.append("  Date: ");
      caption.append(m_RptDate);

      row = m_Sheet.createRow(rowNum++);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitle);
      cell.setCellValue(new XSSFRichTextString(caption.toString()));

      //
      // Skip an extra row before creating the captions and reset the alignment
      // before adding the column headings
      row = m_Sheet.createRow(++rowNum);

      try {
         if ( row != null ) {
            for ( int i = 0; i < m_CellStyles.length; i++ ) {
               cell = row.createCell(i);
               cell.setCellStyle(styleCaption);
            }

            row.getCell(col++).setCellValue(new XSSFRichTextString("Item"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Cust Sku"));
            row.getCell(col).setCellValue(new XSSFRichTextString("UPC-Primary"));
            m_Sheet.setColumnWidth(col++, 4000);
          	row.getCell(col++).setCellValue(new XSSFRichTextString("In Catalog"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Mfgr. Part No."));
            row.getCell(col).setCellValue(new XSSFRichTextString("Item Description"));
            m_Sheet.setColumnWidth(col++, 8000);
            row.getCell(col).setCellValue(new XSSFRichTextString("Vendor Name"));
            m_Sheet.setColumnWidth(col++, 8000);
            row.getCell(col++).setCellValue(new XSSFRichTextString("Shelf Pack"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("NBC"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Ship Unit"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Retail Unit"));
            row.getCell(col).setCellValue(new XSSFRichTextString("Customer Cost"));
            m_Sheet.setColumnWidth(col++, 4000);
            row.getCell(col).setCellValue(new XSSFRichTextString("Customer Retail"));
            m_Sheet.setColumnWidth(col++, 4000);
            row.getCell(col++).setCellValue(new XSSFRichTextString("Units Sold: Rolling 12 Mos"));
            row.getCell(col).setCellValue(new XSSFRichTextString("Dollars Sold: Rolling 12 Mos"));
            m_Sheet.setColumnWidth(col++, 4000);
            row.getCell(col++).setCellValue(new XSSFRichTextString("Flc"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Freight/Item"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Emery Cost"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Base Cost"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("A Mkt Retail"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("B Mkt Retail"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("C Mkt Retail"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("D Mkt Retail"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Sens Code"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Status"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("RMS#"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Vendor ID"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("DC"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("QOH"));
            row.getCell(col++).setCellValue(new XSSFRichTextString("Setup Date"));
         }
      }

      finally {
         row = null;
         cell = null;
         fontTitle = null;
         styleTitle = null;
         styleCaption = null;
         caption = null;
      }

      return ++rowNum;
   }

   /**
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    *
    * @return The formatted row of the spreadsheet.
    */
   private XSSFRow createRow(int rowNum)
   {
      XSSFRow row = null;
      XSSFCell cell = null;

      if ( m_Sheet != null ) {
         row = m_Sheet.createRow(rowNum);

         //
         // set the type and style of the cell.
         if ( row != null ) {
            for ( int i = 0; i < m_CellStyles.length; i++ ) {
               cell = row.createCell(i);
               cell.setCellStyle(m_CellStyles[i]);
            }
         }
      }

      return row;
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
         
         prepareStatements();
         created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[CustPriceQuote]", ex);
      }

      finally {
         closeStatements();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }
   
   /**
    * Gets the customer cost for a specific item
    *
    * @param itemId The item to check the cost on.
    * @return The customer's cost for the item.
    */
   private double getCustCost(int itemEaId)
   {
      double cost = 0.0;
      ResultSet rset = null;
      
      try {
         m_CustCost.setString(1, m_CustId);
         m_CustCost.setInt(2, itemEaId);
         rset = m_CustCost.executeQuery();

         if ( rset.next() )
            cost = rset.getDouble(1);
      }

      catch ( SQLException ex ) {
         cost = -1.0;
         try {
				m_EdbConn.rollback();
			} catch (SQLException e) {
			}
      }

      finally {
         closeRSet(rset);
         rset = null;
      }
      
      return cost;
   }

   /**
    * Gets the customer retail price for a specific item
    *
    * @param itemId The item to check the price on.
    * @return The retail price for the item and customer.
    */
   private double getCustRetail(String itemId)
   {
      double retail = 0.0;
      ResultSet rset = null;

      if ( itemId != null && itemId.length() == 7 ) {
         try {
            m_CustRetail.setString(1, m_CustId);
            m_CustRetail.setString(2, itemId);
            rset = m_CustRetail.executeQuery();

            if ( rset.next() )
               retail = rset.getDouble(1);
         }

         catch ( SQLException ex ) {
            retail = -1.0;
            try {
   				m_EdbConn.rollback();
   			} catch (SQLException e) {
   			}
         }

         finally {
            closeRSet(rset);
            rset = null;
         }
      }

      return retail;
   }

   /**
    * Gets the customer ID for a specific customer's parent company, if one exists
    *
    * @param custId   the customer number for which we want to get the parent id
    * @return         the customer number for the parent of the customer associated with custId
    *                 if one exists, otherwise custID
    * @throws SQLException
    */
   private String getParentId(String custId) throws SQLException
   {
      String parent = custId;
      m_CustParent.setString(1, custId);

      ResultSet rs = m_CustParent.executeQuery();

      try {
         if ( rs.next() )
            parent = rs.getString(1);
      }

      finally {
         closeRSet(rs);
         rs = null;
      }

      return parent;
   }

   /**
    * Retrieves the price related data for an item regardless if it's ACE or Emery.
    * @param itemId
    * @param whsId
    * @param priceData - An array to hold the price data.
    * @throws SQLException
    *
    * Note - Using an array for now to bypass possible class loading issues.
    */
   private void getPricing(int ejdItemId, int whsId, double[] priceData)
   {
      ResultSet rs = null;
      
      try {         
         m_ItemPrices.setInt(1, ejdItemId);
         m_ItemPrices.setInt(2, whsId);
         rs = m_ItemPrices.executeQuery();
            
         if ( rs.next() ) {            
            priceData[0] = rs.getDouble(1);  // buy
            priceData[1] = rs.getDouble(2);  // sell
            priceData[2] = rs.getDouble(3);  // reta
            priceData[3] = rs.getDouble(4);  // retb
            priceData[4] = rs.getDouble(5);  // retc
            priceData[5] = rs.getDouble(6);  // retd
            priceData[6] = rs.getInt(7);     // sen_code_id
         }
      }
      
      catch ( Exception ex ) {         
         log.error("[CustPriceQuote]", ex);
         try {
				m_EdbConn.rollback();
			} catch (SQLException e) {
			}
      }
   }

   /**
    * Gets the rms id for a specific item
    *
    * @return The bull vendor id if there is one, or just the vendor id if there is not.
    * @throws SQLException
    */
   private String getRmsId(String itemId) throws SQLException
   {
      ResultSet rset = null;
      String rmsId = "";

      if ( itemId != null && itemId.length() == 7 ) {
         try {
            m_RmsId.setString(1, itemId);
            rset = m_RmsId.executeQuery();

            if ( rset.next() )
               rmsId = rset.getString(1);
         }

         finally {
            closeRSet(rset);
            rset = null;
         }
      }

      return rmsId;
   }

   /**
    * Prepares the sql queries for execution.
    *
    */
   private void prepareStatements() throws Exception
   {
      StringBuffer sql = new StringBuffer(256);

      if ( m_EdbConn != null ) {
         sql.setLength(0);
         sql.append("select parent_id from customer where customer_id = ? and parent_id is not null");
         m_CustParent = m_EdbConn.prepareStatement(sql.toString());

         sql.setLength(0);
         sql.append("select warehouse_id ");
         sql.append("from cust_warehouse ");
         sql.append("where customer_id = ? ");
         sql.append("union ");
         sql.append("select warehouse_id ");
         sql.append("from warehouse ");
         sql.append("join ace_rsc on ace_rsc.ace_rsc_id = warehouse.ace_rsc_id ");
         sql.append("join ace_cust_xref on ace_cust_xref.ace_rsc = ace_rsc.sap_site_cd and customer_id = ? ");
         m_CustWhs = m_EdbConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select sum(qty_shipped) as qty_shipped, sum(ext_sell) as ext_sell, sum(ext_cost) as ext_cost ");
         sql.append("from inv_dtl ");
         sql.append("where ");
         sql.append("item_nbr = ? and invoice_date between add_months(now(), -12) and now() ");

         //
         // If there was a vendor filter, then add that to the invoice sql.
         // We need both the emery and bull (array) vendor ids.
         if ( m_VndId > 0 )
            sql.append("and vendor_nbr = ? ");

         m_InvDtl = m_EdbConn.prepareStatement(sql.toString());

         //
         // RMS id query
         sql.setLength(0);
         sql.append("select max(rms_id) as rms_id from rms_item where item_id = ?");
         m_RmsId = m_EdbConn.prepareStatement(sql.toString());
         
         //
         // cust cost query
         sql.setLength(0);
         sql.append("select price from ejd_cust_procs.get_sell_price(?, ?)");
         m_CustCost = m_EdbConn.prepareStatement(sql.toString());

         //
         // cust retail query
         sql.setLength(0);
         sql.append("select ejd.retail_price_procs.getretailprice(?, ?) as price");
         m_CustRetail = m_EdbConn.prepareStatement(sql.toString());
        
         sql.setLength(0);
         sql.append("select buy, sell, retail_a, retail_b, retail_c, retail_d, sen_code_id ");
         sql.append("from ejd_item_price ");
         sql.append("where ejd_item_id = ? and warehouse_id = ? ");
         m_ItemPrices = m_EdbConn.prepareStatement(sql.toString());
         
         //
         // Make sure this is last because it uses other sql statements to get data.
         m_CustData = m_EdbConn.prepareStatement(buildSql());         
      }
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    *
    * Because it's possible that this report can be called from some other system, the
    * best way to deal with params is to not go by the order, but by the name.
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fname = new StringBuffer();
      String tm = Long.toString(System.currentTimeMillis()).substring(3);
      int pcount = params.size();
      Param param = null;

      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);

         if ( param.name.equals("rpttype") )
            m_FilterType = Short.parseShort(param.value);

         if ( param.name.equals("custid") )
            m_CustId = param.value;

         if ( param.name.equals("vendor") && param.value.trim().length() > 0 )
            m_VndId = Integer.parseInt(param.value);

         if ( param.name.equals("flc") )
            m_FlcId = param.value;

         if ( param.name.equals("dc") )
            m_DC = param.value;

         if ( param.name.equals("itemeaid") )
            m_ItemEaId = param.value;

         if ( param.name.equals("nrha") )
            m_NrhaId = param.value;

         if ( param.name.equals("msi") )
            m_Msi = param.value;

         if ( param.name.equals("msidate") )
            m_MsiDate = param.value;

         if ( param.name.equals("rptdate") )
            m_RptDate = param.value;
      }
      
      //
      // Build the file name.
      fname.append(tm);
      fname.append("-");
      fname.append(m_RptProc.getUid());
      fname.append("cpq.xlsx");
      m_FileNames.add(fname.toString());
   }

   /**
    * Sets up the styles for the cells based on the column data.  Does any other inititialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {
      XSSFCellStyle styleText;      // Text right justified
      XSSFCellStyle styleInt;       // Style with 0 decimals
      XSSFCellStyle styleDbl4;      // double with 4 decimals "0.0000"
      XSSFCellStyle styleMoney2;    // Money ($#,##0.00_);[Red]($#,##0.00)
      XSSFCellStyle styleMoney4;    // Money ($#,##0.0000_);[Red]($#,##0.00)
      XSSFCellStyle stylePct;       // Style with 0 decimals + %
      XSSFCellStyle styleDate;      // d-mmm-yy
      XSSFDataFormat format = null;

      styleText = m_Wrkbk.createCellStyle();
      format = m_Wrkbk.createDataFormat();

      //styleText.setFont(m_FontData);
      styleText.setAlignment(HorizontalAlignment.LEFT);

      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);

      styleDbl4 = m_Wrkbk.createCellStyle();
      styleDbl4.setAlignment(HorizontalAlignment.RIGHT);
      styleDbl4.setDataFormat(format.getFormat("0.0000"));

      styleMoney2 = m_Wrkbk.createCellStyle();
      styleMoney2.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney2.setDataFormat((short)8);

      styleMoney4 = m_Wrkbk.createCellStyle();
      styleMoney4.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney4.setDataFormat(format.getFormat("$#,##0.0000_);[Red]($#,##0.0000)"));

      stylePct = m_Wrkbk.createCellStyle();
      stylePct.setAlignment(HorizontalAlignment.RIGHT);
      stylePct.setDataFormat((short)9);

      styleDate = m_Wrkbk.createCellStyle();
      styleDate.setAlignment(HorizontalAlignment.LEFT);
      styleDate.setDataFormat(format.getFormat("d-mmm-yy"));
      
      m_CellStyles = new XSSFCellStyle[] {
         styleText,    // col 0 item id
         styleText,    // col 1 cust sku
         styleText,    // col 2 upc
         styleText,    // col 3 in catalog
         styleText,    // col 4 mfg part
         styleText,    // col 5 item desc
         styleText,    // col 6 vnd name
         styleInt,     // col 7 shelf pack
         styleText,    // col 8 nbc
         styleInt,     // col 9 ship unit
         styleInt,     // col 10 ret unit
         styleMoney4,  // col 11 cust cost
         styleMoney2,  // col 12 cust retail
         styleInt,     // col 13 units sold
         styleMoney2,  // col 14 $ sold
         styleText,    // col 15 flc
         styleDbl4,    // col 16 freight
         styleMoney2,  // col 17 emery cost
         styleMoney2,  // col 18 base cost
         styleMoney2,  // col 19 retail a
         styleMoney2,  // col 20 retail b
         styleMoney2,  // col 21 retail c
         styleMoney2,  // col 22 retail d
         styleText,    // col 23 sen code
         styleText,    // col 24 status
         styleText,    // col 25 rms id
         styleText,    // col 26 vnd id
         styleText,    // col 27 DC name
         styleInt,     // col 28 qty on hand
         styleDate     // col 29 item setup date.
      };
      
      styleText = null;
      styleInt = null;
      styleMoney2 = null;
      styleMoney4 = null;
      stylePct = null;
      format = null;
   }
   
   /*
   // Main for testing purposes
   public static void main(String args[]) {
   	CustPriceQuote cpq = new CustPriceQuote();
      
      Param p1 = new Param();
      p1.name = "custid";
      p1.value = "004871";
      Param p2 = new Param();
      p2.name = "flc";
      p2.value = "2620";
      Param p3 = new Param();
      p3.name = "rpttype";
      p3.value = "2";
      Param p4 = new Param();
      p4.name = "rptdate";
      p4.value = "11/27/2017";
      Param p5 = new Param();
      p5.name = "msidate";
      p5.value = "171127";
      Param p6 = new Param();
      p6.name = "vendor";
      p6.value = "0";
      //Param p7 = new Param();
      //p7.name = "dc";
      //p7.value = null;
      Param p8 = new Param();
      p8.name = "msi";
      p8.value = "";
      Param p9 = new Param();
      p9.name = "item";
      p9.value = "";
      Param p10 = new Param();
      p10.name = "nrha";
      p10.value = "";
      
      
      ArrayList<Param> params = new ArrayList<Param>();
      params.add(p1);
      params.add(p2);
      params.add(p3);
      params.add(p4);
      params.add(p5);
      params.add(p6);
      //params.add(p7);
      params.add(p8);
      params.add(p9);
      params.add(p10);
      
      cpq.m_FilePath = "C:\\Exp\\";
      
   	java.util.Properties connProps = new java.util.Properties();
   	connProps.put("user", "ejd");
   	connProps.put("password", "boxer");
   	try {
   		cpq.m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33:5444/emery_jensen",connProps);
      
   		cpq.setParams(params);
   		cpq.createReport();
   	} catch (Exception e) {
   		e.printStackTrace();
   	}
   }*/
}
