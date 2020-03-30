/**
 * File: UnitDollarSales.java
 * Description: The unit dollar sales excel spreadsheet.
 *
 * @author Jeffrey Fisher
 * @author Paul Davidson
 *
 * $Revision: 1.31 $
 *
 * Create Date: 05/20/2005
 * Last Update: $Id: UnitDollarSales.java,v 1.31 2013/11/19 20:10:23 jfisher Exp $
 *
 * History
 *    $Log: UnitDollarSales.java,v $
 *    Revision 1.31  2013/11/19 20:10:23  jfisher
 *    Fixed the reference to the item.in_catalog flag.
 *
 *    Revision 1.30  2013/09/25 14:43:04  jfisher
 *    Updated to use current buyer not historical buyer.
 *
 *    Revision 1.29  2013/09/11 14:25:42  tli
 *    Converted the facilityId to facilityName when needed
 *
 *    Revision 1.28  2013/09/09 18:33:38  tli
 *    Replace SkuQty web service call with item_qty_view
 *
 *    Revision 1.27  2012/10/05 14:11:19  jfisher
 *    Changes to deal with the timeout on the sku quantity web service.
 *
 *    Revision 1.26  2012/08/29 19:53:02  jfisher
 *    Switched web service calls from Wasp to Axis2
 *
 *    Revision 1.25  2012/05/05 06:06:17  pberggren
 *    Removed redundant loading of system properties.
 *
 *    Revision 1.24  2012/05/03 07:55:10  prichter
 *    Fix to web service ip address
 *
 *    Revision 1.23  2012/05/03 04:31:44  pberggren
 *    Added server.properties call to force report to .57
 *
 *    Revision 1.22  2011/11/09 06:14:43  npasnur
 *    Modified the query to pull units/dollars ordered by emery from po_dtl instead of inv_dtl
 *
 *    Revision 1.21  2011/09/24 22:01:16  npasnur
 *    Modified the column widths to make it easy to understand the report.
 *
 *    Revision 1.20  2011/09/24 21:37:30  npasnur
 *    Added new column(USA) to identify items that are MADE IN USA.
 *
 *    Revision 1.19  2011/07/19 00:15:28  npasnur
 *    Added new columns to the report to display units ordered and buyers name
 *
 *    Revision 1.18  2010/09/02 02:12:47  epearson
 *    Fixed a bug in identifying the primary vendor.
 *
 *    Revision 1.17  2010/07/25 03:07:23  epearson
 *    Added Item Type field to identify vendors as primary or secondary
 *
 *    Revision 1.16  2009/08/24 17:05:43  smurdock
 *    added Pittston Available Quantity
 *
 *    Revision 1.15  2009/04/15 09:02:49  pdavidson
 *    Fixed bug with getting vendor id values
 *
 *    Revision 1.14  2009/04/02 03:48:44  pdavidson
 *    Show units and dollars sold from our warehouses
 *
 *    Revision 1.13  2009/04/02 03:45:26  pdavidson
 *    Show units and dollars sold from our warehouses
 *
 *    Revision 1.12  2009/04/01 01:23:23  pdavidson
 *    Updated for new report requirements.
 *    See: http://webws1/dotproject/index.php?m=tasks&a=view&task_id=1598
 *
 *    Revision 1.11  2009/03/19 20:45:06  jfisher
 *    Added the sale type to the where clause to remove transfers
 *
 *    Revision 1.10  2009/02/18 16:53:10  jfisher
 *    Fixed depricated methods after poi upgrade
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class UnitDollarSales extends Report
{

   private static short BASE_COLS = 40;
   private static short PT_ROLLING = 0;

   private String m_BegDate;                // Begin date for reporting, if period type not rolling 12
   private boolean m_CustXRef;              //
   private String m_EndDate;                // End date for reporting, if period type not rolling 12
   private String m_FlcId;                  // Comma delimited list of FLC ids
   private String m_ItemId;                 // Comma delimited list of item ids
   private String m_NrhaId;                 // Comma delimited list of NRHA ids
   private short m_PeriodType;              // Use rolling 12 months or specific time window
   private String m_VndId;                  // Comma delimited list of vendor ids
   private ArrayList<String> m_XRefList;    //

   private PreparedStatement m_CrossRef;
   private PreparedStatement m_ItemSales;

   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;

   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;

   private PreparedStatement m_ItemDCQty;

    /**
    * default constructor
    */
   public UnitDollarSales()
   {
      super();

      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      m_XRefList = new ArrayList<String>();
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

      if ( m_XRefList != null )
         m_XRefList.clear();

      m_Sheet = null;
      m_Wrkbk = null;
      m_CellStyles = null;
      m_XRefList = null;
      m_ItemDCQty = null;
      super.finalize();
   }

   /**
    * Adds the customer item cross reference data to the end of the spread sheet.
    * @param curRow The current row in the spreadsheet.
    * @param startCol The column to start adding data to.
    * @param item The item to cross reference.
    */
   private void addXRefCols(XSSFRow curRow, int startCol, String item) throws SQLException
   {
      ResultSet xref = null;
      String custId = null;
      String custName = null;
      String custSku = null;
      XSSFRow caption = null;
      XSSFCell cell = null;
      int index;
      int col;

      m_CrossRef.setString(1, item);
      xref = m_CrossRef.executeQuery();

      try {
         while ( xref.next() ) {
            custId = xref.getString(1);
            custName = xref.getString(2);
            custSku = xref.getString(3);

            index = m_XRefList.indexOf(custId);

            if ( index == -1 ) {
               m_XRefList.add(custId);
               col = ((m_XRefList.size()-1) + startCol);
               caption = m_Sheet.getRow(2);
               cell = caption.createCell(col);
               cell.setCellValue(new XSSFRichTextString(custName));
            }
            else
               col = (index + startCol);

            cell = curRow.createCell(col);
            cell.setCellValue(new XSSFRichTextString(custSku));
         }
      }

      finally {
         cell = null;
         caption = null;
         custId = null;
         custName = null;
         custSku = null;

         closeRSet(xref);
      }
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
      ResultSet itemSales = null;
      int colCnt = BASE_COLS;
      int rowNum = 1;
      boolean result = false;
      String item = "";
      String vendorId = "";
      String primaryVendorId = "";
      String upc = null;
      String usaItem = "";
      int qtyShipped = 0;
      double sold = 0.0;
      double cost = 0.0;
      double marginPct = 0.0;
      double margin = 0.0;
      double qtyOrdered = 0.0;
      double ordered = 0.0;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         rowNum = createCaptions();
         itemSales = m_ItemSales.executeQuery();

         while ( itemSales.next() && m_Status == RptServer.RUNNING ) {
            row = createRow(rowNum, colCnt);
            item = itemSales.getString("item_id");
            vendorId = itemSales.getString("vendor_id");
            primaryVendorId = itemSales.getString("primary_vendor");
            upc = itemSales.getString("upc_code");

            setCurAction("processing item: " + item);

            qtyOrdered = itemSales.getInt("tot_qty_ordered");
            qtyShipped = itemSales.getInt("tot_qty_shipped");
            ordered = itemSales.getDouble("tot_dollars_ordered");
            sold = itemSales.getDouble("tot_dollars_sold");
            cost = itemSales.getDouble("tot_dollars_cost");
            margin = sold - cost;

            if ( sold > 0 )
               marginPct = margin/sold;
            else
               marginPct = 0.0;

            //
            //09/23/2011. Naresh
            //MADE IN USA items
            usaItem = itemSales.getString("usa_item");
            usaItem = (usaItem == null? "": "USA");

            row.getCell(0).setCellValue(new XSSFRichTextString(itemSales.getString("vendor_name")));
            row.getCell(1).setCellValue(new XSSFRichTextString(itemSales.getString("vendor_id")));
            row.getCell(2).setCellValue(new XSSFRichTextString(primaryVendorId.equals(vendorId) ? "Primary" : "Secondary"));
            row.getCell(3).setCellValue(new XSSFRichTextString(item));
            row.getCell(4).setCellValue(new XSSFRichTextString(usaItem));
            row.getCell(5).setCellValue(new XSSFRichTextString(itemSales.getString("vendor_item_num")));
            row.getCell(6).setCellValue(new XSSFRichTextString(itemSales.getString("stock_pack")));
            row.getCell(7).setCellValue(new XSSFRichTextString(itemSales.getString("ship_unit")));
            row.getCell(8).setCellValue(new XSSFRichTextString(itemSales.getString("retail_unit")));
            row.getCell(9).setCellValue(new XSSFRichTextString(itemSales.getString("retail_pack")));
            row.getCell(10).setCellValue(new XSSFRichTextString(upc != null ? upc : ""));
            row.getCell(11).setCellValue(new XSSFRichTextString(itemSales.getString("description")));
            row.getCell(12).setCellValue(itemSales.getDouble("buy"));
            row.getCell(13).setCellValue(itemSales.getDouble("sell"));
            row.getCell(14).setCellValue(itemSales.getDouble("retail_a"));
            row.getCell(15).setCellValue(itemSales.getDouble("retail_b"));
            row.getCell(16).setCellValue(itemSales.getDouble("retail_c"));
            row.getCell(17).setCellValue(itemSales.getDouble("retail_d"));
            row.getCell(18).setCellValue(qtyShipped);
            row.getCell(19).setCellValue(sold);
            row.getCell(20).setCellValue(itemSales.getInt("portland_qty_shipped"));
            row.getCell(21).setCellValue(itemSales.getDouble("portland_dollars_sold"));
            row.getCell(22).setCellValue(itemSales.getInt("pittston_qty_shipped"));
            row.getCell(23).setCellValue(itemSales.getDouble("pittston_dollars_sold"));
            row.getCell(24).setCellValue(new XSSFRichTextString(itemSales.getString("sen_code_id")));
            row.getCell(25).setCellValue(marginPct);
            row.getCell(26).setCellValue(margin);
            row.getCell(27).setCellValue(getAvailQty(item, "PORTLAND"));//01
            row.getCell(28).setCellValue(getAvailQty(item, "PITTSTON"));//04
            row.getCell(29).setCellValue(new XSSFRichTextString(itemSales.getString("flc_id")));
            row.getCell(30).setCellValue(new XSSFRichTextString(itemSales.getString("nbc")));
            row.getCell(31).setCellValue(new XSSFRichTextString(itemSales.getInt("in_catalog") == 0 ? "no" : "yes"));
            row.getCell(32).setCellValue(new XSSFRichTextString(itemSales.getString("velocity")));
            row.getCell(33).setCellValue(qtyOrdered);
            row.getCell(34).setCellValue(ordered);
            row.getCell(35).setCellValue(itemSales.getDouble("portland_qty_ordered"));
            row.getCell(36).setCellValue(itemSales.getDouble("pittston_qty_ordered"));
            row.getCell(37).setCellValue(itemSales.getDouble("portland_dollars_ordered"));
            row.getCell(38).setCellValue(itemSales.getDouble("pittston_dollars_ordered"));
            row.getCell(39).setCellValue(new XSSFRichTextString(itemSales.getString("buyer_name")));

            if ( m_CustXRef )
               addXRefCols(row, 40, item);

            rowNum++;
            upc = null;
         }

         m_Wrkbk.write(outFile);
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         RptServer.log.error("[UnitDollarSales]", ex);
      }

      finally {
         row = null;
         closeRSet(itemSales);

         try {
            outFile.close();
         }

         catch( Exception e ) {
            RptServer.log.error("[UnitDollarSales]", e);
         }

         outFile = null;
      }

      return result;
   }

   /**
    * Builds the sql based on the type of filter requested by the user.
    * @return A complete sql statement.
    */
   private String buildSql()
   {
      //boolean condition = false;
      StringBuffer sql = new StringBuffer();

      sql.append("select distinct");
      sql.append("   vendor.vendor_id, vendor.name vendor_name, item_entity_attr.item_id, vendor_item_num, item_entity_attr.description, ");
      sql.append(" 	 item_entity_attr.vendor_id as primary_vendor, ejd_item_price.buy as buy, sell, ");
      sql.append("   retail_a, retail_b, retail_c, retail_d, sen_code_id, ejd_item.flc_id, ship_unit.unit as ship_unit, ");
      sql.append("   retail_unit.unit as retail_unit, retail_pack, cat.stock_pack, ejd_item_whs_upc.upc_code, ");
      sql.append("   decode(broken_case.description, 'ALLOW BROKEN CASES', 'N', 'Y') as nbc, ");
      sql.append("   cat.in_catalog, /*velocity*/'-' as velocity, buyer.name as buyer_name, ia.item_id as usa_item, ");
      sql.append(" 	 coalesce(po_det.tot_qty_ordered, 0) as tot_qty_ordered, ");
      sql.append("   coalesce(invd.tot_qty_shipped, 0) as tot_qty_shipped, ");
      sql.append("   coalesce(po_det.tot_dollars_ordered, 0) as tot_dollars_ordered, ");
      sql.append("   coalesce(invd.tot_dollars_sold, 0) as tot_dollars_sold, ");
      sql.append("   coalesce(invd.tot_dollars_cost, 0) as tot_dollars_cost, ");
      sql.append("   coalesce(po_det.portland_qty_ordered, 0) as portland_qty_ordered, ");
      sql.append("   coalesce(invd.portland_qty_shipped, 0) as portland_qty_shipped, ");
      sql.append(" 	 coalesce(po_det.portland_dollars_ordered, 0) as portland_dollars_ordered, ");
      sql.append("   coalesce(invd.portland_dollars_sold, 0) as portland_dollars_sold, ");
      sql.append("   coalesce(invd.portland_dollars_cost, 0) as portland_dollars_cost, ");
      sql.append(" 	 coalesce(po_det.pittston_qty_ordered, 0) as pittston_qty_ordered, ");
      sql.append("   coalesce(invd.pittston_qty_shipped, 0) as pittston_qty_shipped, ");
      sql.append(" 	 coalesce(po_det.pittston_dollars_ordered, 0) as pittston_dollars_ordered, ");
      sql.append("   coalesce(invd.pittston_dollars_sold, 0) as pittston_dollars_sold, ");
      sql.append("   coalesce(invd.pittston_dollars_cost, 0) as pittston_dollars_cost ");
      sql.append("from ");
      sql.append("   item_entity_attr ");
      sql.append("join ( ");
      sql.append("   select ejd_item_id, sum(in_catalog) as in_catalog, warehouse_id, stock_pack ");
      sql.append("   from ejd_item_warehouse ");
      sql.append("   where warehouse_id in (1,2) ");
      sql.append("   group by ejd_item_id, warehouse_id, stock_pack ");
      sql.append(") cat on cat.ejd_item_id = item_entity_attr.ejd_item_id ");
      sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = item_entity_attr.ejd_item_id ");
      sql.append("   and ejd_item_price.warehouse_id = cat.warehouse_id ");
     // sql.append("   item_price.sell_date = ( ");
     // sql.append("   select max(ip.sell_date) from item_price ip ");
     // sql.append("   where ip.sell_date <= trunc(sysdate) and ");
     // sql.append("   ip.item_id = item.item_id) ");
     //sql.append("join item_velocity on item.velocity_id = item_velocity.velocity_id ");  temporarily leave this out until the report can be converted to use at the item_warehouse level
      sql.append("join vendor_item_ea_cross on item_entity_attr.item_ea_id = vendor_item_ea_cross.item_ea_id ");
      sql.append("join vendor on vendor_item_ea_cross.vendor_id = vendor.vendor_id ");
      sql.append("join vendor_dept on vendor_dept.vendor_id = vendor.vendor_id ");
      sql.append("join emery_dept on emery_dept.dept_id = vendor_dept.dept_id ");
      sql.append("join buyer on buyer.buyer_id = emery_dept.buyer_id ");
      sql.append("join ship_unit on item_entity_attr.ship_unit_id = ship_unit.unit_id ");
      sql.append("join retail_unit on item_entity_attr.ret_unit_id = retail_unit.unit_id ");
      sql.append("join ejd_item on item_entity_attr.ejd_item_id = ejd_item.ejd_item_id ");
      sql.append("join broken_case on ejd_item.broken_case_id = broken_case.broken_case_id ");
      sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = ejd_item.ejd_item_id and primary_upc = 1 ");
      sql.append("join item_type on item_entity_attr.item_type_id = item_type.item_type_id and item_type.itemtype = 'STOCK' ");

      if ( m_NrhaId != null && m_NrhaId.length() > 0 ) {  // if filtering by nrha need to join with appropriate tables to get there
         sql.append("join flc on ejd_item.flc_id = flc.flc_id ");
         sql.append("join mdc on flc.mdc_id = mdc.mdc_id ");
         sql.append("join nrha on mdc.nrha_id = nrha.nrha_id ");
      }

      sql.append("   left outer join ( ");
      sql.append("      select ");
      sql.append("        item_ea_id, ");
      sql.append("        vendor_nbr, ");
      sql.append("        sum(qty_shipped) tot_qty_shipped, ");
      sql.append("        sum(ext_sell) tot_dollars_sold, ");
      sql.append("        sum(ext_cost) tot_dollars_cost, ");

      // This was the most efficient way to get at this data from within the same query.  If we add another
      // warehouse in future, which is unlikely, it will have to be added to here, if that ever happens.
      sql.append("        sum(decode(warehouse, 'PORTLAND', qty_shipped, 0)) portland_qty_shipped, ");
      sql.append("        sum(decode(warehouse, 'PORTLAND', ext_sell, 0)) portland_dollars_sold, ");
      sql.append("        sum(decode(warehouse, 'PORTLAND', ext_cost, 0)) portland_dollars_cost, ");
      sql.append("        sum(decode(warehouse, 'PITTSTON', qty_shipped, 0)) pittston_qty_shipped, ");
      sql.append("        sum(decode(warehouse, 'PITTSTON', ext_sell, 0)) pittston_dollars_sold, ");
      sql.append("        sum(decode(warehouse, 'PITTSTON', ext_cost, 0)) pittston_dollars_cost ");
      sql.append("      from ");
      sql.append("        inv_dtl where sale_type = 'WAREHOUSE' and ");

      // Set the period
      if ( m_PeriodType == PT_ROLLING )
         sql.append("(invoice_date > add_months(current_date, -12)) and (invoice_date <= current_date) ");
      else {
         sql.append(String.format("invoice_date >= to_date('%s', 'mm/dd/yyyy') ", m_BegDate));
         sql.append(String.format("and invoice_date <= to_date('%s', 'mm/dd/yyyy') ", m_EndDate));
      }

      sql.append("     group by item_ea_id, vendor_nbr ");  // group by item# and vendor#, as want to see sales by vendor
      sql.append("   ) invd on item_entity_attr.item_ea_id = invd.item_ea_id and vendor.vendor_id = invd.vendor_nbr ");

      //
      // 11/08/2011. Naresh
      // Emery PO Details
      sql.append("   left outer join ( ");
      sql.append("      select ");
      sql.append("         item_ea_id, vendor_id, ");
      sql.append("         sum(qty_ordered) tot_qty_ordered, ");
      sql.append("         sum(emery_cost) tot_dollars_ordered, ");
      sql.append("         sum(decode(po_dtl.warehouse, '01', qty_ordered, 0)) portland_qty_ordered, ");
      sql.append("         sum(decode(po_dtl.warehouse, '01', emery_cost, 0)) portland_dollars_ordered, ");
      sql.append("         sum(decode(po_dtl.warehouse, '04', qty_ordered, 0)) pittston_qty_ordered, ");
      sql.append("         sum(decode(po_dtl.warehouse, '04', emery_cost, 0)) pittston_dollars_ordered ");
      sql.append("      from ");
      sql.append("         po_hdr, po_dtl ");
      sql.append("      where ");
      sql.append("         po_hdr.po_hdr_id = po_dtl.po_hdr_id and ");

      // Set the period
      if ( m_PeriodType == PT_ROLLING )
         sql.append("      (po_hdr.po_date > add_months(current_date, -12)) and (po_hdr.po_date <= current_date) ");
      else {
         sql.append(String.format("po_hdr.po_date >= to_date('%s', 'mm/dd/yyyy') ", m_BegDate));
         sql.append(String.format("and po_hdr.po_date <= to_date('%s', 'mm/dd/yyyy') ", m_EndDate));
      }

      sql.append("      group by item_ea_id, vendor_id ");  // group by item# and vendor#, as want to see sales by vendor
      sql.append("      ) po_det on item_entity_attr.item_ea_id = po_det.item_ea_id and vendor.vendor_id = po_det.vendor_id  ");

      //
      // 09/23/2011. Naresh
      // Identifies MADE IN USA items
      sql.append("left outer join ( " );
      sql.append("  select item_id, ejd_item_attribute.ejd_item_id, attribute_value_id ");
      sql.append("  from ejd_item_attribute join item_entity_attr on ejd_item_attribute.ejd_item_id = item_entity_attr.ejd_item_id ");
      sql.append("and attribute_value_id in ( select attribute_value_id from attribute a , attribute_value av where a.attribute_id = av.attribute_id and av.value = 'MADE IN USA') ");
      sql.append(" )ia on item_entity_attr.ejd_item_id = ia.ejd_item_id ");
      
      //sql.append("item_attribute ia on item.item_id = ia.item_id and ");
      //sql.append("ia.attribute_value_id in ( select attribute_value_id from attribute a , attribute_value av where a.attribute_id = av.attribute_id and av.value = 'MADE IN USA') ");

      sql.append("where item_entity_attr.item_type_id < 8");
      
      if ( m_VndId != null && m_VndId.length() > 0 ) {
         sql.append(" and vendor.vendor_id in (");
         sql.append(m_VndId);
         sql.append(") ");

         //condition = true;
      }

      if ( m_NrhaId != null && m_NrhaId.length() > 0 ) {
         sql.append((" and ") + "nrha.nrha_id in (");
         sql.append(m_NrhaId);
         sql.append(") ");

         //condition = true;
      }

      if ( m_FlcId != null && m_FlcId.length() > 0 ) {
         sql.append((" and ") + "ejd_item.flc_id in (");
         sql.append(m_FlcId);
         sql.append(") ");

         //condition = true;
      }

      if ( m_ItemId != null && m_ItemId.length() > 0  ) {
         sql.append((" and ") + "item_entity_attr.item_id in (");
         sql.append(m_ItemId);
         sql.append(") ");
      }

      sql.append("order by vendor.name, item_id");
      return sql.toString();
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_ItemSales);
      closeStmt(m_CrossRef);
      closeStmt(m_ItemDCQty);
   }

   /**
    * Creates the report title and the captions.
    */
   private int createCaptions()
   {
      XSSFFont fontTitle;
      XSSFCellStyle styleTitle;   // Bold, centered
      XSSFRow row = null;
      XSSFCell cell = null;
      int rowNum = 0;
      StringBuffer caption = new StringBuffer("Unit and Dollar Sales Report: ");

      if ( m_Sheet == null )
         return 0;

      fontTitle = m_Wrkbk.createFont();
      fontTitle.setFontHeightInPoints((short)10);
      fontTitle.setFontName("Arial");
      fontTitle.setBold(true);

      styleTitle = m_Wrkbk.createCellStyle();
      styleTitle.setFont(fontTitle);
      styleTitle.setAlignment(HorizontalAlignment.CENTER);

      //
      // set the report title
      row = m_Sheet.createRow(rowNum);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitle);

      if ( m_PeriodType == PT_ROLLING )
         caption.append("Rolling 12 Months");
      else {
         caption.append(m_BegDate);
         caption.append(" - ");
         caption.append(m_EndDate);
      }

      cell.setCellValue(new XSSFRichTextString(caption.toString()));

      rowNum = 2;
      row = m_Sheet.createRow(rowNum);

      try {
         if ( row != null ) {
            for ( int i = 0; i < BASE_COLS; i++ ) {
               cell = row.createCell(i);
               cell.setCellStyle(styleTitle);
            }

            row.getCell(0).setCellValue(new XSSFRichTextString("Vendor Name"));
            m_Sheet.setColumnWidth(0, 8000);
            row.getCell(1).setCellValue(new XSSFRichTextString("Vendor ID"));
            row.getCell(2).setCellValue(new XSSFRichTextString("Item Type"));
            row.getCell(3).setCellValue(new XSSFRichTextString("Item #"));
            row.getCell(4).setCellValue(new XSSFRichTextString("USA"));
            row.getCell(5).setCellValue(new XSSFRichTextString("Mfgr. Part No."));
            m_Sheet.setColumnWidth(5, 8000);
            row.getCell(6).setCellValue(new XSSFRichTextString("Stock Pack"));
            m_Sheet.setColumnWidth(6, 3000);
            row.getCell(7).setCellValue(new XSSFRichTextString("Ship Unit"));
            m_Sheet.setColumnWidth(7, 3000);
            row.getCell(8).setCellValue(new XSSFRichTextString("Retail Unit"));
            m_Sheet.setColumnWidth(8, 3000);
            row.getCell(9).setCellValue(new XSSFRichTextString("Dealer Pack"));
            m_Sheet.setColumnWidth(9, 3000);
            row.getCell(10).setCellValue(new XSSFRichTextString("UPC-Primary"));
            m_Sheet.setColumnWidth(10, 5000);
            row.getCell(11).setCellValue(new XSSFRichTextString("Item Description"));
            m_Sheet.setColumnWidth(11, 14000);
            row.getCell(12).setCellValue(new XSSFRichTextString("Emery Cost"));
            m_Sheet.setColumnWidth(12, 3000);
            row.getCell(13).setCellValue(new XSSFRichTextString("Base Cost"));
            m_Sheet.setColumnWidth(13, 3000);
            row.getCell(14).setCellValue(new XSSFRichTextString("A Mkt Retail"));
            m_Sheet.setColumnWidth(14, 3000);
            row.getCell(15).setCellValue(new XSSFRichTextString("B Mkt Retail"));
            m_Sheet.setColumnWidth(15, 3000);
            row.getCell(16).setCellValue(new XSSFRichTextString("C Mkt Retail"));
            m_Sheet.setColumnWidth(16, 3000);
            row.getCell(17).setCellValue(new XSSFRichTextString("D Mkt Retail"));
            m_Sheet.setColumnWidth(17, 3000);
            row.getCell(18).setCellValue(new XSSFRichTextString("Units Sold"));
            row.getCell(19).setCellValue(new XSSFRichTextString("Dollars Sold"));
            m_Sheet.setColumnWidth(19, 4000);
            row.getCell(20).setCellValue(new XSSFRichTextString("Units Sold Portland"));
            m_Sheet.setColumnWidth(20, 4500);
            row.getCell(21).setCellValue(new XSSFRichTextString("Dollars Sold Portland"));
            m_Sheet.setColumnWidth(21, 5000);
            row.getCell(22).setCellValue(new XSSFRichTextString("Units Sold Pittston"));
            m_Sheet.setColumnWidth(22, 4500);
            row.getCell(23).setCellValue(new XSSFRichTextString("Dollars Sold Pittston"));
            m_Sheet.setColumnWidth(23, 5000);
            row.getCell(24).setCellValue(new XSSFRichTextString("Sensitivity Code"));
            m_Sheet.setColumnWidth(24, 4000);
            row.getCell(25).setCellValue(new XSSFRichTextString("Emery Margin%"));
            m_Sheet.setColumnWidth(25, 4000);
            row.getCell(26).setCellValue(new XSSFRichTextString("Emery Margin$"));
            m_Sheet.setColumnWidth(26, 4000);
            row.getCell(27).setCellValue(new XSSFRichTextString("Portland On Hand"));
            m_Sheet.setColumnWidth(27, 4000);
            row.getCell(28).setCellValue(new XSSFRichTextString("Pittston On Hand"));
            m_Sheet.setColumnWidth(28, 4000);
            row.getCell(29).setCellValue(new XSSFRichTextString("FLC"));
            row.getCell(30).setCellValue(new XSSFRichTextString("NBC"));
            row.getCell(31).setCellValue(new XSSFRichTextString("Catalog"));
            row.getCell(32).setCellValue(new XSSFRichTextString("Velocity"));
            row.getCell(33).setCellValue(new XSSFRichTextString("Units Ordered"));
            m_Sheet.setColumnWidth(33, 4500);
            row.getCell(34).setCellValue(new XSSFRichTextString("Dollars Ordered"));
            m_Sheet.setColumnWidth(34, 5000);
            row.getCell(35).setCellValue(new XSSFRichTextString("Units Ord Portland"));
            m_Sheet.setColumnWidth(35, 4500);
            row.getCell(36).setCellValue(new XSSFRichTextString("Units Ord Pittston"));
            m_Sheet.setColumnWidth(36, 4500);
            row.getCell(37).setCellValue(new XSSFRichTextString("Dollars Ord Portland"));
            m_Sheet.setColumnWidth(37, 5000);
            row.getCell(38).setCellValue(new XSSFRichTextString("Dollars Ord Pittston"));
            m_Sheet.setColumnWidth(38, 5000);
            row.getCell(39).setCellValue(new XSSFRichTextString("Buyer's Name"));
            m_Sheet.setColumnWidth(39, 8000);

         }
      }

      finally {
         row = null;
         cell = null;
         fontTitle = null;
         styleTitle = null;
         caption = null;
      }

      return ++rowNum;
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
         RptServer.log.error("[UnitDollarSales]", ex);
      }

      finally {
         closeStatements();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Gets the available (Portland) quantity from fascor.  This is done through querying view
    * that directly copied from fascor.  It returns the available qty after the
    * appropriate adjustments have been made.
    *
    * @param item The item to check.
    *
    * @return The quantity available of the item in fascor
    * @throws Exception
    */
   private int getAvailQty(String item, String warehouse) throws Exception
   {
	   int qty = 0;
	   ResultSet rset = null;

	   if ( item != null && item.length() == 7 ) {
	      try {
	         m_ItemDCQty.setString(1, item);
         	 m_ItemDCQty.setString(2, warehouse);
	         rset = m_ItemDCQty.executeQuery();

	         if ( rset.next() )
	            qty = rset.getInt("avail_qty");
	      }
	      
	      finally {
	         closeRSet(rset);
	         rset = null;
	      }
	   }

	   return qty;
   }

   /**
    * Prepares the sql queries for execution.
    */
   private void prepareStatements() throws Exception
   {
      StringBuffer sql = new StringBuffer();

      if ( m_EdbConn != null ) {
         m_ItemSales = m_EdbConn.prepareStatement(buildSql());

         sql.append("select customer.customer_id, name, customer_sku ");
         sql.append("from customer ");
         sql.append("join item_ea_cross on item_ea_cross.customer_id = customer.customer_id ");
         sql.append("join item_entity_attr on item_entity_attr.item_ea_id = item_ea_cross.item_ea_id ");
         sql.append("where item_entity_attr.item_id = ?");
         m_CrossRef = m_EdbConn.prepareStatement(sql.toString());

         sql.setLength(0);
         sql.append("select avail_qty ");
         sql.append("from ejd_item_warehouse ");
         sql.append("join warehouse on warehouse.warehouse_id = ejd_item_warehouse.warehouse_id ");
         sql.append("join item_entity_attr on item_entity_attr.ejd_item_id = ejd_item_warehouse.ejd_item_id ");
         sql.append("where item_entity_attr.item_id = ? and warehouse.name = ?");
         m_ItemDCQty = m_EdbConn.prepareStatement(sql.toString());
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

         if ( param.name.equals("prdtype") )
            m_PeriodType = Short.parseShort(param.value);

         if ( param.name.equals("nrha") )
            m_NrhaId = param.value;

         if ( param.name.equals("flc") )
            m_FlcId = param.value;

         if ( param.name.equals("vendor") && param.value.trim().length() > 0 )
            m_VndId = param.value;

         if ( param.name.equals("item") )
            m_ItemId = param.value;

         if ( param.name.equals("xref") )
            m_CustXRef = Boolean.parseBoolean(param.value);

         if ( param.name.equals("begdate") )
            m_BegDate = param.value;

         if ( param.name.equals("enddate") )
            m_EndDate = param.value;
      }

      //
      // Build the file name.
      fname.append(tm);
      fname.append("-");
      fname.append(m_RptProc.getUid());
      fname.append("uds.xlsx");
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
      XSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00)
      XSSFCellStyle stylePct;       // Style with 0 decimals + %

      styleText = m_Wrkbk.createCellStyle();
      //styleText.setFont(m_FontData);
      styleText.setAlignment(HorizontalAlignment.LEFT);

      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);

      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short)8);

      stylePct = m_Wrkbk.createCellStyle();
      stylePct.setAlignment(HorizontalAlignment.RIGHT);
      stylePct.setDataFormat((short)9);

      m_CellStyles = new XSSFCellStyle[] {
         styleText,    // col 0 vnd name
         styleText,    // col 1 vnd id
         styleText,	   // col 1b primary vendor
         styleText,    // col 3 item
         styleText,    // col 4 usa
         styleText,    // col 5 mfg part#
         styleInt,     // col 6 stock pack
         styleText,    // col 7 ship unit
         styleText,    // col 8 retail unit
         styleInt,     // col 9 retail pack
         styleText,    // col 10 upc
         styleText,    // col 11 item desc
         styleMoney,   // col 12 buy
         styleMoney,   // col 13 sell
         styleMoney,   // col 14 retail a
         styleMoney,   // col 15 retail b
         styleMoney,   // col 16 retail c
         styleMoney,   // col 17 retail d

         styleInt,     // col 18 qty shipped
         styleMoney,   // col 19 $ sold

         styleInt,     // col 20 portland qty shipped
         styleMoney,   // col 21 portland $ sold
         styleInt,     // col 22 pittston qty shipped
         styleMoney,   // col 23 pittston $ sold

         styleInt,     // col 24 sen code
         stylePct,     // col 25 margin pct
         styleMoney,   // col 26 margin
         styleInt,     // col 27 qty avail Portland
         styleInt,     // col 28 qty avail Pittson
         styleText,    // col 29 flc
         styleText,    // col 30 nbc
         styleText,    // col 31 in catalog
         styleText,     // col 32 velocity code
         styleText,    // col 33 units ord
         styleText,     // col 34 $ ord
         styleText,    // col 35 units ord Portland
         styleText,     // col 36 units ord Pittston
         styleText,    // col 37 $ ord Portland
         styleText,     // col 38 $ ord Pittston
         styleText,    // col 39 buyers name
         styleText     // col 40 new column
      };

      styleText = null;
      styleInt = null;
      styleMoney = null;
      stylePct = null;
   }
}
