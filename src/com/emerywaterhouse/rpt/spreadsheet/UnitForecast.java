/**
 * File: UnitForecast.java
 * Description: Unit Forecasting report containing data on current sales trends,
 *    promotional sales history and same period replenishment numbers.
 *
 *    This report includes a separation of sales numbers for replenishment,
 *    best price and promo sales.  Show columns for different types (demand type)
 *    of promotions like in the Prescient application.  Also shows sales numbers
 *    over the selected time period for a chosen list of packets.  The packet
 *    filter does not affect how many records are returned, but rather augments
 *    the existing results.
 *
 *    For promo demand type: 4 columns for units, and 4 columns for $$ (1 for each promo type displayed in Prescient).
 *
 *    Uses the same inputs as the Unit-Dollar Sales report, except for the additional packet selection.
 *
 * @author Paul Davidson
 *
 * $Revision: 1.26 $
 *
 * Create Date: 4/6/09
 * Last Update: $Id: UnitForecast.java,v 1.26 2013/09/09 18:33:38 tli Exp $
 *
 * History:
 *    $Log: UnitForecast.java,v $
 *    Revision 1.26  2013/09/09 18:33:38  tli
 *    Replace SkuQty web service call with item_qty_view
 *
 *    Revision 1.25  2012/08/29 19:53:02  jfisher
 *    Switched web service calls from Wasp to Axis2
 *
 *    Revision 1.24  2012/07/19 19:38:00  jfisher
 *    in_catalog at the warehouse level changes
 *
 *    Revision 1.23  2012/05/05 06:06:04  pberggren
 *    Removed redundant loading of system properties.
 *
 *    Revision 1.22  2012/05/03 07:55:10  prichter
 *    Fix to web service ip address
 *
 *    Revision 1.21  2012/05/03 04:45:17  pberggren
 *    Added server.properties call to force report to .57
 *
 *    Revision 1.20  2011/09/24 20:56:14  npasnur
 *    Added new column(USA) to identify items that are MADE IN USA.
 *
 *    Revision 1.19  2011/08/06 06:11:50  jfisher
 *    Added discontinued item filter
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import com.emerywaterhouse.rpt.helper.EmeryRptSheet;
import com.emerywaterhouse.rpt.helper.EmeryRptSheet.Column;
import com.emerywaterhouse.rpt.helper.EmeryRptWorkbook;
import com.emerywaterhouse.rpt.helper.EmeryRptWorkbook.Format;
import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class UnitForecast extends Report
{
   /**
    * This class handles the customer xref data.
    * It allows for one call to be made to the database and allows for
    * easy future updates to the report. -epearson
    */
   public class CustXRefData{
      // lookup table for customer name based on id
      private final Map<String,String> m_CustIdNameMap;

      // lookup table for customer sku based on item id and customer
      private final Map<String,String> m_ItemCustSkuMap;

      public CustXRefData(ResultSet rs) throws SQLException{
         m_CustIdNameMap = new HashMap<String,String>();
         m_ItemCustSkuMap = new HashMap<String,String>();

         String itemId;
         String custId;
         String custName;
         String custSku;

         while(rs.next()){
            itemId = rs.getString("item_id");
            custId = rs.getString("customer_id");
            custName = rs.getString("name");
            custSku = rs.getString("customer_sku");

            m_CustIdNameMap.put(custId,custName);
            m_ItemCustSkuMap.put(createKey(itemId, custId), custSku);
         }
      }

      /**
       * Combines itemid and customer name to create a unique key
       *
       * @param itemId
       * @param custName
       * @return			a key
       */
      private String createKey(String itemId, String custId){
         return itemId + custId;
      }

      /**
       * Returns the customer names as an array
       *
       * @param index
       * @return		customer id
       */
      public String[] getCustIds(){
         return m_CustIdNameMap.keySet().toArray(new String[0]);
      }

      /**
       * Returns the customer names as an array
       *
       * @param index
       * @return		customer name
       */
      public String[] getCustNames(){
         return m_CustIdNameMap.values().toArray(new String[0]);
      }

      /**
       * Get a customer sku given the customer name and item id
       *
       * @param itemId
       * @param custName
       * @return	the corresponding customer sku or null
       */
      public String getCustSku(String itemId, String custId){
         return m_ItemCustSkuMap.get(createKey(itemId, custId));
      }
   }

   private static short PT_ROLLING = 0;

   private static final int DEFAULT_COL_WIDTH = 10;
   private static final int NAME_COL_WIDTH = 20;
   private static final int DESCRIP_COL_WIDTH = 40;
   private static final int QTY_COL_WIDTH = 8;
   private static final int CURRENCY_COL_WIDTH = 15;

   private String m_BegDate;                // Begin date for reporting, if period type not rolling 12
   private boolean m_CustXRef;              //
   private boolean m_ShowDiscItems;         // Flag to determine whether to include discontinued items.
   private String m_EndDate;                // End date for reporting, if period type not rolling 12
   private String m_FlcId;                  // Comma delimited list of FLC ids
   private String m_ItemId;                 // Comma delimited list of item ids
   private String m_NrhaId;                 // Comma delimited list of NRHA ids
   private String m_Packets;                // Comma delimited list of packets ids as a string value
   private String[] m_PacketArray;          // Produced from splitting up the m_Packets string value
   private short m_PeriodType;              // Use rolling 12 months or specific time window
   private String m_VndId;                  // Comma delimited list of vendor ids
   private String m_WhseName;               // Warehouse name parameter, e.g. PORTLAND, PITTSTON
   private String m_WhseFac;                // Warehouse facility id parameter, e.g. 01, 04

   private PreparedStatement m_CrossRef;    //
   private PreparedStatement m_ItemSales;   // Main query to get item sales history and other extraneous data
   private final EmeryRptWorkbook m_Wrkbk;  // POI workbook object wrapper
   private EmeryRptSheet m_Sheet;           // POI worksheet object wrapper
   private Column[] m_Fields; 				  // fields/column headers
   private CustXRefData m_CustXRefData;     // customer xref data
   private String[] m_CustXRefNames;        // customer xref names (used as header captions)
   private String[] m_CustXRefIds;          // customer ids (used to lookup customer sku)
   
   private PreparedStatement m_ItemDCQty;

   /**
    * Default constructor
    */
   public UnitForecast()
   {
      super();

      m_Wrkbk = new EmeryRptWorkbook();
      m_Packets = "";
      m_PacketArray = new String[]{}; // Create initial empty array so code doesn't blow up later
      m_WhseName = null;
      m_WhseFac = null;

      //
      // By default always include discontinued items.
      m_ShowDiscItems = true;
   }

   /**
    * Executes the queries and builds the report workbook
    *
    * @return true if the file was built, false if not.
    * @throws FileNotFoundException
    */
   private boolean buildWorkbook() throws FileNotFoundException
   {
      FileOutputStream outFile = null;
      ResultSet itemSales = null;
      int col = 0;
      boolean result = false;
      String item = "";
      String nbc;
      String upc = null;
      int qtyShipped = 0;
      double sold = 0.0;
      double cost = 0.0;
      double marginPct = 0.0;
      double margin = 0.0;
      String primaryVendorId = null;
      String usaItem = "";

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      // Build Title Caption
      StringBuffer caption = new StringBuffer("Unit Forecast Report: ");
      if ( m_PeriodType == PT_ROLLING )
         caption.append("Rolling 12 Months");
      else {
         caption.append(m_BegDate);
         caption.append(" - ");
         caption.append(m_EndDate);
      }

      if ( m_WhseName != null )
         caption.append(" - " + m_WhseName);

      try {
         m_Sheet = m_Wrkbk.createSheet("UnitForecast", caption.toString(), m_Fields);

         //
         // We need a warehouse to get the in_catalog flag.  If it's null (for both warehouses), then
         // default to Portland for now.  Will need to change
         m_ItemSales.setString(1, m_WhseName != null ? m_WhseName : "PORTLAND");
         itemSales = m_ItemSales.executeQuery();

         while ( itemSales.next() && m_Status == RptServer.RUNNING ) {
            col = 0;

            item = itemSales.getString("item_id");
            upc = itemSales.getString("upc_code");

            nbc = itemSales.getString("nbc");
            nbc = (nbc == null? "": nbc);

            setCurAction("processing item: " + item);

            qtyShipped = itemSales.getInt("tot_promo_qty_shipped");
            sold = itemSales.getDouble("tot_promo_dollars_sold");
            cost = itemSales.getDouble("tot_promo_dollars_cost");
            margin = sold - cost;

            primaryVendorId = itemSales.getString("primary_vendor");

            if ( sold > 0 )
               marginPct = margin/sold;
            else
               marginPct = 0.0;

            //
            //09/23/2011. Naresh
            //MADE IN USA items
            usaItem = itemSales.getString("usa_item");
            usaItem = (usaItem == null? "": "USA");

            m_Sheet.addRow();

            m_Sheet.setField(col++, itemSales.getString("name"), false);
            m_Sheet.setField(col++, itemSales.getString("vendor_id"), false);
            m_Sheet.setField(col++, item, false);
            m_Sheet.setField(col++, usaItem, false);
            m_Sheet.setField(col++, itemSales.getString("vendor_item_num"), false);
            m_Sheet.setField(col++, itemSales.getString("stock_pack") + nbc, false);
            m_Sheet.setField(col++, itemSales.getString("ship_unit"), false);
            m_Sheet.setField(col++, itemSales.getString("retail_unit"), false);
            m_Sheet.setField(col++, itemSales.getString("retail_pack"), false);
            m_Sheet.setField(col++, (upc != null ? upc : ""), false);
            m_Sheet.setField(col++, itemSales.getString("description"), false);
            m_Sheet.setField(col++, itemSales.getDouble("buy"), Format.Currency2d);
            m_Sheet.setField(col++, itemSales.getDouble("sell"), Format.Currency2d);
            m_Sheet.setField(col++, itemSales.getDouble("retail_a"), Format.Currency2d);
            m_Sheet.setField(col++, itemSales.getDouble("retail_b"), Format.Currency2d);
            m_Sheet.setField(col++, itemSales.getDouble("retail_c"), Format.Currency2d);
            m_Sheet.setField(col++, itemSales.getDouble("retail_d"), Format.Currency2d);

            m_Sheet.setField(col++, itemSales.getInt("tot_qty_shipped"));
            m_Sheet.setField(col++, itemSales.getDouble("tot_dollars_sold"), Format.Currency2d);

            m_Sheet.setField(col++, itemSales.getInt("tot_replen_qty_shipped"));
            m_Sheet.setField(col++, itemSales.getDouble("tot_replen_dollars_sold"), Format.Currency2d);

            m_Sheet.setField(col++, itemSales.getInt("tot_bestpr_qty_shipped"));
            m_Sheet.setField(col++, itemSales.getDouble("tot_bestpr_dollars_sold"), Format.Currency2d);

            m_Sheet.setField(col++, qtyShipped);
            m_Sheet.setField(col++, sold, Format.Currency2d);

            m_Sheet.setField(col++, itemSales.getInt("events_qty_shipped"));
            m_Sheet.setField(col++, itemSales.getInt("flyers_qty_shipped"));
            m_Sheet.setField(col++, itemSales.getInt("pools_qty_shipped"));
            m_Sheet.setField(col++, itemSales.getInt("other_qty_shipped"));

            m_Sheet.setField(col++, itemSales.getDouble("events_dollars_sold"), Format.Currency2d);
            m_Sheet.setField(col++, itemSales.getDouble("flyers_dollars_sold"), Format.Currency2d);
            m_Sheet.setField(col++, itemSales.getDouble("pools_dollars_sold"), Format.Currency2d);
            m_Sheet.setField(col++, itemSales.getDouble("other_dollars_sold"), Format.Currency2d);

            //
            // Add the packet columns if the user selected any packets
            // Each packet has 4 columns: 2 columns for promo units and sales, and 2 for bestprice units and sales
            for ( String pkt : m_PacketArray ) {
               m_Sheet.setField(col++, itemSales.getInt("pkt_" + pkt + "_units"));
               m_Sheet.setField(col++, itemSales.getDouble("pkt_" + pkt + "_dollars"), Format.Currency2d);
               m_Sheet.setField(col++, itemSales.getInt("bestpr_" + pkt + "_units"));
               m_Sheet.setField(col++, itemSales.getDouble("bestpr_" + pkt + "_dollars"), Format.Currency2d);
            }

            m_Sheet.setField(col++, itemSales.getString("sen_code_id"), false);
            m_Sheet.setField(col++, marginPct, Format.Percent);
            m_Sheet.setField(col++, margin, Format.Currency2d);
            m_Sheet.setField(col++, getAvailQty(item));
            m_Sheet.setField(col++, itemSales.getString("flc_id"), false);
            m_Sheet.setField(col++, nbc, false);
            m_Sheet.setField(col++, itemSales.getInt("in_catalog") == 1 ? "yes" : "no", false);
            m_Sheet.setField(col++, itemSales.getString("velocity"), false);

            // Add customer xref data if flag is true
            if ( m_CustXRef ){
               for(int i=0; i<m_CustXRefNames.length; i++){
                  m_Sheet.setField(	col++,
                        m_CustXRefData.getCustSku(item, m_CustXRefIds[i]),
                        false	);
               }
            }

            // set Item Type field (primary or secondary)
            m_Sheet.setField(	col++,
                  primaryVendorId.equals(itemSales.getString("vendor_id")) ? "Primary" : "Secondary",
                        false	);

            upc = null;
         }

         m_Wrkbk.write(outFile);
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         RptServer.log.error("[UnitForecast]", ex);
      }

      finally {
         closeRSet(itemSales);

         try {
            outFile.close();
         }

         catch( Exception e ) {
            RptServer.log.error("[UnitForecast]", e);
         }

         outFile = null;
      }

      return result;
   }

   /**
    * Builds the SQL for the item sales prepared statement
    *
    * @return String - A complete SQL statement.
    */
   private String buildItemSalesSql()
   {
      boolean condition = false;
      StringBuffer sql = new StringBuffer();

      sql.append("select ");
      sql.append("   vendor.vendor_id, vendor.name, item_entity_attr.item_id, vendor_item_num, item_entity_attr.description, ");
      sql.append("   buy, sell, item_entity_attr.vendor_id as primary_vendor, ");
      sql.append("   retail_a, retail_b, retail_c, retail_d, sen_code_id, ");
      sql.append("   ejd_item.flc_id, ship_unit.unit ship_unit, retail_unit.unit retail_unit, retail_pack, stock_pack, upc.upc_code, ");
      sql.append("   decode(broken_case.description, 'ALLOW BROKEN CASES', '', 'N') nbc, ejd_item_warehouse.in_catalog, velocity, ia.item_id as usa_item, ");
      sql.append("   nvl(invd.tot_qty_shipped, 0) tot_qty_shipped, ");
      sql.append("   nvl(invd.tot_dollars_sold, 0) tot_dollars_sold, ");
      sql.append("   nvl(invd_replen.tot_replen_qty_shipped, 0) tot_replen_qty_shipped, ");
      sql.append("   nvl(invd_replen.tot_replen_dollars_sold, 0) tot_replen_dollars_sold, ");
      sql.append("   nvl(invd_bestprice.tot_bestpr_qty_shipped, 0) tot_bestpr_qty_shipped, ");
      sql.append("   nvl(invd_bestprice.tot_bestpr_dollars_sold, 0) tot_bestpr_dollars_sold, ");
      sql.append("   nvl(invd_promo.tot_promo_qty_shipped, 0) tot_promo_qty_shipped, ");
      sql.append("   nvl(invd_promo.tot_promo_dollars_sold, 0) tot_promo_dollars_sold, ");
      sql.append("   nvl(invd_promo.tot_promo_dollars_cost, 0) tot_promo_dollars_cost, ");
      sql.append("   nvl(events_qty_shipped, 0) events_qty_shipped, ");
      sql.append("   nvl(pools_qty_shipped, 0) pools_qty_shipped, ");
      sql.append("   nvl(flyers_qty_shipped, 0) flyers_qty_shipped, ");
      sql.append("   nvl(other_qty_shipped, 0) other_qty_shipped, ");
      sql.append("   nvl(events_dollars_sold, 0) events_dollars_sold, ");
      sql.append("   nvl(pools_dollars_sold, 0) pools_dollars_sold, ");
      sql.append("   nvl(flyers_dollars_sold, 0) flyers_dollars_sold, ");
      sql.append("   nvl(other_dollars_sold, 0) other_dollars_sold ");

      //
      // Now select the packet columns, if any packets were chosen.
      // For each packet select the promo sales excluding best price, and then the bestprice sales in separate columns.
      // Note the leading comma in the 1st "sql.append" below.  This is required.
      for ( String pkt : m_PacketArray ) {
         sql.append(",nvl(invd_pkt_" + pkt + ".dollars_sold, 0) pkt_" + pkt + "_dollars, ");
         sql.append("nvl(invd_pkt_" + pkt + ".qty_shipped, 0) pkt_" + pkt + "_units, ");
         sql.append("nvl(invd_bestpr_" + pkt + ".dollars_sold, 0) bestpr_" + pkt + "_dollars, ");
         sql.append("nvl(invd_bestpr_" + pkt + ".qty_shipped, 0) bestpr_" + pkt + "_units ");
      }

      sql.append("from item_entity_attr ");
      sql.append("join warehouse on warehouse.name = ? ");
      sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
      sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id and ejd_item_warehouse.warehouse_id = warehouse.warehouse_id ");
      
      //
      // Filter out discontinued items
      if ( !m_ShowDiscItems )
         sql.append("join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id and (disposition <> 'NOBUY-NOSELL' and disposition <> 'DELETE' ) ");

      sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id and ejd_item_price.warehouse_id = warehouse.warehouse_id ");
      sql.append("join item_velocity on ejd_item_warehouse.velocity_id = item_velocity.velocity_id ");
      sql.append("join vendor_item_cross on item_entity_attr.item_id = vendor_item_cross.item_id ");
      sql.append("join vendor on vendor_item_cross.vendor_id = vendor.vendor_id ");
      sql.append("join ship_unit on item_entity_attr.ship_unit_id = ship_unit.unit_id ");
      sql.append("join retail_unit on item_entity_attr.ret_unit_id = retail_unit.unit_id ");
      sql.append("join broken_case on ejd_item.broken_case_id = broken_case.broken_case_id ");
      sql.append("left outer join (select item_id, upc_code from item_upc where primary_upc = 1) upc on item_entity_attr.item_id = upc.item_id ");
      sql.append("join item_type on item_entity_attr.item_type_id = item_type.item_type_id and item_type.itemtype = 'STOCK' ");

      if ( m_NrhaId != null && m_NrhaId.length() > 0 ) {  // if filtering by nrha need to join with appropriate tables to get there
         sql.append("join flc on ejd_item.flc_id = flc.flc_id ");
         sql.append("join mdc on flc.mdc_id = mdc.mdc_id ");
         sql.append("join nrha on mdc.nrha_id = nrha.nrha_id ");
      }

      //
      // Get the total sales numbers
      sql.append("left outer join ( ");
      sql.append("   select item_nbr, vendor_nbr, sum(qty_shipped) tot_qty_shipped, sum(ext_sell) tot_dollars_sold ");
      sql.append("   from  inv_dtl ");
      sql.append("   where sale_type = 'WAREHOUSE' and ");

      if ( m_WhseName != null ) 
         sql.append("warehouse = '" + m_WhseName + "' and ");

      if ( m_PeriodType == PT_ROLLING )
         sql.append("(invoice_date > (now - interval '1 year')) and (invoice_date <= now) ");
      else {
         sql.append(String.format("invoice_date >= to_date('%s', 'mm/dd/yyyy') ", m_BegDate));
         sql.append(String.format("and invoice_date <= to_date('%s', 'mm/dd/yyyy') ", m_EndDate));
      }

      sql.append("   group by item_nbr, vendor_nbr ");  // group by item# and vendor#, as want to see sales by vendor
      sql.append(") invd on item_entity_attr.item_id = invd.item_nbr and vendor.vendor_id = invd.vendor_nbr ");

      //
      // Get the replenishment total sales numbers.
      // This excludes regular promotional and best-price redirected promo and qty buy sales.
      sql.append("left outer join ( ");
      sql.append("   select item_nbr, vendor_nbr, sum(qty_shipped) tot_replen_qty_shipped, sum(ext_sell) tot_replen_dollars_sold ");
      sql.append("   from  inv_dtl ");
      sql.append("   where ");
      sql.append("      sale_type = 'WAREHOUSE' and ");
      sql.append("      promo_nbr is null and ");
      sql.append("      not exists (select * from sa.invdtl_id_bestpr iib where iib.inv_dtl_id = inv_dtl.inv_dtl_id) and ");
      
      if ( m_WhseName != null ) {
         sql.append("warehouse = '" + m_WhseName + "' and ");
      }
      
      if ( m_PeriodType == PT_ROLLING )
         sql.append("(invoice_date > (now - interval '1 year')) and (invoice_date <= now) ");
      else {
         sql.append(String.format("invoice_date >= to_date('%s', 'mm/dd/yyyy') ", m_BegDate));
         sql.append(String.format("and invoice_date <= to_date('%s', 'mm/dd/yyyy') ", m_EndDate));
      }
      sql.append("   group by item_nbr, vendor_nbr ");  // group by item# and vendor#, as want to see sales by vendor
      sql.append(") invd_replen on item_entity_attr.item_id = invd_replen.item_nbr and vendor.vendor_id = invd_replen.vendor_nbr ");

      //
      // Get all best-price related sales numbers.  This includes both promo and qty buy best price sales.
      sql.append("left outer join ( ");
      sql.append("   select ");
      sql.append("      item_nbr, vendor_nbr, sum(qty_shipped) tot_bestpr_qty_shipped, sum(ext_sell) tot_bestpr_dollars_sold ");
      sql.append("   from inv_dtl ");
      sql.append("   join sa.invdtl_id_bestpr iib on iib.inv_dtl_id = inv_dtl.inv_dtl_id ");
      sql.append("   where ");
      sql.append("      sale_type = 'WAREHOUSE' and ");
      
      if ( m_WhseName != null )
         sql.append("warehouse = '" + m_WhseName + "' and ");
      
      if ( m_PeriodType == PT_ROLLING )
         sql.append("(invoice_date > (now - interval '1 year')) and (invoice_date <= now) ");
      else {
         sql.append(String.format("invoice_date >= to_date('%s', 'mm/dd/yyyy') ", m_BegDate));
         sql.append(String.format("and invoice_date <= to_date('%s', 'mm/dd/yyyy') ", m_EndDate));
      }
      sql.append("   group by item_nbr, vendor_nbr ");  // group by item# and vendor#, as want to see sales by vendor
      sql.append(") invd_bestprice on item_entity_attr.item_id = invd_bestprice.item_nbr and vendor.vendor_id = invd_bestprice.vendor_nbr ");

      //
      // Get the promotional sales numbers broken down by demand type.
      // Make sure best price promo sales are excluded as that artificially inflates the promo demand sales numbers.
      sql.append("   left outer join ( ");
      sql.append("      select ");
      sql.append("         item_nbr, ");
      sql.append("         vendor_nbr, ");
      sql.append("         sum(qty_shipped) tot_promo_qty_shipped, ");
      sql.append("         sum(ext_sell) tot_promo_dollars_sold, ");
      sql.append("         sum(ext_cost) tot_promo_dollars_cost, ");
      sql.append("         sum(decode(demand_type.description, 'SHOWS', qty_shipped, 0)) events_qty_shipped, ");
      sql.append("         sum(decode(demand_type.description, 'POOLS/STOCKING DEALER', qty_shipped, 0)) pools_qty_shipped, ");
      sql.append("         sum(decode(demand_type.description, 'FLYERS/CIRCULARS', qty_shipped, 0)) flyers_qty_shipped, ");
      sql.append("         sum(decode(demand_type.description, 'OTHER PROMOTIONS', qty_shipped, 0)) other_qty_shipped, ");
      sql.append("         sum(decode(demand_type.description, 'SHOWS', ext_sell, 0)) events_dollars_sold, ");
      sql.append("         sum(decode(demand_type.description, 'POOLS/STOCKING DEALER', ext_sell, 0)) pools_dollars_sold, ");
      sql.append("         sum(decode(demand_type.description, 'FLYERS/CIRCULARS', ext_sell, 0)) flyers_dollars_sold, ");
      sql.append("         sum(decode(demand_type.description, 'OTHER PROMOTIONS', ext_sell, 0)) other_dollars_sold");
      sql.append("      from inv_dtl ");
      sql.append("      join promotion on inv_dtl.promo_nbr = promotion.promo_id ");
      sql.append("      join demand_type on promotion.demand_type_id = demand_type.demand_type_id ");
      sql.append("      where ");
      sql.append("         sale_type = 'WAREHOUSE' and ");
      sql.append("         not exists (select * from sa.invdtl_id_bestpr iib where iib.inv_dtl_id = inv_dtl.inv_dtl_id) and ");
      
      if ( m_WhseName != null ) {
         sql.append("inv_dtl.warehouse = '" + m_WhseName + "' and ");
      }
      
      if ( m_PeriodType == PT_ROLLING )
         sql.append("(invoice_date > (now - interval '1 year')) and (invoice_date <= now) ");
      else {
         sql.append(String.format("invoice_date >= to_date('%s', 'mm/dd/yyyy') ", m_BegDate));
         sql.append(String.format("and invoice_date <= to_date('%s', 'mm/dd/yyyy') ", m_EndDate));
      }
      
      sql.append("   group by item_nbr, vendor_nbr ");  // group by item# and vendor#, as want to see sales by vendor
      sql.append(") invd_promo on item_entity_attr.item_id = invd_promo.item_nbr and vendor.vendor_id = invd_promo.vendor_nbr ");

      //
      // Get the promotional and bestprice sales FOR EACH PACKET REQUESTED
      for ( String pkt : m_PacketArray ) {
         // Promotional sales (excluding bestprice) for current packet
         sql.append("left outer join ( ");
         sql.append("   select item_nbr, vendor_nbr, sum(qty_shipped) qty_shipped, sum(ext_sell) dollars_sold ");
         sql.append("   from inv_dtl ");
         sql.append("   join promotion on inv_dtl.promo_nbr = promotion.promo_id and promotion.packet_id = '" + pkt + "' ");
         sql.append("   where ");
         sql.append("      sale_type = 'WAREHOUSE' and ");
         sql.append("      not exists (select * from sa.invdtl_id_bestpr iib where iib.inv_dtl_id = inv_dtl.inv_dtl_id) and ");
         
         if ( m_WhseName != null ) {
            sql.append("inv_dtl.warehouse = '" + m_WhseName + "' and ");
         }
         
         if ( m_PeriodType == PT_ROLLING )
            sql.append("(invoice_date > (now - interval '1 year')) and (invoice_date <= now) ");
         else {
            sql.append(String.format("invoice_date >= to_date('%s', 'mm/dd/yyyy') ", m_BegDate));
            sql.append(String.format("and invoice_date <= to_date('%s', 'mm/dd/yyyy') ", m_EndDate));
         }
         
         sql.append("   group by item_nbr, vendor_nbr ");  // group by item# and vendor#, as want to see sales by vendor
         sql.append(") invd_pkt_" + pkt + " on item_entity_attr.item_id = invd_pkt_" + pkt + ".item_nbr and vendor.vendor_id = invd_pkt_" + pkt + ".vendor_nbr ");

         // Bestprice sales for current packet
         sql.append("   left outer join ( ");
         sql.append("      select ");
         sql.append("         item_nbr, ");
         sql.append("         vendor_nbr, ");
         sql.append("         sum(qty_shipped) qty_shipped, ");
         sql.append("         sum(ext_sell) dollars_sold ");
         sql.append("      from ");
         sql.append("         inv_dtl ");
         sql.append("      join promotion on inv_dtl.promo_nbr = promotion.promo_id and promotion.packet_id = '" + pkt + "' ");
         sql.append("      join sa.invdtl_id_bestpr iib on iib.inv_dtl_id = inv_dtl.inv_dtl_id ");
         sql.append("      where ");
         sql.append("         sale_type = 'WAREHOUSE' and ");
         
         if ( m_WhseName != null ) {
            sql.append("inv_dtl.warehouse = '" + m_WhseName + "' and ");
         }
         
         if ( m_PeriodType == PT_ROLLING )
            sql.append("(invoice_date > (now - interval '1 year')) and (invoice_date <= now) ");
         else {
            sql.append(String.format("invoice_date >= to_date('%s', 'mm/dd/yyyy') ", m_BegDate));
            sql.append(String.format("and invoice_date <= to_date('%s', 'mm/dd/yyyy') ", m_EndDate));
         }
         
         sql.append("      group by item_nbr, vendor_nbr ");  // group by item# and vendor#, as want to see sales by vendor
         sql.append("   ) invd_bestpr_" + pkt + " on item_entity_attr.item_id = invd_bestpr_" + pkt + ".item_nbr and vendor.vendor_id = invd_bestpr_" + pkt + ".vendor_nbr ");
      }

      //
      // 09/23/2011. Naresh
      // Identifies MADE IN USA items
      sql.append("left outer join item_attribute ia on item_entity_attr.item_id = ia.item_id and ");
      sql.append("ia.attribute_value_id in ( select attribute_value_id from attribute a , attribute_value av where a.attribute_id = av.attribute_id and av.value = 'MADE IN USA') ");
      sql.append("where ");

      if ( m_VndId != null && m_VndId.length() > 0 ) {
         sql.append("vendor.vendor_id in (");
         sql.append(m_VndId);
         sql.append(") ");

         condition = true;
      }

      if ( m_NrhaId != null && m_NrhaId.length() > 0 ) {
         sql.append((condition? " and ": "") + "nrha.nrha_id in (");
         sql.append(m_NrhaId);
         sql.append(") ");

         condition = true;
      }

      if ( m_FlcId != null && m_FlcId.length() > 0 ) {
         sql.append((condition? " and ": "") + "ejd_item.flc_id in (");
         sql.append(m_FlcId);
         sql.append(") ");

         condition = true;
      }

      if ( m_ItemId != null && m_ItemId.length() > 0  ) {
         sql.append((condition ? " and ": "") + "item_entity_attr.item_id in (");
         sql.append(m_ItemId);
         sql.append(") ");
      }

      sql.append("order by vendor.name, item_id");
      
      return sql.toString();
   }

   /**
    * Build the SQL statement for the customer cross reference data.
    *
    * @return	a SQL statement
    */
   private String buildCustXRefSql(){
      StringBuffer sql = new StringBuffer();
      Boolean condition = false;

      sql.append("select item_entity_attr.item_id, item_ea_cross.customer_sku, customer.customer_id, customer.name ");
      sql.append("from customer ");
      sql.append("join item_ea_cross on customer.customer_id = item_ea_cross.customer_id ");
      sql.append("join item_entity_attr on item_ea_cross.item_ea_id = item_entity_attr.item_ea_id ");

      // Filter by Vendor Id
      if( m_VndId != null && m_VndId.length() > 0 )
         sql.append("join vendor_item_ea_cross on vendor_item_ea_cross.item_ea_id = item_entity_attr.item_ea_id ");

      // Filtering by NHRA Id or FLC Id
      if ( (m_NrhaId != null && m_NrhaId.length() > 0) || (m_FlcId != null && m_FlcId.length() > 0)) {
         sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
         sql.append("join flc on ejd_item.flc_id = flc.flc_id ");
      }

      // Filtering by NHRA Id
      if ( (m_NrhaId != null && m_NrhaId.length() > 0) ) {
         sql.append("join mdc on flc.mdc_id = mdc.mdc_id ");
         sql.append("join nrha on mdc.nrha_id = nrha.nrha_id ");
      }

      // Construct WHERE clause
      sql.append("where ");

      // Filter by Vendor Id
      if ( m_VndId != null && m_VndId.length() > 0 ) {
         sql.append("vendor_item_ea_cross.vendor_id in (");
         sql.append(m_VndId);
         sql.append(") ");

         condition = true;
      }

      // Filter by NRHA Id
      if ( m_NrhaId != null && m_NrhaId.length() > 0 ) {
         sql.append((condition? " and ": "") + "nrha.nrha_id in (");
         sql.append(m_NrhaId);
         sql.append(") ");

         condition = true;
      }

      // Filter by FLC Id
      if ( m_FlcId != null && m_FlcId.length() > 0 ) {
         sql.append((condition? " and ": "") + "ejd_item.flc_id in (");
         sql.append(m_FlcId);
         sql.append(") ");

         condition = true;
      }

      // Filter by Item Id
      if ( m_ItemId != null && m_ItemId.length() > 0  ) {
         sql.append((condition? " and ": "") + "item_entity_attr.item_id in (");
         sql.append(m_ItemId);
         sql.append(") ");
      }

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

         if(m_CustXRef){
            initCustXRefData();
         }

         initFields();

         created = buildWorkbook();
      }

      catch ( Exception ex ) {
         RptServer.log.fatal("[UnitForecast]", ex);
      }

      finally {
         closeStatements();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Runs a query and prepares the cust xref data for adding
    * to the report
    * @throws SQLException
    */
   private void initCustXRefData() throws SQLException {
      ResultSet rs;

      rs = m_CrossRef.executeQuery();

      m_CustXRefData = new CustXRefData(rs);

      m_CustXRefNames = m_CustXRefData.getCustNames();
      m_CustXRefIds = m_CustXRefData.getCustIds();
   }

   /**
    * Initializes the fields which will be used
    * as column headers and by the worksheet
    */
   private void initFields() {
      ArrayList<Column> fields = new ArrayList<Column>();

      fields.add(new Column("Vendor Name", NAME_COL_WIDTH));
      fields.add(new Column("Vendor ID", DEFAULT_COL_WIDTH));
      fields.add(new Column("Item #", DEFAULT_COL_WIDTH));
      fields.add(new Column("USA", QTY_COL_WIDTH));
      fields.add(new Column("Mfgr. Part No.", NAME_COL_WIDTH));
      fields.add(new Column("Stock Pack", QTY_COL_WIDTH));
      fields.add(new Column("Ship Unit", QTY_COL_WIDTH));
      fields.add(new Column("Retail Unit", QTY_COL_WIDTH));
      fields.add(new Column("Dealer Pack", QTY_COL_WIDTH));
      fields.add(new Column("UPC-Primary", 14));
      fields.add(new Column("Item Description", DESCRIP_COL_WIDTH));
      fields.add(new Column("Emery Cost", CURRENCY_COL_WIDTH));
      fields.add(new Column("Base Cost", CURRENCY_COL_WIDTH));
      fields.add(new Column("A Mkt Retail", CURRENCY_COL_WIDTH));
      fields.add(new Column("B Mkt Retail", CURRENCY_COL_WIDTH));
      fields.add(new Column("C Mkt Retail", CURRENCY_COL_WIDTH));
      fields.add(new Column("D Mkt Retail", CURRENCY_COL_WIDTH));
      fields.add(new Column("Total Units Sold", QTY_COL_WIDTH));
      fields.add(new Column("Total Dollars Sold", CURRENCY_COL_WIDTH));
      fields.add(new Column("Replenishment Units Sold", 15));
      fields.add(new Column("Replenishment Dollars Sold", CURRENCY_COL_WIDTH));
      fields.add(new Column("BestPrice Units Sold", QTY_COL_WIDTH));
      fields.add(new Column("BestPrice Dollars Sold", CURRENCY_COL_WIDTH));
      fields.add(new Column("Promo Units Sold", QTY_COL_WIDTH));
      fields.add(new Column("Promo Dollars Sold", CURRENCY_COL_WIDTH));
      fields.add(new Column("Events Units Sold", QTY_COL_WIDTH));
      fields.add(new Column("Flyers Units Sold", QTY_COL_WIDTH));
      fields.add(new Column("Pools Units Sold", QTY_COL_WIDTH));
      fields.add(new Column("Other Units Sold", QTY_COL_WIDTH));
      fields.add(new Column("Events Dollars Sold", CURRENCY_COL_WIDTH));
      fields.add(new Column("Flyers Dollars Sold", CURRENCY_COL_WIDTH));
      fields.add(new Column("Pools Dollars Sold", CURRENCY_COL_WIDTH));
      fields.add(new Column("Other Dollars Sold", CURRENCY_COL_WIDTH));

      //
      // Add the packet column titles, if the user selected any packets.
      // 4 columns per packet: 2 for promo units and dollars and 2 for bestprice units and dollars.
      for ( String pkt : m_PacketArray ) {
         fields.add(new Column("Pkt " + pkt + " Units", DEFAULT_COL_WIDTH));
         fields.add(new Column("Pkt " + pkt + " Dollars", CURRENCY_COL_WIDTH));
         fields.add(new Column("BestPr " + pkt + " Units", DEFAULT_COL_WIDTH));
         fields.add(new Column("BestPr " + pkt + " Dollars", CURRENCY_COL_WIDTH));
      }

      fields.add(new Column("Sensitivity Code", DEFAULT_COL_WIDTH));
      fields.add(new Column("Promo Margin%", DEFAULT_COL_WIDTH));
      fields.add(new Column("Promo Margin$", DEFAULT_COL_WIDTH));
      fields.add(new Column("Units On Hand", DEFAULT_COL_WIDTH));
      fields.add(new Column("FLC", DEFAULT_COL_WIDTH));
      fields.add(new Column("NBC", DEFAULT_COL_WIDTH));
      fields.add(new Column("Catalog", DEFAULT_COL_WIDTH));
      fields.add(new Column("Velocity", DEFAULT_COL_WIDTH));

      if(m_CustXRef){
         for(String s : m_CustXRefNames){
            fields.add(new Column(s, NAME_COL_WIDTH));
         }
      }

      // Add Item Type (Primary/Secondary for vendor) to the end of the report
      // per user request
      fields.add(new Column("Item Type", DEFAULT_COL_WIDTH));

      m_Fields = fields.toArray(new Column[0]);
   }

   /**
    * Gets the available quantity from fascor.  This is done through a web service
    * that directly connects to fascor.  It returns the available qty after the
    * appropriate adjustments have been made.
    *
    * @param item String - The item to check.
    *
    * @return The quantity available of the item in fascor
    * @throws Exception
    */
   private int getAvailQty(String item) throws Exception
   {
	   int qty = 0;
	   ResultSet rset = null;
	      
	   if ( item != null && item.length() == 7 ) {
		   try {
			   m_ItemDCQty.setString(1, item);
			   m_ItemDCQty.setString(2, "PORTLAND");
			   m_ItemDCQty.setString(3, "PITTSTON");
			   rset = m_ItemDCQty.executeQuery();

			   if ( rset.next() )
				   qty = rset.getInt("available_qty");
		   }
		   finally {
			   closeRSet(rset);
			   rset = null;
		   }
	   }

      return qty < 0 ? 0 : qty;
   }
  
   /**
    * Prepares the sql queries for execution.
    */
   private void prepareStatements() throws Exception
   {
      if ( m_EdbConn != null ) {
         m_ItemSales = m_EdbConn.prepareStatement(buildItemSalesSql());

         if(m_CustXRef){
            m_CrossRef = m_EdbConn.prepareStatement(buildCustXRefSql());
         }
         
         m_ItemDCQty = m_EdbConn.prepareStatement("select sum(avail_qty) as available_qty from ejd_item_warehouse  " +
          		"join item_entity_attr on item_entity_attr.ejd_item_id = ejd_item_warehouse.ejd_item_id " +
                 "where item_entity_attr.item_id = ? "+
          		"and warehouse_id in (select warehouse_id from warehouse where name = ? or name = ? ) "); 
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

      for ( Param param : params ) {
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

         if ( param.name.equals("packet") )
            m_Packets = param.value;

         //
         // Split up the packet elements in an array.
         // Just do it once here as this is accessed all over the place.
         if ( m_Packets != null && m_Packets.length() > 0 )
            m_PacketArray = m_Packets.split(",");

         if ( param.name.equals("warehousename") )
            m_WhseName = param.value;

         if ( param.name.equals("warehousefac") )
            m_WhseFac = param.value;

         if (param.name.equals("showdisc") )
            m_ShowDiscItems = Boolean.parseBoolean(param.value);
      }

      //
      // Default this to portland if nothing was sent in.
      if ( m_WhseFac == null )
         m_WhseFac = "01";

      //
      // Build the file name.
      fname.append(tm);
      fname.append("_");
      fname.append(m_RptProc.getUid());


      String tmp = Long.toString(System.currentTimeMillis());
      
      //if ( m_VndId != null && m_VndId.length() > 0 )
      //   fname.append(String.format("_%s", m_VndId));

      fname.append(tmp.substring(tmp.length()-5, tmp.length()));
      
      fname.append("_uf.xls");
      m_FileNames.add(fname.toString());
   }
}
