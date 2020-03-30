/**
 * File: RMSExport.java
 * Description: No description on original report
 *    Rewrite to allow running in the new report server.
 *    The original author was Jacob Heric.
 *
 * @author Jacob Heric
 * @author Jeffrey Fisher
 *
 * Create Date: 05/19/2005
 * Last Update: $Id: RMSExport.java,v 1.17 2014/03/17 18:37:50 epearson Exp $
 *
 * History
 *    $Log: RMSExport.java,v $
 *    Revision 1.17  2014/03/17 18:37:50  epearson
 *    updated characters to UTF-8
 *
 *    Revision 1.16  2013/01/02 17:34:03  prichter
 *    Pittston expansion - modified the query to handle both warehouses.
 *
 *    Revision 1.15  2012/07/19 19:37:11  jfisher
 *    in_catalog at the warehouse level changes
 *
 *    Revision 1.14  2008/10/29 20:58:18  jfisher
 *    Fixed potential null warnings.
 *
 *    Revision 1.13  2008/07/15 01:55:58  pdavidson
 *    Adjusted where velocity originates (R12 qty shipped).
 *    Added column to display all warehouses of current item.
 *
 *    03/25/2005 - Added log4j logging. jcf
 *
 *    03/11/2005 - Changed the setup_date format to yyyy/MM/dd per CR#587 - jcf
 *
 *    05/03/2004 - Removed the usage of the m_DistList member variable.  This variable gets cleaned up before it can be
 *       used in the email webservice. - jcf
 *
 *    04/07/2004 - Applied Email class changes. - jcf
 *
 *    12/22/2003 - Modified the email distribution list processing to handle the new xml request format. Also removed
 *       some unused imports, unused variables and a bogus exception. - jcf
 *
 *    03/19/2002 - Changed the way the connection to Oracle is established.  The program now uses the
 *       getOracleConn() method to retrieve a new connection object.  The connection pool is no longer
 *       used. Also added the cleanup method.- jcf
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;


public class RMSExport extends Report
{  
   private PreparedStatement 
      m_rmsItem,
      m_rmsAceItem;

   /**
    * default constructor
    */
   public RMSExport()
   {
      super();
      m_FileNames.add("rmsItem.dat");
   }

   /**
    * Executes the queries and builds the output file
    *
    * @return boolean true if the file was built without errors, false if not.
    */
   private boolean buildOutputFile()
   {
      StringBuffer Line = new StringBuffer();

      FileOutputStream OutFile = null;
      ResultSet 
         rmsItemData = null,
         rmsAceItemData = null;

      String VendorId = null;
      String VendorName = null;
      String ItemId = null;
      String VendorItemId = null;
      String ItemDesc = null;
      String Buy = null;
      String Sell = null;
      String RetailA = null;
      String RetailB = null;
      String RetailC = null;
      String RetailD = null;
      String Unit = null;
      String RetailPack = null;
      String StockPack = null;
      String UPC = null;
      String Sensitivity = null;
      String Price_Method = null;
      String FlcId = null;
      String InCatalog = null;
      String Disp = null;
      String NRHA = null;
      String Velocity = null;
      String NBC = null;
      String BuyerID = null;
      String RMSID = null;
      String UnitSales = null;
      String SetupDate = null;
      String dcNames = null;
      boolean Result = false;

      try {
         setCurAction("creating/opening output file " + m_FilePath + m_FileNames.get(0));
         OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

         //
         // Build the Captions
         setCurAction("inserting header line into file");
         Line.append("Vendor#\tVendor\tItem#\tManf#\tDescription\tBuy\tSell\tRetailA\tRetailB\tRetailC\t" +
                     "RetailD\tUnit\tRetail Pack#\tStock Pack\tUPC\tSenCode\tPriceMethod\tFLC\tInCatalog\tDisp\tNRHA\t" +
                     "Velocity\tNBC\tBuyer\tRMS\tSales History\tSetup Date\tWarehouses\r\n");

         if ( m_Status == RptServer.RUNNING ) {
            setCurAction("executing report queries");
            rmsItemData = m_rmsItem.executeQuery();
            rmsAceItemData = m_rmsAceItem.executeQuery();
            setCurAction("report queries executed ");
         }
         
         for (ResultSet resultSet : new ResultSet[]{rmsItemData, rmsAceItemData}) {
            if ( resultSet != null ) {
               while ( resultSet.next() && m_Status == RptServer.RUNNING ) {
                  VendorId = "";
                  VendorName = "";
                  ItemId = "";
                  VendorItemId = "";
                  ItemDesc = "";
                  Buy = "";
                  Sell = "";
                  RetailA = "";
                  RetailB = "";
                  RetailC = "";
                  RetailD = "";
                  Unit = "";
                  RetailPack = "";
                  StockPack = "";
                  Sensitivity = "";
                  Price_Method = "";
                  UPC = "";
                  FlcId = "";
                  InCatalog = "";
                  RetailD = "";
                  Disp = "";
                  NRHA = "";
                  Velocity = "";
                  NBC = "";
                  BuyerID = "";
                  RMSID = "";
                  UnitSales = "";
                  SetupDate = "";
                  dcNames = "";
   
                  VendorId = resultSet.getString("vendor_id");
                  VendorName = resultSet.getString("name");
                  ItemId = resultSet.getString("item_id");
                  setCurAction("inserting line for item " + ItemId);
                  VendorItemId = resultSet.getString("vendor_item_num");
                  ItemDesc = resultSet.getString("description");
                  Buy = resultSet.getString("buy");
                  Sell = resultSet.getString("sell");
                  RetailA = resultSet.getString("retail_a");
                  RetailB = resultSet.getString("retail_b");
                  RetailC = resultSet.getString("retail_c");
   
                  RetailD = resultSet.getString("retail_d");
                  if ( RetailD == null )
                     RetailD = "";
   
                  Unit = resultSet.getString("unit");
                  RetailPack = resultSet.getString("retail_pack");
                  StockPack = resultSet.getString("stock_pack");
   
                  UPC = resultSet.getString("upc_code");
                  if ( UPC == null)
                     UPC = "";
   
                  Sensitivity = resultSet.getString("sen_code_id");
                  if ( Sensitivity == null )
                     Sensitivity = "";
   
                  Price_Method = resultSet.getString("price_method");
                  FlcId = resultSet.getString("flc_id");
                  InCatalog = resultSet.getString("in_catalog");
                  Disp = resultSet.getString("disposition");
                  NRHA = resultSet.getString("nrha");
   
                  Velocity = resultSet.getString("velocity");
                  if ( Velocity == null )
                     Velocity = "";
   
                  NBC = resultSet.getString("nbc");
                  if ( NBC == null )
                     NBC = "";
   
                  BuyerID = resultSet.getString("dept_num");
   
                  RMSID = resultSet.getString("rms_id");
                  if( RMSID == null )
                     RMSID = "";
   
                  UnitSales = resultSet.getString("Unit_Sales");
                  if ( UnitSales == null )
                     UnitSales = "";
   
                  SetupDate = resultSet.getString("setup_date");
                  if ( SetupDate == null )
                     SetupDate = "";
   
                  dcNames = resultSet.getString("dc_names");
                  if ( dcNames == null )
                     dcNames = "";
   
                  Line.append(VendorId + "\t");
                  Line.append(VendorName + "\t");
                  Line.append(ItemId + "\t");
                  Line.append(VendorItemId + "\t");
                  Line.append(ItemDesc + "\t");
                  Line.append(Buy + "\t");
                  Line.append(Sell + "\t");
                  Line.append(RetailA + "\t");
                  Line.append(RetailB + "\t");
                  Line.append(RetailC + "\t");
                  Line.append(RetailD + "\t");
                  Line.append(Unit + "\t");
                  Line.append(RetailPack + "\t");
                  Line.append(StockPack + "\t");
                  Line.append(UPC + "\t");
                  Line.append(Sensitivity + "\t");
                  Line.append(Price_Method + "\t");
                  Line.append(FlcId + "\t");
                  Line.append(InCatalog + "\t");
                  Line.append(Disp + "\t");
                  Line.append(NRHA + "\t");
                  Line.append(Velocity + "\t");
                  Line.append(NBC + "\t");
                  Line.append(BuyerID + "\t");
                  Line.append(RMSID + "\t");
                  Line.append(UnitSales + "\t");
                  Line.append(SetupDate + "\t");
                  Line.append(dcNames + "\t");
                  Line.append("\r\n");
   
                  OutFile.write(Line.toString().getBytes());
                  Line.delete(0, Line.length());
                  setCurAction("finished line for item " + ItemId);
               }
   
               setCurAction("closing recordset");
               DbUtils.closeDbConn(null, null, resultSet);
               resultSet = null;
               
               Result = true;
            }
         }
      }

      catch ( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }
      
      finally {
      	if ( OutFile != null ) {
      	   try {
      	   	OutFile.close();
      	   }
      	   catch ( Exception e ) {
      	   	log.error("exception", e);
      	   }

            OutFile = null;      	   
      	}
      }

      return Result;
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not garunteed to know when finalization occurs.
    */
   protected void cleanup()
   {
      closeStatements();
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      try {
         for (PreparedStatement statement : new PreparedStatement[]{m_rmsItem, m_rmsAceItem}) {
            if ( statement != null ) {
               statement.close();
               statement = null;
            }
         }
      }
      catch ( Exception ex ) {
         log.error(ex);
      }
   }

   /**
    * Create the report.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
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
         log.fatal("exception:", ex);
      }

      finally {
         cleanup();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Prepares the sql queries for execution.
    *
    * @return boolean true if the statements were prepared, false if not.
    */
   private boolean prepareStatements()
   {
      boolean isPrepared = false;
      

      if ( m_OraConn != null ) {
         try {
            m_rmsItem = prepareItemQuery();
            m_rmsAceItem = prepareAceItemQuery();
            isPrepared = true;
         }

         catch (Exception ex) {
            log.fatal("exception:", ex);
         }
      }

      return isPrepared;
   }

   private PreparedStatement prepareItemQuery() throws Exception {
      StringBuffer sql = new StringBuffer(1024);
      //
      // JF 7/11/2012 re-write of the original Paul query to add in the item_warehouse table for the
      // in_catalog flag and also to use the non oracle specific join syntax

      sql.append("select ");
      sql.append("vendor.vendor_id, vendor.name, item.item_id, vendor_item_num, ");
      sql.append("item.description, item_price.buy, item_price.sell, retail_a, retail_b, ");
      sql.append("retail_c, retail_d, ship_unit.unit, item.retail_pack, item.stock_pack, ");
      sql.append(" item_upc.upc_code, item_price.sen_code_id, item_price_method.price_method, item.flc_id, ");
      sql.append("decode(cat.in_catalog, 0, 'N', 'Y') in_catalog, item_disp.disposition, ");
      sql.append("flc_procs.GETNRHAID(item.flc_id) nrha, ");

      //
      // PD 7/11/08 - get item velocity, which in this case is the R12 qty shipped
      // The query goes back 13 months to get a full 12 months back, since monthlyitemsales
      // won't have data for the current month (until the month end proc runs).  So running
      // the query sometime in July 2008 will get total shipped from 06/07-06/08.
      sql.append("(select sum(units_shipped) ");
      sql.append("from sa.monthlyitemsales outermis ");
      sql.append("where item_nbr = item.item_id and mis_id >= ( ");
      sql.append("   select ");
      sql.append("      max(mis.mis_id) ");
      sql.append("   from ");
      sql.append("      sa.monthlyitemsales mis ");
      sql.append("   where ");
      sql.append("      mis.item_nbr = outermis.item_nbr and ");
      sql.append("      mis.sale_month = extract(month from add_months(sysdate, -13)) and ");
      sql.append("      mis.sale_year = extract(year from add_months(sysdate, -13))) ");
      sql.append(") as velocity, ");

      sql.append("decode(broken_case.description, 'ALLOW BROKEN CASES', null, 'N') NBC, ");
      sql.append("vendor_buyers.DEPT_NUM, rms_item.RMS_ID, ");
      sql.append("(select sum(units_shipped) from sa.monthlyitemsales where ");
      sql.append("item_nbr = item.item_id and months_between(last_day(sysdate), ");
      sql.append("to_date(year_month, 'yyyymm')) <= 12) Unit_Sales, to_char(setup_date, 'yyyy/MM/dd') setup_date, ");
      //
      // PD 7/10/08 -added what's below to get warehouse names
      // ** NOTE: using undocumented, unsupported wm_concat function **

      sql.append("(select  ");
      sql.append("    wmsys.wm_concat(' '||warehouse.name) "); 
      sql.append(" from  ");
      sql.append("    item itm, item_warehouse, warehouse "); 
      sql.append(" where  ");
      sql.append("    itm.item_id = item.item_id and "); 
      sql.append("    itm.item_id = item_warehouse.item_id and "); 
      sql.append("    item_warehouse.warehouse_id = warehouse.warehouse_id and ");
      sql.append("    item_warehouse.warehouse_id in (1, 2, 11) ");
      sql.append(" group by itm.item_id) as dc_names  ");
      sql.append("from  ");
      sql.append("   item  ");

      sql.append("join ( ");
      sql.append("   select item_id, sum(in_catalog) in_catalog ");
      sql.append("   from item_warehouse ");
      sql.append("   group by item_id ");
      sql.append(") cat on cat.item_id = item.item_id ");

      sql.append("join item_type on item_type.item_type_id = item.item_type_id and itemtype in ('STOCK', 'MISC') ");
      sql.append("join vendor on vendor.vendor_id = item.vendor_id ");
      sql.append("join vendor_item_cross on vendor_item_cross.item_id = item.item_id and vendor_item_cross.vendor_id = item.vendor_id ");
      sql.append("join vendor_buyers on vendor_buyers.vendor_id = item.vendor_id ");
      sql.append("join ship_unit on ship_unit.unit_id = item.ship_unit_id ");
      sql.append("join item_disp on item_disp.disp_id = item.disp_id ");
      sql.append("join item_velocity on item_velocity.velocity_id = item.velocity_id ");
      sql.append("join broken_case on broken_case.broken_case_id = item.broken_case_id ");
      sql.append("left outer join rms_item on rms_item.item_id = item.item_id ");
      sql.append("join item_price on item_price.item_id = item.item_id and  ");
      sql.append("   item_price.sell_date = ( ");
      sql.append("      select max(sell_date)  ");
      sql.append("      from item_price ");
      sql.append("      where sell_date <= trunc(sysdate) and item_id = item.item_id ");
      sql.append("   ) ");
      sql.append("join item_price_method on item_price_method.method_id = item_price.method_id ");
      sql.append("left outer join item_upc on item_upc.item_id = item.item_id and primary_upc = 1 ");
      sql.append("where ");
      sql.append("   rms_item.rms_id in ( ");
      sql.append("      select max(rms_id)  ");
      sql.append("      from rms_item ");
      sql.append("      where item.item_id = rms_item.item_id ");
      sql.append("   ) or rms_item.rms_id is null ");
      
      return m_OraConn.prepareStatement(sql.toString());
   }
   
   private PreparedStatement prepareAceItemQuery() throws Exception {
      StringBuffer sql = new StringBuffer(1024);
      
      sql.append("select ");
      sql.append("vendor.vendor_id, vendor.name, item.item_id, vendor_item_num, ");
      sql.append("item.description, ace_item_price.reg_cost buy, ace_item_price.sell sell, retail_a, retail_b, ");
      sql.append("retail_c, retail_d, ship_unit.unit, item.retail_pack, item.stock_pack, ");
      sql.append("item_upc.upc_code, null sen_code_id, null price_method, item.flc_id, ");
      sql.append("decode(item_warehouse.in_catalog, 0, 'N', 'Y') in_catalog, item_disp.disposition, "); 
      sql.append("flc_procs.GETNRHAID(item.flc_id) nrha, ");
      sql.append("(select sum(qty_shipped) from inv_dtl where item_nbr = item.item_id and invoice_date <= trunc(sysdate) - 365) as velocity, ");
      sql.append("decode(broken_case.description, 'ALLOW BROKEN CASES', null, 'N') NBC, ");
      sql.append("vendor_buyers.DEPT_NUM, rms_item.RMS_ID, ");
      sql.append("(select sum(units_shipped) from sa.monthlyitemsales where ");
      sql.append("item_nbr = item.item_id and months_between(last_day(sysdate), ");
      sql.append("to_date(year_month, 'yyyymm')) <= 12) Unit_Sales, to_char(setup_date, 'yyyy/MM/dd') setup_date, ");
      sql.append("warehouse.name as dc_names ");      
      sql.append("from "); 
      sql.append("   item ");
      sql.append("join item_type on item_type.item_type_id = item.item_type_id and itemtype = 'ACE' ");
      sql.append("join vendor on vendor.vendor_id = item.vendor_id ");
      sql.append("join vendor_item_cross on vendor_item_cross.item_id = item.item_id and vendor_item_cross.vendor_id = item.vendor_id ");
      sql.append("left join vendor_buyers on vendor_buyers.vendor_id = item.vendor_id ");
      sql.append("join ship_unit on ship_unit.unit_id = item.ship_unit_id ");
      sql.append("join item_warehouse on item_warehouse.item_id = item.item_id ");
      sql.append("join warehouse on warehouse.warehouse_id = item_warehouse.warehouse_id and warehouse.name = 'WILTON' ");
      sql.append("join item_disp on item_disp.disp_id = item_warehouse.disp_id and disposition = 'BUY-SELL' ");
      sql.append("join item_velocity on item_velocity.velocity_id = item.velocity_id ");
      sql.append("join broken_case on broken_case.broken_case_id = item.broken_case_id ");
      sql.append("left outer join rms_item on rms_item.item_id = item.item_id ");
      sql.append("join ace_item_price on ace_item_price.item_id = item.item_id and ace_item_price.ace_rsc_id = 11 ");
      sql.append("left outer join item_upc on item_upc.item_id = item.item_id and primary_upc = 1 ");
      sql.append("where ");
      sql.append("   (rms_item.rms_id in ( ");
      sql.append("      select max(rms_id) ");
      sql.append("      from rms_item ");
      sql.append("      where item.item_id = rms_item.item_id ");
      sql.append("   ) or rms_item.rms_id is null) ");
            
      return m_OraConn.prepareStatement(sql.toString());
   }
}
