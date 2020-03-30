/**
 * File: SpruceCatalog.java
 * Description: Exports the catalog data in a flat file.  Based on the Rockysoft format with an added
 *      web item description field.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 2016/12/14
 * Last Update: 2016/12/14 jfisher
 *
 * History
 *    Initial production version
 */
package com.emerywaterhouse.rpt.export;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Arrays;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;


public class SpruceCatalog extends Report
{
   private static final int recordLength = 562;

   private PreparedStatement m_PurchData;
   private PreparedStatement m_ItemData;
   private PreparedStatement m_PriceData;
   private PreparedStatement m_UpcData;
   private PreparedStatement m_CustWhs;

   //
   // Params
   private String m_CustId;      // The customer number for the report data to be run against.
   private int m_Dc;             // The customer's distribution center.
   private boolean m_Overwrite;  // Overwrite the file flag
   /**
    *
    */
   public SpruceCatalog()
   {
      super();

      m_CustId = "";
      m_Dc = 0;
      m_Overwrite = false;
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      m_CustId = null;

      super.finalize();
   }

   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream outFile = null;
      boolean result = false;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         result = buildCatalogFile(outFile);
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[Spruce]", ex);
      }

      finally {
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[Spruce]", e);
         }

         outFile = null;
      }

      return result;
   }

   /**
    * Builds the catalog export in the Spruce (flat file) format.
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildCatalogFile(FileOutputStream outFile) throws Exception
   {
      boolean result = false;
      StringBuffer line = new StringBuffer();
      char[] filler = new char[recordLength];
      ResultSet itemData = null;
      ResultSet priceData = null;
      String itemId = null;
      String desc = null;
      String shortDesc = null;
      String upc = null;
      String altUpc = null;
      String vndSku = null;
      String vndId = null;
      String vndName = null;
      String brkCase = null;
      String retPack = null;
      String packOf = null;
      String cost = null;
      String retail = null;
      String length = null;
      String width = null;
      String height = null;
      String weight = null;
      String cube = null;
      String uom = null;
      String flc = null;
      String mdc = null;
      String nrha = null;
      String brandName = null;
      String noun = null;
      String modifier = null;
      String purHist = null;
      String disco = " ";
      String shipUnit = null;
      int itemEaId = 0;

      //
      // If the DC has not been set, get it based on the
      // customer's warehouse settings.
      if ( m_Dc == 0 )
         m_Dc = getWarehouseId(m_CustId);

      //
      // Set the customer id for the cost and retail price calculations and the warehouse id for the items.
      // The warehouse id is not the fascor id, but the eis warehouse id number. (1 or 2)
      m_ItemData.setInt(1, m_Dc);      
      m_ItemData.setString(2, m_CustId);

      itemData = m_ItemData.executeQuery();

      try {
         Arrays.fill(filler, ' ');

         while ( itemData.next() && m_Status == RptServer.RUNNING ) {
            itemId = itemData.getString("item_id");
            itemEaId = itemData.getInt("item_ea_id");
         	
         	// see if we can get pricing.
         	// if we can't, just continue to the next item instead of putting in an item without a price in the catalog.
         	m_PriceData.setString(1, m_CustId);
         	m_PriceData.setInt(2, itemEaId);
         	m_PriceData.setString(3, m_CustId);
         	m_PriceData.setString(4, itemId);
         	
            try {
               priceData = m_PriceData.executeQuery();
               
               if ( !priceData.next() ) {
                  continue; // skip to the next item, we didn't get any pricing for this one.
               }
            } 
            
            catch (Exception e) { // problem getting a price            	
            	m_EdbConn.rollback();
            	continue;
            }
         	
            line.setLength(0);
            line.append(filler);

            setCurAction(String.format("processing item %s for customer %s", itemId, m_CustId));

            desc = itemData.getString("description");
            upc = itemData.getString("upc_code");
            vndSku = itemData.getString("vendor_item_num");
            vndId = itemData.getString("vendor_id");
            vndName = itemData.getString("name");
            brkCase = itemData.getString("broken_case");
            retPack = itemData.getString("retail_pack");
            packOf = itemData.getString("packof");
            cost = String.format("%1.3f",priceData.getDouble("cost"));
            retail = String.format("%1.2f",priceData.getDouble("retail"));
            length = String.format("%1.3f",itemData.getDouble("length"));
            width = String.format("%1.3f",itemData.getDouble("width"));
            height = String.format("%1.3f",itemData.getDouble("height"));
            weight = String.format("%1.3f",itemData.getDouble("weight"));
            cube = String.format("%1.3f",itemData.getDouble("cube"));
            uom = itemData.getString("uom");
            flc = itemData.getString("flc_id");
            mdc = itemData.getString("mdc_id");
            nrha = itemData.getString("nrha_id");
            brandName = itemData.getString("brand_name");
            noun = itemData.getString("noun");
            modifier = itemData.getString("modifier");
            altUpc = getAltUpc(itemId);
            purHist = getPurchHist(itemId, m_CustId);
            shipUnit = itemData.getString("ship_unit");

            //
            // The spec calls for just 25 characters for the description.
            if ( desc.length() > 25 )
               shortDesc = desc.substring(0, 25);

            //
            // Make sure there's no null text in the document and we also
            // need a length.
            if ( upc == null || upc.length() == 0 )
               upc = " ";

            if ( vndSku == null || vndSku.length() == 0 )
               vndSku = " ";

            if ( brandName == null || brandName.length() == 0 )
               brandName = " ";

            if ( noun == null || noun.length() == 0 )
               noun = " ";

            if ( modifier == null || modifier.length() == 0 )
               modifier = " ";

            line.replace(0, itemId.length(), itemId);             //   7
            line.replace(7, 7 + upc.length(), upc);               //  20
            line.replace(27, 27 + altUpc.length(), altUpc);       //  20
            line.replace(47, 47 + shortDesc.length(), shortDesc); //  25
            line.replace(72, 72 + vndSku.length(), vndSku);       //  30
            line.replace(102, 102 + vndId.length(), vndId);       //   7
            line.replace(109, 109 + vndName.length(), vndName);   //  75
            line.replace(184, 184 + brkCase.length(), brkCase);   //   1
            line.replace(185, 185 + retPack.length(), retPack);   //   6
            line.replace(191, 191 + packOf.length(), packOf);     //   3
            line.replace(194, 194 + cost.length(), cost);         //   8
            line.replace(202, 202 + retail.length(), retail);     //   7
            line.replace(209, 209 + length.length(), length);     //  10
            line.replace(219, 219 + width.length(), width);       //  10
            line.replace(229, 229 + height.length(), height);     //  10
            line.replace(239, 239 + weight.length(), weight);     //  10
            line.replace(249, 249 + cube.length(), cube);         //  10
            line.replace(259, 259 + uom.length(), uom);           //   3
            line.replace(262, 262 + nrha.length(), nrha);         //   2
            line.replace(264, 264 + mdc.length(), mdc);           //   3
            line.replace(267, 267 + flc.length(), flc);           //   4
            line.replace(271, 271 + noun.length(), noun);         //  80
            line.replace(351, 351 + modifier.length(), modifier); //  80
            line.replace(431, 431 + purHist.length(), purHist);   //   7
            line.replace(438, 438 + disco.length(), disco);       //   1
            line.replace(439, 439 + desc.length(), desc);         // 120
            while (line.length() < recordLength) {
            	line.append(" "); // descriptions are variable length and the array of spaces isn't filling even though it has a record length of 563?
            }
            line.replace(560, 560 + shipUnit.length(), shipUnit);
            
            line.append("\r\n");

            outFile.write(line.toString().getBytes());
            line.setLength(0);
         }

         outFile.write(line.toString().getBytes());
         result = true;
      }
      
      catch ( SQLException ex ) {
         log.error("[Spruce]", ex);
      }

      finally {
         closeRSet(itemData);
         closeRSet(priceData);
         itemData = null;
         priceData = null;

         itemId = null;
         desc = null;
         upc = null;
         altUpc = null;
         vndSku = null;
         vndId = null;
         vndName = null;
         brkCase = null;
         retPack = null;
         packOf = null;
         cost = null;
         retail = null;
         length = null;
         width = null;
         height = null;
         weight = null;
         cube = null;
         uom = null;
         flc = null;
         mdc = null;
         nrha = null;
         brandName = null;
         noun = null;
         modifier = null;
         purHist = null;

         outFile.close();
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
      closeStmt(m_PurchData);
      closeStmt(m_PriceData);
      closeStmt(m_UpcData);
      closeStmt(m_CustWhs);

      m_ItemData = null;
      m_PurchData = null;
      m_PriceData = null;
      m_UpcData = null;
      m_CustWhs = null;
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
         log.fatal("[Spruce]", ex);
      }

      finally {
         closeStatements();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Returns an alternate UPC for an item.
    *
    * @param itemId The item to get the alternate UPC for.
    * @return The UPC
    *
    * @throws SQLException
    */
   private String getAltUpc(String itemId) throws SQLException
   {
      String altUpc = " ";
      ResultSet rs = null;

      m_UpcData.setString(1, itemId);
      m_UpcData.setString(2, m_CustId);
      m_UpcData.setInt(3, m_Dc);
      rs = m_UpcData.executeQuery();

      try {
         if ( rs.next() ) {
            altUpc = rs.getString(1);
         }
      }

      catch ( SQLException ex ) {
         log.error("[Spruce]", ex);
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return altUpc;
   }

   /**
    * Returns the 12 month purchase history of an item for a customer
    *
    * @param itemId The item to check
    * @param custId The customer the purchased the item
    *
    * @return The units ordered as a string
    * @throws SQLException
    */
   private String getPurchHist(String itemId, String custId) throws SQLException
   {
      String purchHist = "0";
      ResultSet rs = null;

      m_PurchData.setString(1, custId);
      m_PurchData.setString(2, itemId);
      rs = m_PurchData.executeQuery();

      try {
         if ( rs.next() ) {
            purchHist = rs.getString(1);
         }
      }
      
      catch ( SQLException ex ) {
         log.error("[Spruce]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return purchHist;
   }

   /**
    * Gets the warehouse id for the given customer.
    * @param custId
    * @return The warehouse id.
    * @throws SQLException
    */
   private int getWarehouseId(String custId) throws SQLException
   {
      int id = 1;
      ResultSet rs = null;

      m_CustWhs.setString(1, custId);
      rs = m_CustWhs.executeQuery();

      try {
         if ( rs.next() ) {
            id = rs.getInt(1);
         }
      }
      
      catch ( SQLException ex ) {
         log.error("[Spruce]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return id;
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
            sql.append("   item_entity_attr.item_ea_id, item_entity_attr.item_id, nvl(web_item_ea.web_descr, item_entity_attr.description) as description, ");
            sql.append("   upc_code, viec.vendor_item_num, vendor.vendor_id, vendor.name, ");
            sql.append("   img_url_sm, img_url_md, img_url_lg, ");
            sql.append("   decode(bc.description, 'ALLOW BROKEN CASES', 'Y', 'N') as broken_case, ");
            sql.append("   ejd_item_warehouse.stock_pack as \"packof\", ");
            sql.append("   length as length, width, height, ejd_item.weight, cube, ");
            sql.append("   retail_unit.unit as uom, ejd_item.flc_id, mdc.mdc_id, mdc.nrha_id, ");
            sql.append("   item_entity_attr.retail_pack, web_item_ea.brand_name, noun, modifier, ship_unit.unit as ship_unit ");
            sql.append("from ");
            sql.append("   item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id and warehouse_id = ? ");
            sql.append("join web_item_ea on web_item_ea.item_ea_id = item_entity_attr.item_ea_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("left outer join vendor_item_ea_cross viec on viec.vendor_id = item_entity_attr.vendor_id and ");
            sql.append("   viec.item_ea_id = item_entity_attr.item_ea_id ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = ejd_item.ejd_item_id and ejd_item_whs_upc.primary_upc = 1 ");
            sql.append("   and ejd_item_whs_upc.warehouse_id = ejd_item_warehouse.warehouse_id ");            
            sql.append("join broken_case bc on bc.broken_case_id = ejd_item.broken_case_id ");
            sql.append("join retail_unit on retail_unit.unit_id = item_entity_attr.ret_unit_id ");
            sql.append("join ship_unit on ship_unit.unit_id = item_entity_attr.ship_unit_id ");
            sql.append("join flc on flc.flc_id = ejd_item.flc_id ");
            sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
            sql.append("join nrha on nrha.nrha_id = mdc.nrha_id ");
            sql.append("where ");
            sql.append("   item_entity_attr.item_type_id = 1 and nvl(ejd_item_warehouse.in_catalog, 0) = 1 ");
            sql.append("union ");
            sql.append("select ");
            sql.append("   item_entity_attr.item_ea_id ,item_entity_attr.item_id, ");
            sql.append("   nvl(web_item_ea.web_descr, item_entity_attr.description) as description, ejd_item_whs_upc.upc_code, ");
            sql.append("   viec.vendor_item_num, item_entity_attr.vendor_id, vendor.name, ");
            sql.append("   img_url_sm, img_url_md, img_url_lg, ");
            sql.append("   decode(bc.description, 'ALLOW BROKEN CASES', 'Y', 'N') as broken_case, ");
            sql.append("   ejd_item_warehouse.stock_pack as \"packof\", ");
            sql.append("   ace_item_rsc.length, ace_item_rsc.width, ace_item_rsc.height, ace_item_rsc.weight, ");
            sql.append("   ace_item_rsc.cube, retail_unit.unit as uom, ejd_item.flc_id, mdc.mdc_id, mdc.nrha_id, ");
            sql.append("   item_entity_attr.retail_pack, web_item_ea.brand_name, noun, modifier, retail_unit.unit ");
            sql.append("from ");
            sql.append("   ace_cust_xref ");
            sql.append("join ace_rsc on ace_rsc.sap_site_cd = ace_cust_xref.ace_rsc ");
            sql.append("join ace_item_rsc on ace_item_rsc.ace_rsc_id = ace_rsc.ace_rsc_id ");
            sql.append("join ace_item_xref on ace_item_xref.ace_xref_id = ace_item_rsc.ace_xref_id ");
            sql.append("join item_entity_attr on item_entity_attr.item_id = ace_item_xref.item_id and item_entity_attr.item_type_id = 9 ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join warehouse on warehouse.ace_rsc_id = ace_rsc.ace_rsc_id ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.warehouse_id = warehouse.warehouse_id ");
            sql.append("join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_whs_id = ejd_item_warehouse.ejd_item_whs_id ");
            sql.append("join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id ");
            sql.append("left outer join vendor_item_ea_cross viec on viec.vendor_id = item_entity_attr.vendor_id and viec.item_ea_id = item_entity_attr.item_ea_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join web_item_ea on web_item_ea.item_ea_id = item_entity_attr.item_ea_id ");
            sql.append("join retail_unit on retail_unit.unit_id = item_entity_attr.ret_unit_id ");
            sql.append("join broken_case bc on bc.broken_case_id = ejd_item.broken_case_id ");
            sql.append("join flc on flc.flc_id = ejd_item.flc_id ");
            sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
            sql.append("where customer_id = ? ");
            sql.append("and disposition in ('BUY-SELL') ");
            sql.append("order by item_id ");
            m_ItemData = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select "); 
            sql.append("   (select price from ejd_cust_procs.get_sell_price(?, ?)) as cost, "); // cust id, item_ea_id
            sql.append("   ejd.retail_price_procs.getretailprice(?, ?) as retail "); // cust id, item id
            m_PriceData = m_EdbConn.prepareStatement(sql.toString());            
            
            //
            // Item attributes
            sql.setLength(0);
            sql.append("select nvl(sum(qty_ordered), 0) as qty_ord ");
            sql.append("from inv_dtl ");
            sql.append("where cust_nbr = ? and item_nbr = ? and ");
            sql.append("tran_type = 'SALE' and sale_type in ('WAREHOUSE', 'ACE DIRECT') and ");
            sql.append("invoice_date > trunc(now()) - 360 ");
            m_PurchData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select * ");
            sql.append("from ejd_item_whs_upc ");
            sql.append("where ejd_item_id = ");
            sql.append("   (select ejd_item_id from item_entity_attr where item_ea_id = ");
            sql.append("   (select code from ejd_item_procs.get_item_ea_id(?,?))) "); // itemid, custid
            sql.append("and primary_upc = 0 ");
            sql.append("and warehouse_id = ? "); // warehouse_id
            m_UpcData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select warehouse_id from cust_warehouse where customer_id = ? and whs_priority = 1 ");
            m_CustWhs = m_EdbConn.prepareStatement(sql.toString());

            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("[Spruce]", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("[Spruce].prepareStatements - null oracle connection");

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

         //
         // Note, DC will normally be set by the query.  Allow overriding here.
         if ( param.name.equalsIgnoreCase("dc") )
            m_Dc = Integer.parseInt(param.value);

         if ( param.name.equalsIgnoreCase("cust") || param.name.equalsIgnoreCase("custid") )
            m_CustId = param.value;

         if ( param.name.equalsIgnoreCase("overwrite") )
            m_Overwrite = param.value.equalsIgnoreCase("true") ? true : false;
      }

      //
      // Some customers want the same file name each time.  If that's the case, we
      // need to overwrite what we have.  Also Spruce only wants the name to be catalog.txt
      if ( !m_Overwrite ) {
         fileName.append(Long.toString(System.currentTimeMillis()));
         fileName.append("-");
      }

      fileName.append(String.format("CATALOG%s.TXT", m_CustId));
      m_FileNames.add(fileName.toString());
   }
   
   
   /*
   public static void main(String[] args) {
   	SpruceCatalog cat = new SpruceCatalog();

      Param p1 = new Param();
      p1.name = "cust";
      p1.value = "001716";
      Param p2 = new Param();
      p2.name = "dc";
      p2.value = "1";
      Param p3 = new Param();
      p3.name = "overwrite";
      p3.value = "false";
      ArrayList<Param> params = new ArrayList<Param>();
      params.add(p1);
      params.add(p2);
      params.add(p3);
      
      cat.m_FilePath = "C:\\EXP\\";
      
   	java.util.Properties connProps = new java.util.Properties();
   	connProps.put("user", "ejd");
   	connProps.put("password", "boxer");
   	try {
   		cat.m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33:5444/emery_jensen",connProps);   		
   		cat.m_EdbConn.setAutoCommit(false);
   	} catch (Exception e) {
   		e.printStackTrace();
   	}
      
      cat.setParams(params);
      cat.createReport();
   }*/
   
}
