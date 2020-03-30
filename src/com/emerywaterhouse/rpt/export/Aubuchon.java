/**
 * 
 */
package com.emerywaterhouse.rpt.export;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

/**
 * @author JFisher
 *
 */
public class Aubuchon extends Report 
{
   private static final String SKU_ELEMENT        = "   <sku>%s</sku>\r\n";
   private static final String COST_ELEMENT       = "   <cost>%s</cost>\r\n";
   private static final String VENDITMNBR_ELEMENT = "   <vendorItemNumber>%s</vendorItemNumber>\r\n";
   private static final String DESCR_ELEMENT      = "   <itemDescription>%s</itemDescription>\r\n";
   private static final String UPC_ELEMENT        = "   <upcCode>%s</upcCode>\r\n";
   private static final String MINORD_ELEMENT     = "   <minOrderQuantity>%s</minOrderQuantity>\r\n";
   private static final String VELO_ELEMENT       = "   <velocityCode>%s</velocityCode>\r\n";

   /**
    * Environment enumeration
    */
   public enum CatalogType {
      txt,
      xls,
      xml
   };

   public enum CatalogTarget {
      emery,
      ace
   }

   private final String DELIMITER = "\t"; // aubuchon will also accept ~ and ^

   private CatalogType m_CatType;
   private String m_CustId;
   private int m_Dc;             // The customer's distribution center.
   private int m_Rsc  ;          // The customer's ACE distribution center.

   private PreparedStatement m_EmeryItemData;
   private PreparedStatement m_AceItemData;
   private PreparedStatement m_GetSellPrice;
   
   /**
    *
    */
   public Aubuchon()
   {
      super();

      m_CatType = CatalogType.txt;
      m_CustId = "";
      m_Dc = 1;
      m_Rsc = 11;
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
   private boolean buildOutputFile(CatalogTarget target) throws FileNotFoundException
   {
      FileOutputStream outFile = null;
      boolean result = false;

      // Aubuchon requires separate file names depending on which warehouse
      String portFilename = m_FilePath + m_FileNames.get(0);
      String wiltFilename = m_FilePath + m_FileNames.get(1);

      if (target == CatalogTarget.emery) {
         outFile = new FileOutputStream(portFilename, false);
      } 
      else {
         outFile = new FileOutputStream(wiltFilename, false);
      }

      try {
         result = buildCatalogFile(outFile, target);
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[Aubuchon]", ex);
      }

      finally {
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[Aubuchon]", e);
         }

         outFile = null;
      }

      return result;
   }

   /**
    * Builds the catalog export in the specified  format.
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildCatalogFile(FileOutputStream outFile, CatalogTarget target) throws Exception
   {
      boolean result = false;

      switch ( m_CatType ) {
         case txt: {
            result = buildTextFile(outFile, target);
            break;
         }
   
         case xls: {
            result = buildExcelFile(outFile);
            break;
         }
   
         case xml: {
            result = buildXmlFile(outFile, target);
            break;
         }
      }

      return result;
   }

   /**
    * 
    * @param outFile
    * @return
    * @throws Exception
    */
   private boolean buildExcelFile(FileOutputStream outFile) throws Exception
   {
      boolean result = false;

      // TODO

      return result;
   }

   /**
    * 
    * @param outFile
    * @return
    * @throws Exception
    */
   private boolean buildXmlFile(FileOutputStream outFile, CatalogTarget target) throws Exception
   {
      boolean result = false;
      StringBuffer xml = new StringBuffer();
      ResultSet itemData = null;
      ResultSet priceData = null;
      String itemId;
      int itemEaId;
      String cost = "";
      String vendorItemNum;
      String description;
      String upcCode;
      String minOrderQty;
      String velocity;

      if ( target == CatalogTarget.emery ) {
         m_EmeryItemData.setString(1, m_CustId);
         m_EmeryItemData.setInt(2, m_Dc);
         itemData = m_EmeryItemData.executeQuery();
      } 
      else {
         m_AceItemData.setString(1, m_CustId);
         m_AceItemData.setInt(2, m_Rsc);
         itemData = m_AceItemData.executeQuery();
      }

      try {
         xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n");
         xml.append("<catalog xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" ");
         xml.append("xmlns =\"http://www.emeryonline.com/catalog\" ");
         xml.append("xsi:schemaLocation=\"http://www.emeryonline.com/catalog ");
         xml.append("http://www.emeryonline.com/catalog/catalogexp.xsd\" >\r\n");
         outFile.write(xml.toString().getBytes());
         xml.setLength(0);

         while ( itemData.next() ) {
            itemId = itemData.getString("item_id");
            itemEaId = itemData.getInt("item_ea_id");
            vendorItemNum = itemData.getString("vendor_item_num");
            description = itemData.getString("description").replaceAll("[^A-Za-z0-9 ]", "");
            upcCode = itemData.getString("upc_code");
            minOrderQty = itemData.getString("min_order_qty");
            velocity = itemData.getString("velocity");

            try {
               m_GetSellPrice.setString(1, m_CustId);
               m_GetSellPrice.setInt(2, itemEaId);
               priceData = m_GetSellPrice.executeQuery();

               if (priceData.next()) {
                  cost = priceData.getString("cost");
               }
            } 
            
            catch (Exception e) {
               //log.warn("[Aubuchon] Could not get price for item: " + itemId); // This was way too spammy.
               m_EdbConn.rollback();
            } 
            
            finally {
               try {
                  priceData.close();
               } 
               
               catch (Exception e) {
               }
            }

            xml.append("<catalogItem>\r\n");
            xml.append(String.format(SKU_ELEMENT, new Object[] { itemId }));
            xml.append(String.format(COST_ELEMENT, new Object[] { cost }));
            xml.append(String.format(VENDITMNBR_ELEMENT, new Object[] { vendorItemNum }));
            xml.append(String.format(DESCR_ELEMENT, new Object[] { description }));
            xml.append(String.format(UPC_ELEMENT, new Object[] { upcCode }));
            xml.append(String.format(MINORD_ELEMENT, new Object[] { minOrderQty }));
            xml.append(String.format(VELO_ELEMENT, new Object[] { velocity }));
            xml.append("</catalogItem>\r\n");

            outFile.write(xml.toString().getBytes());
            xml.setLength(0); 
         }

         xml.append("</catalog>");
         outFile.write(xml.toString().getBytes());
         result = true;
      }

      finally {
         itemData.close();
      }

      return result;
   }

   /**
    * 
    * @param outFile
    * @return
    * @throws Exception
    */
   private boolean buildTextFile(FileOutputStream outFile, CatalogTarget target) throws Exception
   {
      boolean result = false;
      StringBuffer line = new StringBuffer();
      ResultSet itemData = null;
      ResultSet priceData = null;
      String itemId;
      String cost;
      String vendorItemNum;
      String description;
      String upcCode;
      String minOrderQty;
      String velocity;

      if ( target == CatalogTarget.emery ) {
         m_EmeryItemData.setInt(1, m_Dc);
         itemData = m_EmeryItemData.executeQuery();
      } 
      else {
         m_AceItemData.setInt(1, m_Rsc);
         itemData = m_AceItemData.executeQuery();
      }

      try {
         while ( itemData.next() ) {
            itemId = itemData.getString("item_id");
            vendorItemNum = itemData.getString("vendor_item_num");
            description = itemData.getString("description");
            upcCode = itemData.getString("upc_code");
            minOrderQty = itemData.getString("min_order_qty");
            velocity = itemData.getString("velocity");
            cost = "";

            try {
               m_GetSellPrice.setString(1, m_CustId);
               m_GetSellPrice.setString(2, itemId);
               priceData = m_GetSellPrice.executeQuery();
            
               if (priceData.next())
                  cost = priceData.getString("cost");
               else
                  throw new Exception("No rows returned");
            } 
            
            catch (Exception e) {
               log.warn("[Aubuchon] Could not get price for item: " + itemId);
            } 
            
            finally {
               priceData.close();
            }

            line.append(itemId);
            line.append(DELIMITER);
            line.append(cost);
            line.append(DELIMITER);
            line.append(vendorItemNum);
            line.append(DELIMITER);
            line.append(description.replaceAll("[^A-Za-z0-9 ]", ""));
            line.append(DELIMITER);
            line.append(upcCode);
            line.append(DELIMITER);
            line.append(minOrderQty);
            line.append(DELIMITER);
            line.append(velocity);
            line.append(DELIMITER);
            line.append("\r\n");

            outFile.write(line.toString().getBytes());
            line.setLength(0);           
         }

         outFile.write(line.toString().getBytes());
         result = true;
      }

      finally {
         itemData.close();
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_EmeryItemData);     
      closeStmt(m_AceItemData);
      closeStmt(m_GetSellPrice);
      
      m_EmeryItemData = null;
      m_AceItemData = null;
      m_GetSellPrice = null;
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
         m_EdbConn.setAutoCommit(false);

         // Aubuchon requires 2 catalog files. We'll create them sequentially. 
         // Should there be an error with generation the first, we'll abort on attempting to create the second.

         if ( prepareStatements() )
            created = buildOutputFile(CatalogTarget.emery); // Build the portland / pittston file

         if (created && m_Rsc != 0) { // it is possible to run without an ACE warehouse - originally, we only sent text portland and text pittston files
            created = buildOutputFile(CatalogTarget.ace); // Build the ace file
         }
      }

      catch ( Exception ex ) {
         log.fatal("[Aubuchon]", ex);
      }

      finally {
         closeStatements();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
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

            sql.append("select item_entity_attr.item_id, ");
            sql.append("item_entity_attr.item_ea_id, ");
            sql.append("vendor_item_ea_cross.vendor_item_num, ");
            sql.append("nvl(bmi_item.web_descr, item_entity_attr.description) as description, ");
            sql.append("nvl(ejd_item_whs_upc.upc_code, ' ') as upc_code, ");
            sql.append("decode(broken_case.description, 'ALLOW BROKEN CASES', 1, ejd_item_warehouse.stock_pack) as min_order_qty, ");
            sql.append("item_velocity.velocity, ");           
            sql.append("ace_item_xref.direct_match ");           
            sql.append("from item_entity_attr ");
            sql.append("join customer on customer.customer_id = ? ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join item_type on item_type.item_type_id = item_entity_attr.item_type_id and itemtype = 'STOCK' "); // Note! This line differs from the ace query
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join warehouse on warehouse.warehouse_id = ejd_item_warehouse.warehouse_id ");
            sql.append("join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id and disposition in ('BUY-SELL', 'NOBUY') ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_whs_id = ejd_item_warehouse.ejd_item_whs_id and ejd_item_whs_upc.primary_upc = 1 ");
            sql.append("left outer join bmi_item on bmi_item.item_id = item_entity_attr.item_id ");
            sql.append("left outer join vendor_item_ea_cross on vendor_item_ea_cross.vendor_id = item_entity_attr.vendor_id and vendor_item_ea_cross.item_ea_id = item_entity_attr.item_ea_id ");
            sql.append("join broken_case on broken_case.broken_case_id = ejd_item.broken_case_id ");
            sql.append("join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id ");
            sql.append("left join ace_item_xref on item_entity_attr.item_id = ace_item_xref.item_id ");
            sql.append("where ejd_item_warehouse.warehouse_id = ? and ");
            sql.append("      (ace_item_xref.direct_match is null or ace_item_xref.direct_match = false) "); // Note! This line differs from the ace query.
            sql.append("order by item_entity_attr.item_id ");
            m_EmeryItemData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select item_entity_attr.item_id, ");
            sql.append("item_entity_attr.item_ea_id, ");
            sql.append("vendor_item_ea_cross.vendor_item_num, ");
            sql.append("nvl(bmi_item.web_descr, item_entity_attr.description) as description, ");
            sql.append("nvl(ejd_item_whs_upc.upc_code, ' ') as upc_code, ");
            sql.append("ejd_item_warehouse.stock_pack as min_order_qty, "); // WILTON does not check nbc, always use stock pack
            sql.append("item_velocity.velocity ");           
            sql.append("from item_entity_attr ");
            sql.append("join customer on customer.customer_id = ? ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join item_type on item_type.item_type_id = item_entity_attr.item_type_id and itemtype = 'ACE' ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join warehouse on warehouse.warehouse_id = ejd_item_warehouse.warehouse_id ");
            sql.append("join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id and disposition = 'BUY-SELL' ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_whs_id = ejd_item_warehouse.ejd_item_whs_id and ejd_item_whs_upc.primary_upc = 1 ");
            sql.append("left outer join bmi_item on bmi_item.item_id = item_entity_attr.item_id ");
            sql.append("left outer join vendor_item_ea_cross on vendor_item_ea_cross.vendor_id = item_entity_attr.vendor_id and vendor_item_ea_cross.item_ea_id = item_entity_attr.item_ea_id ");
            sql.append("join broken_case on broken_case.broken_case_id = ejd_item.broken_case_id ");
            sql.append("join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id ");
            sql.append("where ejd_item_warehouse.warehouse_id = ? ");
            sql.append("order by item_entity_attr.item_id ");
            m_AceItemData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select price as cost from ejd_cust_procs.get_sell_price(?, ?) "); // customer_id, item_ea_id
            m_GetSellPrice = m_EdbConn.prepareStatement(sql.toString());

            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("[Aubuchon]", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("[Aubuchon].prepareStatements - null EDB connection");

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

         if ( param.name.equalsIgnoreCase("acedc") )
            m_Rsc = Integer.parseInt(param.value);

         if ( param.name.equalsIgnoreCase("cust") || param.name.equalsIgnoreCase("custid") )
            m_CustId = param.value;

         //
         // Set the catalog type.  Defaults to text and xml is handled by the regular catalog 
         // export at this time.  If it doesn't match then leave it as the default.
         if ( param.name.equalsIgnoreCase("cattype") ) {
            if ( param.value.equals("text") ) {
               m_CatType = CatalogType.txt;
            } 
            else { 
               if ( param.value.equals("excel" ) ) {
                  m_CatType = CatalogType.xls;
               } 
               else { // must be xml
                  m_CatType = CatalogType.xml;
               }
            }
         }
      }

      // 'REPLACE' will be modified at file generation to display whether the file contains emery or ace items
      fileName.append("REPLACE-emery-catalog");

      if (m_CatType.equals(CatalogType.txt)) {
         fileName.append(".txt");
      } 
      else {
         if (m_CatType.equals(CatalogType.xls) )
            fileName.append(".xls");
         else 
            fileName.append(".xml");         
      }
      
      m_FileNames.add(fileName.toString().replace("REPLACE", "Portland"));
      m_FileNames.add(fileName.toString().replace("REPLACE", "Wilton"));
   }

}
