/**
 * File: MobileDbBuilder.java
 * Description: Builds the database files for mobile platforms.
 *
 * @author Jeff Fisher
 * @author Erik Pearson
 *
 * Create Date: 02/25/2011
 * Last Update: $Id: MobileDbBuilder.java,v 1.8 2014/07/25 18:24:46 jfisher Exp $
 *
 * History:
 *
 * $Log: MobileDbBuilder.java,v $
 * Revision 1.8  2014/07/25 18:24:46  jfisher
 * Fixed the bin label query so use the customer warehouse instead of just portland.  Added a param for warehouse for the item db
 *
 * Revision 1.7  2012/07/11 17:30:56  jfisher
 * in_catalog modification
 *
 * Revision 1.6  2012/01/05 16:26:21  epearson
 * Added binlabel and compconvert db builders
 *
 * Revision 1.5  2011/06/28 14:25:15  epearson
 * Added missing 'break' statement when creating the data file
 *
 * Revision 1.4  2011/03/24 13:16:48  epearson
 * Added BinLabel database build code.
 *
 * Revision 1.3  2011/03/14 10:50:48  epearson
 * Added cvs logging tag
 *
 *
 */
package com.emerywaterhouse.rpt.export;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

/**
 * Builds the database files for mobile platforms.
 *
 */
public class MobileDbBuilder extends Report {
   private int m_AppId;
   private String m_AccessKey;
   private String m_DeviceId;
   private String m_DbName;
   @SuppressWarnings("unused")
   private String m_DbPath;
   private String m_Uid;
   @SuppressWarnings("unused")
   private String m_Pwd;
   private ArrayList<String> m_ParamList;
   private int m_WhsId;

   /**
    *
    */
   public MobileDbBuilder()
   {
      m_ParamList = new ArrayList<String>();
      m_AccessKey = "";
      m_DbName = "";
      m_DbPath = "";
      m_AppId = -1;
      m_DeviceId = "";
      m_WhsId = 1;
   }

   /**
    * Cleanup any allocated resources.
    *
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      if (m_ParamList != null) {
         m_ParamList.clear();
      }

      m_ParamList = null;
      m_DbName = null;
      m_DbPath = null;

      super.finalize();
   }

   /**
    * Builds the Competitor Conversion SQLite Datafile
    *
    * @return true, if successful, otherwise false
    */
   public boolean buildCompConvertDb()
   {
      boolean result = false;
      Connection sc = null;
      Statement ostmt = null;
      Statement sstmt = null;
      PreparedStatement ustmt1 = null;
      PreparedStatement ustmt2 = null;
      ResultSet rs = null;
      String sql = null;
      int i = 0;

      try {
         setCurAction("Registering db driver");
         DriverManager.registerDriver(new org.sqlite.JDBC());
         sc = DriverManager.getConnection(String.format("jdbc:sqlite:%s",
                  m_FilePath + m_DbName));
         sc.setAutoCommit(false);

         //
         // Create the competitor table
         setCurAction("creating competitor table");
         sstmt = sc.createStatement();
         sstmt.execute("create table competitor "
                  + "(comp_id integer primary key, comp_name text)");
         DbUtils.closeDbConn(null, sstmt, null);

         //
         // Create the competitor cross ref table
         setCurAction("creating competitor_xref table");
         sstmt = sc.createStatement();
         sstmt.execute("create table competitor_xref "
                  + "(comp_id integer, emery_item_id text, comp_item_id text, "
                  + "constraint pk_comp_xref primary key (comp_id, emery_item_id, "
                  + "comp_item_id), foreign key (comp_id) "
                  + "references competitor(comp_id))");
         DbUtils.closeDbConn(null, sstmt, null);

         ostmt = m_EdbConn.createStatement();
         ustmt1 = sc.prepareStatement("insert into competitor " +
                  "values (?, ?)");
         ustmt2 = sc.prepareStatement("insert into competitor_xref " +
                  "values (?, ?, ?)");

         //
         // Load the competitor table
         sql = "select id, name from competitor";

         setCurAction("adding competitor records");
         rs = ostmt.executeQuery(sql);

         while (rs.next() && m_Status == RptServer.RUNNING) {
            i++;
            ustmt1.setInt(1, rs.getInt(1));
            ustmt1.setString(2, rs.getString(2));

            ustmt1.addBatch();

            if (i == 100) {
               ustmt1.executeBatch();
               sc.commit();
               i = 0;
            }
         }

         //
         // Make sure we get the last batch of records.
         if (i > 0) {
            ustmt1.executeBatch();
            sc.commit();
         }

         //
         // Load the competitor_xref table
         i = 0;
         DbUtils.closeDbConn(null, ostmt, rs);

         sql = "select comp_id, emery_item_id, comp_item_id " +
                  "from competitor_xref";

         setCurAction("adding competitor xref records");
         ostmt = m_EdbConn.createStatement();
         rs = ostmt.executeQuery(sql);

         while (rs.next() && m_Status == RptServer.RUNNING) {
            i++;
            ustmt2.setInt(1, rs.getInt(1));
            ustmt2.setString(2, rs.getString(2));
            ustmt2.setString(3, rs.getString(3));

            ustmt2.addBatch();

            if (i == 100) {
               ustmt2.executeBatch();
               sc.commit();
               i = 0;
            }
         }

         if (i > 0) {
            ustmt2.executeBatch();
            sc.commit();
         }

         //
         // Update the version table with the latest db information.
         if (m_Status == RptServer.RUNNING) {
            updateDbVer();
         }

         result = true;
      }

      catch (Exception ex) {         
         log.error("[MobileDbBuilder]", ex);
      }

      finally {
         DbUtils.closeDbConn(sc, ustmt1, null);
         DbUtils.closeDbConn(null, ustmt2, null);
         sc = null;
         ustmt1 = null;
         ustmt2 = null;
         sstmt = null;

         DbUtils.closeDbConn(null, ostmt, rs);
         ostmt = null;
         sql = null;
         rs = null;
      }

      return result;
   }

   /**
    * Builds the database for the emery link program.
    */
   private boolean buildEmeryLinkDb()
   {
      boolean result = false;
      boolean isAceCust = m_WhsId > 2;
      Connection sc = null;
      Statement sstmt = null;
      PreparedStatement ostmt = null;      
      PreparedStatement ustmt1 = null;
      PreparedStatement ustmt2 = null;
      PreparedStatement ustmt3 = null;
      ResultSet rs = null;
      StringBuffer sql = new StringBuffer();
      String curItem = null;
      String prevItem = "";
      int i = 0;

      try {
         setCurAction("Registering db driver");
         DriverManager.registerDriver(new org.sqlite.JDBC());
         sc = DriverManager.getConnection(String.format("jdbc:sqlite:%s", m_FilePath + m_DbName));
         sc.setAutoCommit(true);

         //
         // Create the item table
         setCurAction("creating item table");
         
         sstmt = sc.createStatement();
         sstmt.execute(
            "create table item (item_id text, stock_pack integer, broken_case text, description text , " +
            "constraint pk_item primary key (item_id) )"
         );

         sstmt.execute("create index idx_itm_desc on item (description)");
         
         //
         // Create the cross ref table
         setCurAction("creating cross_ref table");
         
         sstmt.execute(
            "create table cross_ref ( " +
            "customer_id text, item_id text, cust_sku text, " + 
            "constraint pk_crossref primary key (customer_id, item_id, cust_sku) )"
         );
         
         sstmt.execute("create index idx_cr_item on cross_ref (item_id)");
         sstmt.execute("create index idx_cr_cust on cross_ref (customer_id)");         
         sstmt.execute("create index idx_cr_cust_sku on cross_ref (cust_sku)");
       
         //
         // Create the upc table.
         setCurAction("creating upc table");
         
         sstmt.execute(
            "create table item_upc ( " + 
            "item_id text, upc_code text, primary_upc integer, " + 
            "constraint pk_upc primary key (item_id, upc_code, primary_upc))"
         );

         sstmt.execute("create index idx_up_item on item_upc (item_id)");
         sstmt.execute("create index idx_up_upc on item_upc (upc_code)");

         DbUtils.closeDbConn(null, sstmt, null);

         ustmt1 = sc.prepareStatement("insert into item values (?, ?, ?, ?)");
         ustmt2 = sc.prepareStatement("insert into cross_ref values (?, ?, ?)");
         ustmt3 = sc.prepareStatement("insert into item_upc values (?, ?, ?)");

         sql.append("select ");
         sql.append("item_entity_attr.item_id, eiw.stock_pack, decode(bc.description, 'ALLOW BROKEN CASES', 'Y', 'N') brk_case, bmi_item.web_descr, item_type_id ");
         sql.append("from item_entity_attr ");
         sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
         sql.append("join bmi_item on bmi_item.item_id = item_entity_attr.item_id ");
         sql.append("join ejd_item_warehouse eiw on eiw.ejd_item_id = item_entity_attr.ejd_item_id and warehouse_id = ? and eiw.in_catalog = 1 ");
         sql.append("join broken_case bc on bc.broken_case_id = ejd_item.broken_case_id ");
                  
         if ( !isAceCust ) {
            sql.append("where item_type_id = 1 ");
            sql.append("union ");
            sql.append("select ");
            sql.append("item_entity_attr.item_id, eiw.stock_pack, decode(bc.description, 'ALLOW BROKEN CASES', 'Y', 'N') brk_case, bmi_item.web_descr, item_type_id ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join bmi_item on bmi_item.item_id = item_entity_attr.item_id ");
            sql.append("join ejd_item_warehouse eiw on eiw.ejd_item_id = item_entity_attr.ejd_item_id and warehouse_id = 11 and eiw.in_catalog = 1 ");
            sql.append("join broken_case bc on bc.broken_case_id = ejd_item.broken_case_id ");
            sql.append("where item_type_id = 9 ");
         }
         else
            sql.append("where item_type_id = 8 ");
         
         sql.append("order by item_id, item_type_id ");
                  
         setCurAction("adding item records");
         
         ostmt = m_EdbConn.prepareStatement(sql.toString());
         ostmt.setInt(1, m_WhsId);
         rs = ostmt.executeQuery();

         log.info("[mobile db] adding item records");
         
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            i++;
            curItem = rs.getString(1);
            setCurAction("adding item: " + curItem);
            
            if ( !curItem.equals(prevItem)) {
               ustmt1.setString(1, curItem);
               ustmt1.setInt(2, rs.getInt(2));
               ustmt1.setString(3, rs.getString(3));
               ustmt1.setString(4, rs.getString(4));
                           
               try {               
                  ustmt1.executeUpdate();
               }
               
               catch ( Exception ex ) {
                  log.error("[mobile db] adding item record: " + ex.getMessage()); // don't do anything.
                  
                  //
                  // SQLite is POS
                  DbUtils.closeDbConn(null, ustmt1, null);
                  ustmt1 = sc.prepareStatement("insert into item values (?, ?, ?, ?)");
               }
            }
            
            prevItem = curItem;
         }

         //
         // Load the item_cross table
         i = 0;
         DbUtils.closeDbConn(null, ostmt, rs);
         sql.setLength(0);

         sql.append("select distinct customer_id, item_entity_attr.item_id, customer_sku ");
         sql.append("from item_ea_cross ");

         if ( !isAceCust )
            sql.append("join item_entity_attr on item_entity_attr.item_ea_id = item_ea_cross.item_ea_id and item_type_id in (1, 9) ");
         else
            sql.append("join item_entity_attr on item_entity_attr.item_ea_id = item_ea_cross.item_ea_id and item_type_id = 8 ");

         sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and " );

         if ( !isAceCust )
            sql.append("         ejd_item_warehouse.warehouse_id in ( ?, 11) and ejd_item_warehouse.in_catalog = 1 ");
         else
            sql.append("         ejd_item_warehouse.warehouse_id = ? and ejd_item_warehouse.in_catalog = 1 ");

         sql.append("order by customer_id, item_id");

         setCurAction("adding cross reference records");
         
         ostmt.close();
         ostmt = m_EdbConn.prepareStatement(sql.toString());
         ostmt.setInt(1, m_WhsId);
         rs = ostmt.executeQuery();

         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            i++;
            ustmt2.setString(1, rs.getString(1));
            ustmt2.setString(2, rs.getString(2));
            ustmt2.setString(3, rs.getString(3));

            ustmt2.addBatch();

            if (i == 100) {
               ustmt2.executeBatch();               
               i = 0;
            }
         }

         if ( i > 0 )
            ustmt2.executeBatch();

         //
         // Load the item_upc table
         i = 0;
         DbUtils.closeDbConn(null, ostmt, rs);
         sql.setLength(0);

         sql.append("select distinct item_entity_attr.item_id, upc_code, primary_upc ");
         sql.append("from ejd_item_whs_upc ");
         sql.append("join ejd_item on ejd_item.ejd_item_id = ejd_item_whs_upc.ejd_item_id ");

         if ( !isAceCust )
            sql.append("join item_entity_attr on item_entity_attr.ejd_item_id =  ejd_item.ejd_item_id and item_type_id in (1, 9) ");
         else
            sql.append("join item_entity_attr on item_entity_attr.ejd_item_id =  ejd_item.ejd_item_id and item_type_id = 8 ");

         sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ");

         if ( !isAceCust )
            sql.append("         ejd_item_warehouse.warehouse_id in (?, 11) and ejd_item_warehouse.in_catalog = 1 ");
         else
            sql.append("         ejd_item_warehouse.warehouse_id = ? and ejd_item_warehouse.in_catalog = 1 ");

         sql.append("order by upc_code, item_id");

         setCurAction("adding upc records");
         log.info("[mobile db] adding upc records");
         
         ostmt.close();
         ostmt = m_EdbConn.prepareStatement(sql.toString());
         ostmt.setInt(1, m_WhsId);
         rs = ostmt.executeQuery();

         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            i++;
            ustmt3.setString(1, rs.getString(1));
            ustmt3.setString(2, rs.getString(2));
            ustmt3.setInt(3, rs.getInt(3));

            ustmt3.addBatch();

            if (i == 200) {
               ustmt3.executeBatch();               
               i = 0;
            }
         }

         if ( i > 0 )
            ustmt3.executeBatch();

         //
         // Update the version table with the latest db information.
         if ( m_Status == RptServer.RUNNING )
            updateDbVer();

         result = true;
      }

      catch (Exception ex) {
         log.error("[mobile db] ", ex);
      }

      finally {
         DbUtils.closeDbConn(sc, ustmt1, null);
         DbUtils.closeDbConn(null, ustmt2, null);
         DbUtils.closeDbConn(null, ustmt3, null);
         sc = null;
         ustmt1 = null;
         ustmt2 = null;
         ustmt3 = null;
         sstmt = null;

         DbUtils.closeDbConn(null, ostmt, rs);
         ostmt = null;
         sql = null;
         rs = null;
      }

      return result;
   }

   /**
    * Builds the sqlite database for the BinLabel Application
    *
    * @return result
    */
   public boolean buildBinLabelDb()
   {
      Connection sc = null;
      PreparedStatement ostmt = null;
      Statement sstmt = null;
      PreparedStatement ustmt1 = null;
      ResultSet rs = null;
      PreparedStatement custStmt = null;
      ResultSet custRs = null;
      StringBuffer sql = new StringBuffer();
      int i = 0;
      int j = 0;
      boolean result = false;
      String customerId = null;

      try {
         // Get the customer id using the username parameter
         custStmt = m_PgConn.prepareStatement(
                  "select customer_id from web_partner "
                           + "join web_user on WEB_USER.PARTNER_ID = web_partner.partner_id "
                           + "where WEB_USER.USER_NAME = ? "
                  );

         custStmt.setString(1, m_Uid);
         custRs = custStmt.executeQuery();


         // if the web user is a customer, get the customer id
         // otherwise, return false
         if ( custRs.next() ) {
            customerId = custRs.getString(1);

            if ( customerId != null ) {
               setCurAction("Registering db driver");

               DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
               DriverManager.registerDriver(new org.sqlite.JDBC());

               sc = DriverManager.getConnection(String.format("jdbc:sqlite:%s", m_FilePath + m_DbName));
               sc.setAutoCommit(false);

               //
               // Create the item table
               sstmt = sc.createStatement();
               sstmt.execute(
                        "create table item ( " +
                                 "item_id text, upc text, vendor_item_num text, description text, " +
                                 "flc_id text, nrha_id text, stock_pack integer, ship_unit text, nbc text, " +
                                 "cost numeric, cost_code text, retail numeric, vendor text, cust_sku text, "  +
                                 "PRIMARY KEY (item_id, upc, stock_pack )" +
                                 ")"
                        );
               sstmt.execute("create index idx_item on item (item_id)");
               sstmt.execute("create index idx_upc on item (upc)");
               sstmt.execute("create index idx_desc on item (description)");
               sstmt.execute("create index idx_cust_sku on item (cust_sku)");
               sc.commit();

               DbUtils.closeDbConn(null, sstmt, null);

               sql.append("select distinct");
               sql.append("   item_entity_attr.item_id, ");
               sql.append("   decode(iUPC.upc_code, null, '000000000000', iUPC.upc_code) as UPC, ");
               sql.append("   vendor_item_ea_cross.vendor_item_num, ");
               sql.append("   item_entity_attr.description, ");
               sql.append("   ejd_item.flc_id, ");
               sql.append("   ejd.flc_procs.getnrhaid(flc_id) NRHA_ID, ");
               sql.append("   eiw.stock_pack as STOCK_PACK, ");
               sql.append("   ship_unit.unit as STOCK_UNIT, ");
               sql.append("   decode(ejd_item.broken_case_id, 1, '', 'NBC') as NBC, ");
               sql.append("   ejd.cust_procs.getsellprice(?, item_entity_attr.item_id) as COST, ");
               sql.append("   ejd.service_procs.Code_Price(?, ejd.cust_procs.GetSellPrice(?, item_entity_attr.item_id)) as COST_CODE, ");
               sql.append("   ejd.cust_procs.GetRetailPrice(?, item_entity_attr.item_id) as RETAIL, ");
               sql.append("   vend.name as VENDOR_NAME, ");
               sql.append("   ejd.cust_procs.getsku(?, item_entity_attr.item_id) as CUSTOMER_SKU, ");
               sql.append("   item_entity_attr.item_ea_id ");
               sql.append("from ");
               sql.append("   item_entity_attr ");
               sql.append("join ejd_item_warehouse eiw on eiw.ejd_item_id = item_entity_attr.ejd_item_id ");
               sql.append("join cust_warehouse on cust_warehouse.warehouse_id = eiw.warehouse_id and ");
               sql.append("      cust_warehouse.customer_id = ? and whs_priority = 1 ");
               sql.append("join (select vendor_id, name from vendor) vend on item_entity_attr.vendor_id = vend.vendor_id ");
               sql.append("left outer join ( ");
               sql.append("   select ejd_item_id, upc_code , warehouse_id ");
               sql.append("   from ejd_item_whs_upc ");
               sql.append("   where primary_upc = 1 ");
               sql.append(") iUPC on item_entity_attr.ejd_item_id = iUPC.ejd_item_id and iUPC.warehouse_id = eiw.warehouse_id ");
               sql.append("left outer join vendor_item_ea_cross on item_entity_attr.item_ea_id = vendor_item_ea_cross.item_ea_id and ");
               sql.append("   item_entity_attr.vendor_id = vendor_item_ea_cross.vendor_id ");
               sql.append("join ship_unit on item_entity_attr.ship_unit_id = ship_unit.unit_id ");
               sql.append("join item_type on item_type.item_type_id = item_entity_attr.item_type_id and itemtype in ('STOCK', 'VIRTUAL') ");
               sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
               sql.append("where ");
               sql.append("   eiw.disp_id = 1 and item_entity_attr.item_type_id in (1, 2) and ");
               sql.append("   ejd_item.flc_id <> 9998 and ");
               sql.append("   exists( ");
               sql.append("      select * ");
               sql.append("      from ejd_item_price ");
               sql.append("      where ejd_item_id = item_entity_attr.ejd_item_id and ");
               sql.append("         approved_by is not null ");
               sql.append("   ) ");

               System.out.println("processing item records");
               setCurAction("processing item records");

               ostmt = m_EdbConn.prepareStatement(sql.toString());
               ostmt.setString(1, customerId);
               ostmt.setString(2, customerId);
               ostmt.setString(3, customerId);
               ostmt.setString(4, customerId);
               ostmt.setString(5, customerId);
               ostmt.setString(6, customerId);

               rs = ostmt.executeQuery();
               setCurAction("inserting item data records");

               //
               // Insert into the sqlite db
               ustmt1 = sc.prepareStatement("insert into item values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?)");

               while ( rs.next() ) {
                  i++;
                  ustmt1.setString(1, rs.getString(1));     // item_id
                  ustmt1.setString(2, rs.getString(2));     // upc
                  ustmt1.setString(3, rs.getString(3));     // vnd item
                  ustmt1.setString(4, rs.getString(4));     // desc
                  ustmt1.setString(5, rs.getString(5));     // flc id
                  ustmt1.setString(6, rs.getString(6));     // nrha id
                  ustmt1.setInt(7, rs.getInt(7));           // stock pack
                  ustmt1.setString(8, rs.getString(8));     // unit
                  ustmt1.setString(9, rs.getString(9));     // nbc
                  ustmt1.setDouble(10, rs.getDouble(10));   // cost
                  ustmt1.setString(11, rs.getString(11));   // cost code
                  ustmt1.setDouble(12, rs.getDouble(12));   // retail
                  ustmt1.setString(13, rs.getString(13));   // vnd name
                  ustmt1.setString(14, rs.getString(14));   // cust sku
                  ustmt1.setInt(15, rs.getInt(15));   // item_ea_id

                  ustmt1.addBatch();

                  if ( i == 100 ) {
                     ustmt1.executeBatch();
                     sc.commit();
                     j+=i;
                     i = 0;
                  }
               }

               //
               // Make sure we get the last batch of records.
               if ( i > 0 ) {
                  ustmt1.executeBatch();
                  sc.commit();
               }

               setCurAction(String.format("processed %d records", j+i));

               //
               // Update the version table with the latest db information.
               if (m_Status == RptServer.RUNNING) {
                  updateDbVer();
               }

               result = true;
            }
         }
      }

      catch( Exception ex ) {
         log.error("[mobile db]", ex);
      }

      finally {
         DbUtils.closeDbConn(sc, ustmt1, null);
         DbUtils.closeDbConn(null, custStmt, custRs);
         sc = null;
         ustmt1 = null;
         sstmt = null;
         custRs = null;
         custStmt = null;

         DbUtils.closeDbConn(null, ostmt, rs);
         ostmt = null;
         sql = null;
         rs = null;
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements() {
      // closeStmt(m_ItemData);
   }

   @Override
   public boolean createReport() {
      boolean created = false;
      boolean okToProcess = false;
      m_Status = RptServer.RUNNING;

      try {
         if (m_AccessKey != null && m_AccessKey.length() > 0)
            okToProcess = true;
         else
            log.fatal("[mobile db] Missing access key, unable to process mobile db request");

         if ( okToProcess ) {
            if ( m_AppId > -1 )
               okToProcess = true;
            else
               log.fatal("[mobile db] Missing application id, unable to process mobile db request");
         }

         if ( okToProcess ) {
            m_EdbConn = m_RptProc.getEdbConn();

            switch ( m_AppId ) {               
               //
               // Emery Orders Application
               case 0: { 
                  created = buildEmeryLinkDb();
                  break;
               }

               //
               // BinLabel Application
               case 1: {
                  created = buildBinLabelDb();
                  break;
               }

               //
               // Competitor Conversion
               case 2: {
                  created = buildCompConvertDb();
                  break;
               }

               default: {
                  log.warn("[mobile db] invalid application id; mobile db request not processed");
               }
            }
         }
      }

      catch (Exception ex) {
         log.fatal("[mobile db]", ex);
      }

      finally {
         closeStatements();

         if (m_Status == RptServer.RUNNING)
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Sets the parameters of this report.
    *
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params) {
      StringBuffer fileName = new StringBuffer();
      int pcount = params.size();
      int j = 0;
      Param param = null;

      for (int i = 0; i < pcount; i++) {
         param = params.get(i);

         if (param.name.equals("dbname"))
            m_DbName = param.value;

         if (param.name.equals("dbpath"))
            m_DbPath = param.value;

         if (param.name.equals("uid"))
            m_Uid = param.value;

         if (param.name.equals("pwd"))
            m_Pwd = param.value;

         if (param.name.equals("accesskey"))
            m_AccessKey = param.value;

         if (param.name.equals("devid"))
            m_DeviceId = param.value;

         if (param.name.equals("appid"))
            m_AppId = Integer.parseInt(param.value);

         if (param.name.equals("whsid"))
            m_WhsId = Integer.parseInt(param.value);

         //
         // Variable parameter list processing. Only add the parameters to
         // the list and let each process handle them.
         if (param.name.equals("p" + j)) {
            m_ParamList.add(param.value);
            j++;
         }
      }

      fileName.append(m_DbName);
      m_FileNames.add(fileName.toString());
   }

   /**
    * Updates the internal database with the latest version information.
    *
    * @param conn The edb connection
    * @param data The array of data values. Results will be updated.
    * @param key The access key used as the owner of the database version.
    */
   private void updateDbVer() 
   {
      StringBuffer sql = new StringBuffer();
      PreparedStatement stmt = null;

      if ( m_EdbConn != null ) {
         try {
            sql.append("insert into b2b_mobile_data(name, location, userid, pwd, owner, device_id, app_id, warehouse_id) ");
            sql.append("values(?,?,?,?,?,?,?,?)");
            stmt = m_EdbConn.prepareStatement(sql.toString());

            stmt.setString(1, m_DbName);
            stmt.setString(2, "ftp.emeryonline.com:data");

            // TODO Set uid/pwd based on report parameters
            stmt.setString(3, "mobile");
            stmt.setString(4, "DJfTzlVW");

            stmt.setString(5, m_AccessKey);
            stmt.setString(6, m_DeviceId);
            stmt.setInt(7, m_AppId);
            stmt.setInt(8, m_WhsId);
            stmt.executeUpdate();
         }

         catch (Exception ex) {
            log.error("[mobile db]", ex);
         }

         finally {
            DbUtils.closeDbConn(null, stmt, null);
            stmt = null;
            sql = null;
         }
      }
   }
}
