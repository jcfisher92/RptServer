/**
 * File: SkuQty.java
 * Description: Export routine for getting the sku quantities.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 04/29/2009
 * Last Update: $Id: SkuQty.java,v 1.4 2014/02/14 17:37:20 jfisher Exp $
 *
 * History
 *    $Log: SkuQty.java,v $
 *    Revision 1.4  2014/02/14 17:37:20  jfisher
 *    removed discontinued items
 *
 *    Revision 1.3  2013/12/03 14:14:26  jfisher
 *    Tweaked the query to only pull one warehouse if a customer is assinged to two.
 *
 *    Revision 1.2  2013/12/03 14:07:52  jfisher
 *    Fixed the query to pull the data from the item_qty_view and use the warehouse the customer is assigned to.
 *
 *    Revision 1.1  2009/06/23 13:21:53  jfisher
 *    Initial Add - production version
 *
 *
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class SkuQty extends Report
{
   private PreparedStatement m_ItemData;
   private PreparedStatement m_WhsList;

   private String m_CustId;
   private String m_Dc;
   private boolean m_Overwrite;
   private boolean m_IsRscCust;
   private int m_StockWhs;
   private int m_Rsc;

   /**
    * default constructor
    */
   public SkuQty()
   {
      super();

      m_MaxRunTime = RptServer.HOUR * 12;
      m_CustId = "";
      m_Dc = "01";
      m_Overwrite = false;      
      m_IsRscCust = false;
      m_StockWhs = 0;
      m_Rsc = 0;
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      m_ItemData = null;
      m_WhsList = null;
      m_Dc = null;
      m_CustId = null;

      super.finalize();
   }

   /**
    * Builds the output file based on the query selection criteria.
    *
    * @return  boolean
    *    true if the file was created.
    *    false if there was some sort of error.
    */
   private boolean buildOutputFile()
   {
      StringBuffer line = new StringBuffer();
      FileOutputStream outFile = null;
      ResultSet itemData = null;
      String itemId;
      String itemType;
      int qty;
      boolean result = false;

      try {
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
         
         //
         // Generate the correct warehouse IDs based on customer warehouse and priority.
         getWhsList();         
         itemType = m_IsRscCust ? "ACE" : "EXPANDED ASST";
         
         m_ItemData.setInt(1, m_StockWhs);  // Note, this may be 0 if this is an RSC only customer.
         m_ItemData.setString(2, itemType);
         m_ItemData.setInt(3, m_Rsc);         
         itemData = m_ItemData.executeQuery();

         while ( itemData.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            itemId = itemData.getString(1);
            qty = itemData.getInt(2);
            itemType = itemData.getString(3);

            setCurAction("processing item: " + itemId);
            line.append(String.format("%s, %d, %s\r\n", itemId, qty, itemType.equals("EXPANDED ASST") ? "ACE" : itemType));
            outFile.write(line.toString().getBytes());
         }

         result = true;
      }

      catch ( Exception ex ) {
         log.error("[SkuQty]", ex);
      }

      finally {
         line = null;

         if ( outFile != null ) {
            try {
               outFile.close();
               outFile = null;
            }

            catch ( IOException ex ) {
               log.error("[SkuQty]", ex);
            }
         }
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_ItemData);
      closeStmt(m_WhsList);
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
         log.fatal("[SkuQty]", ex);
      }

      finally {
         closeStatements();
         
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Overrides the base class method to return the customer id the extract was run for.
    * @see com.emerywaterhouse.rpt.server.Report#getCustId()
    */
   @Override
   public String getCustId()
   {
      return m_CustId;
   }

   private void getWhsList() throws SQLException
   {      
      ResultSet rs = null;
      int count = 0;
      int whsId = 1;
      int whsPriority;
      
      m_WhsList.setString(1, m_CustId);      
      rs = m_WhsList.executeQuery();
      
      while ( rs.next() ) {
         whsId = rs.getInt(1);
         whsPriority = rs.getInt(2);
         
         if ( whsId < 3 ) {
            m_StockWhs = whsId;
            
            if ( whsPriority == 1 )
               m_IsRscCust = false;
         }
         else {
            m_Rsc = whsId;
          
            if ( whsPriority == 1 )
               m_IsRscCust = true;
         }
         
         count++;
      }
      
      //
      // If there is only just one cust warehouse record and it's not an RSC 
      // add in expanded assortment/wilton items.
      if ( count == 1 ) {
         if ( !m_IsRscCust || m_Rsc == 0 )
            m_Rsc = 11;
      }      
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(512);
      boolean isPrepared = false;
      String AceOrExp = "";

      if ( m_EdbConn != null ) {
         try {             
            sql.setLength(0);
            sql.append("select warehouse_id, whs_priority ");
            sql.append("from cust_warehouse ");
            sql.append("where customer_id = ? ");
            sql.append("order by whs_priority desc ");
            m_WhsList = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select item_id, qoh, item_type.itemtype ");
            sql.append("from item_entity_attr ");
            sql.append("join item_type on item_type.item_type_id = item_entity_attr.item_type_id and itemtype = 'STOCK' ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.warehouse_id = ? ");
            sql.append("join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id and disposition = 'BUY-SELL' ");
            sql.append("union ");
            sql.append("select item_id, qoh, item_type.itemtype ");
            sql.append("from item_entity_attr ");
            sql.append("join item_type on item_type.item_type_id = item_entity_attr.item_type_id and itemtype = ? ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.warehouse_id = ? ");
            sql.append("join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id and disposition = 'BUY-SELL' ");
            sql.append("order by item_id");
            
            m_ItemData = m_EdbConn.prepareStatement(sql.toString());
            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("[SkuQty]", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("[SkuQty] missing database connection");

      return isPrepared;
   }

   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    *
    * Note - The DC param is not used in the query any longer but has to be kept in the file name because
    *    we have b2b partners that pull the file and expect the name to be the same as it previously was. - jcf 12/03/2013
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      String fmt1 = "%d-%s-skuqty-%s.csv";
      String fmt2 = "%s-skuqty-%s.csv";
      StringBuffer fileName = new StringBuffer();
      int pcount = params.size();
      Param param = null;

      try {
         for (int i = 0; i < pcount; i++) {
            param = params.get(i);

            if (param.name.equals("dc"))
               m_Dc = param.value;

            if (param.name.equals("cust"))
               m_CustId = param.value;

            if ( param.name.equals("overwrite") )
               m_Overwrite = param.value.equalsIgnoreCase("true") ? true : false;
         }

         //
         // Some customers want the same file name each time.  If that's the case, we
         // need to overwrite what we have.
         if ( m_Overwrite )
            fileName.append(String.format(fmt2, m_CustId, m_Dc));
         else
            fileName.append(String.format(fmt1, System.currentTimeMillis(), m_CustId, m_Dc));

         m_FileNames.add(fileName.toString());
      }

      finally {
         fileName = null;
         param = null;
         fmt1 = null;
         fmt2 = null;
      }
   }   
}
