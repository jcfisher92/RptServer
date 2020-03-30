/**
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
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

/**
 * @author JFisher
 *
 */
public class SkuQtyEx extends Report
{
   private PreparedStatement m_ItemData;
   private PreparedStatement m_CustWhs;

   private String m_CustId;
   private boolean m_Overwrite;
      
   /**
    * 
    */
   public SkuQtyEx() 
   {
      super();

      m_MaxRunTime = RptServer.HOUR * 12;
      m_CustId = "";
      m_Overwrite = false;
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
      String whs;
      int qty;
      boolean result = false;

      try {
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);       
         itemData = m_ItemData.executeQuery();

         while ( itemData.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            itemId = itemData.getString(1);
            qty = itemData.getInt(2);
            whs = itemData.getString(3);

            setCurAction("processing item: " + itemId);
            line.append(String.format("%s, %d, %s\r\n", itemId, qty, whs));
            
            outFile.write(line.toString().getBytes());
         }

         result = true;
      }

      catch ( Exception ex ) {
         log.error("[SkuQtyEx]", ex);
      }

      finally {
         line = null;

         if ( outFile != null ) {
            try {
               outFile.close();
               outFile = null;
            }

            catch ( IOException ex ) {
               log.error("[SkuQtyEx]", ex);
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
      closeStmt(m_CustWhs);
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
         log.fatal("[SkuQtyEx]", ex);
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
   
   /**
    * 
    */
   public String getWhsList()
   {
      //
      // Stub function until we have time to make the dynamic.
      return "'PORTLAND', 'WILTON'";
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(512);
      boolean isPrepared = false;

      if ( m_EdbConn != null ) {
         try {
            sql.append("select item_id, qoh, warehouse.name as warehouse ");
            sql.append("from ejd_item_warehouse ");
            sql.append("join warehouse on warehouse.warehouse_id = ejd_item_warehouse.warehouse_id ");            
            sql.append("join item_entity_attr on item_entity_attr.ejd_item_id = ejd_item_warehouse.ejd_item_id ");
            sql.append("join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id and disposition in ('BUY-SELL', 'NOBUY') ");
            sql.append("join item_type on item_type.item_type_id = item_entity_attr.item_type_id and itemtype in ('STOCK', 'ACE') ");                        
            sql.append("where warehouse.name in (").append(getWhsList()).append(") ");
            sql.append("order by warehouse, item_id");
            
            m_ItemData = m_EdbConn.prepareStatement(sql.toString());
            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("[SkuQtyEx]", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("[SkuQtyEx] missing database connection");

      return isPrepared;
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    *
    * Note - 
    */   
   public void setParams(ArrayList<Param> params)
   {
      String fmt1 = "%d-emery-skuqty.csv";
      StringBuffer fileName = new StringBuffer();
      int pcount = params.size();
      Param param = null;

      try {
         for (int i = 0; i < pcount; i++) {
            param = params.get(i);
            
            if (param.name.equals("cust"))
               m_CustId = param.value;

            if ( param.name.equals("overwrite") )
               m_Overwrite = param.value.equalsIgnoreCase("true") ? true : false;
         }

         //
         // Some customers want the same file name each time.  If that's the case, we
         // need to overwrite what we have.
         if ( !m_Overwrite )            
            fileName.append(String.format(fmt1, System.currentTimeMillis()));
         else
            fileName.append("emery-skuqty.csv");   

         m_FileNames.add(fileName.toString());
      }

      finally {
         fileName = null;
         param = null;
         fmt1 = null;
      }
   }
}
