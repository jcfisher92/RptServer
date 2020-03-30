/**
 * File: ItemGroups.java
 * Description: Extracts the item group data.
 *    
 * @author Jeffrey Fisher
 * 
 * Create Date: 05/03/2006
 * Last Update: $Id: ItemGroups.java,v 1.2 2008/10/29 21:31:00 jfisher Exp $
 * 
 * History:
 *     
 */
package com.emerywaterhouse.rpt.inxpo;

import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;


public class ItemGroups extends Report
{
   private PreparedStatement m_ItemGroups;
   
   /**
    * Default constructor
    */
   public ItemGroups()
   {
      super();
      
      m_FileNames.add("itemgrp.txt");      
      m_MaxRunTime = RptServer.HOUR;
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    * @throws Throwable 
    */
   public void finalize() throws Throwable
   {      
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
      StringBuffer line = new StringBuffer(1024);   
      FileOutputStream outFile = null;
      boolean result = false;
      ResultSet itemGroup = null;
      
      try {
         setCurAction("creating/opening output file " + m_FilePath + m_FileNames.get(0));
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

         if ( m_Status == RptServer.RUNNING ) {
            setCurAction("starting item group export");
            itemGroup = m_ItemGroups.executeQuery();            
                  
            while ( itemGroup.next() && m_Status == RptServer.RUNNING ) {            
               line.append(itemGroup.getString(1) + "\t");     // distributor name
               line.append(itemGroup.getString(2) + "\t");     // description
               line.append(itemGroup.getString(3) + "\t");     // display order
               line.append(itemGroup.getString(4) + "\r\n");   // grouping code
               
               outFile.write(line.toString().getBytes());
               line.delete(0, line.length());            
            }
         
            setCurAction("finished exporting item groups");
            closeRSet(itemGroup);
            result = true;
         }
      }

      catch ( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }
      
      return result;
   }
   
   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    */
   protected void cleanup()
   {
      closeStmt(m_ItemGroups);
      
      m_ItemGroups = null;
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
    */
   private boolean prepareStatements() throws Exception
   {      
      StringBuffer sql = new StringBuffer();
      boolean prepared = false;

      if ( m_OraConn != null ) {
         sql.append("select distributer_name, description, display_order, grouping_code ");
         sql.append("from inxpo.inventory_item_grouping ");
         sql.append("order by display_order");         
         m_ItemGroups = m_OraConn.prepareStatement(sql.toString());
                  
         prepared = true;
      }
      
      return prepared;
   }
}
