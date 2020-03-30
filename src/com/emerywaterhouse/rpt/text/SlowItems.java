/**
 * File: SlowItems.java
 * Description: Slow selling items report
 *    Rewrite to allow running in the new report server.
 *    Original author was Jacob Heric
 *
 * @author Jacob Heric
 * @author Jeffrey Fisher
 *
 * Create Date: 05/20/2005
 * Last Update: $Id: SlowItems.java,v 1.7 2009/12/10 22:58:18 smurdock Exp $
 * 
 * History
 *    03/25/2005 - Added log4j logging. jcf
 *    
 *    07/06/2004 - Altered report criteria (altered query). No flc restrictions, no disposition restrictions. JBH.
 *    
 *    05/03/2004 - Removed the usage of the m_DistList member variable.  This variable gets cleaned up before it can be
 *       used in the email webservice. - jcf
 *
 *    04/07/2004 - Applied Email class changes. - jcf
 *
 *    12/19/2003 - Changed the way the email param is retrieved and sent.  Also removed some unused imports and
 *       local variables. - jcf
 *
 *    03/19/2002 - Changed the way the connection to Oracle is established.  The program now uses the
 *       getOracleConn() method to retrieve a new connection object.  The connection pool is no longer
 *       used. Also added the cleanup method.- jcf
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;


public class SlowItems extends Report
{
   private PreparedStatement m_allItems;
   
   /**
    * default constructor
    */
   public SlowItems()
   {
      super();
      m_FileNames.add("SlowItem.dat");
   }

   /**
    * Executes the queries and builds the output file
    * @return boolean true if the file was created false if not. 
    * @throws FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {      
      StringBuffer Line = new StringBuffer(1024);
      FileOutputStream OutFile = null;
      ResultSet ItemData = null;
      String itemID;
      int qty = 0;
      int totQty = 0;
      double extCost = 0.0;
      double totExtCost = 0.0;
      boolean Result = false;
      String fileName = m_FilePath + m_FileNames.get(0);

      setCurAction("creating/opening output file: " + fileName);
      OutFile = new FileOutputStream(fileName, false);

      //date formatter from java.utils
      SimpleDateFormat formatter = new SimpleDateFormat("MM/dd/yyyy");

      //
      // Build the Captions
      setCurAction("inserting header ");
      Line.append("Item Number\tVendor\tItem Desc.\tItem Disp.\tSetup Date\tQty. On Hand\tCost\tExt. Cost\t");
      Line.append("Buyer Number\tBuyer Name\r\n");

      try {
         setCurAction("executing report query ");
         ItemData = m_allItems.executeQuery();
         setCurAction("report query executed  ");

         while ( ItemData.next() && m_Status == RptServer.RUNNING ) {
            itemID = ItemData.getString("item_id");
            setCurAction("inserting line for item " + itemID);
            Line.append(itemID + "\t");
            Line.append(ItemData.getString("name") + "\t");
            Line.append(ItemData.getString("description") + "\t");
            Line.append(ItemData.getString("disposition") + "\t");
            Line.append(formatter.format(ItemData.getDate("setup_date")).toString() + "\t");

            qty = ItemData.getInt("Actual_Qty");
            totQty = totQty + qty;
            Line.append(qty + "\t");

            Line.append(ItemData.getDouble("emery_cost") + "\t");

            extCost = ItemData.getDouble("extended_cost");
            totExtCost = totExtCost + extCost;
            Line.append(extCost + "\t");

            Line.append(ItemData.getString("buyer_num") + "\t");
            Line.append(ItemData.getString("buyer_name") + "\t");

            Line.append("\r\n");
            OutFile.write(Line.toString().getBytes());
            Line.delete(0, Line.length());
            setCurAction("finished line for item " + itemID);

         }

         setCurAction("adding summary line (total cost and extended cost)");
         Line.append("\t\t\t\t\tTotal Qty.\t\tTotal Ext. Cost\r\n");
         Line.append("\t\t\t\t\t" + totQty + "\t\t" + totExtCost + "\r\n");
         OutFile.write(Line.toString().getBytes());
         Line.delete(0, Line.length());

         setCurAction("closing recordset");
         ItemData.close();
         Result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         
         log.error("exception: ", ex);
      }

      finally {
         Line = null;

         try {
            setCurAction("closing output file: " + fileName);
            OutFile.close();
         }

         catch( Exception e ) {
            log.error(e);
         }

         OutFile = null;
      }

      return Result;
   }

   /**
    * The SQL for Slow Items must query the monthlyitemsales table
    * for items not sold in the last six months, we could do that
    * easily using Oracle SQL date functions, i.e.:
    *   select item_nbr from monthlyitemsales m
    *   where to_date(m.year_month, 'yyyymm') >= last_day(add_months(sysdate, -6)))
    *   
    * but using these functions prevents the use of indexes, which
    * is essential for tables as big as monthlyitemsales.
    * So, we are building the six month condition manually.
    * p.s I ripped this algorithm off from j. fisher 
    */
   private String buildSalesSQL(int month, int year) throws Exception
   {
      final String Months[] = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"};
      int i;

      StringBuffer SQL = new StringBuffer();

      if ( month < 1 || month > 12 )
         throw new Exception("invalid month parameter");

      SQL.append("select item.item_id, v.name, item.description, id.disposition, ");
      SQL.append("item.setup_date, la.Actual_Qty, item_price_procs.TODAYS_BUY(item.item_id) emery_cost, ");
      SQL.append("(la.Actual_Qty * item_price_procs.TODAYS_BUY(item.item_id)) extended_cost, ");
      SQL.append("vb.dept_num buyer_num, vb.buyer_name ");
      SQL.append("from item_disp id, vendor v, vendor_buyers vb, item, " );
    // SQL.append("(select sum(\"Actual_Qty\") Actual_Qty, \"SKU\" SKU from loc_allocation group by \"SKU\" ");
     //SQL.append("having sum(\"Actual_Qty\") > 0 ) la ");
//    beginning of new loc_allocation stuff
      SQL.append("(select sku, sum(actual_qty) actual_qty from ");
      SQL.append("((select sum(\"Actual_Qty\") Actual_Qty, \"SKU\" SKU from loc_allocation@fas01 group by \"SKU\" ");
      SQL.append("having sum(\"Actual_Qty\") > 0 ) ");
      SQL.append("union all ");
      SQL.append("(select sum(\"Actual_Qty\") Actual_Qty, \"SKU\" SKU from loc_allocation@fas04 group by \"SKU\"  ");
      SQL.append("having sum(\"Actual_Qty\") > 0 )) ");
      SQL.append("group by sku) la ");
//    end of new loc_allocation stuff     
      SQL.append("where not exists (select item_nbr from monthlyitemsales m ");
      SQL.append("where m.invoice_month in ( ");

      //build the string backwards
      //Notice:  this is a thirteen month condition because it includes the
      //current month
      for ( i = 0; i <= 12; i++ ) {
         if ( month == 0 ) {
            month = 12;
            year--;
         }
         //
         SQL.append("'" + Months[month-1] + "/" + Integer.toString(year) + "'");

         if ( i < 12 )
            SQL.append(",");

         month--;
      }

      SQL.append(" ) ");
      SQL.append("and m.item_nbr = item.item_id) and item.disp_id = id.disp_id and " );
      SQL.append("item.item_id = la.SKU and item.vendor_id = v.vendor_id and ");
      SQL.append("item.vendor_id = vb.vendor_id and ");
      SQL.append("trunc(item.setup_date) <= trunc(last_day(add_months(sysdate, - 13))) ");
      SQL.append("order by item.item_id ");      

      return SQL.toString();
   };

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not garunteed to know when finalization occurs.
    */
   protected void cleanup()
   {
      closeStatements();

      m_allItems = null;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      try {
         if ( m_allItems != null )
            m_allItems.close();
      }

      catch ( Exception ex ) {
         log.error(ex);
      }
   }
   
   /**
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
    * @return boolean True if the statement was prepared, false if not.
    */
   private boolean prepareStatements()
   {
      boolean isPrepared = false;
      
      if ( m_OraConn != null ) {
         //date formatter from java.utils
         SimpleDateFormat formatter = new SimpleDateFormat("MM/yyyy");
   
         //Get Month, make it an integer
         int Month = Integer.parseInt(formatter.format(new java.util.Date()).toString().substring(0, 2));
   
         //Get Year, make it an integer
         int Year = Integer.parseInt(formatter.format(new java.util.Date()).toString().substring(3));
   
         try {            
            m_allItems = m_OraConn.prepareStatement(buildSalesSQL(Month, Year));
            isPrepared = true;
         }
         
         catch ( Exception ex ) {
            log.fatal("exception:", ex);
         }
      }
      
      return isPrepared;
   }
}
