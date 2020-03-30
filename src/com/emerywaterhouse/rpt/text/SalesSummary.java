/**
 * File: SalesSummary.java
 * Description: This invokes the monthly sales summary procedure on Oracle.
 *    Rewritten to work with the new report server.
 *    Original author was Jacob Heric
 *
 * @author Jacob Heric
 * @author Jeffrey Fisher
 *
 * Create Date: 05/19/2005
 * Last Update: $Id: SalesSummary.java,v 1.7 2008/10/30 16:49:56 jfisher Exp $
 * 
 * History
 *    $Log: SalesSummary.java,v $
 *    Revision 1.7  2008/10/30 16:49:56  jfisher
 *    Fixed issue with closing a statement
 *
 *    Revision 1.6  2008/10/30 16:47:33  jfisher
 *    Removed unused member var and added the log cvs tag
 *
 */
package com.emerywaterhouse.rpt.text;

import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.GregorianCalendar;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;


public class SalesSummary extends Report
{  
   /**
    * default constructor
    */
   public SalesSummary()
   {
      super();
   }
   
   /**
    * This method executes the Job on Oracle with a SQL statement
    * 
    * @return boolean, true if the report was successfully run, false if not.
    */
   public boolean sum()
   {
      Statement stmt = null;
      GregorianCalendar clndr = null;      
      SimpleDateFormat fmt = new SimpleDateFormat("MM/yyyy");
      boolean result = false;
      
      try{      
         //
         //get date last month
         clndr = new GregorianCalendar();
         clndr.setTimeInMillis(System.currentTimeMillis());
         clndr.set(Calendar.MONTH, clndr.get(Calendar.MONTH) - 1);

         if ( m_OraConn != null && m_Status == RptServer.RUNNING ) {
            stmt = m_OraConn.createStatement();
            
            try {
               setCurAction("Starting MonthEnd Sales Summary Job");
               stmt.execute("begin sa.monthly_sales_summary('" + fmt.format(clndr.getTime()) + "'); end;");
            }
            
            finally {
               DbUtils.closeDbConn(null, stmt, null);
            }
            
            setCurAction("Finished MonthEnd Sales Summary Job");             
            result = true;
         }
         else
            log.fatal("SalesSummary: null oracle connection");
      }
            
      catch (Exception e){
         log.error("exception:", e);
         m_ErrMsg.append("The job had the following Error: \r\n");
         m_ErrMsg.append(e.getClass().getName() + "\r\n" + e.getMessage());         
      }
      
      return result;
   }
   
   /**
    * Run the report.
    * 
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         m_OraConn = m_RptProc.getOraConn();
         created = sum();          
      }
      
      catch ( Exception ex ) {
         log.fatal("exception:", ex);         
      }
      
      finally {
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }
            
      return created;
   }

}
