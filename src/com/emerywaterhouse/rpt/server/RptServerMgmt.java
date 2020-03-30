/**
 * File: RptServerMgmt.java
 * Description: 
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 06/01/2005
 * Last Update: $Id: RptServerMgmt.java,v 1.4 2008/10/30 16:54:41 jfisher Exp $
 * 
 * History
 */
package com.emerywaterhouse.rpt.server;

import java.io.BufferedReader;
import java.io.FileReader;


public class RptServerMgmt implements RptServerMgmtMBean
{   
   private RptServer m_Server;
   
   /**
    * default constructor
    */
   public RptServerMgmt()
   {
      super();      
   }

   /**
    * Creates an instance of the RptServerMgmt bean.  Sets a reference to the report server.
    * 
    * @param server - Reference to the currently running report server.
    */
   public RptServerMgmt(RptServer server)
   {
      m_Server = server;
   }
   
   /**
    * Cleanup allocated resources.
    * 
    * @see java.lang.Object#finalize()
    */
   public void finalize() throws Throwable
   {      
      m_Server = null;
      
      super.finalize();
   }
   
   /**
    * @see com.emerywaterhouse.rpt.server.RptServerMgmtMBean#getDescription()
    */
   public String getDescription()
   {
      return "Report server managment bean";
   }

   /**
    * @see com.emerywaterhouse.rpt.server.RptServerMgmtMBean#getName()
    */
   public String getName()
   {
      return getClass().getName();
   }
   
   /**
    * @see com.emerywaterhouse.rpt.server.RptServerMgmtMBean#stopServer()
    */
   public void stopServer()
   {      
      if ( m_Server != null )
         m_Server.shutdown(RptServer.STOPPED);
   }

   /**
    * Stops a specified report.
    * 
    * @param rptId The internal ID of the report that is to be stopped.
    */
   public void stopReport(long rptId)
   {
      if ( m_Server != null )
         m_Server.stopReport(rptId);
   }
   
   /**
    * Displays the banned list of email addresses.
    * @return The list of banned email addresses
    */
   public String viewBannedList()
   {
      StringBuffer html = new StringBuffer(1024);
      String line = null;
      BufferedReader fr = null;
      
      //
      // Setup the banner
      html.append("<div style=\"height: 22px; margin: 0; padding-top: 6px;");
      html.append("padding-left: 8px; padding-bottom: 4px;");
      html.append("background-color: #4682b4;");
      html.append("vertical-align: middle; text-align: left;");
      html.append("font-family: Arial, Helvetica, sans-serif; font-weight: bold; font-size: 13px;");
      html.append("color: white\"> Emery &#124; Waterhouse &nbsp; &nbsp; Banned Email List</div>\r\n");
      
      //
      // Setup the div for the body of text.
      html.append("<div style=\"font-family: Arial, Helvetica, sans-serif; font-size: 12px; text-align: left; vertical-align: top\">\r\n");
      html.append("<p><br>\r\n");
      
      try {
         html.append("<strong>List Of Banned Email Address</strong><br>\r\n");
         fr = new BufferedReader(new FileReader("bemail.lst"));
         line = fr.readLine();
         
         while ( line != null ) {         
            html.append(line);
            line = fr.readLine();
         }
         
         html.append("</div>");
      }
      
      catch ( Exception ex ) {
         RptServer.log.error("exception", ex);
      }
      
      
      return html.toString();
   }
   
   /**
    * Gets the status information from the report server.
    * @see com.emerywaterhouse.rpt.server.RptServerMgmtMBean#viewStatusInfo()
    */
   public String viewStatusInfo()
   {
      String info = null;
      
      if ( m_Server != null )
        info =  m_Server.getStatusInfo();
      else
         info = "no server available";
      
      return info;
   }
}
