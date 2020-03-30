/**
 * File: RptServerMgmtMBean.java
 * Description: MBean interface for managing the report server.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 06/01/2005
 * Last Update: $Id: RptServerMgmtMBean.java,v 1.3 2005/08/25 13:55:36 jfisher Exp $
 * 
 * History
 */
package com.emerywaterhouse.rpt.server;


public interface RptServerMgmtMBean
{
   public String getDescription();
   public String getName();
      
   /**
    * Stops the report server.    
    */
   public void stopServer();
   
   /**
    * Stops a specified report.
    * 
    * @param rptId The internal ID of the report that is to be stopped.
    */
   public void stopReport(long rptId);
   
   /**
    * Displays a list of banned email addresses
    * @return The banned email list stored on disk.
    */
   public String viewBannedList();
   
   /**
    * Shows the status of all of the currently running reports.
    * @return Status information about the server which includes running reports, etc.
    */
   public String viewStatusInfo();
}
