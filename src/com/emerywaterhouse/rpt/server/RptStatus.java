/**
 * File: RptStatus.java
 * Description: Simple class for moving data elements from one location to another.  Used to hold the
 *    status information from a running report and report processor.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 06/29/2005
 * Last Update: $Id: RptStatus.java,v 1.3 2007/05/18 14:51:39 jfisher Exp $
 * 
 * History
 *    $Log: RptStatus.java,v $
 *    Revision 1.3  2007/05/18 14:51:39  jfisher
 *    added the log tag
 *
 */
package com.emerywaterhouse.rpt.server;


public class RptStatus
{
   public String currentAction;
   public long internalId;
   public String rptName;
   public long maxRunTime;
   public long runTime;
   public long startTime;
   public long threadId;
   public String uid;
   
   /**
    * default constructor
    */
   public RptStatus()
   {
      super();
      
      currentAction = "";
      rptName = "";
      uid = "";
   }

}
