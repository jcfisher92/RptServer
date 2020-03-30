/**
 * File: Report.java
 * Description: This is the report base class. All reports that run through the server
 *    must descend from this class.  It contains a run method that accepts the parameters used
 *    to create the report.  Eac report class determines how to use the parameters.
 *
 * @author Jeffrey Fisher
 *
 * Create Data: 03/31/2005
 * Last Update: $Id: Report.java,v 1.16 2010/01/07 15:07:02 jfisher Exp $
 * 
 * History
 */
package com.emerywaterhouse.rpt.server;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

import org.apache.log4j.Logger;

import com.emerywaterhouse.websvc.Param;


public abstract class Report
{
   protected static final String ACCESS_KEY = "dfc87f3309372eae7fa04ea4f324627ba0639ea2e127ece31987b98d";
   
   protected String m_CurAction;             // Records the current processing taking place.
   protected StringBuffer m_ErrMsg;          // An error message buffer that gets sent back to the processor.
   protected ArrayList<String> m_FileNames;  // The name of the file on disk.
   protected String m_FilePath;              // The path to the report
   protected RptProcessor m_RptProc;         // A reference to the report processing object managing the report.
   protected long m_MaxRunTime;              // The maximum time in milliseconds it will take to run before an warning.
   protected short m_Status;                 // The thread run status.  One of the above constants.
   protected long m_StartTime;               // The time in milliseconds when the report started.   
   
   protected Connection m_EdbConn;            // The connection to postgres
   protected Connection m_OraConn;            // The connection to Oracle, which should be rarely used.
   protected Connection m_FasConn;           // The connection to fascor.
   protected Connection m_PgConn;            // The connection to postgres
   protected Connection m_SageConn;          // The connection to Sage300/AccPac
   
   //
   // Log4j logger
   protected static Logger log = Logger.getLogger(Report.class);
   
   /**    
    * default constructor
    */
   public Report()
   {
      super();
      
      m_CurAction = "initializing";      
      m_ErrMsg = new StringBuffer();
      m_MaxRunTime = RptServer.HOUR * 3;
      m_StartTime = System.currentTimeMillis();
      m_Status = RptServer.INIT;
      m_FilePath = System.getProperty("rpt.dir", "/reports/");      
            
      m_FileNames = new ArrayList<String>();
   }
   
   /**
    * Clean up.
    * Note - the connection to the database is controlled in the report processor
    *    class.
    * 
    * @see java.lang.Object#finalize()
    */
   public void finalize() throws Throwable
   {
      m_ErrMsg = null;
      m_RptProc = null;      
      m_FilePath = null;
      m_CurAction = null;
      
      m_EdbConn = null;
      m_OraConn = null;
      m_FasConn = null;
      m_PgConn = null;
      m_SageConn = null;
      
      if ( m_FileNames != null ) {
         m_FileNames.clear();
         m_FileNames = null;
      }
      
      super.finalize();
   }
   
   /**
    * A convenience method for reports to close a resultset without having to 
    * deal with exceptions.
    * 
    * @param rset A reference to a resultset that needs to be closed.
    */
   protected void closeRSet(ResultSet rset)
   {
      if ( rset != null ) {
         try {
            rset.close();
         }

         catch ( Exception ex ) {
            
         }
      }
   }
   
   /**
    * Convenience method for closing a SQL statement.
    * 
    * @param stmt The stateme to close.
    */
   protected void closeStmt(Statement stmt)
   {
      if ( stmt != null ) {
         try {
            stmt.close();
         }

         catch ( Exception ex ) {
            
         }
      }
   }
      
   /**
    * This is the method that all descendant classes must implement.  This actually
    * starts the report creation.
    *  
    * @return boolean.  True if the report was successfully created, false if not.
    */
   public abstract boolean createReport();
   
   /**
    * Logs a fatal exception
    * @param ex The exception to log.
    */
   public void error(Exception ex)
   {
      if ( ex != null )
         log.error("exception: ", ex);
   }
   
   /**
    * Logs a fatal exception
    * @param ex The exception to log.
    */
   public void fatal(Exception ex)
   {
      if ( ex != null )
         log.fatal("fatal exception: ", ex);
   }
   
   /**
    * @return Returns the curAction.
    */
   public String getCurAction()
   {
      return m_CurAction;
   }
   
   /**
    * Helper function used by the report processor for logging.  This needs to be
    * overridden in the descendant classes.  If a report has a customer id param it should
    * be returned in this method.
    * 
    * @return A customer id if the report has one.
    */
   public String getCustId()
   {
      return "";
   }
   
   /**
    * Returns an error message if the report failed.
    * 
    * @return The failure message from report processing.
    */
   public final String getErrMsg()
   {
      return m_ErrMsg.toString();
   }
   
   /**
    * Returns the number of files associated with this report.  It's the number
    * of entries in the list.
    * 
    * @return The number of files.
    */
   public final int getFileCount()
   {
      return m_FileNames.size();
   }
   
   /**
    * Returns the name of the report file as it is stored on disk.
    * 
    * @param index The filename to get at the specific index.
    * @return The report file name.
    */
   public final String getFileName(int index)
   {
      return m_FileNames.get(index);
   }
   
   /**
    * Returns the path to the report file.
    * 
    * @return The filepath to the report.
    */
   public final String getFilePath()
   {
      return m_FilePath;
   }
   
   /**
    * @return The maximum time, in milliseconds, that the report can run before it
    *    has exceeded the expected time to finish.
    */
   public final long getMaxRunTime()
   {
      return m_MaxRunTime;
   }
   
   /**
    * @return Returns the startTime.
    */
   public final long getStartTime()
   {
      return m_StartTime;
   }
   
   /**
    * Gets the current status of the report.
    * @return The report status.
    */
   public final short getStatus()
   {
      return m_Status;
   }
   
   /**
    * Sets the current action for the report.  This is used to trasmit progress
    * information to clients.
    * 
    * @param action The action to set.
    */
   public void setCurAction(String action)
   {
      if ( action != null )
         m_CurAction = action;
   }
   
   /**
    * Allows another process to set the error message text for this report.
    * @param msg An error message to set.
    */
   public void setErrMsgText(String msg)
   {
      if ( msg != null )
         m_ErrMsg.append(msg);
   }
   
   /**
    * Empty base class method for setting the parameters of the report.  Descendant classes
    * need to override this method if there are any report parameters.  This gets called by the
    * RptProcessor class.
    * @param params
    */
   public void setParams(ArrayList<Param> params)
   {
      
   }
   
   /**
    * Sets the report's report processor object.
    * 
    * @param rptProc reference to set.
    */
   public void setRptProcessor(RptProcessor rptProc)
   {
      if ( rptProc != null ) {
         m_RptProc = rptProc;
         
         try {
            m_EdbConn = rptProc.getEdbConn();
         } 
         
         catch (SQLException e) {         
            ;
         }
      }
   }
   
   /**
    * Sets the report status
    * @param status The status to set on the report.
    */
   public void setStatus(short status)
   {
      m_Status = status;
   }
}
