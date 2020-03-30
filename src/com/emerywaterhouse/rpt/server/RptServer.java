/**
 * File: RptServer.java.
 * Description: This is the main class that acts as the report server.  It is very much like an
 *    adapter in that it takes report requests off of the bus and then processes them.
 *
 * @author Jeffrey Fisher
 *
 * Create Data: 03/29/2005
 * Last Update: $Id$
 *
 * History
 */
package com.emerywaterhouse.rpt.server;

import java.io.FileInputStream;
import java.net.MalformedURLException;
import java.util.Properties;

import javax.management.MBeanServer;
import javax.management.MBeanServerFactory;
import javax.management.ObjectName;

import org.apache.log4j.AsyncAppender;
import org.apache.log4j.Logger;
import org.apache.log4j.xml.DOMConfigurator;

import com.emerywaterhouse.loader.EmLoader;
import com.emerywaterhouse.utils.DataSender;
import com.sun.jdmk.comm.HtmlAdaptorServer;


public class RptServer implements Runnable
{
   public static final String dataProcessing = "dataprocessing@emeryonline.com";
   public static final String emailFrom      = "system@emeryonline.com";
   public static final String callPhone      = "2077492592@vzwpix.com";

   public static final String[] errorList = {
      dataProcessing,
      callPhone
   };

   public static final String[] testList = {"programming@emeryonline.com"};

   public final static short INIT    = 0;
   public final static short RUNNING = 1;
   public final static short STOPPED = 2;
   public final static short ABORT   = 3;
   public final static short FATAL   = 4;

   public static final int HOUR   = 3600000;
   public static final int MINUTE = 60000;
   public static final int SECOND = 1000;

   private static final String PROP_FILE = "server.properties";

   private RptMonitor m_RptMonitor;             // The report queue processor
   private short m_Status;                      // The current run status of the program
   private Thread m_Thread;                     // Internal thread.
   private MBeanServer m_MBServer;              // MBean server for report mgmt mbean
   private HtmlAdaptorServer m_HtmlServer;      // Suns html protocal adaptor for jmx.

   //
   // Static so it's only instantiated once.
   private static EmLoader m_Loader;            // Custom class loader.
   private static Environment m_Env;     /** Environment variable used to determine where to send the report */

   //
   // Log4j logger
   public static Logger log = Logger.getLogger(RptServer.class.getName());

   /**
    * Environment enumeration
    */
   public static enum Environment {
      Test,
      Production
   };

   //
   // Static initialization.  Forces log4j to be initialized before anything else.
   {
      initLog();
   }

   /**
    * default constructor
    */
   public RptServer()
   {
      super();

      m_Status = INIT;
      m_Thread = new Thread(this, "RptServer");
      m_Thread.setDaemon(true);

      m_RptMonitor = new RptMonitor();
      m_MBServer = MBeanServerFactory.createMBeanServer();
      m_HtmlServer = new HtmlAdaptorServer();   // default port is 8082

      if ( System.getProperty("server.mode", "test").equals("test") )
         m_Env = Environment.Test;
      else
         m_Env = Environment.Production;
   }

   /**
    * Clean up any allocated resources.
    *
    *  @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      m_Thread = null;
      m_RptMonitor = null;
      m_MBServer = null;
      m_HtmlServer = null;

      super.finalize();
   }

   /**
    * Gets the current environment the server is running in; test or production.
    * @return The Environment enum value.
    */
   public static Environment getEnv()
   {
      return m_Env;
   }

   /**
    * Gets the custom class loader.
    *
    * @return An instance of the EmLoader.
    * @throws MalformedURLException
    */
   public static synchronized EmLoader getLoader() throws MalformedURLException
   {
      if ( m_Loader == null ) {
         m_Loader = new EmLoader();
         m_Loader.addLocation("file:/");
         m_Loader.addLocation("file:/usr/local/rptserver/");
         m_Loader.addLocation("file:/usr/local/rptserver/com/emerywaterhouse/rpt/");
         m_Loader.addLocation("file:/usr/local/rptserver/com/emerywaterhouse/rpt/spreadsheet/");
         m_Loader.addLocation("file:/usr/local/rptserver/com/emerywaterhouse/rpt/alert/");
      }

      return m_Loader;
   }

   /**
    * Gets the status information from the server and running reports.
    * Note - Currently there is not a custom html adapter and the sun provided adapter uses an
    *    html dtd that is a little old.  Also there is no way to link in a style sheet.  This means
    *    we have to use some odd html.
    *
    * @return an html page containing the status data from the server, the list of currently
    *    running reports, and the report status.
    */
   public synchronized String getStatusInfo()
   {
      StringBuffer html = new StringBuffer(1024);

      //
      // Setup the banner
      html.append("<div style=\"height: 22px; margin: 0; padding-top: 6px;");
      html.append("padding-left: 8px; padding-bottom: 4px;");
      html.append("background-color: #4682b4;");
      html.append("vertical-align: middle; text-align: left;");
      html.append("font-family: Arial, Helvetica, sans-serif; font-weight: bold; font-size: 13px;");
      html.append("color: white\"> Emery &#124; Waterhouse &nbsp; &nbsp; Report Server Status Information</div>\r\n");

      //
      // Setup the div for the body of text.
      html.append("<div style=\"font-family: Arial, Helvetica, sans-serif; font-size: 12px; text-align: left; vertical-align: top\">\r\n");
      html.append("<p><br>\r\n");
      html.append("Maximum Concurrent Reports: ").append(m_RptMonitor.getMaxRptCount());
      html.append("<br><br>\r\n");

      //
      // The monitor has a list of running reports.  We let it handle the querying of status
      // info so that we don't run into thread locking issues.
      html.append(m_RptMonitor.getRptStatus());
      html.append("</div>");

      return html.toString();
   }

   /**
    * Initialize the logger.
    */
   private void initLog()
   {
      AsyncAppender asyncAppender = null;

      try {
         loadProperties();
         DOMConfigurator.configure("logcfg.xml");
         //log.setAdditivity(false);
         asyncAppender = (AsyncAppender)Logger.getRootLogger().getAppender("ASYNC");
         asyncAppender.setBufferSize(15);
         asyncAppender.setLocationInfo(true);
      }

      catch( Exception ex ) {

      }
   }

   /**
    * Loads the mbeans for the report server and sets any attributes on those beans that need to
    * be set.  The mbeans are used for managing the report server.
    *
    * @throws Exception if there are problems with the mbean loading
    */
   private void loadMBeans() throws Exception
   {
      ObjectName objName = null;
      RptServerMgmt mgtBean = new RptServerMgmt(this);

      //
      // Load up the rpt server management class and set the server reference
      // so the bean can call server methods.
      objName = new ObjectName("RptServerMgmt:name=ServerMgmt");
      m_MBServer.registerMBean(mgtBean, objName);

      objName = new ObjectName("HtmlAdaptorServer:name=HtmlAdaptor");
      m_MBServer.registerMBean(m_HtmlServer, objName);
   }

   /**
    * Loads the system properties for the server.
    */
   private void loadProperties()
   {
      String propFileName = null;
      FileInputStream propFile = null;
      Properties p = new Properties(System.getProperties());
      String sepChar = p.getProperty("file.separator");
      String appDir = p.getProperty("user.dir");

      try {
         //
         // Set the path to the properties file for the application.
         propFileName = appDir + sepChar + PROP_FILE;
         propFile = new FileInputStream(propFileName);
         p.load(propFile);
         System.setProperties(p);
      }

      catch ( Exception ex ) {
         ;
      }

      finally {
         if ( propFile != null ) {
            try {
               propFile.close();
               propFile = null;
            }

            catch ( Exception ex ) {

            }
         }

         p = null;
         appDir = null;
         sepChar = null;
         propFileName = null;
      }
   }

   /**
    * Sends an email notification to the MIS department.
    *
    * @param msg - The email message to send.
    */
   public static void notifyMis(String msg)
   {
      String[] recips = null;
      String subj = "System Notification";

      if ( msg != null ) {
         switch ( m_Env ) {
            case Test: {
               subj = "[TEST] " + subj;
               recips = testList;
               break;
            }

            case Production: {
               recips = errorList;
               break;
            }
         }

         try {
            DataSender.smtp(emailFrom, recips, subj, String.format("Report Server \r\n %s", msg));
         }

         catch (Exception ex ) {
            log.error("exception: ", ex);
         }

         finally {
           recips = null;
           subj = null;
         }
      }
   }

   /**
    * Implements the runnable interface.  This allows the program to monitor when to shut down
    * Independent of the other processes that are running.
    */
   @Override
   public void run()
   {
      try {
         log.info("report server started");

         //
         // Process incoming commands and send out status updates.
         while ( m_Status == RUNNING ) {
            Thread.sleep(300);
         }
      }

      catch ( InterruptedException ex ) {
         Thread.currentThread().interrupt();
      }

      catch ( Exception ex ) {
         log.fatal("[RptServer]", ex);
         shutdown(FATAL);
      }
   }

   /**
    * Sends an email notification to the recipients in the EMail input parameter.
    * @param recips List of recipients
    * @param subj The email subject
    *
    * @param msg The email message
    */
   public static void sendEmailNotification(String[] recips, String subj, String msg)
   {
      if ( msg != null && ( recips != null && recips.length > 0) ) {
         if ( subj == null || subj.length() == 0 )
            subj = "Report Notification";

         try {
            DataSender.smtp(emailFrom, recips, subj, msg);
         }

         catch (Exception ex ) {
            log.error("[RptServer]", ex);
         }
      }
   }

   /**
    * Shuts down the report server and all sub processes.
    * @param mode The mode to shut down the server
    */
   public void shutdown(short mode)
   {
      if ( m_Status == RUNNING ) {
         if ( mode < STOPPED )
            log.warn("invalid shutdown mode");
         else {
            log.info("shutting down adapter");
            m_Status = mode;

            //
            // Shutdown the report processor.
            m_RptMonitor.stop();
            m_Thread.interrupt();
            m_HtmlServer.stop();
         }
      }
   }

   /**
    * Starts the Report server.
    *
    * @throws Exception
    */
   public void start() throws Exception
   {
      //
      // Start the report queue processing.
      try {
         loadMBeans();

         if ( m_Status == FATAL || m_Status == ABORT )
            throw new Exception("fatal error, aborting startup");

         log.info("starting report server");
         m_RptMonitor.start();               // for messaging and monitor reports

         m_HtmlServer.setPort(Integer.parseInt(System.getProperty("jmx.port", "8082")));
         m_HtmlServer.start();               // JMX console
      }

      catch ( Exception ex ) {
         RptServer.log.fatal("[RptServer]", ex);
         throw ex;
      }

      //
      // Set the status and start the thread for this process.
      m_Status = RUNNING;
      m_Thread.start();
   }

   /**
    * A pass through method to the report monitor which has references to the report processes.
    *
    * @param rptId The internal ID of the runnting report.
    */
   public void stopReport(long rptId)
   {
      m_RptMonitor.stopReport(rptId);
   }

   //
   // Entry point into the server.  Creates and starts the report server.
   public static void main(String[] args)
   {
      RptServer server = new RptServer();

      try {
         server.start();
      }

      catch ( Exception ex ) {
         server.shutdown(RptServer.FATAL);
      }
   }
}
