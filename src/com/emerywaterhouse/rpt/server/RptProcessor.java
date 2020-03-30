/**
 * File: RptProcessor.java
 * Description: Handles the ProcessReportRequest BOD.  Parses the BOD to get the report class and
 *    parameters.  Launches the report class for creating the report.  The report processor is responsible
 *    for dealing with the distribution list, zipping the files, and sending emails.
 *
 * @author Jeffrey Fisher
 *
 * Create Data: 03/30/2005
 * Last Update: $Id: RptProcessor.java,v 1.31 2014/07/25 18:26:07 jfisher Exp $
 *
 * History
 *    $Log: RptProcessor.java,v $
 *    Revision 1.31  2014/07/25 18:26:07  jfisher
 *    Fixed a stream close bug.
 *
 *    Revision 1.30  2013/02/07 13:54:09  jfisher
 *    Fixed an issue with processing ftp strings without a ":" and logging mods
 *
 */
package com.emerywaterhouse.rpt.server;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.StringReader;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;

import oracle.jdbc.driver.OracleConnection;

import org.apache.log4j.Logger;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserFactory;

import com.emerywaterhouse.rpt.alert.Screamer;
import com.emerywaterhouse.utils.DataSender;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.utils.Zip;
import com.emerywaterhouse.websvc.Param;

public class RptProcessor implements Runnable
{
   private static long START_ID = 0;

   private final String BOD_VERB       = "Process";
   private final String ERR_MSG_HDR    = "the %s report had the following errors: \r\n";
   private final String FTP_TAG        = "Ftp";
   private final String FTP_URL_TAG    = "FtpUrl";
   private final String PARAM_TAG      = "Param";
   private final String PWD_TAG        = "Password";
   private final String RECIP_TAG      = "Recipient";
   private final String RPT_BOD_TYPE   = "ProcessReportRequest";
   private final String RPT_CLASS_TAG  = "ReportClass";
   private final String RPT_NAME_TAG   = "ReportName";
   private final String RPT_REQ_TAG    = "ReportRequest";
   private final String HTTP_TAG       = "Http";

   private final String SMTP_FROM      = "noreply@emeryonline.com";
   private final String STD_ERR_MSG    = "An unknown error occurred while running the report. Check the logs for details";
   private final String UID_TAG        = "UserId";
   private final String USER_ERR_MSG   = "There was an error creating the report.  Contact computer operations support.";

   private boolean m_Attachment;                // Flag that determines whether the report is sent as an attachment.
   private ArrayList<String> m_BList;           // Banned email list.
   private String m_Bod;                        // The process report request bod.
   private String m_BodType;                    // Contains the type of bod.
   private boolean m_Confirm;                   // Flag whether to send an email or confBOD.
   private ArrayList<RptRecipient> m_DistList;  // The email distribution list.
   private StringBuffer m_EmailMsg;             // The email message to send to the dist list.
   
   private String m_FtpPwd;                     // The password for the ftp account.
   private String m_FtpUrl;                     // The url for the ftp location of the reports.
   private String m_FtpUid;                     // The ftp user id.
   private RptMonitor m_Monitor;                // A reference to the report monitor class
   private boolean m_RptNotFound;               // Flag for whether a report class was missing.
   private long m_Id;                           // An id for identifying an instance of a report processor. 
   private String m_Pwd;                        // Password of user requesting report
   private Report m_Rpt;                        // A reference to the current report that is processing.
   private String m_RptClass;                   // The name of the report class to instantiate;
   private String m_RptName;                    // The name of the report.
   private ArrayList<Param> m_RptParams;        // The parameters that get passed to the report.
   private Thread m_Thread;                     // The thread that runs the processing.
   private String m_Uid;                        // The user id of the person requesting the report.
   private XmlPullParser m_Xpp;                 // XML pull parser for parsing responses
   private boolean m_Zipped;                    // Flag that determines whether the file should be zipped.
   private String m_HttpMethod;              // The method the REST web service should use to obtain the report output
   private String m_HttpUrl;              // URI of the REST web service servlet
   private String m_HttpUid;              // The user id of the REST web service
   private String m_HttpPwd;              // The password of the REST web service
   private String m_HttpAccessKey;           // The access key required by the REST web service

   private Connection m_FasConn;                // The connection to fascor.
   private OracleConnection m_OraConn;          // The db connection.   
   private Connection m_PgConn;                 // The connection to a Postgres database.
   private Connection m_EdbConn;                // Connection to EDB
   private Connection m_SageConn;               // Connection to Sage300/AccPac
   
   //
   // Log4j logger
   private static Logger log = Logger.getLogger(RptProcessor.class.getName());

   /**
    * default constructor.  Hidden to prevent actual use.
    */
   protected RptProcessor()
   {
      super();
   }

   /**
    * Creates a RptProcessor object with a bod and an instance to the report monitoring class.
    *
    * @param monitor An instance of the RptMonitor class.
    * @param bod The ProcessReportRequest BOD.
    */
   public RptProcessor(RptMonitor monitor, String bod)
   {
      super();

      m_Pwd = "";
      m_Uid = "";
      m_Id = START_ID++;
      m_HttpUrl = null;
      m_HttpUid = "";
      m_HttpPwd = "";
      m_HttpAccessKey = "";
      m_RptNotFound = false;

      XmlPullParserFactory factory = null;

      try {
         setMonitor(monitor);
         setBOD(bod);

         factory = XmlPullParserFactory.newInstance();
         factory.setNamespaceAware(true);
         m_Xpp = factory.newPullParser();

         m_DistList = new ArrayList<RptRecipient>();
         m_RptParams = new ArrayList<Param>();
         m_BList = new ArrayList<String>();

         m_EmailMsg = new StringBuffer();
         setFtpUrl(System.getProperty("rpt.url"));
      }

      catch ( Exception ex ) {
         log.fatal("[RptProcessor]", ex);
      }

      finally {
         factory = null;
      }
   }

   /**
    * Cleanup
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      m_Bod = null;
      m_FtpPwd = null;
      m_FtpUrl = null;
      m_FtpUid = null;
      m_Monitor = null;
      m_Thread = null;
      m_Xpp = null;
      m_Uid = null;
      m_Pwd = null;
      m_RptClass = null;
      m_RptName = null;
      m_Rpt = null;
      m_HttpUrl = null;
      m_HttpUid = null;
      m_HttpPwd = null;
      m_HttpAccessKey = null;

      if ( m_EmailMsg != null )
         m_EmailMsg = null;

      if ( m_DistList != null ) {
         m_DistList.clear();
         m_DistList = null;
      }

      if ( m_RptParams != null ) {
         m_RptParams.clear();
         m_RptParams = null;
      }

      if ( m_BList != null ) {
         m_BList.clear();
         m_BList = null;
      }

      if ( m_OraConn != null ) {
         m_OraConn.close();
         m_OraConn = null;
      }

      if ( m_FasConn != null ) {
         m_FasConn.close();
         m_FasConn = null;
      }

      if ( m_PgConn != null ) {
         m_PgConn.close();
         m_PgConn = null;
      }

      if ( m_EdbConn != null ) {
         m_EdbConn.close();
         m_EdbConn = null;
      }
      
      if ( m_SageConn != null ) {
         m_SageConn.close();
         m_SageConn = null;
      }
     
      super.finalize();
   }

   /**
    * Builds the email text.  Depends on whether attachments or ftping the files.
    * If the report didn't set a custom message then apply the stock message
    *
    * Note - the ftp url is expected to be in unix format and we have to remove the
    *    colon so the ftp location in the email is correct.
    *
    * @param fileCount The number of files that are being sent.
    */
   private void buildEmailText(int fileCount)
   {
      StringBuffer ftpUrl = new StringBuffer();
      String tmp = null;
      String dirCheck = "export/ftp/";
      int idx = 0;

      try {
         if ( m_EmailMsg.length() == 0 ) {
            m_EmailMsg.append("The " + m_RptName + " report has finished running.\r\n");

            if ( m_Attachment ) {
               m_EmailMsg.append("The following report files have been attached:\r\n");

               for ( int i = 0; i < fileCount; i++ )
                  m_EmailMsg.append(m_Rpt.getFileName(i) + "\r\n");
            }
            else {
               idx = m_FtpUrl.indexOf(':');

               if ( idx > -1 ) {
                  ftpUrl.append(m_FtpUrl.substring(0, idx));
                  tmp = m_FtpUrl.substring(idx+1);

                  //
                  // Check to see if there is a directory separator.
                  idx = tmp.indexOf('/');
                  if (idx != 0 )
                     ftpUrl.append("/");

                  ftpUrl.append(tmp);

                  //
                  // Check to make sure we have the file separator.  In this case
                  // we need to account for the 0 based index.
                  idx = ftpUrl.lastIndexOf("/")+1;

                  if ( idx != ftpUrl.length() )
                     ftpUrl.append("/");

                  //
                  // Check to see if we have the section of the url
                  // the points to the emery export dir.  This has to be removed.
                  idx = ftpUrl.indexOf(dirCheck);
                  if ( idx != -1 )
                     ftpUrl.delete(idx, idx + dirCheck.length());
               }
               else
                  ftpUrl.append(m_FtpUrl).append("/");

               m_EmailMsg.append("The following report files are ready for you to pick up:\r\n");

               for ( int i = 0; i < fileCount; i++ ) {
                  m_EmailMsg.append("ftp://");
                  m_EmailMsg.append(ftpUrl);
                  m_EmailMsg.append(m_Rpt.getFileName(i));
                  m_EmailMsg.append("\r\n");
               }
            }

            m_EmailMsg.append("\r\n\r\nYou received this email because you were on the report distribution list.");
            m_EmailMsg.append("\r\nIf you believe this is an error contact the MIS department");
         }
      }

      finally {
         ftpUrl = null;
         tmp = null;
      }
   }

   /**
    * Checks the report ACL.  First checks for the presence of the ACL and then checks
    * for the ability to run the report.
    *
    * @return true if the user has access, false if not.
    * @throws SQLException
    */
   private boolean checkAcl() throws SQLException
   {
      String rptIdQry = "select eis_rpt_id from eis_report where rpt_class = '%s'";
      String aclCntQry = "select count(*) from eis_report_acl where eis_rpt_id = %d";
      String aclAccessQry = "select count(*) from eis_report_acl where user_id = '%s'";
      boolean canAccess = false;
      Statement stmt = null;
      ResultSet rs = null;
      int aclCount = 0;
      int rptId = -1;

      if ( m_RptClass != null ) {
         stmt = m_EdbConn.createStatement();
         rs = stmt.executeQuery(String.format(rptIdQry,  m_RptClass));

         if ( rs.next() )
            rptId = rs.getInt(1);

         rs.close();
         rs = stmt.executeQuery(String.format(aclCntQry, rptId));

         if ( rs.next() )
            aclCount = rs.getInt(1);

         rs.close();

         //
         // Check for an acl entry, if there is one check the access.
         if ( aclCount > 0 ) {
            rs = stmt.executeQuery(String.format(aclAccessQry, m_Uid.toUpperCase()));
            if ( rs.next() )
               canAccess = rs.getInt(1) > 0;

            rs.close();
            stmt.close();
         }
         else
            canAccess = true;

         if ( !canAccess ) {
            StringBuffer msg = new StringBuffer();
            msg.append("user: ");
            msg.append(m_Uid);
            msg.append(" attempted to run ");
            msg.append(m_RptName);
            msg.append(", they have not been granted access.");
            log.warn("[RptProcessor] " + msg.toString());
            logReport("attempt to run with no access");
         }
      }

      return canAccess;
   }

   /**
    * Checks the email recipients against a list of known banned email addresses.
    *
    * @return true if a banned email address was found, false if not.
    * @throws Exception
    */
   private boolean checkBannedList() throws Exception
   {
      String tmp = null;
      StringBuffer baddr = new StringBuffer();
      boolean banned = false;
      Iterator<RptRecipient> iter = null;

      //
      // Check the distribution list against the banned list.  If there
      // are banned emails, remove them and create an email message notifying the
      // IT staff
      loadBannedList();
      baddr.setLength(0);

      iter = m_DistList.iterator();
      while ( iter.hasNext() ) {
         tmp = iter.next().email;

         if ( m_BList.contains(tmp)) {
            iter.remove();
            baddr.append(tmp);
            baddr.append("\r\n");
         }

         tmp = null;
      }

      if ( baddr.length() > 0 ) {
         banned = true;
         m_EmailMsg.append("The " + m_RptName);
         m_EmailMsg.append(" report has been requested to be sent to a banned email address.\r\n");
         m_EmailMsg.append("\r\n\r\n");
         m_EmailMsg.append("requesting user id: " + m_Uid + "\r\n");
         m_EmailMsg.append("band email address(s):\r\n");
         m_EmailMsg.append(baddr.toString());
         m_EmailMsg.append("\r\n\r\n");
         m_EmailMsg.append("non banned email address(s):\r\n");

         for ( int i = 0; i < m_DistList.size(); i++ ) {
            m_EmailMsg.append(m_DistList.get(i).email);
            m_EmailMsg.append("\r\n");
         }

         RptServer.notifyMis(m_EmailMsg.toString());
         m_EmailMsg.setLength(0);
         log.warn("[RptProcessor] " + m_Uid + " requested " + m_RptName + " with a banned email address");
         logReport("attempt to send to banned email address");
      }

      return banned;
   }

   /**
    * Creates the email message that is used when a report has been aborted.
    *
    * @return The abort message.
    */
   private String createAbortMsg()
   {
      m_EmailMsg.setLength(0);
      m_EmailMsg.append("The ");
      m_EmailMsg.append(m_RptName);
      m_EmailMsg.append(" report you submitted has been stopped by computer operations or a programmer.\r\n");
      m_EmailMsg.append("There may have been problems with the report.  Contact computer operations ");
      m_EmailMsg.append("for more information or try resubmitting the report request.\r\n");

      return m_EmailMsg.toString();
   }

   /**
    * Creates an error message that gets sent to the error distribution list.
    * Takes the error message from the report if there is one and then calls the other
    * createErrMsg() to finish formatting the text.
    *
    * @return Formatted email message text.
    * @see com.emerywaterhouse.rpt.server.RptProcessor#createErrMsg(String)
    */
   private String createErrMsg()
   {
      String errMsg = null;

      //
      // Get the information from the report if there is any.
      if ( m_Rpt != null )
         errMsg = m_Rpt.getErrMsg();

      if ( errMsg == null || errMsg.length() == 0 )
         errMsg = STD_ERR_MSG;

      return createErrMsg(errMsg);
   }

   /**
    * Formats an error message that is sent to the error distribution list.
    *
    * @param errMsg The error message that should be included in the email.
    * @return The formatted email text.
    */
   private String createErrMsg(String errMsg)
   {
      StringBuffer msg = new StringBuffer();

      if ( errMsg == null )
         errMsg = STD_ERR_MSG;

      //
      // Build the email message.
      msg.append(String.format(ERR_MSG_HDR, m_RptName));
      msg.append(errMsg);
      msg.append("\r\n\r\n");
      msg.append("This is the distribution list for the report:\r\n");

      for ( int i = 0; i < m_DistList.size(); i++ ) {
         msg.append(m_DistList.get(i).name);
         msg.append("\t");
         msg.append(m_DistList.get(i).email);
         msg.append("\r\n");
      }

      return msg.toString();
   }
   
   /**
    * Creates an Oracle db connection and stores it for use by the descendant
    * report classes.  This method must be called by the descendant class that is going
    * to be connection to Oracle.  The descendant class is not responsible for closing the
    * connection.
    *
    * @throws SQLException When an error occurs.
    */
   private void createOraConn() throws SQLException
   {
      String mode = System.getProperty("server.mode", "test");
      String prop = "db.ora." + mode;
      String url = System.getProperty(prop + ".url");
      String uid = System.getProperty(prop + ".uid");
      String pwd = System.getProperty(prop + ".pwd");

      if ( m_OraConn == null ) {
         DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
         m_OraConn = (OracleConnection)DriverManager.getConnection (url, uid, pwd);
      }
   }

   /**
    * Creates a Postres db connection and stores it for use by the descendant
    * report classes.  This method must be called by the descendant class that is going
    * to be connecting to Postgres.  The descendant class is not responsible for closing the
    * connection.
    *
    * @throws SQLException When an error occurs.
    */
   private void createPgConn() throws SQLException
   {
      String mode = System.getProperty("server.mode", "test");
      String prop = "db.pg." + mode;
      String url = System.getProperty(prop + ".url");
      String uid = System.getProperty(prop + ".uid");
      String pwd = System.getProperty(prop + ".pwd");

      if ( m_PgConn == null ) {
         DriverManager.registerDriver(new org.postgresql.Driver());
         m_PgConn = DriverManager.getConnection (url, uid, pwd);
      }
   }
   
   /**
    * Creates a Sage300 db connection and stores it for use by the descendant
    * report classes.  This method must be called by the descendant class that is going
    * to be connecting to SQL Server.  The descendant class is not responsible for closing the
    * connection.
    *
    * @throws SQLException When an error occurs.
    */
   private void createSageConn() throws SQLException
   {
      String mode = System.getProperty("server.mode", "test");
      String prop = "db.sage." + mode;
      String url = System.getProperty(prop + ".url");
      String uid = System.getProperty(prop + ".uid");
      String pwd = System.getProperty(prop + ".pwd");

      if ( m_SageConn == null ) {
         DriverManager.registerDriver(new net.sourceforge.jtds.jdbc.Driver());
         m_SageConn = DriverManager.getConnection (url, uid, pwd);
      }
   }


   /**
    * Instantiate a Report object.
    *
    * @param className The name of the report class to instantiate
    * @return An instance of the a Report object.
    */
   private Report createRptClass(String className)
   {
      Class<?> c = null;
      Object obj = null;
      Report rpt = null;

      try {
         if ( m_RptClass != null ) {
            try {
               c = RptServer.getLoader().loadClass(m_RptClass);
            }
            
            catch ( ClassNotFoundException ex ) {
               //
               // If the report class couldn't be found or there was some other error loading the report class
               // create a default report that will contain some error information and alert the requester.
               if ( c == null ) {
                  c = RptServer.getLoader().loadClass("com.emerywaterhouse.rpt.alert.Screamer");            
                  rpt = (Report)c.newInstance();
                  m_RptNotFound = true;
               }
            }

            if ( c != null && !m_RptNotFound) {
               obj = c.newInstance();

               if ( obj instanceof Report )
                  rpt = (Report)obj;
               else
                  log.error(String.format("[RptProcessor] %s is not an instance of the class Report", m_RptClass));
            }
            else
               log.error("[RptProcessor] unable to instantiate " + m_RptClass);
         }
         else
            log.error("[RptProcessor] missing report class from the bod");        
      }
      
      catch ( Exception ex ) {
         log.error("[RptProcessor]", ex);
      }

      finally {
         obj = null;
         c = null;
      }

      return rpt;
   }

   /**
    * FTPs the report file to the ftp server.  The url is in the form of ipaddr:dir
    * The process pulls out the path of the file name since it's not used on the ftp server.
    *
    * @param fileName The name of the file on the server the report was created on, including the path.
    * @param isAscii Determines whether the ftp should be ascii or binary
    *
    * @throws Exception When an ftp error occurs.
    */
   public void ftpRptFile(String fileName, boolean isAscii) throws Exception
   {
      File rptFile = null;
      String sepChar = System.getProperty("file.separator", "/");

      try {
         rptFile = new File(fileName);
         fileName = fileName.substring(fileName.lastIndexOf(sepChar)+1, fileName.length());
         log.info("[RptProcessor] ftping " + fileName + " to url: " + m_FtpUrl);

         if ( m_FtpUid == null )
            m_FtpUid = System.getProperty("ftp.uid");

         if ( m_FtpPwd == null )
            m_FtpPwd = System.getProperty("ftp.pwd");

         DataSender.ftp(m_FtpUrl, m_FtpUid, m_FtpPwd, fileName, rptFile, isAscii);
      }

      finally {
         rptFile = null;
      }
   }

   /**
    * Gets the distribution list.  The DataSender class takes the email address list
    * as an array of strings.  This translates the internal distribution list into the
    * proper format.
    *
    * @return The distribution list as an array of strings.
    */
   public String[] getDistList()
   {
      int size = m_DistList.size();
      String[] list = new String[size];

      for ( int i = 0; i < size; i++ )
         list[i] = m_DistList.get(i).email;

      return list;
   }

   /**
    * Creates a Enterprise db connection and stores it for use by the descendant
    * report classes.  This method must be called by the descendant class that is going
    * to be connecting to Enterprise.  The descendant class is not responsible for closing the
    * connection.
    *
    * @throws SQLException When an error occurs.
    */
   public Connection getEdbConn() throws SQLException
   {           
      return getEdbConn("emery_jensen");
   }
   
   /**
    * Creates a Enterprise db connection and stores it for use by the descendant
    * report classes.  This method must be called by the descendant class that is going
    * to be connecting to Enterprise.  The descendant class is not responsible for closing the
    * connection.
    *
    * @throws SQLException When an error occurs.
    */
   public Connection getEdbConn(String database) throws SQLException
   {
      if ( database != null && database.length() > 0 ) {
         if ( m_EdbConn == null ) {
            String mode = System.getProperty("server.mode", "test");
            String prop = "db.edb." + mode;
            String url = System.getProperty(prop + ".read.url") + database;      
            String uid = System.getProperty(prop + ".uid");
            String pwd = System.getProperty(prop + ".pwd");
            
            DriverManager.registerDriver(new com.edb.Driver());
            m_EdbConn = DriverManager.getConnection (url, uid, pwd);
            m_EdbConn.setAutoCommit(false);
         }
      }
      else
         throw new SQLException("missing database parameter");
      
      return m_EdbConn;
   }
   
   /**
    * Returns the report processor's fascor connection.
    *
    * @return A reference to the report processor's fascor connection object.
    * @throws SQLException
    */
   public Connection getFasConn() throws SQLException
   {
      if ( m_FasConn == null ) {
         DriverManager.registerDriver(new net.sourceforge.jtds.jdbc.Driver());
         m_FasConn = DriverManager.getConnection ("jdbc:jtds:sqlserver://10.128.0.20:1433/DC01EWH", "DC01EWH", "DC01EWH");
      }

      return m_FasConn;
   }

   /**
    * Returns the report processor's fascor connection.
    * @param dc
    *
    * @return A reference to the report processor's fascor connection object.
    * @throws SQLException
    */
   public Connection getFasConn(String dc) throws SQLException
   {
      String url = "jdbc:jtds:sqlserver://10.128.0.20:1433/DC%sEWH";
      String uid = "DC%sEWH";
      String pwd = "DC%sEWH";

      if ( m_FasConn == null ) {
         try {
            DriverManager.registerDriver(new net.sourceforge.jtds.jdbc.Driver());
            m_FasConn = DriverManager.getConnection(
               String.format(url, dc),
               String.format(uid, dc),
               String.format(pwd, dc)
            );
         }

         finally {
            url = null;
            uid = null;
            pwd = null;
         }
      }

      return m_FasConn;
   }

   /**
    * Returns the Ftp password
    *
    * @return String - the Ftp password
    */
   public String getFtpPwd()
   {
      return m_FtpPwd;
   }

   /**
    * Returns the ftp server url
    *
    * @return String - the ftp server url
    */
   public String getFtpUrl()
   {
      return m_FtpUrl;
   }

   /**
    * Returns the ftp user id
    *
    * @return String - the ftp user id
    */
   public String getFtpUid()
   {
      return m_FtpUid;
   }

   /**
    * Returns the access key required by the REST web service
    *
    * @return String - the REST web service access key
    */
   public String getHttpAccessKey()
   {
      return m_HttpAccessKey;
   }

   /**
    * Returns the password of the REST web service
    *
    * @return String - the REST web service password
    */
   public String getHttpPwd()
   {
      return m_HttpPwd;
   }

   /**
    * Returns the user id of the REST web service
    *
    * @return String - the REST web service user id
    */
   public String getHttpUid()
   {
      return m_HttpUid;
   }

   /**
    * Returns the HTTP uri when notification must be sent to a servlet
    *
    * @return String - the HTTP uri
    */
   public String getHttpUrl()
   {
      return m_HttpUrl;
   }

   /**
    * Return the id of the processor.
    *
    * @return The internal processor id used for identifiy this process.
    */
   public long getId()
   {
      return m_Id;
   }

   /**
    * Returns the report processor's oracle connection.
    *
    * @return A reference to the report processor's oracle connection object.
    * @throws SQLException
    */
   public OracleConnection getOraConn() throws SQLException
   {
      if ( m_OraConn == null )
         createOraConn();

      return m_OraConn;
   }

   /**
    * Returns the report processor's postgres connection.
    *
    * @return A reference to the report processor's postgres connection object.
    * @throws SQLException
    */
   public Connection getPgConn() throws SQLException
   {
      if ( m_PgConn == null )
         createPgConn();

      return m_PgConn;
   }
   
   
   /**
    * Returns the report processor's Sage300 connection.
    *
    * @return A reference to the report processor's Sage300 connection object.
    * @throws SQLException
    */
   public Connection getSageConn() throws SQLException
   {
      if ( m_SageConn == null )
         createSageConn();

      return m_SageConn;
   }


   /**
    * Gets the status of the running report and the report processor.
    *
    * @return A newly created RptStatus objet.
    */
   public RptStatus getRptStatus()
   {
      RptStatus status = new RptStatus();

      if ( m_Rpt != null ) {
         status.currentAction = m_Rpt.getCurAction();
         status.internalId = m_Id;
         status.rptName = m_RptName;
         status.maxRunTime = m_Rpt.getMaxRunTime();
         status.runTime = System.currentTimeMillis() - m_Rpt.getStartTime();
         status.startTime = m_Rpt.getStartTime();
         status.threadId = m_Thread.getId();
         status.uid = m_Uid;
      }

      return status;
   }

   /**
    * Returns the user id that requested the report.
    *
    * @return The userid of the report requester.
    */
   public String getUid()
   {
      return m_Uid;
   }

   /**
    * returns whether the report should be compressed or not.
    * @return boolean true if the report is compressed, false if not.
    */
   public boolean getZipped()
   {
      return m_Zipped;
   }

   /**
    * Loads the list of banned email addresses from disk and stores them in the
    * ArrayList for later use.
    *
    * @throws FileNotFoundException
    * @throws IOException
    */
   private void loadBannedList() throws FileNotFoundException, IOException
   {
      String line = null;
      BufferedReader fr = new BufferedReader(new FileReader("bemail.lst"));

      try {
         line = fr.readLine();

         while ( line != null ) {
            m_BList.add(line);
            line = fr.readLine();
         }
      }

      finally {
         fr.close();
      }
   }

   /**
    * Logs the report request in the db report log.  This just give more information than in the
    * server log file.
    *
    * @param msg Extra message text that can be logged.
    * @throws SQLException
    */
   private void logReport(String msg) throws SQLException
   {
      Connection conn = null;
      Statement stmt = null;
      String sql = "insert into eis_report_log (user_id, rpt_name, customer_id, recipients, msg) values ('%s','%s','%s','%s','%s')";
      String recips = "";
      String custId = "";

      String mode = System.getProperty("server.mode", "test");
      String prop = "db.edb." + mode;
      String url = System.getProperty(prop + ".write.url") + "emery_jensen";      
      String uid = System.getProperty(prop + ".uid");
      String pwd = System.getProperty(prop + ".pwd");
      
      DriverManager.registerDriver(new com.edb.Driver());
      conn = DriverManager.getConnection (url, uid, pwd);
      conn.setAutoCommit(true);
      conn.setReadOnly(false);
            
      if ( conn != null ) {
         try {
            if ( msg == null )
               msg = "";

            if ( m_Rpt != null )
               custId = m_Rpt.getCustId();

            for ( int i = 0; i < m_DistList.size(); i++ ) {
               if ( i > 0 )
                  recips += ",";

               //
               // In the infinite wisdom of operations, they gave users email addresses with
               // an apostrophe.  We need to escape these.
               recips += m_DistList.get(i).email.replace("'", "''");
            }

            stmt = conn.createStatement();
            stmt.executeUpdate(String.format(sql, m_Uid, m_RptName, custId, recips, msg));
         }

         finally {
            DbUtils.closeDbConn(conn, stmt, null);
            stmt = null;
            conn = null;
         }
      }
   }

   /**
    * Parses the Report BOD to pull out params, distro list, etc.  All the parsing is done in this
    * method instead of other functions.  The amount of xml to parse is small and solves a couple of
    * thorny issues with attributes.
    *
    * @throws Exception When a parser exception occurs.
    */
   private void parseBOD() throws Exception
   {
      int eventType;
      String curTag = "";
      boolean done = false;
      RptRecipient rcp = null;
      String curText = null;
      Param param = null;

      m_Xpp.setInput(new StringReader(m_Bod));
      eventType = m_Xpp.getEventType();

      while ( !done ) {
         switch ( eventType ) {
            case XmlPullParser.START_DOCUMENT:
            break;

            case XmlPullParser.END_DOCUMENT:
               done = true;
            break;

            //
            // Check the bod type which will be the first tag.  If it's not
            // correct abort.
            case XmlPullParser.START_TAG:
               curTag = m_Xpp.getName();

               if ( m_BodType == null ) {
                  m_BodType = curTag;

                  if ( !m_BodType.equals(RPT_BOD_TYPE) )
                     throw new Exception("invalid bod");
               }
               else {
                  //
                  // Pull out whether to confirm.  This will be via email and not a ConfirmBOD
                  if ( curTag.equals(BOD_VERB) ) {
                     m_Confirm = m_Xpp.getAttributeValue(1).equals("Always");
                     break;
                  }

                  if ( curTag.equals(RPT_REQ_TAG) ) {
                     m_Zipped = m_Xpp.getAttributeValue(0).equalsIgnoreCase("yes");
                     m_Attachment = m_Xpp.getAttributeValue(1).equalsIgnoreCase("yes");
                     break;
                  }

                  //
                  // Get any report params
                  if ( curTag.equals(PARAM_TAG) ) {
                     param = new Param();
                     param.name = m_Xpp.getAttributeValue(0);
                     param.type = m_Xpp.getAttributeValue(1);
                     param.value = m_Xpp.getAttributeValue(2);
                     m_RptParams.add(param);
                     param = null;
                     break;
                  }

                  //
                  // Get the attributes for the recipient.
                  if ( curTag.equals(RECIP_TAG) ) {
                     rcp = new RptRecipient();
                     rcp.name = m_Xpp.getAttributeValue(0);
                     rcp.email = m_Xpp.getAttributeValue(1);
                     m_DistList.add(rcp);
                     rcp = null;
                     break;
                  }

                  if ( curTag.equals(FTP_TAG) ) {
                     m_FtpUrl = m_Xpp.getAttributeValue(0);
                     m_FtpUid = m_Xpp.getAttributeValue(1);
                     m_FtpPwd = m_Xpp.getAttributeValue(2);
                     break;
                  }

                  if ( curTag.equals(HTTP_TAG) ) {
                     for ( int i = 0; i < m_Xpp.getAttributeCount(); i++ ) {
                        if ( m_Xpp.getAttributeName(i).equals("url") ) {
                           m_HttpUrl = m_Xpp.getAttributeValue(i);
                           continue;
                        }

                        if ( m_Xpp.getAttributeName(i).equals("method") ) {
                           m_HttpMethod = m_Xpp.getAttributeValue(i);
                           continue;
                        }

                        if ( m_Xpp.getAttributeName(i).equals("uid") ) {
                           m_HttpUid = m_Xpp.getAttributeValue(i);
                           continue;
                        }

                        if ( m_Xpp.getAttributeName(i).equals("pwd") ) {
                           m_HttpPwd = m_Xpp.getAttributeValue(i);
                           continue;
                        }

                        if ( m_Xpp.getAttributeName(i).equals("accesskey") ) {
                           m_HttpAccessKey = m_Xpp.getAttributeValue(i);
                           continue;
                        }
                     }
                  }
               }
            break;

            case XmlPullParser.END_TAG:
               curTag = "";
            break;

            case XmlPullParser.TEXT:
               //
               // If there is no tag name, skip everything.
               if ( curTag != null && curTag.length() > 0 ) {
                  curText = m_Xpp.getText();

                  //
                  // Set the name of the report
                  if ( curTag.equals(RPT_NAME_TAG) ) {
                     m_RptName = curText;
                     break;
                  }

                  //
                  // Set the name of the class file.
                  if ( curTag.equals(RPT_CLASS_TAG) ) {
                     m_RptClass = curText;
                     break;
                  }

                  //
                  // Set the user id.
                  if ( curTag.equals(UID_TAG) ) {
                     setUid(curText);
                     break;
                  }

                  //
                  // set the passord
                  if ( curTag.equals(PWD_TAG) ) {
                     m_Pwd = curText;
                     break;
                  }

                  if (curTag.equals(FTP_URL_TAG) ) {
                     setFtpUrl(curText);
                     break;
                  }
               }
            break;
         }

         if ( eventType != XmlPullParser.END_DOCUMENT )
            eventType = m_Xpp.next();
      }
   }

   /**
    * Public interface to start the bod processing.  Starts the thread if it has been assigned
    * which will start the processing of the BOD document.  This method will return before the
    * processing is finished which will allow the report listener to pull another report off of
    * the queue while others are being processed.
    */
   public void processBOD()
   {
      m_Thread = new Thread(this, "RptProcessor" + m_Id);
      m_Thread.setDaemon(true);
      m_Thread.start();
   }

   /**
    * Implements the runnable interface.  Parses the bod for the report parameters and
    * then instantiates the report class.  This finishes when the report finishes.  Handles
    * parsing the BOD and generating the email messages, sending the files, etc.  Also checks
    * for banned email addresses.  If a banned email address is found, the report is not run.
    *
    * Note - the report is not run in a separate thread.  It's treated just like a method.
    *
    * @see java.lang.Runnable#run()
    * @see com.emerywaterhouse.rpt.server.RptProcessor#parseBOD()
    */
   @Override
   public void run()
   {
      String fileName = null;
      int fileCount = 1;
      String[] attachments = null;
      boolean hasBanned = false;
      boolean canAccess = false;
      boolean ascii = false;
      Param param = null;

      try {
         getEdbConn();   // required to check ACL and some other stuff.  Need to migrate this to EDB.
         parseBOD();                  
         canAccess = checkAcl();
         hasBanned = checkBannedList();

         if ( canAccess && !hasBanned ) {
            //
            // Once we've parsed out everything we need.  We can create
            // the report object and then let it run until completion.
            try {
               log.info("[RptProcessor] starting " + m_RptName + " for " + m_Uid);
               m_Rpt = createRptClass(m_RptClass);

               //
               // Run the report
               if ( m_Rpt != null ) {
                  m_Rpt.setRptProcessor(this);

                  //
                  // If the report class wasn't found, add an additional parameter to the parameter list.
                  // This will be used by the alert class when building a message.  Also turn off the confirmation
                  // flag and the attachment flag.  The alert process will handle notification.
                  if ( m_RptNotFound ) {
                     param = new Param();
                     param.name = "reportname";
                     param.type = "string";
                     param.value = m_RptName;
                     m_RptParams.add(param);
                     
                     m_Confirm = false;
                     m_Attachment = false;
                     m_Zipped = false;
                  }
                  
                  if ( m_RptParams.size() > 0 )
                     m_Rpt.setParams(m_RptParams);
                  
                  //
                  // Log the request in the db.  We need to do this one after the params
                  // so we can get the customer number if there is one.
                  logReport("");

                  //
                  // Once the report is created, send it to the recipients.
                  // Note - the report may have multiple files.  Also we only send xls files or
                  //   zipped files as binary files.  We may have to change this later.
                  if ( m_Rpt.createReport() && m_Rpt.getStatus() != RptServer.ABORT ) {
                     fileCount = m_Rpt.getFileCount();
                     attachments = fileCount > 0 ? new String[fileCount] : null;

                     for ( int i = 0; i < fileCount; i++ ) {
                        fileName = m_Rpt.getFilePath() + m_Rpt.getFileName(i);

                        if ( m_Zipped ) {
                           fileName = zipRptFile(fileName);
                           m_Rpt.m_FileNames.set(i, fileName);
                        }

                        if ( m_Attachment )
                           attachments[i] = fileName;
                        else {
                           ascii = !m_Zipped && !fileName.endsWith(".xls");
                           ftpRptFile(fileName, ascii);
                        }

                        if ( m_HttpUrl != null ) {
                           sendHttpNotification(fileName);
                        }
                     }

                     if ( m_Confirm && m_DistList.size() > 0 ) {
                        buildEmailText(fileCount);

                        //
                        // Send the email outside of the zipping/ftp loop.  Don't want multiple emails.
                        // Note - Have to use datasender because the web service doesn't do attachments.
                        if ( m_Attachment )
                              DataSender.smtp(SMTP_FROM, getDistList(), m_RptName, m_EmailMsg.toString(),
                                 attachments, DataSender.ATTACH_MULTI);
                        else
                           DataSender.smtp(SMTP_FROM, getDistList(), m_RptName, m_EmailMsg.toString());
                     }
                     
                     log.info("[RptProcessor] finished " + m_RptName + " for " + m_Uid);
                  }
                  else {
                     if ( m_Rpt.getStatus() == RptServer.ABORT )
                        RptServer.sendEmailNotification(getDistList(), m_RptName, createAbortMsg());
                     else {
                        RptServer.notifyMis(createErrMsg());
                        RptServer.sendEmailNotification(getDistList(), m_RptName, USER_ERR_MSG);
                     }
                  }
               }
               else
                  log.fatal("[RptProcessor] null report class, uknown creation failure");
            }

            catch ( Exception ex ) {
               log.error("[RptProcessor]", ex);
               RptServer.notifyMis(createErrMsg(ex.getMessage()));
               RptServer.sendEmailNotification(getDistList(), m_RptName, USER_ERR_MSG);
            }
         }
      }

      catch ( Exception ex ) {
         log.fatal("[RptProcessor]", ex);
         RptServer.notifyMis(createErrMsg(ex.getMessage()));
         RptServer.sendEmailNotification(getDistList(), m_RptName, USER_ERR_MSG);
      }

      finally {
         m_Monitor.decRptCount();
         m_Monitor.remRptProc(this);
         m_Rpt = null;
         
         try {
            m_EdbConn.commit();
         } 
         
         catch (SQLException ex) {         
            log.error("[RptProcessor]", ex);
         }
         
         DbUtils.closeDbConn(m_EdbConn, null, null);
         DbUtils.closeDbConn(m_OraConn, null, null);
         DbUtils.closeDbConn(m_PgConn, null, null);
         DbUtils.closeDbConn(m_FasConn, null, null);
         
         m_EdbConn = null;
         m_OraConn = null;
         m_PgConn = null;
         m_FasConn = null;
      }
   }


   /**
    * Issues an HTTP request that notifies a servlet that the file is available.
    * The location of the file - ftp server url, user id, and password
    * are included as parameters in the request.
    *
    * @param fileName The name of the file on the server the report was created on, including the path.
    *
    * @throws Exception If the HTTP request could not be made
    */
   public void sendHttpNotification(String fileName) throws Exception
   {
      String sepChar = System.getProperty("file.separator", "/");
      StringBuffer tmp = new StringBuffer();

      fileName = fileName.substring(fileName.lastIndexOf(sepChar)+1, fileName.length());

      if ( m_HttpMethod.trim().equalsIgnoreCase("ftp") ) {
         if ( m_FtpUid == null )
            m_FtpUid = System.getProperty("ftp.uid");

         if ( m_FtpPwd == null )
            m_FtpPwd = System.getProperty("ftp.pwd");

         tmp.append(m_HttpUrl);
         tmp.append("?filename=");
         tmp.append(fileName);
         tmp.append("&method=");
         tmp.append(m_HttpMethod);

         if ( m_HttpUid != null && m_HttpUid.trim().length() > 0 ) {
            tmp.append("&wsuid=");
            tmp.append(m_HttpUid);
         }

         if ( m_HttpPwd != null && m_HttpPwd.trim().length() > 0 ) {
            tmp.append("&wspwd=");
            tmp.append(m_HttpPwd);
         }

         if ( m_HttpAccessKey != null && m_HttpAccessKey.trim().length() > 0 ) {
            tmp.append("&accesskey=");
            tmp.append(m_HttpAccessKey);
         }

         if ( m_FtpUrl != null && m_FtpUrl.trim().length() > 0 ) {
            tmp.append("&ftpurl=");
            tmp.append(m_FtpUrl);
         }

         if ( m_FtpUid != null && m_FtpUid.trim().length() > 0 ) {
            tmp.append("&ftpuid=");
            tmp.append(m_FtpUid);
         }

         if ( m_FtpPwd != null && m_FtpPwd.trim().length() > 0 ) {
            tmp.append("&ftppwd=");
            tmp.append(m_FtpPwd);
         }

         if ( m_Zipped )
            tmp.append("&zipped=true");

         //
         // Send the http request.  Just for grins, include the original report
         // request bod in the body of the http request.  Who knows.  It might come
         // in handy.
         log.info("[RptProcessor] Sending http request: " + tmp.toString());
         DataSender.http(tmp.toString(), m_Bod);
      }

      if ( m_HttpMethod.equalsIgnoreCase("stream") ) {
         // TODO implement stream from the request payload
      }

      if ( m_HttpMethod.equalsIgnoreCase("filecopy") ) {
         // TODO implement just moving the file from the source to destination
      }
   }


   /**
    * Sets the internal BOD member.
    *
    * @param bod The bod.
    * @throws Exception when the bod var is null.
    */
   public void setBOD(String bod) throws Exception
   {
      if ( bod != null)
         m_Bod = bod;
      else
         throw new Exception("attempt to set bod to null");
   }

   /**
    * Sets the email message.  This allows the reports to provide a custom email msg if the
    * standard email message doesn't work.  It should only be used for specific non standard
    * email messages.
    *
    * @param msg The email message.
    */
   public void setEmailMsg(String msg)
   {
      if ( msg != null ) {
         m_EmailMsg.setLength(0);
         m_EmailMsg.append(msg);
      }
   }

   /**
    * Sets the ftp url.  Checks to make sure it's not an empty url and also
    * makes sure there is a trailing slash or backslash.
    *
    * @param url The ftp url.
    */
   private void setFtpUrl(String url)
   {
      String sepChar = "/";

      if ( url != null && url.length() > 0 ) {
         m_FtpUrl = url;

         if ( !m_FtpUrl.endsWith(sepChar) )
            m_FtpUrl += sepChar;
      }
   }

   /**
    * sets the REST web service access key
    *
    * @param key String - the REST web service access key
    */
   public void setHttpAccessKey(String key)
   {
      m_HttpAccessKey = key;
   }

   /**
    * Sets the REST web service password
    *
    * @param pwd String - the REST web service password
    */
   public void setHttpPwd(String pwd)
   {
      m_HttpPwd = pwd;
   }

   /**
    * Sets the REST web service user id
    *
    * @param uid String - the REST web service user id
    */
   public void setHttpUid(String uid)
   {
      m_HttpUid = uid;
   }

   /**
    * For use when a servlet must be notified that the output file is available
    *
    * @param url String - the url of the HTTP request
    */
   public void setHttpUri(String url)
   {
      m_HttpUrl = url;
   }

   /**
    * Sets the monitor var.
    *
    * @param monitor The monitor reference to set.
    * @throws Exception when the monitor var is null.
    */
   public void setMonitor(RptMonitor monitor) throws Exception
   {
      if ( monitor != null ) {
         m_Monitor = monitor;
         m_Monitor.addRptProc(this);
      }
      else
         throw new Exception("report monitor can't be set to null");
   }

   /**
    * Pass through function for setting the status of a running report.
    *
    * @param status The new status that is to be set.
    */
   public void setRptStatus(short status)
   {
      synchronized ( m_Rpt ) {
         if ( m_Rpt != null )
            m_Rpt.setStatus(status);

         if ( status == RptServer.STOPPED || status == RptServer.ABORT )
            if ( Thread.currentThread().getId() == m_Thread.getId() )
               m_Thread.interrupt();
      }
   }

   /**
    * Sets the user id field.
    *
    * @param uid The user id to set.
    */
   public void setUid(String uid)
   {
      if ( uid != null )
         m_Uid = uid;
      else
         m_Uid = "";
   }

   /**
    * Zips the report file in winzip format.
    * @param inFile the path and file name of the file to compress.
    * @return Returns the name of the file with the zip extension.
    *
    * @throws IOException
    */
   public String zipRptFile(String inFile) throws IOException
   {
      Zip zip;
      File rptFile = null;
      String outFile;

      zip = new Zip(0);
      rptFile = new File(inFile);

      try {
         outFile = inFile.substring(0, inFile.lastIndexOf('.')) + ".zip";
         zip.zipFile(inFile, outFile, false);
         rptFile.delete();
      }

      finally {
         zip = null;
         rptFile = null;
      }

      return outFile;
   }

   /**
    * Internal place holder class for a distribution list recipient.  Holds
    * the email name and email address of the recipient.  This object is used
    * so that the data is bound together in the distribution list.
    *
    * Note - since this is a place holder class only, all members are public and
    *    the m_ naming convention is not used.
    */
   private class RptRecipient
   {
      public String name;
      public String email;

      /**
       * default constructor
       */
      RptRecipient()
      {
         super();

         name = "";
         email = "";
      }
   }
}
