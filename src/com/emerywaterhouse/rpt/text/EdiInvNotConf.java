package com.emerywaterhouse.rpt.text;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.GregorianCalendar;

import org.apache.log4j.Logger;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DataSender;
import com.emerywaterhouse.utils.DbUtils;

public class EdiInvNotConf extends Report {
	
   private PreparedStatement m_EdiInvoices;
   private GregorianCalendar m_Date;
   protected final static Logger log = Logger.getLogger("com.emerywaterhouse.EDIInvoicesNotConfirmed");
   
   private static final String[] rcptList = {"programming@emeryonline.com"};

   /**
    * Environment enumeration
    */
   public enum Environment {
      Test,
      Production
   };

   private Environment m_Env;                     /** Environment variable used to determine test or production */
   private static String m_From;                  /** The email from property */

   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {      
      StringBuffer Line = new StringBuffer();
      FileOutputStream OutFile = null;
      ResultSet ediInvoiceData = null;
      
      boolean Result = false;
      m_Date = new GregorianCalendar();
      StringBuffer name = new StringBuffer();
      
      name.append("ediInvNotConf");
      name.append('-');
      name.append(m_Date.get(Calendar.YEAR));
      name.append(m_Date.get(Calendar.MONTH)+1);
      name.append(m_Date.get(Calendar.DATE));
      name.append(".dat");
      
      m_FileNames.add(name.toString());
      String fileName = m_FilePath + m_FileNames.get(0);
      
      log.info("[EdiInvNotConf] creating output file: " + fileName);
      OutFile = new FileOutputStream(fileName, false);

      //date formatter from java.utils
      SimpleDateFormat formatter = new SimpleDateFormat("MM/dd/yyyy");

      //
      // Build the Captions      
      Line.append("Invoice Number\tAck_Status\t");
      Line.append("Time_sent\tWF_ID\r\n");
      int count = 0;
      
      try {
         setCurAction("executing report query ");
         ediInvoiceData = m_EdiInvoices.executeQuery();
         setCurAction("report query executed  ");

         while ( ediInvoiceData.next() && m_Status == RptServer.RUNNING ) {
            count++;
            
            Line.append(ediInvoiceData.getString("invoice_num") + "\t\t");
            Line.append(ediInvoiceData.getString("ack_status") + "\t\t");
            Line.append(formatter.format(ediInvoiceData.getDate("time_sent")).toString() + "\t");
            Line.append(ediInvoiceData.getString("wf_id"));

            Line.append("\r\n");
            OutFile.write(Line.toString().getBytes());

         }
         ediInvoiceData.close();
         
         log.info("[EdiInvNotConf] notify Mis for not confirmed EDI invoices");
         
      //   if(count > 0)
         notifyMis(Line.toString());
         
         Result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         
         log.error("[EdiInvNotConf] exception: ", ex);
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
    * Resource cleanup
    */
   private void closeStatements()
   {
   	DbUtils.closeDbConn(m_OraConn, m_EdiInvoices, null);
   	m_EdiInvoices = null;
   	m_OraConn = null;
   }

	/* (non-Javadoc)
	 * @see com.emerywaterhouse.rpt.server.Report#createReport()
	 */
	@Override
	public boolean createReport()
	{
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      m_Env = Environment.Production;

      if ( System.getProperty("server.mode", "test").equalsIgnoreCase("test") )
         m_Env = Environment.Test;
      else
         m_Env = Environment.Production;
      
      try {         
         m_OraConn = m_RptProc.getOraConn();   
         if ( prepareStatements() ) 
            created = buildOutputFile();
      }
      
      catch ( Exception ex ) {
         log.fatal("exception:", ex);
      }
      
      finally {
         closeStatements(); 
         
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }
      
      return created;
	}
		
   /**
    * Prepares the sql queries for execution.
    */
	private boolean prepareStatements()
	{      
	   StringBuffer sql = new StringBuffer();
	   boolean isPrepared = false;

	   if ( m_OraConn != null ) {
	      try {

	         sql.setLength(0);
	         sql.append("select wf_id, invoice_num, time_sent, ack_status from ( ");
	         sql.append("   select ");
	         sql.append("   cs_inv.wf_id, to_number(cs_inv.value) invoice_num, to_date('19700101','YYYYMMDD')+ (cs_time.value/86400000)-(5/24) time_sent, ");
	         sql.append("   cs_ack.value ack_status ");
	         sql.append("from sterlb2bi.correlation_set cs_inv ");     
	         sql.append("  join sterlb2bi.correlation_set cs_810 on cs_810.wf_id = cs_inv.wf_id ");
	         sql.append("  join sterlb2bi.correlation_set cs_time on cs_inv.wf_id = cs_time.wf_id ");
	         sql.append("  join sterlb2bi.correlation_set cs_ack on cs_ack.wf_id = cs_time.wf_id ");
	         sql.append("  where ");
	         sql.append("     cs_inv.name = 'EM_InvoiceNbr' and cs_810.name = 'TransactionSetID' and  ");
	         sql.append("     cs_810.value = '810' and cs_time.name = 'TransactionDateTime' and  ");
	         sql.append("     cs_ack.name = 'GroupAckStatus' ");
	         sql.append(")  ");
	         sql.append("where time_sent >= trunc(sysdate) - 14 and time_sent < trunc(sysdate) - 1  ");  
	         sql.append("  and ack_status = 'WAITING' ");	      

	         m_EdiInvoices = m_OraConn.prepareStatement(sql.toString());
	         isPrepared = true;
	      }

	      catch ( SQLException ex ) {
	         log.error("[EdiInvNotConf] exception:", ex);
	      }

	      finally {
	         sql = null;
	      }         
	   }
	   else
	      log.error("[EdiInvNotConf] - null oracle connection");

	   return isPrepared;
	}

   /**
    * Sends an email notification to the MIS department.
    *
    * @param msg - The email message to send.
    */
   public void notifyMis(String msg)
   {
      String[] recips = null;
      String subj = "Not Confirmed EDI Invoices Sent to LMBA/LMC in last two weeks";

      m_From = System.getProperty("mail.from", "noreply@emeryonline.com");

      if ( msg != null ) {
         switch ( m_Env ) {
            case Test: {
               subj = "[TEST] " + subj;
               recips = rcptList;
               break;
            }

            case Production: {
               recips = rcptList;
               break;
            }
         }

         try {
            DataSender.smtp(m_From, recips, subj, String.format("%s", msg));
         }        
         catch (Exception ex ) {
            log.error("[JmsProcessor]", ex);
         }

         finally {
           recips = null;
           subj = null;
         }
      }
   }
 
  public static void main(String[] args) {
        
      EdiInvNotConf ediInvNotConf = new EdiInvNotConf();
            
      ediInvNotConf.m_Status = RptServer.RUNNING;
      ediInvNotConf.m_Env = Environment.Test;
      java.util.Properties connProps = new java.util.Properties();
      connProps.put("user", "eis_emery");
      connProps.put("password", "boxer");
      
      StringBuffer Line = new StringBuffer();
      
      try {
         Connection eisConn  = java.sql.DriverManager.getConnection("jdbc:oracle:thin:@10.128.0.9:1521:GROK", connProps);
                  
         ediInvNotConf.m_OraConn = eisConn;
         ediInvNotConf.prepareStatements();
         //ediInvoicesNotConfirmed.buildOutputFile();
         //
         // Build the Captions      
         Line.append("Invoice Number\tAck_Status\t");
         Line.append("Time_sent\tWF_ID\r\n");
         ResultSet ediInvoiceData;
         
         //date formatter from java.utils
         SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-yyyy");

         try {
            ediInvoiceData = ediInvNotConf.m_EdiInvoices.executeQuery();

            while ( ediInvoiceData.next()) {
               
               Line.append(ediInvoiceData.getString("invoice_num") + "\t\t");
               Line.append(ediInvoiceData.getString("ack_status") + "\t\t");
               Line.append(formatter.format(ediInvoiceData.getDate("time_sent")).toString() + "\t");
               Line.append(ediInvoiceData.getString("wf_id") + "\t");

               Line.append("\r\n");
               
            }
            ediInvoiceData.close();
            
           // ediInvNotConf.notifyMis(Line.toString());
            System.out.println(Line.toString());
         }
      
         catch ( Exception ex ) {
            log.fatal("[EdiInvNotConf] exception:", ex);
         }
      
         finally {
            ediInvNotConf.closeStatements();         
         }
         
      } catch (SQLException e) {
         // TODO Auto-generated catch block
         e.printStackTrace();
      }   finally {
            ediInvNotConf.closeStatements();         
      }
   }
}


