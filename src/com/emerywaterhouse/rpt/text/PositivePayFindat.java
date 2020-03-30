/**
 * File: PositivePay.java
 * Description: Fixed length TD Banknorth check verification file built using data from the AccPac payment details 
 *    table: findat.aptcpr.  Code is based on Paul Davidson's source code DaRemittance.java, which is based on
 *     Jeff Fisher's source code in DAFile.java.
 *
 * @author Seth Murdock
 *
 * Create Date: 03/28/2008
 * Last Update: $Id: PositivePay.java,v 1.2 2008/03/28 23:25:23 jfisher Exp $
 *
 * History:
 *    $Log: PositivePay.java,v $
 *    Revision 1.2  2008/03/28 23:25:23  jfisher
 *    added cvs tags and fixed some indentation
 *
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.text.SimpleDateFormat;





import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class PositivePayFindat extends Report 
{
   private String m_batch_start;
   private String m_batch_end;   
   private PreparedStatement m_PositivePayData;   // For remittance detail file records
      
   /**
    * default constructor
    */
   public PositivePayFindat()
   {
      super();
     
      m_MaxRunTime = RptServer.HOUR * 2;
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    */
   @Override
   public void finalize() throws Throwable
   {      
      m_PositivePayData = null;
      
      super.finalize();
   }
   
   /**
    * Builds the Positive Pay file based on the query selection criteria.
    *
    * @return  boolean
    *    true if the file was created.
    *    false if there was some sort of error.
    */
   private boolean buildOutputFile()
   {
	  String AccountNbr = "0241178905";
      String BankNbr = "910010000000";
      String HeaderTranCode = "BH";
      String IssueTranCode = "40";
      String IssueSeqNbr = "999";;
      String IssueFlag = "0";
      String FillerSevenZeroes = "0000000";
      String FillerFortySix = "                                              ";
      String batchnbr = "";
      StringBuffer EmailText = new StringBuffer(1024);
      char[] WholeIssueFiller = new char[80];
      char[] WholeHeaderFiller = new char[80];
      StringBuffer line = new StringBuffer(1024);
      int BatchCount = 0;
      double CheckTotal = 0;
      FileOutputStream outFile = null;
      ResultSet IssueData = null;
      int thisbatch = -1;
      
      boolean Result = false;

      try {
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
         m_PositivePayData.setString(1,m_batch_start);        
         m_PositivePayData.setString(2,m_batch_end);        

         //   wrtie the check data line by line, write header on batch number change
         IssueData = m_PositivePayData.executeQuery();

         while ( IssueData.next() && m_Status == RptServer.RUNNING ) {
            if (!(IssueData.getInt("cntbtch") == thisbatch)) {
               BatchCount++;
               Arrays.fill(WholeHeaderFiller, ' ');
               line.setLength(0);
               line.append(WholeHeaderFiller);

               // Header Bank Number code (column 1, length=12)
               line.replace(0, 12, BankNbr);

               // Header account number (column 2, length=10) 
               line.replace(12, 22, AccountNbr);

               // Header Tran Code(column 3, length=2)
               line.replace(22, 24, HeaderTranCode);

               // Header Filler(column 4, length=7)
               line.replace(24, 31, FillerSevenZeroes);

               // Header Filler Code(column 5, length=46)
               line.replace(31, 77, FillerFortySix);

               // Header Tran Code(column 4, length=7)
               batchnbr = IssueData.getString("cntbtch");
               if (batchnbr.length() > 3) // TDBank takes only a 3 digit batch number, we will use the last three digits
                  batchnbr = batchnbr.substring(batchnbr.length() -3, batchnbr.length() );
               line.replace((80 - batchnbr.length()), 80,  batchnbr);

               line.append("\r\n");

               outFile.write(line.toString().getBytes());

               thisbatch = IssueData.getInt("cntbtch"); // keep trak of number of batches (only for email info -- bank does not care)        

            }  //end of header write

            //on to the Issue write   
            // Prefill the char array with spaces.
            line.setLength(0);

            Arrays.fill(WholeIssueFiller, ' ');

            line.setLength(0);
            line.append(WholeIssueFiller);

            // Issue Bank Number code (column 1, length=12)
            line.replace(0, 12, BankNbr);

            // Issue account number (column 2, length=10) 
            line.replace(12, 22, AccountNbr);

            // Issue Tran Code(column 3, length=2)
            line.replace(22, 24, IssueTranCode);

            // Issue Date(column 4, length=6)
            line.replace(24, 30, IssueData.getString("chk_date"));

            // Issue Amount (column 5, length=11)
            line.replace(30, 41, IssueData.getString("chk_amt"));

            // Keep track of sum of checks (only for email info -- bank does not care)
            CheckTotal = CheckTotal + IssueData.getLong("chk_amt");

            // Issue Check Number (column 6, length=10)
            line.replace(41, 51, IssueData.getString("chk_num"));

            // Issue Sequence Number(column 7, length=3)
            line.replace(51, 54, IssueSeqNbr);

            // Issue Payee (column 8, length=13)
            line.replace(54, 67, IssueData.getString("payee"));

            // Issue Filler (column 9, length=12)
            line.replace(67, 79, "            ");  //This should be FillerTwelve in a more perfect world

            // Issue Flag (column 10, length=1)
            line.replace(79, 80, IssueFlag);

            line.append("\r\n");
            outFile.write(line.toString().getBytes());
         }
         EmailText.setLength(0);

         //we override the email text message here, mostly so whoever is running the report will get the check total
         EmailText.append("The Positive Pay report has finished running.\r\n");
         EmailText.append("The report file has been attached.\r\n");
         EmailText.append("\r\n");
         EmailText.append("Batches: ");
         EmailText.append(BatchCount);
         EmailText.append("\r\n");
         EmailText.append("Total: ");
         EmailText.append(CheckTotal/100);


         m_RptProc.setEmailMsg(EmailText.toString()); 

         Result = true;
      }

      catch ( Exception ex ) {
         log.error("exception:", ex);
      }

      finally {
         // Close remittance data result sets
         closeRSet(IssueData);

         // Close file output stream
         if ( outFile != null ) {
            try {
               outFile.close();
            }
            catch ( IOException ioe )
            {}

            outFile = null;
         }   
      }
      
      return Result;
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_PositivePayData);
   }
   
   /**
    * Implements the base class abstract method.  Creates a connection to Oracle for data.
    * Then builds the output file.
    * 
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    * @see com.emerywaterhouse.rpt.text.DARemittance#buildOutputFile()
    */
   @Override
   public boolean createReport()
   {      
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {
         m_OraConn = m_RptProc.getOraConn();
         
         if ( prepareStatements() ) {
            created = buildOutputFile();
         }
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
    * Doing way too much formatting here.  So bite me.
    * @throws  SQLException
    */
   private boolean prepareStatements() throws SQLException
   {
      StringBuffer sql = new StringBuffer();
      
      if ( m_OraConn == null )
         return false;
      
      sql.append("select ");
      sql.append("   cntbtch, ");      
      sql.append("   substr(idrmit, -10) as chk_num, ");
      sql.append("   to_char(to_date(to_char(datermit),'YYYYMMDD'),'MMDDYY') as chk_date, ");
      sql.append("   lpad(to_char(round(amtrmit * 100)),11,'0') as chk_amt, ");
      sql.append("   rpad(substr(namermit,1,13),13,' ') as payee ");
      sql.append("from ");
      sql.append("   findat.aptcr ");
      sql.append("where ");
      sql.append("   cntbtch >= ? and ");
      sql.append("   cntbtch <=  ? ");
      sql.append("order by cntbtch, chk_num ");
      
      m_PositivePayData = m_OraConn.prepareStatement(sql.toString());
            
      return true;
   }
   
   /**
    * Sets the parameters for the report.
    *    param(0) = batch number
    *    
    * @param params ArrayList<Param> - list of report parameters.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fname = new StringBuffer();
      int pcount = params.size();                     
      Param param = null;                             
      SimpleDateFormat formatter = new SimpleDateFormat ("yyyyMMddHHmmss");
      Date day = new Date();
       
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
         if ( param.name.equals("batchstart") )
            m_batch_start  = param.value;
         
         if ( param.name.equals("batchend") )
            m_batch_end  = param.value;
      }
      
      //
      // Build the file name.
      fname.append(formatter.format( day ));
      fname.append("-");      
      fname.append("PositivePay-fdat.txt");
      m_FileNames.add(fname.toString());
   }
}