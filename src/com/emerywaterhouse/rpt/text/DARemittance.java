/**
 * File: DARemittance.java
 * Description: Fixed length DA remittance file built using data from the AccPac payment details 
 *    table: emeryd.aptcp.  Code is based on Jeff Fisher's source code in DAFile.java.
 *
 * @author Paul Davidson
 *
 * Create Data: 12/07/2005
 * Last Update: $Id: DARemittance.java,v 1.12 2008/02/19 19:02:41 pdavidson Exp $
 * 
 * History:
 *    $Log: DARemittance.java,v $
 *    Revision 1.12  2008/02/19 19:02:41  pdavidson
 *    Fixed formatting of negative numbers in formatAmount() method
 *
 *    Revision 1.11  2006/06/23 16:33:13  pdavidso
 *    Reworked for preventing duplicate records when joining with inv header (apibh)
 *
 *    Revision 1.10  2006/06/23 16:12:46  pdavidso
 *    Reworked for preventing duplicate records when joining with inv header (apibh)
 *
 *    Revision 1.9  2006/02/28 13:52:56  jfisher
 *    Put the logger instance here for the sub classes to use.
 *
 *    Revision 1.8  2006/01/19 21:19:38  pdavidso
 *    Joined vendor ids of payment details and invoice header to prevent duplicate records.
 *
 *    Revision 1.7  2006/01/19 19:36:16  pdavidso
 *    Getting program# from correct opt field.  Fixed dup records behaviour.
 *
 *    Revision 1.4  2005/12/14 20:39:46  pdavidso
 *    Fixed null pointer exception bugs
 *
 *    Revision 1.3  2005/12/13 19:43:28  pdavidso
 *    Added code for writing trailer record
 *
 *    Revision 1.2  2005/12/12 23:04:00  pdavidso
 *    Added function to format amount values
 *
 *    Revision 1.1  2005/12/08 00:03:30  pdavidso
 *    Initial commit.
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class DARemittance extends Report
{
   private long m_BatchNo;
   private PreparedStatement m_RemitDetailData;   // For remittance detail file records
   private PreparedStatement m_RemitTrailerData;  // For remitttance trailer file record
   private PreparedStatement m_RebateOpt;         // Gets optional rebate value
   private PreparedStatement m_InvHdr;            // Gets the invoice header data
   
   /**
    * default constructor
    */
   public DARemittance()
   {
      super();
      
      m_FileNames.add("da_remit.txt");
      m_MaxRunTime = RptServer.HOUR * 2;
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    */
   @Override
   public void finalize() throws Throwable
   {      
      m_RemitDetailData = null;
      m_RemitTrailerData = null;
      m_RebateOpt = null;
      m_InvHdr = null;
      
      super.finalize();
   }
   
   /**
    * Builds the DA remittance file based on the query selection criteria.
    *
    * @return  boolean
    *    true if the file was created.
    *    false if there was some sort of error.
    */
   private boolean buildOutputFile()
   {
      long cntBtch = 0;
      long cntItem = 0;
      String dateInvC = null;
      String discountAmt;
      SimpleDateFormat dtFmt;      
      char[] filler = new char[85];
      String grossAmt;
      StringBuffer line = new StringBuffer(1024);
      int lineCount = 0;
      String lineCountStr;
      FileOutputStream outFile = null;
      ResultSet invHdrData = null;
      ResultSet remitDetailData = null;
      ResultSet remitTrailerData = null;
      ResultSet rebateOptData = null;
      String rebateCode;
      String transmissionDate;
      String vndProgNum;
      String vndInvNum;
      String vndName;
      String vndInvDate;
      String wireAmt;
      
      boolean Result = false;

      try {
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
         m_RemitDetailData.setLong(1, m_BatchNo);
         remitDetailData = m_RemitDetailData.executeQuery();

         // Prefill the char array with spaces.
         Arrays.fill(filler, ' ');

         while ( remitDetailData.next() && m_Status == RptServer.RUNNING ) {
            lineCount++;
            line.setLength(0);
            line.append(filler);

            // Record code (column 1, length=1)
            line.replace(0, 1, "1");
            
            // Member number (column 2, length=3) 
            line.replace(1, 4, "480");

            // Vendor program# (column 3, length=6)
            vndProgNum = remitDetailData.getString("vendprogno");
            if ( vndProgNum != null && vndProgNum.length() > 0 ) {
               if ( vndProgNum.length() > 6 )
                  vndProgNum = vndProgNum.substring(0, 6);
               
               line.replace(4, 4+vndProgNum.length(), vndProgNum);
            }
            
            // Vendor invoice# (column 4, length=22)
            vndInvNum = remitDetailData.getString("IDINVC");
            if ( vndInvNum != null && vndInvNum.length() > 0 ) {
               if ( vndInvNum.length() > 22 )
                  vndInvNum = vndInvNum.substring(0, 22);
            
               line.replace(10, 10+vndInvNum.length(), vndInvNum);
            }
            
            // Vendor name (column 5, length=20)
            vndName = remitDetailData.getString("VENDNAME");
            if ( vndName != null && vndName.length() > 0 ) {
               if ( vndName.length() > 20 )
                  vndName = vndName.substring(0, 20);
            
               line.replace(32, 32+vndName.length(), vndName);
            }
            
            // Gross invoice amount (column 6, length=12, e.g. format="+/-00012112.00")
            grossAmt = remitDetailData.getString("grossamt");
            grossAmt = formatAmount(grossAmt);
            line.replace(52, 52+grossAmt.length(), grossAmt);
            
            // Discount amount (column 7, length=12, e.g. format="+/-00012112.00")
            discountAmt = remitDetailData.getString("discountamt");
            discountAmt = formatAmount(discountAmt);
            line.replace(64, 64+discountAmt.length(), discountAmt);
            
            try {
               dateInvC = null;
               
               m_InvHdr.setString(1, remitDetailData.getString("IDINVC"));
               m_InvHdr.setString(2, remitDetailData.getString("IDVEND"));
               invHdrData = m_InvHdr.executeQuery();
               
               if (invHdrData.next() ) {
                  cntBtch = invHdrData.getLong("CNTBTCH");
                  cntItem = invHdrData.getLong("CNTITEM");
                  dateInvC = invHdrData.getString("DATEINVC");
               }
            }
            
            finally {
               closeRSet(invHdrData);
            }

            // Rebate transaction code (column 8, length=1)
            try {
               rebateCode = null;
               
               m_RebateOpt.setLong(1, cntBtch);
               m_RebateOpt.setLong(2, cntItem);
               rebateOptData = m_RebateOpt.executeQuery();
               
               if (rebateOptData.next() ) {
                  rebateCode = rebateOptData.getString("VALUE");
               }
                  
               if ( rebateCode != null && rebateCode.length() > 0 ) {
                  if ( rebateCode.length() > 1 )
                     rebateCode = rebateCode.substring(0, 1);
               
                  line.replace(76, 77, rebateCode);
               }
            }
            
            finally {
               closeRSet(rebateOptData);
            }
            
            // Vendor invoice date (column 9, length=8, format=mmddccyy)
            vndInvDate = dateInvC;
            if ( vndInvDate != null && vndInvDate.length() > 0 ) {
               // Date value is in ccyymmdd format in the table, switch
               // this to mmddccyy format in the file
               vndInvDate = vndInvDate.substring(4, vndInvDate.length()) + vndInvDate.substring(0, 4);
               line.replace(77, 85, vndInvDate);
            }
            
            // Chop off anything that might have gone past our line limit.
            if (line.length() > 85 )
               line.setLength(85);

            line.append("\r\n");
            outFile.write(line.toString().getBytes());
         }

         // Build trailer record
         if ( lineCount > 0 && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            line.append(filler);

            // Record code (column 1, length=1)
            line.replace(0, 1, "2");
            
            // Member number (column 2, length=3) 
            line.replace(1, 4, "480");
          
            // Transmission date (column 3, length=8)
            dtFmt = new SimpleDateFormat("MMddyyyy");
            transmissionDate = dtFmt.format(new Date(System.currentTimeMillis()));
            line.replace(4, 12, transmissionDate);
            
            // Total detail records sent (column 4, length=6)
            line.replace(12, 18, "000000");
            lineCountStr = Integer.toString(lineCount);
            line.replace(18-lineCountStr.length(), 18, lineCountStr);
            
            m_RemitTrailerData.setLong(1, m_BatchNo);
            remitTrailerData = m_RemitTrailerData.executeQuery();
            
            if ( remitTrailerData.next() ) {
               // Total gross amount (column 5, length=12, e.g. format="+/-00012112.00")
               grossAmt = remitTrailerData.getString("totgrossamt");
               grossAmt = formatAmount(grossAmt);
               line.replace(18, 18+grossAmt.length(), grossAmt);
               
               // Total discount amount (column 6, length=12, e.g. format="+/-00012112.00")
               discountAmt = remitTrailerData.getString("totdiscountamt");
               discountAmt = formatAmount(discountAmt);
               line.replace(30, 30+discountAmt.length(), discountAmt);
               
               // Total to-pay amount (column 7, length=12)
               wireAmt = remitTrailerData.getString("totwireamt");
               wireAmt = formatAmount(wireAmt);
               line.replace(42, 42+wireAmt.length(), wireAmt);
            }
            
            // Chop off anything that might have gone past our line limit.
            if ( line.length() > 85 )
               line.setLength(85);

            outFile.write(line.toString().getBytes());
         }
         
         Result = true;
      }

      catch ( Exception ex ) {
         log.error("exception:", ex);
      }

      finally {
         // Close remittance data result sets
         closeRSet(remitDetailData);
         closeRSet(remitTrailerData);
         
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
      closeStmt(m_RemitDetailData);
      closeStmt(m_RemitTrailerData);
      closeStmt(m_RebateOpt);
      closeStmt(m_InvHdr);
   }
   
   /**
    * Implements the base class abstract method.  Creates a connection to Sage SQL Server for data.
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
         m_SageConn = m_RptProc.getSageConn();
         
         if ( prepareStatements() ) {
            setCurAction("Building DA remittance file for batch# " + m_BatchNo);
            created = buildOutputFile();
            setCurAction("DA remittance file build for batch# " + m_BatchNo + " is complete");
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
    * Formats the input value to something like this: +/-00012112.00, for the purposes 
    * of this remittance file.  Note that all amounts are right justified and left zero
    * filled.
    * 
    * @param amt String - Some amount value
    * @return String - Formatted amount.
    */
   private String formatAmount(String amt)
   {
      String amtFiller = "+00000000.00";
      StringBuffer amtBuf = new StringBuffer(12);
      int decLen;
      int decPos;
      boolean negative;
      
      if ( Double.parseDouble(amt) == 0.00 ) {
         return amtFiller;
      }
      
      amtBuf.setLength(0);
      amtBuf.append(amtFiller);
      negative = amt.charAt(0) == '-';
      
      if ( negative ) {
         amtBuf.replace(0, 1, "-");
         amt = amt.substring(1, amt.length());
      }
      
      decPos = amt.indexOf(".");

      if ( decPos == -1 ) {
         amtBuf.replace(9-amt.length(), 9, amt);
      }
      else {
         decLen = amt.substring(decPos+1, amt.length()).length();
         
         if ( decLen == 1 ) {
            amt = amt + "0";
         }
         
         amtBuf.replace(12-amt.length(), 12, amt);
      }
      
      return amtBuf.toString();
   }
   
   /**
    * Prepares the sql queries for execution.
    * @throws  SQLException
    */
   private boolean prepareStatements() throws SQLException
   {
      StringBuffer sql = new StringBuffer();
      
      if ( m_EdbConn == null )
         return false;
      
      sql.append("select "); 
      sql.append("   apveno.VALUE as vendprogno, ");
      sql.append("   aptcp.IDINVC, ");
      sql.append("   aptcp.IDVEND, ");
      sql.append("   apven.VENDNAME, ");
      sql.append("   convert(varchar(20), round((aptcp.AMTPAYM - aptcp.AMTADJTOT + aptcp.AMTERNDISC), 2)) as grossamt, ");
      sql.append("   convert(varchar(20), round((aptcp.AMTERNDISC), 2)) as discountamt ");
      sql.append("from ");
      sql.append("   EMEDAT.dbo.APTCP aptcp ");																																		 
      sql.append("left join (select VENDORID, VALUE from EMEDAT.dbo.APVENO where OPTFIELD = 'DAVNDID') apveno on apveno.VENDORID = aptcp.IDVEND "); 
      sql.append("left join EMEDAT.dbo.APVEN apven on apven.VENDORID = aptcp.IDVEND");																							 
      sql.append("where ");
      sql.append("   aptcp.BATCHTYPE = 'PY' and ");
      sql.append("   aptcp.CNTBTCH = ? ");
      m_RemitDetailData = m_SageConn.prepareStatement(sql.toString());
   
      sql.setLength(0);
      sql.append("select ");  
      sql.append("   convert(varchar(20), round(sum(aptcp.AMTPAYM - aptcp.AMTADJTOT + aptcp.AMTERNDISC), 2)) as totgrossamt, "); 
      sql.append("   convert(varchar(20), round(sum(aptcp.AMTERNDISC), 2)) as totdiscountamt, ");
      sql.append("   convert(varchar(20), round(sum(aptcp.AMTPAYM), 2)) as totwireamt "); 
      sql.append("from "); 
      sql.append("   EMEDAT.dbo.APTCP aptcp ");																																				  
      sql.append("left join (select VENDORID, VALUE from EMEDAT.dbo.APVENO where OPTFIELD = 'DAVNDID') apveno on apveno.VENDORID = aptcp.IDVEND ");  
      sql.append("left join EMEDAT.dbo.APVEN apven on apven.VENDORID = aptcp.IDVEND "); 																						  
      sql.append("where "); 
      sql.append("   aptcp.BATCHTYPE = 'PY' and "); 
      sql.append("   aptcp.CNTBTCH = ? "); 
      m_RemitTrailerData = m_SageConn.prepareStatement(sql.toString());
      
      sql.setLength(0);
      sql.append("select VALUE "); 
      sql.append("from EMEDAT.dbo.APIBHO ");
      sql.append("where OPTFIELD = 'REBATEABLE' and ");
      sql.append("   CNTBTCH = ? and ");
      sql.append("   CNTITEM = ?");
      m_RebateOpt = m_SageConn.prepareStatement(sql.toString());

      sql.setLength(0);
      sql.append("select CNTBTCH, CNTITEM, DATEINVC "); 
      sql.append("from EMEDAT.dbo.APIBH ");
      sql.append("where IDINVC = ? and ");
      sql.append("   IDVEND = ? ");
      m_InvHdr = m_SageConn.prepareStatement(sql.toString());
      
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
      m_BatchNo = Long.parseLong(params.get(0).value);
   }
}