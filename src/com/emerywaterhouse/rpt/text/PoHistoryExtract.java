/**
 * Title:			PoHistoryExtract.java
 * Description:	    Creates a text file in fixed column format of POs closed during a
 *                  given time frame.  
 * Company:			Emery-Waterhouse
 * @author			prichter
 * @version			1.0
 * <p>
 * Create Date: Feb 21, 2007
 * Last Update: $Id: PoHistoryExtract.java,v 1.8 2007/08/02 09:53:46 prichter Exp $
 * <p>
 * History:
 *   $Log: PoHistoryExtract.java,v $
 *   Revision 1.8  2007/08/02 09:53:46  prichter
 *   Fixed an error in the query that was causing an incorrect 1st receipt date under some circumstances.
 *
 *   Revision 1.7  2007/06/26 21:25:14  prichter
 *   Rebuilt the main query to reduce the amount of memory needed and speed up the report.
 *
 *   Revision 1.6  2007/06/21 20:02:37  prichter
 *   Changed the 1st receipt date to the dates that the 0110 (receiver scheduled) was created
 *
 *   Revision 1.5  2007/03/13 16:51:38  prichter
 *   Changed the precision on cube, weight, and cost.  Added PO date.  Fixed a formatting problem caused by vendors with no address.
 *
 *   Revision 1.4  2007/03/09 14:42:10  prichter
 *   Report crashed if fascor data had been purged for a given receiver.
 *
 *   Revision 1.3  2007/03/06 17:46:28  prichter
 *   Fixed a couple of substring index out of bounds problems
 *
 *   Revision 1.2  2007/02/27 19:26:58  prichter
 *   Truncated the state and country fields to 10 and 20 characters.
 *
 *   Revision 1.1  2007/02/27 15:46:02  prichter
 *   Initial add.
 *
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Arrays;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class PoHistoryExtract extends Report
{
   private PreparedStatement m_Address; // Vendor Address
   private PreparedStatement m_Po;      // Main query. Retrieves PO history
   
   private String m_BegDate;            // Begin date parameter
   private String m_EndDate;            // End date parameter
   

   public PoHistoryExtract() {
      super();
   }

   /**
    * Builds the data file using data from the pos system and Fascor
    * @return boolean - true if the extract was built successfully
    */
   private boolean buildOutputFile()
   {
      StringBuffer line = new StringBuffer();
      FileOutputStream outFile = null;
      ResultSet poRs = null;
      ResultSet addrRs = null;
      String tmp = null;
      
      try {         
         
         //
         // Open the output file
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0));
         
         m_Po.setString(1, m_BegDate);
         m_Po.setString(2, m_EndDate);
         m_Po.setString(3, m_BegDate);
         m_Po.setString(4, m_EndDate);
         m_Po.setString(5, m_BegDate);
         m_Po.setString(6, m_EndDate);
         poRs = m_Po.executeQuery();
         
         while ( poRs.next() ) {
            line.setLength(0);
            
            tmp = poRs.getString("po_nbr") + "-" + poRs.getString("rcvr_nbr");
            line.append(rFill(tmp, 20));
            line.append(" ");
            line.append(rFill(poRs.getString("vendor_id"), 6));
            line.append(" ");
            line.append(rFill(poRs.getString("vendor_name"), 75));
            line.append(" ");
            
            // Vendor address
            m_Address.setInt(1, poRs.getInt("vendor_id"));
            addrRs = m_Address.executeQuery();
            if ( addrRs.next() ) {
               line.append(rFill(addrRs.getString("city"), 50));
               line.append(" ");
               line.append(rFill(addrRs.getString("state"), 10));
               line.append(" ");
               line.append(rFill(addrRs.getString("country"), 20));
               line.append(" ");
            }
            
            else {
               char[] Filler = new char[83];
               Arrays.fill(Filler, ' ');
               line.append(Filler);
               Filler = null;
            }
            
            closeRSet(addrRs);
            
            line.append(rFill(poRs.getString("department"), 30));
            line.append(" ");
            
            if ( poRs.getString("carrier") != null ) {
               line.append(rFill(poRs.getString("carrier"), 4));
               line.append(" ");
            }
            
            else
               line.append("     ");

            // Date that the receipt showed up on the loading dock
            if ( poRs.getString("first_receipt") != null ) {
               //  the rcvr_msg timestamps have the format yyyymmddhhmmss
               tmp = poRs.getString("first_receipt").substring(0, 8);
               line.append(tmp.substring(4,6) + "/" + tmp.substring(6,8) + "/" + tmp.substring(2, 4));
               line.append(" ");
            }
            else
               line.append("         ");

            // Date that the receipt complete message was created
            if ( poRs.getString("last_receipt") != null ) {
               //  the rcvr_msg timestamps have the format yyyymmddhhmmss
               tmp = poRs.getString("last_receipt").substring(0, 8);
               line.append(tmp.substring(4,6) + "/" + tmp.substring(6,8) + "/" + tmp.substring(2, 4));
               line.append(" ");
            }
            else
               line.append("         ");

            // Date of the first putaway
            if ( poRs.getString("first_putaway") != null ) {
               //  the rcvr_msg timestamps have the format yyyymmddhhmmss
               tmp = poRs.getString("first_putaway").substring(0, 8);
               line.append(tmp.substring(4,6) + "/" + tmp.substring(6,8) + "/" + tmp.substring(2, 4));
               line.append(" ");
            }
            else
               line.append("         ");

            // Date of the last putaway
            if ( poRs.getString("last_putaway") != null ) {
               //  the rcvr_msg timestamps have the format yyyymmddhhmmss
               tmp = poRs.getString("last_putaway").substring(0, 8);
               line.append(tmp.substring(4,6) + "/" + tmp.substring(6,8) + "/" + tmp.substring(2, 4));
               line.append(" ");
            }
            else
               line.append("         ");

            line.append(lFill(poRs.getString("cost"), 10));
            line.append(" ");
            line.append(lFill(poRs.getString("weight"), 10));
            line.append(" ");
            line.append(lFill(poRs.getString("cube"), 10));
            line.append(" ");
            line.append(lFill(poRs.getString("lines"), 10));
            line.append(" ");
            line.append(poRs.getString("po_date"));

            line.append("\r\n");
            outFile.write(line.toString().getBytes());            
         }
         
         return true;
      }

      catch ( Exception ex ) {
         log.error("Error processing PO-Receiver: " + tmp);
         log.error("exception", ex);
      }
      
      finally {
         if ( outFile != null ) {
            try {
               outFile.close();
            }
            catch ( Exception e ) {
               log.error("exception", e);
            }
            
            outFile = null;
         }
         
         closeRSet(poRs);
         closeRSet(addrRs);
         poRs = null;
         addrRs = null;
      }
      
      return false;
   }
   
   /**
    * Closes and nullifies prepared statements
    *
    */
   private void closeStatements()
   {
      closeStmt(m_Po);
      closeStmt(m_Address);
      m_Po = null;
      m_Address = null;      
   }

   /** 
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
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
    * 
    * @return true if the statements were succssfully prepared
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer();
      boolean isPrepared = false;
      
      if ( m_OraConn != null ) {
         try {
            sql.append("select po_hdr.po_nbr, rcvr_po_hdr.rcvr_nbr, upper(rcvr_hdr.carrier) carrier, \r\n"); 
            sql.append("       po_hdr.vendor_id, vendor.name vendor_name,    \r\n");
            sql.append("       emery_dept.name department,  \r\n");
            sql.append("       to_char(po_hdr.po_date, 'MM/dd/yy') po_date, rcvr_po_hdr.receipt_date, \r\n");   
            sql.append("       sum(round(rcvr_dtl.qty_received * \"Cube\" / 1728,3)) cube,   \r\n");
            sql.append("       sum(round(rcvr_dtl.qty_received * unit_cost,2)) cost,   \r\n");
            sql.append("       sum(round(rcvr_dtl.qty_received * item.weight,3)) weight,   \r\n");
            sql.append("       count(*) lines,   \r\n");
            sql.append("       rcpt.first_receipt, rcpt.last_receipt, \r\n"); 
            sql.append("       putaway.first_putaway, putaway.last_putaway    \r\n");
            sql.append("from po_hdr     \r\n");
            sql.append("join vendor_dept on vendor_dept.vendor_id = po_hdr.vendor_id \r\n");   
            sql.append("join vendor on vendor.vendor_id = po_hdr.vendor_id    \r\n");
            sql.append("join emery_dept on emery_dept.dept_id = vendor_dept.dept_id   \r\n"); 
            sql.append("join rcvr_po_hdr on rcvr_po_hdr.po_hdr_id = po_hdr.po_hdr_id and rcvr_po_hdr.receipt_date is not null \r\n");   
            sql.append("join rcvr_hdr on rcvr_hdr.warehouse = rcvr_po_hdr.warehouse and \r\n");
            sql.append("                 rcvr_hdr.rcvr_nbr = rcvr_po_hdr.rcvr_nbr \r\n");
            sql.append("join rcvr_dtl on rcvr_dtl.rcvr_po_hdr_id = rcvr_po_hdr.rcvr_po_hdr_id \r\n");  
            sql.append("join sku_master on sku_master.\"SKU\" = rcvr_dtl.item_nbr   \r\n");
            sql.append("join item on item.item_id = rcvr_dtl.item_nbr   \r\n");
            sql.append("join (select rcvr_msg.warehouse, rcvr_msg.rcvr_nbr, \r\n");
            sql.append("             min(time_stamp) first_receipt, max(time_stamp) last_receipt \r\n");
            sql.append("      from rcvr_msg \r\n");
            sql.append("      join rcvr_po_hdr on rcvr_po_hdr.warehouse = rcvr_msg.warehouse and \r\n");
            sql.append("           rcvr_po_hdr.rcvr_nbr = rcvr_msg.rcvr_nbr and \r\n");
            sql.append("           rcvr_po_hdr.receipt_date >= to_date(?, 'mm/dd/yyyy') and \r\n");   
            sql.append("           rcvr_po_hdr.receipt_date <= to_date(?, 'mm/dd/yyyy')    \r\n");
            sql.append("      where rcvr_msg.msg_type in ('0130','0110') \r\n");
            sql.append("      group by rcvr_msg.warehouse, rcvr_msg.rcvr_nbr) rcpt \r\n");
            sql.append("      on rcpt.warehouse = rcvr_po_hdr.warehouse and \r\n");
            sql.append("         rcpt.rcvr_nbr = rcvr_po_hdr.rcvr_nbr  \r\n");
            sql.append("join (select rcvr_msg.warehouse, rcvr_msg.rcvr_nbr, rcvr_msg.po_nbr, \r\n");
            sql.append("             min(time_stamp) first_putaway, max(time_stamp) last_putaway \r\n");
            sql.append("      from rcvr_msg \r\n");
            sql.append("      join rcvr_po_hdr on rcvr_po_hdr.warehouse = rcvr_msg.warehouse and \r\n");
            sql.append("           rcvr_po_hdr.rcvr_nbr = rcvr_msg.rcvr_nbr and \r\n");
            sql.append("           rcvr_po_hdr.po_nbr = rcvr_msg.po_nbr and  \r\n");
            sql.append("           rcvr_po_hdr.receipt_date >= to_date(?, 'mm/dd/yyyy') and \r\n");   
            sql.append("           rcvr_po_hdr.receipt_date <= to_date(?, 'mm/dd/yyyy')    \r\n");
            sql.append("      where rcvr_msg.msg_type = '0210'  \r\n");
            sql.append("      group by rcvr_msg.warehouse, rcvr_msg.rcvr_nbr, rcvr_msg.po_nbr) putaway \r\n");
            sql.append("      on putaway.warehouse = rcvr_po_hdr.warehouse and \r\n");
            sql.append("         putaway.rcvr_nbr = rcvr_po_hdr.rcvr_nbr and \r\n");
            sql.append("         putaway.po_nbr = rcvr_po_hdr.po_nbr  \r\n");
            sql.append("where rcvr_po_hdr.receipt_date >= to_date(?, 'mm/dd/yyyy') and \r\n");   
            sql.append("      rcvr_po_hdr.receipt_date <= to_date(?, 'mm/dd/yyyy') and \r\n");   
            sql.append("      po_hdr.status = 'CLOSED' and \r\n");   
            sql.append("      po_hdr.po_nbr not like 'TR%' and \r\n");   
            sql.append("      po_hdr.po_nbr not like 'M%' and \r\n"); 
            sql.append("      rcvr_dtl.qty_received is not null and \r\n"); 
            sql.append("      rcvr_dtl.qty_received > 0 \r\n");   
            sql.append("group by po_hdr.po_nbr, rcvr_po_hdr.rcvr_nbr, rcvr_hdr.carrier, \r\n");   
            sql.append("         po_hdr.vendor_id, vendor.name,  \r\n");
            sql.append("         emery_dept.name, po_hdr.po_date, rcvr_po_hdr.receipt_date, \r\n"); 
            sql.append("         rcpt.first_receipt, rcpt.last_receipt, putaway.first_putaway, putaway.last_putaway \r\n");       
            m_Po = m_OraConn.prepareStatement(sql.toString()); 
            
            sql.setLength(0);
            sql.append("select nvl(substr(city, 1, 50), ' ') city, nvl(substr(state, 1, 10), ' ') state, ");
            sql.append("       nvl(substr(country, 1, 20), ' ') country from vendor_address ");
            sql.append("where vendor_id = ? and description = 'SHIPPING' ");
            m_Address = m_OraConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("exception:", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("PoHistoryExtract.prepareStatements - null oracle connection");
      
      return isPrepared;
   }

   /**
    * Utility to pad the string spaces on the left with to the desired length
    * @param str String - the original string
    * @param len - the desired length
    * @return String - the padded string
    */
   private String lFill(String str, int len) 
   {
      if ( str.length() > len )
         return str.substring(0, len);
      
      char[] Filler = new char[len - str.length()];
      Arrays.fill(Filler, ' ');      
      return String.valueOf(Filler) + str;
   }
   
   /**
    * Utility to pad the string spaces on the right with to the desired length
    * @param str String - the original string
    * @param len - the desired length
    * @return String - the padded string
    */
   private String rFill(String str, int len) 
   {
      if ( str.length() > len )
         return str.substring(0, len);
      
      char[] Filler = new char[len - str.length()];
      Arrays.fill(Filler, ' ');      
      return str + String.valueOf(Filler);
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();      
      String tmp = Long.toString(System.currentTimeMillis());
                  
      fileName.append("po_history-");      
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".txt");
      m_FileNames.add(fileName.toString());
      
      m_BegDate = params.get(0).value.trim();
      m_EndDate = params.get(1).value.trim();
   }

}
