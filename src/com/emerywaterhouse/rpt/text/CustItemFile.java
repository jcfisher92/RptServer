/**
 * File: CustItemFile.java
 * Description: Builds a report of customers inventory based on an msi ack number.  This is the converted
 *    report for the new report server.
 *    Original author was Jeffrey Fisher
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 05/10/2005
 * Last Update: $Id: CustItemFile.java,v 1.11 2010/08/21 22:18:53 jfisher Exp $
 * 
 * History
 *    $Log: CustItemFile.java,v $
 *    Revision 1.11  2010/08/21 22:18:53  jfisher
 *    Mods to handle orders that came in via XML
 *
 *    Revision 1.10  2010/02/07 04:51:22  smurdock
 *    changed default item sell price to customer specific sell price
 *
 *    Revision 1.9  2008/10/29 21:24:12  jfisher
 *    Fixed potential null warnings.
 *
 *    Revision 1.8  2007/01/31 17:06:03  jheric
 *    Fixed bug while setting date, just parse the date use simple date parser.
 *
 *    Revision 1.7  2006/03/03 14:23:00  jfisher
 *    Added the error log entry date as a parameter to the query.  See cr# 822
 *
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileOutputStream;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class CustItemFile extends Report
{
   public final int MSI    = 0;
   public final int ASYNC  = 1;
   public final int BISYNC = 2;
   public final int X12    = 3;
   public final int NA     = 4;

   private PreparedStatement m_ItemData;
   private PreparedStatement m_FileData;
   private PreparedStatement m_HdrId;
   
   private GregorianCalendar m_Date;
   private String m_CustId;
   private String m_Ack;
   private String m_SrcRef;
   private int m_FileType;
      
   /**
    * 
    * default constructor
    */
   public CustItemFile()
   {
      super();
      
      m_Date = new GregorianCalendar();
      m_FileType = MSI;
   }
   

   /**
    * Cleanup when were done.
    *
    * @throws Throwable
    */
   public void finalize() throws Throwable
   {
      super.finalize();
   }

   /**
    * Executes the queries and builds the output file.
    *
    * @return true if the file was created, false if there was an error.
    */
   private boolean buildOutputFile()
   {      
      StringBuffer Line = new StringBuffer(1024);
      NumberFormat nf = NumberFormat.getInstance();
      FileOutputStream OutFile = null;
      boolean Result = false;
      ResultSet HdrId = null;
      ResultSet FileData = null;
      ResultSet ItemData = null;
      Date errDate = null;

      int Qty = 0;
      double Sell = 0.0;
      double ExtSell = 0.0;
      double Retail = 0.0; 															// added double Retail
      String upc = "";
      String ItemId = "";
      String Nrha = "";
      String Flc = "";
      String Desc = "";
      
      //
      // Set error date.
      errDate = new Date(m_Date.getTimeInMillis());      
            
      nf.setMaximumFractionDigits(3);
      
      try {
         OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
         
         //
         // Write the heading first.
         Line.append("NRHA\t");
         Line.append("FLC\t");
         Line.append("UPC\t");
         Line.append("SKU\t");
         Line.append("ITEM DESCRIPTION\t");
         Line.append("QTY\t");
         Line.append("COST\t");
         Line.append("EXT COST\t");
         Line.append("ACK # \t");					// changed to \t to add additional column to header
         Line.append("RETAIL PRICE \r\n");				// added this line to to complete header of spreadsheet

         OutFile.write(Line.toString().getBytes());
         Line.delete(0, Line.length());
         
         //
         // Pull the data from the order_line_error table.
         m_HdrId.setString(1, m_CustId);
         m_HdrId.setDate(2, errDate);
         HdrId = m_HdrId.executeQuery();
         
         while ( HdrId.next() ) {
            m_FileData.setInt(1, HdrId.getInt(1));
            FileData = m_FileData.executeQuery();

            while ( FileData.next() ) {
               upc = FileData.getString("upc"); 
               if (upc == null || upc.isEmpty()) {
                   upc = FileData.getString("cust_sku"); 
               }
               if (upc == null || upc.isEmpty()) {
                   upc = FileData.getString("emery_item_id"); 
               }
               if (upc == null || upc.isEmpty()) {
                   upc = FileData.getString("item_id"); 
               }

               ItemId = FileData.getString("item_id");
               Qty = FileData.getInt("quantity");

               if ( ItemId != null ) {
                   try {
                  m_ItemData.setString(1, m_CustId);
                  m_ItemData.setString(2, m_CustId);
                  m_ItemData.setString(3, ItemId);
                  
                  ItemData = m_ItemData.executeQuery();

                  if ( ItemData.next() ) {
                     Nrha = ItemData.getString("nrha_id");
                     Flc = ItemData.getString("flc_id");
                     Desc = ItemData.getString("description");
                     Sell = ItemData.getDouble("sell");
                     ExtSell = Sell * Qty;
                     Retail = ItemData.getDouble("retail");                             // added value "Retail" to list
                     
                     if( upc == null )
                    	 upc = ItemData.getString("upc_code");
                  }


                     ItemData.close();
                  }

                  catch ( Exception ex ) {                     
                  }
               }
               else {
                  ItemId = "";
                  Nrha = "";
                  Flc = "";
                  Desc = "";
                  Sell = 0.0;
                  ExtSell = 0.0;
                  Retail = 0.0;                             // added line to maintain consistancy
               }

               Line.append(Nrha + "\t");
               Line.append(Flc + "\t");
               Line.append(upc + "\t");
               Line.append(ItemId + "\t");
               Line.append(Desc + "\t");
               Line.append(Qty + "\t");
               Line.append(nf.format(Sell) + "\t");
               Line.append(nf.format(ExtSell) + "\t");
               Line.append(m_Ack + "\t");                      // changed \r\n to \t to add additional retail price to line
               Line.append(nf.format(Retail) + "\r\n");

               OutFile.write(Line.toString().getBytes());
               Line.delete(0, Line.length());
            }

            Line.append(m_Date.getTime() + "\r\n");
            OutFile.write(Line.toString().getBytes());

            try {
               FileData.close();
            }

            catch ( Exception ex ) {               
            }
         }

         Result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("The MSI Customer Item File report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         
         log.fatal("report error", ex);
      }

      finally {
         Line = null;

         if ( OutFile != null ) {
            try {
               OutFile.close();
               OutFile = null;
            }

            catch( Exception e ) {
            }            
         }

         if ( HdrId != null ) {
            try {
               HdrId.close();
               HdrId = null;               
            }

            catch( Exception e ) {               
            }            
         }

         if ( FileData != null ) {
            try {
               FileData.close();
               FileData = null;
            }

            catch( Exception e ) {              
            }            
         }
      }
      

      return Result;
   }

   /**
    * Closes all open statements, closes database connections and calls the system.gc method
    * to notify the vm that it should cleanup.
    */
   protected void cleanup()
   {     
      m_Date = null;
      m_ItemData = null;
      m_FileData = null;
      m_HdrId = null;    

      closeStatements();

      setCurAction("");
      System.gc();
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      try {
         if ( m_ItemData != null )
            m_ItemData.close();
      }

      catch ( Exception ex ) {
      }

      try {
         if ( m_FileData != null )
            m_FileData.close();
      }

      catch ( Exception ex ) {
      }

      try {
         if ( m_HdrId != null )
            m_HdrId.close();
      }

      catch ( Exception ex ) {
         
      }
   }

   /**
    * Creates the file name for the output file based on the customer id, ack number, and
    * the date.
    * <p>
    * The file name format is:
    *    FileType or nothing
    *    Custmer ID
    *    Ack# or source ref
    *    year
    *    month
    *    day
    */
   private void createFileName()
   {
      final String FileTypes[] = {"msi", "async", "bisync", "x12"};
      StringBuffer fileName = new StringBuffer();

      try {
         if ( m_SrcRef != null && m_SrcRef.length() > 0 ) {
            fileName.append(m_CustId);
            fileName.append('-');
            fileName.append(m_SrcRef);
         }
         else {
            fileName.append(FileTypes[m_FileType]);
            fileName.append(m_CustId);
            fileName.append('-');
            fileName.append(m_Ack);
         }
         
         fileName.append('-');
         fileName.append(m_Date.get(Calendar.YEAR));
         fileName.append((m_Date.get(Calendar.MONTH)+1));
         fileName.append(m_Date.get(Calendar.DATE));
         fileName.append(".dat");
      
         m_FileNames.add(fileName.toString());
      }
      
      finally {
         fileName = null;
      }
   }

   /**
    * Handles connecting to the db server and creating the report.
    * 
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {      
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         m_OraConn = m_RptProc.getOraConn();
         m_EdbConn = m_RptProc.getEdbConn();
                  
         if ( prepareStatements() ) {            
            createFileName();
            created = buildOutputFile();
         }
      }
      
      catch ( Exception ex ) {
         log.fatal("exception:", ex);
      }
      
      finally {
        cleanup();
        
        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }
      
      return created;
   }
   
   /**
    * overrides base class method for logging.
    * @return The id of the customer from the params passed to the report.
    * @see com.emerywaterhouse.rpt.server.Report#getCustId()
    */
   public String getCustId()
   {
      return m_CustId;
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements() throws Exception
   {      
      StringBuffer sql = new StringBuffer();
      boolean prepared = false;

      System.out.println("preparing statements");
      
      if ( m_OraConn != null ) {
         try {
            sql.append("select mdc.nrha_id, item.flc_id, item.description, item_upc.upc_code, ");
            sql.append("nvl(cust_procs.GetSellPrice(?,item.item_id), 0) sell ,");                 // added comma at end of this line to continue SELECT items
            sql.append("cust_procs.GetRetailPrice(?,item.item_id) retail ");	           // added this line to calculate and query retail price 
            sql.append("from item, item_upc, flc, mdc ");                                    // added order_line and order_header to FROM 
            sql.append("where item.item_id = ? and item.flc_id = flc.flc_id and mdc.mdc_id = flc.mdc_id ");
            sql.append(" and item_upc.item_id = item.item_id ");     // added link for order_header.item_id = order_line.order_id
            m_ItemData = m_EdbConn.prepareStatement(sql.toString());
            																						
            sql.setLength(0);																		
            sql.append("select emery_item_id, cust_sku, upc, item_id, quantity ");
            sql.append("from order_line_error where ohe_id = ? order by to_number(line_seq)");
            m_FileData = m_EdbConn.prepareStatement(sql.toString());
      
            sql.setLength(0);
            sql.append("select ohe_id from order_header_error where customer_id = ? ");         
            //
            // Transitional code.  Eventually all orders will come through the same place
            // and will only have a source ref.  For now there are multiple formats.  If the 
            // order comes through the internet/mobile, then just use the source ref field.
            // 07/23/2010 jcf
            if ( m_SrcRef != null && m_SrcRef.length() > 0 )
               sql.append(String.format("and source_ref = '%s' and ", m_SrcRef));
            else   
               sql.append(String.format("and source_ref like '%%%s' and ", m_Ack));
            
            sql.append("trunc(error_date) = ?");
            m_HdrId = m_EdbConn.prepareStatement(sql.toString());
                       
            prepared = true;
         }
         
         finally {
            sql = null;
         }
      }
      
      return prepared;
   }
   
   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    * 
    * Pulls the parameters from the param list of the job object.
    * We need at least the first three params: file type, cust#, and Ack#
    */
   public void setParams(ArrayList<Param> params)
   {
      Param param = null;
      SimpleDateFormat smp = new SimpleDateFormat("MM/dd/yyyy");
      int pcount = params.size();
            
      try {
         for ( int i = 0; i < pcount; i++ ) {
            param = params.get(i);
   
            if ( param.name.equals("filetype") ) {
               m_FileType = Integer.parseInt(param.value);
            }
            
            if ( param.name.equals("custid") ) {
               m_CustId = param.value;
            }
            
            if ( param.name.equals("ack") ) {
               m_Ack = param.value;
            }
            
            if ( param.name.equals("ackdate") ) {
               try{
                  m_Date.setTime(smp.parse(param.value));
               }
               
               catch(Exception e) {
                  log.error("exception", e);
               }
            }
            
            if ( param.name.equals("srcref") ) {
               m_SrcRef = param.value;
            }
         }         
      }
      
      finally {         
         param = null;
         smp = null;
      }      
   }
}
