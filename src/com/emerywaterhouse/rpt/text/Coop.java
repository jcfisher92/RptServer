/**
 * File: Coop.java
 * Description: Month end file build for coop data.  Builds the Glidden-Valspar and sales extract
 *    files for the previous month.  These files are input to some program used by the emery co-op
 *    administrator.  The files need to be in DOS format, so make sure that there are carriage
 *    return-line feeds at the end of each line.  Individual fields in the files are tab delimited.
 *    
 *    Rewrite of the original Coop report so that it works with the new report server.  Orginal author is 
 *    Paul Davidson
 * <p>
 * @author Paul Davidson
 * @author Jeffrey Fisher
 *
 * Create Data: 05/09/2005
 * Last Update: $Id: Coop.java,v 1.17 2009/09/03 12:42:49 smurdock Exp $
 * 
 * History
 *    $Log: Coop.java,v $
 *    Revision 1.17  2009/09/03 12:42:49  smurdock
 *    deleted three Scotts vendor ids from warehouse and drop ship reports per Tim Reilley
 *
 *    Revision 1.16  2009/09/01 18:22:11  smurdock
 *    decided they wanted cust 152188 instead of 034711 for weber report.  whateer.
 *
 *    Revision 1.15  2009/08/26 18:37:07  smurdock
 *    added weber report for 2 customers
 *
 *    deleted valspar and glidden for all customers from emery co-op report and delted Weber for 2 customers above -- no more double dipping on co-op
 *
 *    Revision 1.14  2009/07/13 15:07:20  jfisher
 *    Changes for imports into the new coop system.
 *
 *    Revision 1.13  2009/06/23 13:21:19  jfisher
 *    Added a month delta to handle go back more than a single month.  Requested by Tim Reily.
 *
 *    Revision 1.12  2009/06/22 17:17:57  jfisher
 *    Modified the file output based on Tim Reily's request.
 *
 *    Revision 1.11  2006/10/24 17:27:48  pdavidso
 *    Included account 20240 in warehouse vendor purchases query
 *
 *    Revision 1.10  2006/10/10 18:05:04  pdavidso
 *    Added override tag to createReport() method
 *
 *    Revision 1.9  2006/06/01 15:26:06  pdavidso
 *    Changed to use accpac account numbers instead of descriptions when
 *    pulling COGS data directly from Accpac.
 *
 *    Revision 1.8  2006/06/01 14:48:51  pdavidso
 *    Updated dropship cogs query to include all dropship cogs accounts
 *
 *    Revision 1.7  2006/03/21 16:06:38  pdavidso
 *    Mods to negative the amount for the vendor warehouse and dropship
 *    purchase files if they came from a credit note.
 *
 *    Revision 1.6  2006/02/28 13:52:56  jfisher
 *    Put the logger instance here for the sub classes to use.
 *
 *    Revision 1.5  2006/02/07 16:52:40  pdavidso
 *    Updated queries for Promax files (dropship and warehouse) to pull invoice data based on the fiscal period and not the date of the invoice.
 *
 *    Revision 1.4  2006/01/31 17:07:21  pdavidso
 *    Format amounts in whse and dropship files to comply with Promax software
 *
 *    Revision 1.3  2006/01/30 19:48:19  pdavidso
 *    Added methods to build dropship and warehouse vendor purchase files
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class Coop extends Report
{
   private static final String gliddenFmt = "%s Sales\t3000\tGlidden Co-op\t%s\t%s\t%s\t%1.2f\t%s\r\n";
   private static final String weberFmt = "%s Sales\t02\tWeber Co-op\t%s\t%s\t%s\t%1.2f\t%s\r\n";
   private static final String salesFmt = "%s Sales\t01\tEmery Co-op\t%s\t%s\t%s\t%1.2f\t%s\r\n";
   private static final String valsparFmt = "%s Sales\t10038\tValspar Co-op\t%s\t%s\t%s\t%1.2f\t%s\r\n";   
   private static final String dsPoFmt = "PO DS %s\t%d\t%s\t%1.2f\t%s\r\n";
   private static final String whsPoFmt = "PO WH %s\t%d\t%s\t%1.2f\t%s\r\n";
   
   private PreparedStatement m_DropshipCost;  // Dropship cost data
   private PreparedStatement m_GliddenData;   // Sales data against Glidden product
   private PreparedStatement m_WeberData;     // Sales data against Weber product fot custs 229598 and 152188
   private PreparedStatement m_ValsparData;   // Sales data against Valspar product
   private PreparedStatement m_WarehouseCost; // Warehouse cost data
   private PreparedStatement m_SalesData;     // Customer sales data

   // The difference between the current month and the month to run against.
   // Allows for running back dated months.
   private int m_Delta;    
   
   /**
    * Default constructor.  Calls the base class constructor so we get the job setup.
    * Initialize any data members specific to this class.
    */
   public Coop()
   {
      super();
      
      //
      // Default this to one so that everything works as normal.
      m_Delta = 1;
   }

   /**
    * Builds the Glidden coop data file using data from the inv_dtl table.
    */
   private void buildGliddenFile()
   {
      Calendar calendar = new GregorianCalendar();
      java.util.Date date = null;
      SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
      SimpleDateFormat dfm = new SimpleDateFormat("MMMM");
      StringBuffer fileName = new StringBuffer();
      ResultSet rs = null;
      String lastDay = null;
      String salesNum = null;
      StringBuffer line = new StringBuffer(1024);
      FileOutputStream outFile = null;
     
      try {
         //
         // Get the date of the last day of the previous month in mm/dd/yyyy format
         calendar.setTime(new java.util.Date(System.currentTimeMillis()));
         calendar.set(Calendar.MONTH, calendar.get(Calendar.MONTH) - m_Delta);
         calendar.set(Calendar.DAY_OF_MONTH, calendar.getActualMaximum(Calendar.DAY_OF_MONTH));
         date = new java.util.Date(calendar.getTimeInMillis());
         
         lastDay = df.format(date);
         salesNum = dfm.format(date);
         
         //
         // Build the file name
         dfm = new SimpleDateFormat("MM");
         fileName.append("coop-");
         fileName.append(dfm.format(date));
         fileName.append(calendar.get(Calendar.YEAR));
         fileName.append("-glidden.txt");
         m_FileNames.add(fileName.toString());

         //
         // Open the Glidden file for writing.
         outFile = new FileOutputStream(m_FilePath + fileName);
         writeCoopFileHeadings(outFile);
         
         rs = m_GliddenData.executeQuery();

         //
         // Write the Glidden lines
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);

            line.append(
               String.format(gliddenFmt, 
                  salesNum,
                  rs.getString("cust"),
                  rs.getString("name"),
                  rs.getString("sale_type"),
                  rs.getDouble("sales"),
                  lastDay
               )
            );
            
            outFile.write(line.toString().getBytes());        
         }
      }

      catch ( Exception ex ) {
         log.error("exception: ", ex);
      }
      
      finally {
         closeRSet(rs);
         
         //
         // Close file output stream
         if ( outFile != null ) {
            try {
               outFile.close();
               outFile = null;
            }
            
            catch ( Exception e ) {               
            }            
         }

         calendar = null;
         date = null;
         df = null;
         dfm = null;
         fileName = null;
         lastDay = null;
         line = null;
         salesNum = null;
         rs = null;
      }
   }
   
   
   /**
    * Builds (selected - 2 customers only) Weber coop data file using data from the inv_dtl table.
    */
  
   private void buildWeberFile()
   {
      Calendar calendar = new GregorianCalendar();
      java.util.Date date = null;
      SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
      SimpleDateFormat dfm = new SimpleDateFormat("MMMM");
      StringBuffer fileName = new StringBuffer();
      ResultSet rs = null;
      String lastDay = null;
      String salesNum = null;
      StringBuffer line = new StringBuffer(1024);
      FileOutputStream outFile = null;
     
      try {
         //
         // Get the date of the last day of the previous month in mm/dd/yyyy format
         calendar.setTime(new java.util.Date(System.currentTimeMillis()));
         calendar.set(Calendar.MONTH, calendar.get(Calendar.MONTH) - m_Delta);
         calendar.set(Calendar.DAY_OF_MONTH, calendar.getActualMaximum(Calendar.DAY_OF_MONTH));
         date = new java.util.Date(calendar.getTimeInMillis());
         
         lastDay = df.format(date);
         salesNum = dfm.format(date);
         
         //
         // Build the file name
         dfm = new SimpleDateFormat("MM");
         fileName.append("coop-");
         fileName.append(dfm.format(date));
         fileName.append(calendar.get(Calendar.YEAR));
         fileName.append("-weber.txt");
         m_FileNames.add(fileName.toString());

         //
         // Open the Weber file for writing.
         outFile = new FileOutputStream(m_FilePath + fileName);
         writeCoopFileHeadings(outFile);
         
         rs = m_WeberData.executeQuery();

         //
         // Write the Weber lines
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);

            line.append(
               String.format(weberFmt, 
                  salesNum,
                  rs.getString("cust"),
                  rs.getString("name"),
                  rs.getString("sale_type"),
                  rs.getDouble("sales"),
                  lastDay
               )
            );
            
            outFile.write(line.toString().getBytes());        
         }
      }

      catch ( Exception ex ) {
         log.error("exception: ", ex);
      }
      
      finally {
         closeRSet(rs);
         
         //
         // Close file output stream
         if ( outFile != null ) {
            try {
               outFile.close();
               outFile = null;
            }
            
            catch ( Exception e ) {               
            }            
         }

         calendar = null;
         date = null;
         df = null;
         dfm = null;
         fileName = null;
         lastDay = null;
         line = null;
         salesNum = null;
         rs = null;
      }
   }

   
   

   /**
    * Builds the customer sales coop data file using data from the inv_hdr table.
    * (08/26/2009 sm) m_SalesData now uses inv_dtl to exclude certain vendors and customers
    */
   private void buildSalesFile()
   {
      Calendar calendar = new GregorianCalendar();
      java.util.Date date = null;
      SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
      SimpleDateFormat dfm = new SimpleDateFormat("MMMM");
      StringBuffer fileName = new StringBuffer();
      ResultSet rs = null;
      String lastDay = null;
      String salesNum = null;
      StringBuffer line = new StringBuffer(1024);
      FileOutputStream outFile = null;
      
      try {
         //
         // Get the date of the last day of the current month in mm/dd/yyyy format
         calendar.setTime(new java.util.Date(System.currentTimeMillis()));
         calendar.set(Calendar.MONTH, calendar.get(Calendar.MONTH) - m_Delta);
         calendar.set(Calendar.DAY_OF_MONTH, calendar.getActualMaximum(Calendar.DAY_OF_MONTH));
         date = new java.util.Date(calendar.getTimeInMillis());
         
         lastDay = df.format(date);
         salesNum = dfm.format(date);
         
         //
         // Build the file name
         dfm = new SimpleDateFormat("MM");
         fileName.append("coop-");
         fileName.append(dfm.format(date));
         fileName.append(calendar.get(Calendar.YEAR));
         fileName.append("-sales.txt");
         m_FileNames.add(fileName.toString());

         //
         // Open the sales coop file for writing
         outFile = new FileOutputStream(m_FilePath + fileName);
         writeCoopFileHeadings(outFile);

         rs = m_SalesData.executeQuery();

         //
         // Write the Glidden lines
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            line.append(
               String.format(salesFmt, 
                  salesNum,
                  rs.getString("cust"),
                  rs.getString("name"),
                  rs.getString("sale_type"),
                  rs.getDouble("sales"),
                  lastDay
               )
            );
            outFile.write(line.toString().getBytes());            
         }
      }

      catch ( Exception ex ) {
         log.error("exception: ", ex);
      }
      
      finally {
         closeRSet(rs);

         //
         // Close file output stream
         if ( outFile != null ) {
            try {
               outFile.close();
               outFile = null;
            }
            
            catch ( Exception e ) {               
            }
            
         }

         calendar = null;
         date = null;
         df = null;
         dfm = null;
         fileName = null;
         lastDay = null;
         line = null;
         salesNum = null;
         rs = null;
      }
   }
   
   /**
    * Builds dropship purchases cost file using data from the accpac tables.
    * Fixed length file - line length = 33 characters.
    */
   private void buildDropShipFile()
   {
      Calendar calendar = new GregorianCalendar();
      SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
      SimpleDateFormat dfm = new SimpleDateFormat("MMMM");
      StringBuffer fileName = new StringBuffer();      
      StringBuffer line = new StringBuffer(1024);
      java.util.Date date = null;
      FileOutputStream outFile = null;
      ResultSet rs = null;      
      String salesNum = null;
      String poDate = null;

      try {
         //
         // Get the date of the last day of the current month in mm/dd/yyyy format
         calendar.setTime(new java.util.Date(System.currentTimeMillis()));
         calendar.set(Calendar.MONTH, calendar.get(Calendar.MONTH) - m_Delta);
         calendar.set(Calendar.DAY_OF_MONTH, calendar.getActualMaximum(Calendar.DAY_OF_MONTH));
         date = new java.util.Date(calendar.getTimeInMillis());
         
         poDate = df.format(date);
         salesNum = dfm.format(date);
         
         //
         // Build the file name
         dfm = new SimpleDateFormat("MM");
         fileName.append("dropship-");
         fileName.append(dfm.format(date));
         fileName.append(calendar.get(Calendar.YEAR));
         fileName.append("-purchases.txt");
         m_FileNames.add(fileName.toString());

         outFile = new FileOutputStream(m_FilePath + fileName);
         writePoFileHeadings(outFile);
         
         rs = m_DropshipCost.executeQuery();

         //
         // Write the file lines
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            
            line.append(
               String.format(dsPoFmt, 
                  salesNum,
                  rs.getLong("idvend"),
                  rs.getString("name"),
                  rs.getDouble("cost"),
                  poDate
               )
            );
            
            outFile.write(line.toString().getBytes());
         }
      }

      catch ( Exception ex ) {
         log.error("exception: ", ex);
      }
      
      finally {
         closeRSet(rs);

         //
         // Close file output stream
         if ( outFile != null ) {
            try {
               outFile.close();
               outFile = null;
            }
            
            catch ( Exception e ) {               
            }
         }

         salesNum = null;
         poDate = null;
         calendar = null;
         date = null;
         df = null;
         dfm = null;
         fileName = null;
         line = null;
         rs = null;
      }
   }

   /**
    * Builds the Valspar coop file for importing into the JDA application.
    */
   private void buildValsparFile()
   {
      Calendar calendar = new GregorianCalendar();
      java.util.Date date = null;
      SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
      SimpleDateFormat dfm = new SimpleDateFormat("MMMM");      
      StringBuffer fileName = new StringBuffer();
      ResultSet rs = null;
      String lastDay = null;
      String salesNum = null;
      StringBuffer line = new StringBuffer(1024);
      FileOutputStream outFile = null;
      
      try {         
         //
         // Get the date of the last day of the previous month in mm/dd/yyyy format
         calendar.setTime(new java.util.Date(System.currentTimeMillis()));
         calendar.set(Calendar.MONTH, calendar.get(Calendar.MONTH) - m_Delta);
         calendar.set(Calendar.DAY_OF_MONTH, calendar.getActualMaximum(Calendar.DAY_OF_MONTH));
         date = new java.util.Date(calendar.getTimeInMillis());
         
         lastDay = df.format(date);
         salesNum = dfm.format(date);
         
         //
         // Build the file name
         dfm = new SimpleDateFormat("MM");
         fileName.append("coop-");
         fileName.append(dfm.format(date));
         fileName.append(calendar.get(Calendar.YEAR));
         fileName.append("-valspar.txt");
         m_FileNames.add(fileName.toString());

         //
         // Open the Valspar file for writing.
         outFile = new FileOutputStream(m_FilePath + fileName);
         writeCoopFileHeadings(outFile);
         
         rs = m_ValsparData.executeQuery();
         
         while ( rs.next() && m_Status != RptServer.STOPPED ) {
            line.setLength(0);
            line.append(
               String.format(valsparFmt, 
                  salesNum,
                  rs.getString("cust"),
                  rs.getString("name"),
                  rs.getString("sale_type"),
                  rs.getDouble("sales"),
                  lastDay
               )
            );
            
            outFile.write(line.toString().getBytes());
         }
      }

      catch ( Exception ex ) {
         log.error("exception: ", ex);
      }
      
      finally {
         closeRSet(rs);
         
         //
         // Close file output stream
         if ( outFile != null ) {
            try {
               outFile.close();
               outFile = null;
            }
            
            catch ( Exception e ) {               
            }            
         }

         calendar = null;
         date = null;
         df = null;
         dfm = null;
         fileName = null;
         lastDay = null;
         salesNum = null;
         line = null;
         rs = null;
      }
   }

   /**
    * Builds warehouse purchasees cost file using data from the accpac tables.
    * Fixed length file - line length = 33 characters.
    */
   private void buildWarehouseFile()
   {
      Calendar calendar = new GregorianCalendar();
      SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
      SimpleDateFormat dfm = new SimpleDateFormat("MMMM");
      StringBuffer fileName = new StringBuffer();      
      StringBuffer line = new StringBuffer(1024);
      java.util.Date date = null;
      FileOutputStream outFile = null;
      ResultSet rs = null;      
      String salesNum = null;
      String poDate = null;

      try {
         //
         // Get the date of the last day of the current month in mm/dd/yyyy format
         calendar.setTime(new java.util.Date(System.currentTimeMillis()));
         calendar.set(Calendar.MONTH, calendar.get(Calendar.MONTH) - m_Delta);
         calendar.set(Calendar.DAY_OF_MONTH, calendar.getActualMaximum(Calendar.DAY_OF_MONTH));
         date = new java.util.Date(calendar.getTimeInMillis());
         
         poDate = df.format(date);
         salesNum = dfm.format(date);
         
         //
         // Build the file name
         dfm = new SimpleDateFormat("MM");
         fileName.append("warehouse-");
         fileName.append(dfm.format(date));
         fileName.append(calendar.get(Calendar.YEAR));
         fileName.append("-purchases.txt");
         m_FileNames.add(fileName.toString());

         outFile = new FileOutputStream(m_FilePath + fileName);
         writePoFileHeadings(outFile);
         
         rs = m_WarehouseCost.executeQuery();
         
         //
         // Write the file lines
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            
            line.append(
               String.format(whsPoFmt, 
                  salesNum,
                  rs.getLong("idvend"),
                  rs.getString("name"),
                  rs.getDouble("cost"),
                  poDate
               )
            );
            
            outFile.write(line.toString().getBytes());            
         }
      }

      catch ( Exception ex ) {
         log.error("exception: ", ex);
      }
      
      finally {
         closeRSet(rs);
         
         //
         // Close file output stream
         if ( outFile != null ) {
            try {
               outFile.close();
               outFile = null;
            }
            
            catch ( Exception e ) {               
            }
         }
         calendar = null;
         date = null;
         df = null;
         dfm = null;
         fileName = null;
         line = null;
         rs = null;
         poDate = null;
         salesNum = null;
      }
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      //
      // Close Glidden data statement
      if ( m_GliddenData != null ) {
         try {
            m_GliddenData.close();
         }
         catch ( SQLException e )
         {}

         m_GliddenData = null;
      }

      //
      // Close Valspar data statement
      if ( m_ValsparData != null ) {
         try {
            m_ValsparData.close();
         }
         catch ( SQLException e )
         {}

         m_ValsparData = null;
      }

      //
      // Close Sales data statement
      if ( m_SalesData != null ) {
         try {
            m_SalesData.close();
         }
         catch ( SQLException e )
         {}

         m_SalesData = null;
      }
      
      // Close dropship cost statement
      if ( m_DropshipCost != null ) {
         try {
            m_DropshipCost.close();
         }
         catch ( SQLException e )
         {}

         m_DropshipCost = null;
      }
      
      // Close warehouse cost statement
      if ( m_WarehouseCost != null ) {
         try {
            m_WarehouseCost.close();
         }
         catch ( SQLException e )
         {}

         m_WarehouseCost = null;
      }
   }
   
   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {      
      boolean result = false;
      m_Status = RptServer.RUNNING;
      
      try {
         setCurAction("starting coop file reports");
         m_OraConn = m_RptProc.getOraConn();
         
         if ( prepareStatements() ) {
            buildGliddenFile();
            buildWeberFile();
            buildValsparFile();
            buildSalesFile();
            buildDropShipFile();
            buildWarehouseFile();
            
            result = true;
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
      
      return result;
   }
   
   /**
    * Prepares the sql queries for execution.
    *
    * @return boolean true if the statements were prepared, false of not
    * @throws  Exception
    */
   private boolean prepareStatements() throws Exception
   {
      String datestr;
      boolean prepared = false;
      Calendar calendar = new GregorianCalendar();    
      StringBuffer sql = new StringBuffer();
      int month;
      String mthstr;
      int year;
      
      if ( m_OraConn != null ) {
         //
         // Set up the calendar for the previous month.
         calendar.setTime(new java.util.Date(System.currentTimeMillis()));
         calendar.set(Calendar.MONTH, calendar.get(Calendar.MONTH) - m_Delta);
         month = calendar.get(Calendar.MONTH) + 1;
         mthstr = Integer.toString(month);
         year = calendar.get(Calendar.YEAR);
   
         if ( month < 10 ) {
            mthstr = "0" + mthstr;
         }
         
         // Build date string used in SQL below that indicates the 1st of this month
         datestr = mthstr + "01" + year;
            
         //
         // Build the Glidden sales statement
         sql.setLength(0);
         sql.append("select nvl(parent_id, customer_id) cust, name, sale_type, sum(ext_sell) sales ");
         sql.append("from customer ");
         sql.append("join inv_dtl on inv_dtl.cust_nbr = customer.customer_id and ");         
         sql.append("   invoice_date >= to_date('" + datestr + "', 'mmddyyyy') and ");
         sql.append("   invoice_date <= last_day(to_date('" + datestr + "', 'mmddyyyy')) and ");
         sql.append("   vendor_name like 'GLIDDEN%' and ext_sell <> 0 and sale_type = 'WAREHOUSE' ");         
         sql.append("group by nvl(parent_id, customer_id), name, sale_type ");
         sql.append("order by cust");
         m_GliddenData = m_OraConn.prepareStatement(sql.toString());
 
         //
         // Build the Weber sales statement
         sql.setLength(0);
         sql.append("select nvl(parent_id, customer_id) cust, name, sale_type, sum(ext_sell) sales ");
         sql.append("from customer ");
         sql.append("join inv_dtl on inv_dtl.cust_nbr = customer.customer_id and ");         
         sql.append("   invoice_date >= to_date('" + datestr + "', 'mmddyyyy') and ");
         sql.append("   invoice_date <= last_day(to_date('" + datestr + "', 'mmddyyyy')) and ");
         sql.append("   vendor_name like 'WEBER%' and ext_sell <> 0 and sale_type = 'WAREHOUSE' ");         
         sql.append("where customer_id in ('152188','229598') ");         
         sql.append("group by nvl(parent_id, customer_id), name, sale_type ");
         sql.append("order by cust");
         m_WeberData = m_OraConn.prepareStatement(sql.toString());
         
         
         //
         // Build the Valspar sales statement
         sql.setLength(0);
         sql.append("select nvl(parent_id, customer_id) cust, name, sale_type, sum(ext_sell) sales ");
         sql.append("from customer ");
         sql.append("join inv_dtl on inv_dtl.cust_nbr = customer.customer_id and ");
         sql.append("   invoice_date >= to_date('" + datestr + "', 'mmddyyyy') and ");
         sql.append("   invoice_date <= last_day(to_date('" + datestr + "', 'mmddyyyy')) and ");
         sql.append("   vendor_name like 'VALSPAR%' and ext_sell <> 0 and sale_type = 'WAREHOUSE' ");         
         sql.append("group by nvl(parent_id, customer_id), name, sale_type ");
         sql.append("order by cust");
         m_ValsparData = m_OraConn.prepareStatement(sql.toString());
            
         //
         // Build the customers sales statement
         // (08/26/2009 sm) rewritten to NOT include Valspar (all customers), Glidden(all customers)
         // and Weber (two customers only)
         // The above three cases are covered in their own reports.  No more double dipping.
         // hard coded three Scotts vendors out of this report per time Reilley and Abbey Pierson  sm 09/02/2009
         sql.setLength(0);
         sql.append("select nvl(parent_id, customer.customer_id) cust, name, sale_type, sum(ext_sell) sales ");
         sql.append("from customer ");
         sql.append("join cust_market_view cmv on cmv.customer_id = customer.customer_id and ");
         sql.append("   market = 'CUSTOMER TYPE' and class <> 'EMPLOYEE' ");
         sql.append("join inv_dtl on inv_dtl.cust_nbr = customer.customer_id and ");         
         sql.append("inv_dtl.sale_type = 'WAREHOUSE' and ");         
         sql.append("inv_dtl.invoice_date >= to_date('" + datestr + "', 'mmddyyyy') and ");         
         sql.append("inv_dtl.invoice_date <=  last_day(to_date('" + datestr + "', 'mmddyyyy'))  and ");         
         sql.append("inv_dtl.ext_sell <> 0 and inv_dtl.sale_type = 'WAREHOUSE' and ");         
         sql.append("inv_dtl.vendor_nbr not in ");         
         sql.append("   (select vval.vendor_id from vendor vval where vval.name like 'VALSPAR%')  and ");         
         sql.append("inv_dtl.vendor_nbr not in ");         
         sql.append("   (select vglid.vendor_id from vendor vglid where vglid.name  like 'GLIDDEN%') and ");         
         sql.append("inv_dtl.inv_dtl_id  not in  ");         
         sql.append("    (select inv_dtl_id from inv_dtl ");         
         sql.append("         where vendor_nbr in ");         
         sql.append("           (select vendor_id from vendor where vendor_name like 'WEBER%' ");         
         sql.append("              and cust_nbr in ('152188','229598')and ");         
         sql.append("              invoice_date >= to_date('" + datestr + "', 'mmddyyyy') and ");         
         sql.append("              invoice_date <=  last_day(to_date('" + datestr + "', 'mmddyyyy' )) ))  ");         
         sql.append("group by nvl(parent_id, customer.customer_id), name, sale_type ");
         sql.append("order by cust");
         m_SalesData = m_OraConn.prepareStatement(sql.toString());
                  
         //
         // Build dropship cost statement by pulling data directly out of Accpac.
         // Note that the amount is negatived if its a credit note.
         // From the Accpac view documentation for invoice header and the idtrx field:
         //    12 = Invoice - Summary Entered
         //    13 = Invoice - Recurring Charge
         //    22 = Debit Note - Summary Entered
         //    32 = Credit Note - Summary Entered
         //    40 = Interest Charge
         // hard coded three Scotts vendors out of this report per time Reilley and Abbey Pierson  sm 09/02/2009
         sql.setLength(0);
         sql.append("select "); 
         sql.append("   idvend, name, round(sum(decode(idtrx, 32, -amtdist, amtdist)), 2) cost "); 
         sql.append("from "); 
         sql.append("   emeryd.apibd, "); 
         sql.append("   emeryd.apibh, ");  
         sql.append("   emeryd.glamf, ");
         sql.append("   vendor ");
         sql.append("where "); 
         sql.append("   fiscper = '" + mthstr + "' and ");
         sql.append("   fiscyr = '" + year + "' and "); 
         sql.append("   acsegval02 = '30410' and "); 
         sql.append("   apibd.cntbtch = apibh.cntbtch and "); 
         sql.append("   apibd.cntitem = apibh.cntitem and "); 
         sql.append("   apibd.idglacct = glamf.acctfmttd and ");
         sql.append("   apibh.idvend not in ('703463','703462','701940') and ");
         sql.append("   apibh.idvend = vendor.vendor_id ");
         sql.append("group by idvend, name "); 
         sql.append("order by idvend ");         
         m_DropshipCost = m_OraConn.prepareStatement(sql.toString());
         
         //
         // Build warehouse cost statement by pulling data directly out of Accpac.
         // Note that the amount is negatived if its a credit note.
         // From the Accpac view documentation for invoice header and the idtrx field:
         //    12 = Invoice - Summary Entered
         //    13 = Invoice - Recurring Charge
         //    22 = Debit Note - Summary Entered
         //    32 = Credit Note - Summary Entered
         //    40 = Interest Charge
         sql.setLength(0);
         sql.append("select \r\n"); 
         sql.append("   idvend, name, round(sum(decode(idtrx, 32, -amtdist, amtdist)), 2) cost \r\n"); 
         sql.append("from \r\n"); 
         sql.append("   emeryd.apibd, \r\n"); 
         sql.append("   emeryd.apibh, \r\n");
         sql.append("   emeryd.glamf, \r\n");
         sql.append("   vendor \r\n"); 
         sql.append("where \r\n");
         sql.append("   fiscper = '" + mthstr + "' and \r\n");
         sql.append("   fiscyr = '" + year + "' and \r\n");
         sql.append("   acsegval02 in ('10210', '20240') and \r\n"); 
         sql.append("   apibd.cntbtch = apibh.cntbtch and \r\n"); 
         sql.append("   apibd.cntitem = apibh.cntitem and \r\n"); 
         sql.append("   apibd.idglacct = glamf.acctfmttd and \r\n");
         sql.append("   apibh.idvend not in ('703463','703462','701940') and ");         
         sql.append("   apibh.idvend = vendor.vendor_id \r\n");
         sql.append("group by idvend, name \r\n"); 
         sql.append("order by idvend ");
         m_WarehouseCost = m_OraConn.prepareStatement(sql.toString());
                  
         prepared = true;
      }
      
      return prepared;
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {      
      int pcount = params.size();
      Param param = null;
      int tmp;
      
      try {
         for ( int i = 0; i < pcount; i++ ) {
            param = params.get(i);
   
            if ( param.name.equals("delta") ) {
               tmp = Integer.parseInt(param.value);
               
               if ( tmp > 0 )
                  m_Delta = tmp;
               else
                  log.warn("coop report: invalid month delta, must be greater than 0; using default");
            }
         }         
      }
      
      finally {         
         param = null;         
      }
   }
   
   /**
    * Generic method to write the file headings for the coop files.  All the columns 
    * are the same for each of the files.
    * 
    * @param fos The stream object that is being used to write to the file.
    * @throws IOException
    */
   private void writeCoopFileHeadings(FileOutputStream fos) throws IOException 
   {
      StringBuffer line = new StringBuffer();
      
      try {      
         line.append("Sales Number\t");
         line.append("Vendor ID\t");
         line.append("Vendor Name\t");
         line.append("Customer ID\t");
         line.append("Customer Name\t");
         line.append("Sale Type\t");
         line.append("Sales Amt\t");
         line.append("Tran Date\r\n");
         
         fos.write(line.toString().getBytes());
      }
      
      finally {
         line = null;
      }      
   }
   
   /**
    * Method to write the data headings for the PO files.
    * 
    * @param fos The stream object for writing to the file.
    * @throws IOException
    */
   private void writePoFileHeadings(FileOutputStream fos) throws IOException
   {
      StringBuffer line = new StringBuffer();
      
      try {      
         line.append("PO Number\t");
         line.append("Vendor ID\t");
         line.append("Vendor Name\t");
         line.append("PO Amount\t");         
         line.append("PO Date\r\n");
         
         fos.write(line.toString().getBytes());
      }
      
      finally {
         line = null;
      }      
   }
}