/**
 * File: RollingCustSales.java
 * Description: Rolling 12 month customer sales report.
 *    This is the replacement for the previous customer sales report.  It's been changed to fit into the
 *    new report server.
 *
 * @author Jeffrey Fisher
 *
 * Create Data: 04/06/2005
 * Last Update: $Id: RollingCustSales.java,v 1.17 2013/09/25 18:03:42 jfisher Exp $
 *
 * History:
 *    $Log: RollingCustSales.java,v $
 *    Revision 1.17  2013/09/25 18:03:42  jfisher
 *    Additional change to use current buyer not historical buyer.
 *
 *    Revision 1.16  2011/01/08 12:58:43  jfisher
 *    added the "PAINT" class
 *
 *    Revision 1.15  2009/02/18 16:13:18  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    08/26/2005 - Modified the file naming so it would work when multiple users tried to run
 *       the same report at the same time. jcf
 *
 *    04/06/2005 - Modified to work within the new report server structure. - jcf
 *
 *    03/25/2005 - Added log4j logging. jcf
 *
 *    10/25/2004 - Switched ftp server address to name (addressed changed) jbh
 *
 *    05/04/2004 - Added the customer id to the parent id field if the parent id field is empty.  Per KJ request jcf
 *
 *    04/07/2004 - Applied Email class changes. - jcf
 *
 *    12/09/2003 - Modified the way the address list was retrieved to handle changes made by the email webservice class.
 *       Also made changes to handle running the report as a webservice.
 *
 *    10/17/2003 - Added columns to account for all sales types.  Added some formatting for cell width and
 *       also added some for styles.  The style does not seem to work and the docs are based on 1.5 which
 *       currently we do not have installed. - jcf
 *
 *    05/13/2003 - Changed report so that all customer stati are used.  Moved to the use of a
 *       StringBuffer for the sql. - jcf
 *
 *    12/12/2002 - Updated pkg name for new POI 1.5.1.  Removed deprecated method
 *       createCell(short column, int type). PD
 *
 *    02/06/2002 - Removed the oracle connection pool and replaced it with the getOracleConn method.
 *       Added the cleanup routines - jcf
 *
 *    01/02/2002 - Modified the customer type selection criteria to reflect data changes. - jcf
 *
 *    12/28/2001 - Added current processing data. - jcf

 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;


public class RollingCustSales extends Report
{
   private final short BASE_COLS = 15;

   private PreparedStatement m_CustAdder;
   private PreparedStatement m_CustAddr;
   private PreparedStatement m_CustData;
   private PreparedStatement m_CustSales;
   private PreparedStatement m_CustTrip;
   private PreparedStatement m_CustPhone;
   private PreparedStatement m_AdderCount;

   private int m_Month;                      // The month to start in.
   private int m_Year;                       // The year to start in.

   /**
    * default constructor
    */
   public RollingCustSales()
   {
      super();
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      m_CustData = null;
      m_CustSales = null;
      m_CustTrip = null;
      m_CustAdder = null;
      m_CustAddr = null;
      m_CustPhone = null;
      m_AdderCount = null;

      super.finalize();
   }

   /**
    * Executes the queries and builds the output file
    * Note - the file name of the report is set in the preparestatments method because
    *    that's where we have the month and year.
    *
    * @throws FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      ArrayList<String> saleTypes = null;
      XSSFWorkbook wrkBk = null;
      XSSFSheet sheet = null;
      // can only use this on an upgraded version of poi
      //HSSFDataFormat format = null;
      //HSSFCellStyle style = null;
      XSSFRow row = null;
      FileOutputStream outFile = null;
      ResultSet custData = null;
      ResultSet custTrip = null;
      ResultSet custSales = null;
      ResultSet custAdder = null;
      ResultSet custAddr = null;
      ResultSet custPhone = null;
      ResultSet adderCount = null;
      StringBuffer adders = new StringBuffer(256);
      String custId;
      String parentId;
      String mailAddr;
      String shipAddr;
      String city;
      String state;
      String zip;
      String contact;
      String phone;
      String fax;
      String trips;
      int i;
      int colCnt = 1;
      int rowNum = 1;
      boolean result = false;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      wrkBk = new XSSFWorkbook();
      sheet = wrkBk.createSheet();

      // see comments above
      //format = WrkBk.createDataFormat();

      try {
         //
         // Get the maximum number of adders so we have the correct column numbers
         adderCount = m_AdderCount.executeQuery();

         if ( adderCount.next() )
            colCnt = adderCount.getShort(1);

         //
         // Get the sale types for column headings and column counts.  The sale types
         // will also be used as a param for the sales query.
         saleTypes = getSaleTypes();

         colCnt += saleTypes.size();
         colCnt += BASE_COLS;
         createCaptions(sheet, colCnt, saleTypes);

         custData = m_CustData.executeQuery();

         while ( custData.next() && m_Status == RptServer.RUNNING ) {
            row = createRow(sheet, rowNum, colCnt, saleTypes.size());
            custId = custData.getString("customer_id");
            setCurAction("processing customer: " + custId);

            //
            // requested by Karen Jorgensen 5/04/2004.  Wants the parent id field filled with the customer id
            // if there is no data in the parent id field.  This is to accommodate sorting in the spreadsheet.
            parentId = custData.getString("parent_id");
            if ( parentId == null || parentId.length() == 0 )
               parentId = custId;

            m_CustTrip.setString(1, custId);
            m_CustAdder.setString(1, custId);
            m_CustAddr.setString(1, custId);
            m_CustPhone.setString(1, custId);

            custAddr = m_CustAddr.executeQuery();
            custPhone = m_CustPhone.executeQuery();
            custTrip = m_CustTrip.executeQuery();
            custAdder = m_CustAdder.executeQuery();

            mailAddr = "";
            shipAddr = "";
            city = "";
            state = "";
            zip = "";
            contact = "";
            phone = "";
            fax = "";
            trips = "";
            i = 1;

            //
            // Only disply the city, state and zip if it's the shipping address
            // addr1, city, state, postal_code, addrtype
            while ( custAddr.next() ) {
               if ( custAddr.getString(5).equalsIgnoreCase("MAILING") )
                  mailAddr = custAddr.getString(1);
               else {
                  shipAddr = custAddr.getString(1);
                  city = custAddr.getString(2);
                  state = custAddr.getString(3);
                  zip = custAddr.getString(4);
               }
            }

            //
            // Add the contact data
            while ( custPhone.next() ) {
               if ( custPhone.getString("phone_type").equalsIgnoreCase("BUSINESS") ) {
                  contact = custPhone.getString("name");
                  phone = custPhone.getString("phone");
               }
               else {
                  fax = custPhone.getString("phone");

                  if ( contact.length() == 0 )
                     contact = custPhone.getString("name");
               }
            }

            //
            // Create a comma separated list of trips for each customer.
            // If we don't create the list, then we will have multiple customer
            // lines for the same customer.
            while ( custTrip.next() ) {
               if ( i == 1 )
                  trips = custTrip.getString("trip");
               else
                  trips = trips + ", " + custTrip.getString("trip");

               i++;
            }

            //
            // Fill the row with data
            row.getCell(0).setCellValue(new XSSFRichTextString(custId));
            row.getCell(1).setCellValue(new XSSFRichTextString(parentId));
            row.getCell(2).setCellValue(new XSSFRichTextString(custData.getString("name")));
            row.getCell(3).setCellValue(new XSSFRichTextString(mailAddr));
            row.getCell(4).setCellValue(new XSSFRichTextString(shipAddr));
            row.getCell(5).setCellValue(new XSSFRichTextString(city));
            row.getCell(6).setCellValue(new XSSFRichTextString(state));
            row.getCell(7).setCellValue(new XSSFRichTextString(zip));
            row.getCell(8).setCellValue(new XSSFRichTextString(contact));
            row.getCell(9).setCellValue(new XSSFRichTextString(phone));
            row.getCell(10).setCellValue(new XSSFRichTextString(fax));
            row.getCell(11).setCellValue(new XSSFRichTextString(trips));
            row.getCell(12).setCellValue(new XSSFRichTextString(custData.getString("setup_date")));
            row.getCell(13).setCellValue(new XSSFRichTextString(custData.getString("rep_name")));
            row.getCell(14).setCellValue(new XSSFRichTextString(custData.getString("status")));

            //
            // Iterate through the sale types and recored the sales.
            Iterator<String> iter = saleTypes.iterator();
            i = BASE_COLS;

            while ( iter.hasNext() && m_Status == RptServer.RUNNING ) {
               m_CustSales.setString(1, custId);
               m_CustSales.setString(2, iter.next());

               custSales = m_CustSales.executeQuery();

               if ( custSales.next() )
                  row.getCell(i).setCellValue(custSales.getDouble("sales"));

               i++;
            }

            iter = null;

            //
            // Add the sales adders, start at the point where the sales column data
            // ends.
            while ( custAdder.next() && m_Status == RptServer.RUNNING ) {
               adders.append(custAdder.getString("adder") + ", ");
               adders.append(custAdder.getString("adder_value") + ", ");
               adders.append(custAdder.getDouble("percent"));

               row.getCell(i).setCellValue(new XSSFRichTextString(adders.toString()));

               adders.delete(0, adders.length());
               i++;
            }

            DbUtils.closeDbConn(null, null, custAddr);
            DbUtils.closeDbConn(null, null, custAdder);
            DbUtils.closeDbConn(null, null, custSales);
            DbUtils.closeDbConn(null, null, custPhone);
            DbUtils.closeDbConn(null, null, custTrip);

            rowNum++;
         }

         wrkBk.write(outFile);
         wrkBk.close();

         custData.close();
         adderCount.close();
         result = true;
      }

      catch ( Exception ex ) {
         log.error("[RollingCustSales]", ex);
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }

      finally {
         adders = null;
         sheet = null;
         row = null;
         wrkBk = null;

         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[RollingCustSales]", e);
         }

         outFile = null;

         custAddr = null;
         custAdder = null;
         custSales = null;
         custPhone = null;
         custTrip = null;
      }

      return result;
   }

   /**
    * Builds the sql used in the rollling sales query.  We don't know what the month range is until
    * the query is going to be prepared so we have to build it on the fly based on the starting month.
    * <p>
    * @param   month - The ending month.  This should be a full month.
    * @param   year - The year the ending month is in.
    */
   private String buildSalesSQL(int month, int year) throws Exception
   {
      final String Months[] = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"};
      int i;

      StringBuffer sql = new StringBuffer();

      if ( month < 1 || month > 12 )
         throw new Exception("invalid month parameter");

      sql.append("select nvl(sum(dollars_shipped), 0) as sales ");
      sql.append("from sa.customersales ");
      sql.append("where cust_nbr = ? and invoice_month in ( ");

      //
      // Build the string backwards
      for ( i = 0; i <= 11; i++ ) {
         if ( month == 0 ) {
            month = 12;
            year--;
         }

         sql.append("'" + Months[month-1] + "/" + Integer.toString(year) + "'");

         if ( i < 11 )
            sql.append(",");

         month--;
      }

      sql.append(" ) and sale_type = ?");

      return sql.toString();
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_CustData);
      closeStmt(m_CustSales);
      closeStmt(m_CustTrip);
      closeStmt(m_CustAdder);
      closeStmt(m_CustAddr);
      closeStmt(m_CustPhone);
      closeStmt(m_AdderCount);
   }

   /**
    * Builds the captions on the worksheet.
    */
   private void createCaptions(XSSFSheet sheet, int colCnt, ArrayList<String> saleTypes)
   {
      XSSFRow Row = null;
      XSSFCell Cell = null;
      int colPos = BASE_COLS;

      if ( sheet == null )
         return;

      Row = sheet.createRow(0);

      if ( Row != null ) {
         for ( int i = 0; i < colCnt; i++ ) {
            Cell = Row.createCell(i);
            Cell.setCellType(CellType.STRING);
         }

         Row.getCell(0).setCellValue(new XSSFRichTextString("Cust Nbr"));
         Row.getCell(1).setCellValue(new XSSFRichTextString("Parent Acct"));
         Row.getCell(2).setCellValue(new XSSFRichTextString("Cust Name"));
         sheet.setColumnWidth(2, 8000);
         Row.getCell(3).setCellValue(new XSSFRichTextString("Mail Address"));
         sheet.setColumnWidth(3, 7000);
         Row.getCell(4).setCellValue(new XSSFRichTextString("Shipping Address"));
         sheet.setColumnWidth(4, 7000);
         Row.getCell(5).setCellValue(new XSSFRichTextString("City"));
         sheet.setColumnWidth(5, 4000);
         Row.getCell(6).setCellValue(new XSSFRichTextString("State"));
         Row.getCell(7).setCellValue(new XSSFRichTextString("Zip"));
         Row.getCell(8).setCellValue(new XSSFRichTextString("Contact"));
         sheet.setColumnWidth(8, 6000);
         Row.getCell(9).setCellValue(new XSSFRichTextString("Phone"));
         sheet.setColumnWidth(9, 3300);
         Row.getCell(10).setCellValue(new XSSFRichTextString("Fax"));
         sheet.setColumnWidth(10, 3300);
         Row.getCell(11).setCellValue(new XSSFRichTextString("Truck Trip"));
         Row.getCell(12).setCellValue(new XSSFRichTextString("Setup Date"));
         sheet.setColumnWidth(12, 3300);
         Row.getCell(13).setCellValue(new XSSFRichTextString("Tm"));
         sheet.setColumnWidth(13, 6000);
         Row.getCell(14).setCellValue(new XSSFRichTextString("Status"));

         //
         // Create the columns for the sale types.  We never know how many sale types
         // there can be, so this is dynamic.
         Iterator<String> iter = saleTypes.iterator();

         while ( iter.hasNext() ) {
            Row.getCell(colPos).setCellValue(new XSSFRichTextString(iter.next()));
            sheet.setColumnWidth(colPos, 3300);
            colPos++;
         }

         iter = null;

         //
         // Create the adder types.  This is variable so we just get the max amount
         // and tack them on to the spot where the sales data stops.
         for ( int i = colPos; i < colCnt; i++ ) {
            Row.getCell(i).setCellValue(new XSSFRichTextString("Adder" + ((i - colPos) + 1)));
            sheet.setColumnWidth(i, 10000);
         }
      }
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
         m_EdbConn = m_RptProc.getEdbConn();

         if ( prepareStatements() )
            created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[RollingCustSales]", ex);
      }

      finally {
         closeStatements();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Creates a row in the worksheet.
    */
   private XSSFRow createRow(XSSFSheet sheet, int rowNum, int colCnt, int saleTypeCount)
   {
      XSSFRow Row = null;
      XSSFCell Cell = null;
      XSSFCellStyle style = null;
      int numericColStart = BASE_COLS;
      int numericColEnd = BASE_COLS + saleTypeCount-1;

      if ( sheet == null )
         return Row;

      Row = sheet.createRow(rowNum);

      //
      // The warehouse sales column is currently col #13, it's the only non string col.
      // This is bad code but it works.
      if ( Row != null ) {
         for ( int i = 0; i < colCnt; i++ ) {
            if ( i < numericColStart || i > numericColEnd ) {
               Cell = Row.createCell(i);
               Cell.setCellType(CellType.STRING);
               Cell.setCellValue(new XSSFRichTextString(""));
            }
            else {
               Cell = Row.createCell(i);
               Cell.setCellType(CellType.NUMERIC);
               style = Cell.getCellStyle();
               Cell.setCellStyle(style);
               style.setDataFormat((short)8);
               Cell.setCellValue(0.0);

               style = null;
            }
         }
      }

      return Row;
   }

   /**
    * Gets the list of distinct sale types from the customersales table.  The sale types are stored
    * in an array for use later in the program.  The sale types can change over time so they must
    * be retrieved and built into the report dynamically.
    *
    * @return ArrayList the list of sale types
    */
   private ArrayList<String> getSaleTypes()
   {
      ArrayList<String> list = new ArrayList<String>();
      ResultSet set = null;
      Statement stmt = null;

      try {
         stmt = m_EdbConn.createStatement();
         set = stmt.executeQuery("select distinct(sale_type) from customersales order by 1");

         while ( set.next() ) {
            list.add(set.getString(1));
         }
      }

      catch ( Exception ex ) {
         log.error("[RollingCustSales]", ex);
      }

      finally {
         if ( set != null ) {
            try {
               set.close();
               set = null;
            }

            catch ( SQLException ex ) {
               log.error("[RollingCustSales]", ex);
            }
         }

         if ( stmt != null ) {
            try {
               stmt.close();
               stmt = null;
            }

            catch ( SQLException ex ) {
               log.error("[RollingCustSales]", ex);
            }
         }
      }

      return list;
   }

   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements() throws Exception
   {
      StringBuffer sql = new StringBuffer(256);
      boolean prepared = false;

      if ( m_EdbConn != null ) {
         sql.append("select distinct customer.customer_id, customer.name, setup_date, ");
         sql.append("first || ' ' || last as rep_name, customer_status.description as status, parent_id ");
         sql.append("from customer ");
         sql.append("join customer_status on customer_status.cust_status_id = customer.cust_status_id ");
         sql.append("join cust_market_view on cust_market_view.customer_id = customer.customer_id and ");
         sql.append("   cust_market_view.market = 'CUSTOMER TYPE' and  ");
         sql.append("   cust_market_view.class in ('NAT ACCT', 'PRO', 'HDW', 'HOME CTR', 'MAS', 'MISC', 'PAINT') ");
         sql.append("left outer join cust_rep on cust_rep.customer_id = customer.customer_id ");
         sql.append("left outer join emery_rep on emery_rep.er_id = cust_rep.er_id ");
         sql.append("join emery_rep_type on emery_rep_type.rep_type_id = cust_rep.rep_type_id and emery_rep_type.description = 'SALES REP' ");
         sql.append("order by customer.name ");
         m_CustData = m_EdbConn.prepareStatement(sql.toString());

         sql.setLength(0);
         sql.append("select distinct trip_sched.name as trip ");
         sql.append("from trip_stop_sched ");
         sql.append("join trip_sched on trip_sched.ts_id = trip_stop_sched.ts_id ");
         sql.append("where trip_stop_sched.customer_id = ? and is_active = 1 ");
         m_CustTrip = m_EdbConn.prepareStatement(sql.toString());

         sql.setLength(0);
         sql.append("select adder, adder_value, percent ");
         sql.append("from cust_adder_view ");
         sql.append("where customer_id = ? ");
         sql.append("order by adder");
         m_CustAdder = m_EdbConn.prepareStatement(sql.toString());

         sql.setLength(0);
         sql.append("select addr1, city, state, postal_code, addrtype " );
         sql.append("from cust_address_view ");
         sql.append("where customer_id = ? and addrtype in ('MAILING', 'SHIPPING') ");
         sql.append("order by addrtype");
         m_CustAddr = m_EdbConn.prepareStatement(sql.toString());

         sql.setLength(0);         
         sql.append("select ");
         sql.append("first || ' ' || last as name, emery_utils.format_phone(phone_number) as phone, extension, phone_type ");
         sql.append("from cust_contact ");
         sql.append("join emery_contact_phone_view on emery_contact_phone_view.ec_id = cust_contact.ec_id ");
         sql.append("where customer_id = ? and phone_type in ('BUSINESS', 'BUSINESS FAX') ");         
         sql.append("order by phone_type");
         m_CustPhone = m_EdbConn.prepareStatement(sql.toString());

         sql.setLength(0);
         sql.append("select max(cnt) as max_cnt ");
         sql.append("from ( ");
         sql.append("   select customer_id, count(adder) cnt " );
         sql.append("   from cust_adder_view ");
         sql.append("   group by customer_id ");
         sql.append(")");
         m_AdderCount = m_EdbConn.prepareStatement(sql.toString());

         m_CustSales = m_EdbConn.prepareStatement(buildSalesSQL(m_Month, m_Year));
         prepared = true;
      }

      return prepared;
   }

   /**
    * Sets the parameters for the report and also builds the file name.
    *
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fname = new StringBuffer();
      String tm = Long.toString(System.currentTimeMillis()).substring(3);
      //
      // We are for the time being, just going to grab whats in the first two elements.
      // This will change later to add more protection.  A 0 month or year means get the current
      // date from the system
      m_Month = Integer.parseInt(params.get(0).value);
      m_Year = Integer.parseInt(params.get(1).value);

      //
      // Build the file name.
      fname.append(tm);
      fname.append("-");
      fname.append(Integer.toString(m_Month));
      fname.append(Integer.toString(m_Year));
      fname.append("r12cs.xlsx");
      m_FileNames.add(fname.toString());

      //
      // Note - The month for calander is 0 based which means we don't have to move back
      // a month we can leave it alone (we always process the month previous to the curent month).
      // For the file name we need the month that will be processed which is going to be the
      // previous month unless it's set in the params.  This means we have to
      // back up to the previous year.
      if ( m_Month == 0 || m_Year == 0 ) {
         m_Month = Calendar.getInstance().get(Calendar.MONTH);
         m_Year = Calendar.getInstance().get(Calendar.YEAR);

         //
         // The sql builder is 1 based.  When we are at current month 0 (Jan) we
         // need to backup one so the builder can handle it.  Other than that we're
         // ok.
         if ( m_Month == 0 ) {
            m_Month = 12;
            m_Year--;
         }
      }
   }
}
