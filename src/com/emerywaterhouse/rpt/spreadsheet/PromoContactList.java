/**
 * File: PromoContactList.java
 * Description: Promotion contacts list report.
 *    This is based on the Rolling 12 month customer sales report and has similar queries to that.
 *
 * @author Paul Davidson (based on code by Jeff Fisher in RollingCustSales.java)
 * 
 * $Revision: 1.39 $
 * 
 * Create Data: 7/23/2009
 * Last Update: $Id: PromoContactList.java,v 1.39 2011/06/24 03:38:18 jfisher Exp $
 * 
 * History:
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class PromoContactList extends Report
{
   private PreparedStatement m_CustData;          // Main customer contacts statement
   private PreparedStatement m_CustSalesAllLocs;  // R12 sales for all stores in family tree
   private PreparedStatement m_CustSales;         // R12 sales for specific customer#
   private PreparedStatement m_CustTrip;          //
   private boolean m_IsMailingList;               // True if end-user wanted the current mailng list else false
   private PreparedStatement m_MailService;       // Checks if customer is on mail service
   private int m_Month;                           // The month to start in.
   XSSFCellStyle m_StyleText;                     // Text style left justified
   private String m_WhseName;                     // Warehouse name parameter, e.g. PORTLAND, PITTSTON
   private int m_Year;                            // The year to start in.

   /**
    * default constructor
    */
   public PromoContactList()
   {
      super();

      m_IsMailingList = true;
      m_WhseName = null;
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
      m_CustSalesAllLocs = null;
      m_CustSales = null;
      m_CustTrip = null;
      m_MailService = null;
      m_StyleText = null;

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
      boolean bbb;                        // True if customer is BB&B or Lowes
      XSSFCell cell;                      // Temp reference to spreadsheet cell
      int col = 0;                        //
      ResultSet custData = null;          //
      ResultSet custTrip = null;          //
      ResultSet custSalesAllLocs = null;  //
      ResultSet custSales = null;         //
      String custId;                      // Current customer# in report
      String custName;                    //
      String contactName = null;          // Current customer contact name
      String emailBooks = null;           // Send email notices for promos in these books
      boolean hasMailing;                 //
      boolean hasPromoAddress;            // Flag for checking if using regular or promotion mailing address
      ResultSet mailService = null;       // Mail service result set
      FileOutputStream outFile = null;    //
      String promoStreet;                 // Street (for promotional mailing address)
      String paperBooks = null;           // Send paper version of books
      boolean result = false;             //
      int rowNum = 1;                     //
      XSSFRow row = null;                 // Spreadsheet row object
      double salesR12AllLocs;             // Rolling 12 sales total (for all stores if multiple accounts)
      double salesR12;                    // Rolling 12 sales total for current customer#
      String svcComments = null;          // Any comment attached to mail service
      String svcBegDate = null;           // Mail service begin
      String svcEndDate = null;           // Mail service end
      String sendViaMail = null;          // Send promo mailings via mail
      XSSFSheet sheet = null;             // Spreadsheet sheet object
      String tripFascorId;                // String holding trip Fascor ID
      String tripStopNum;                 // Trip stop number
      String tripDay;                     // Trip day
      XSSFWorkbook wrkBk = null;          // Spreadsheet workbook object

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      wrkBk = new XSSFWorkbook();
      sheet = wrkBk.createSheet();

      try {
         m_StyleText = wrkBk.createCellStyle();
         m_StyleText.setAlignment(HorizontalAlignment.LEFT);

         createCaptions(sheet);
         custData = m_CustData.executeQuery();

         while ( custData.next() && m_Status == RptServer.RUNNING ) {
            custId = custData.getString("customer_id");
            custName = custData.getString("name");
            setCurAction("processing customer: " + custId);

            //
            // If BB&B or Lowes just go to the next customer, per Cetta's request
            bbb = (custName.indexOf("LOWES") > -1) || (custName.indexOf("BED BATH & BEYOND") > -1);
            if ( bbb ) {
               continue;
            }

            //
            // Check if customer is on mail service
            m_MailService.setString(1, custId);
            try {
               mailService = m_MailService.executeQuery();

               if ( mailService.next() ) {
                  hasMailing = true;

                  svcComments = mailService.getString("comments");
                  svcBegDate = mailService.getString("beg_date");
                  svcBegDate = (svcBegDate == null? "": svcBegDate);
                  svcEndDate = mailService.getString("end_date");
                  svcEndDate = (svcEndDate == null? "": svcEndDate);
                  sendViaMail = mailService.getString("send_via_mail");
                  sendViaMail = (sendViaMail == null? "": sendViaMail);
               }
               else {
                  hasMailing = false;

                  svcComments = "";
                  svcBegDate = "";
                  svcEndDate = "";
                  sendViaMail = "";
               }
            }
            finally {
               DbUtils.closeDbConn(null, null, mailService);
            }

            //
            // Depending on the report type requested, we have to either show the current mailing list, or
            // those customers that are not on the current mailing list.
            if ( m_IsMailingList ) {
               if ( !hasMailing )
                  continue;  // Show mailing list, so exclude any customer not on the mail service
            }
            else {
               if ( hasMailing )
                  continue; // Exclude customer on existing mailing list, since we want to show all that are NOT on the list
            }

            tripFascorId = "";
            tripStopNum = "";
            tripDay = "";

            m_CustTrip.setString(1, custId);

            try {
               custTrip = m_CustTrip.executeQuery();

               //
               // Get trip Fascor ID, stop number and day.  If customer has mutiple of these records choose first one.
               if ( custTrip.next() ) {
                  tripFascorId = custTrip.getString("fascor_id");
                  tripStopNum = custTrip.getString("stop_num");
                  tripDay = custTrip.getString("day");
               }
            }
            finally {
               DbUtils.closeDbConn(null, null, custTrip);
            }

            //
            // R12 sales totalled for all locations
            m_CustSalesAllLocs.setString(1, custId);
            m_CustSalesAllLocs.setString(2, "WAREHOUSE");
            salesR12AllLocs = 0.0;
            try {
               custSalesAllLocs = m_CustSalesAllLocs.executeQuery();

               if ( custSalesAllLocs.next() )
                  salesR12AllLocs = custSalesAllLocs.getDouble("sales");
            }
            finally {
               DbUtils.closeDbConn(null, null, custSalesAllLocs);
            }

            //
            // R12 sales total for current customer#
            m_CustSales.setString(1, custId);
            m_CustSales.setString(2, "WAREHOUSE");
            salesR12 = 0.0;
            try {
               custSales = m_CustSales.executeQuery();

               if ( custSales.next() )
                  salesR12 = custSales.getDouble("sales");
            }
            finally {
               DbUtils.closeDbConn(null, null, custSales);
            }

            //
            // Check if promo address has a value, else use the regular mailing address
            promoStreet = custData.getString("street");
            hasPromoAddress = promoStreet != null && promoStreet.trim().length() > 0;

            //
            // Promo book preferences
            emailBooks = custData.getString("email_books");
            emailBooks = (emailBooks == null? "": emailBooks);
            paperBooks = custData.getString("paper_books");
            paperBooks = (paperBooks == null? "": paperBooks);

            row = sheet.createRow(rowNum);
            col = 0;

            // Customer ID
            cell = row.createCell(col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(custId));

            // Customer name
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(custData.getString("name")));

            // Status
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(custData.getString("status")));

            // Warehouse
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(custData.getString("whse_name")));

            if ( m_IsMailingList ) {
               // Service comments
               cell = row.createCell(++col);
               cell.setCellStyle(m_StyleText);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(svcComments));

               // Service begin date
               cell = row.createCell(++col);
               cell.setCellStyle(m_StyleText);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(svcBegDate));

               // Service end date
               cell = row.createCell(++col);
               cell.setCellStyle(m_StyleText);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(svcEndDate));

               // Send via mail (Y/N)
               cell = row.createCell(++col);
               cell.setCellStyle(m_StyleText);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(sendViaMail));
            }

            // Promo address flag
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            if ( m_IsMailingList )
               cell.setCellValue(new XSSFRichTextString(hasPromoAddress? "Y": ""));
            else
               cell.setCellValue(new XSSFRichTextString(hasPromoAddress? "Y": "N"));

            if ( m_IsMailingList ) {
               // Contact name
               contactName = custData.getString("first") == null? "": custData.getString("first") + ' ' + custData.getString("last");
               cell = row.createCell(++col);
               cell.setCellStyle(m_StyleText);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(contactName));
            }

            // Street
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(hasPromoAddress? promoStreet: custData.getString("street_mailing")));

            // City
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(hasPromoAddress? custData.getString("city"): custData.getString("city_mailing")));

            // State
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(hasPromoAddress? custData.getString("province"): custData.getString("state_mailing")));

            // Zip
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(hasPromoAddress? custData.getString("postal_code"): custData.getString("zip_mailing")));

            if ( m_IsMailingList ) {
               // Phone#
               cell = row.createCell(++col);
               cell.setCellStyle(m_StyleText);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(custData.getString("phone_number")));

               // Fax#
               cell = row.createCell(++col);
               cell.setCellStyle(m_StyleText);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(custData.getString("fax_number")));

               // Send email notice
               cell = row.createCell(++col);
               cell.setCellStyle(m_StyleText);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(emailBooks));

               // Send paper version
               cell = row.createCell(++col);
               cell.setCellStyle(m_StyleText);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(paperBooks));
            }

            // Trips
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(tripFascorId));

            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(tripStopNum));

            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(tripDay));

            // Customer setup date
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(custData.getString("setup_date")));

            // Rep name
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(custData.getString("rep_name")));

            // Glidden exclusive line flag
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(custData.getString("glid_excl")));

            // Glidden functional discount flag
            cell = row.createCell(++col);
            cell.setCellStyle(m_StyleText);
            cell.setCellType(CellType.STRING);
            cell.setCellValue(new XSSFRichTextString(custData.getString("glid_disc")));

            // R12 sales for single customer
            cell = row.createCell(++col);
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(salesR12);

            // R12 sales for all locations
            cell = row.createCell(++col);
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(salesR12AllLocs);

            rowNum++;
         }

         wrkBk.write(outFile);
         wrkBk.close();

         custData.close();
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }

      finally {
         sheet = null;
         row = null;
         wrkBk = null;
         m_StyleText = null;

         try {
            if ( outFile != null )
               outFile.close();
         }
         catch( Exception e ) {
            log.error("PromoContactList report exception:", e);
         }

         outFile = null;

         custSalesAllLocs = null;
         custTrip = null;
         custData = null;
      }

      return result;
   }

   /**
    * Builds the sql used in the rolling sales query.  Gets total for all locations if more than one.
    * We don't know what the month range is until the query is going to be prepared so we have to build
    * it on the fly based on the starting month.
    * 
    * NOTE: Gets R12 sales for all stores if multiple locations.
    * <p>
    * @param   month - The ending month.  This should be a full month.
    * @param   year - The year the ending month is in.
    */
   private String buildSalesSQLAllLocs(int month, int year) throws Exception
   {
      int i;
      final String Months[] = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"};
      StringBuffer sql = new StringBuffer();

      if ( month < 1 || month > 12 )
         throw new Exception("invalid month parameter");

      sql.append("select nvl(sum(dollars_shipped), 0) as sales from customersales ");
      sql.append("where cust_nbr in (");
      sql.append("   select customer_id ");
      sql.append("   from customer ");
      sql.append("   start with customer_id = cust_procs.findtopparent(?) ");
      sql.append("   connect by parent_id = prior customer_id ");
      sql.append(") ");
      sql.append(" and invoice_month in ( ");

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
    * Builds the rolling sales query for a specific customer.
    * 
    * @param   month - The ending month.  This should be a full month.
    * @param   year - The year the ending month is in.
    */
   private String buildSalesSQL(int month, int year) throws Exception
   {
      int i;
      final String Months[] = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"};
      StringBuffer sql = new StringBuffer();

      if ( month < 1 || month > 12 )
         throw new Exception("invalid month parameter");

      sql.append("select nvl(sum(dollars_shipped), 0) as sales from customersales ");
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
      closeStmt(m_CustSalesAllLocs);
      closeStmt(m_CustSales);
      closeStmt(m_CustTrip);
      closeStmt(m_MailService);
   }

   /**
    * Builds the captions on the worksheet.
    */
   private void createCaptions(XSSFSheet sheet)
   {
      XSSFCell Cell = null;
      int col = 0;
      XSSFRow Row = null;

      if ( sheet == null )
         return;

      Row = sheet.createRow(0);

      if ( Row != null ) {
         Cell = Row.createCell(col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Cust Nbr"));

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Cust Name"));
         sheet.setColumnWidth(col, 8000);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Status"));
         sheet.setColumnWidth(col, 3000);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Warehouse"));
         sheet.setColumnWidth(col, 3000);

         if ( m_IsMailingList ) {
            Cell = Row.createCell(++col);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Mail Service Comment"));
            sheet.setColumnWidth(col, 8000);

            Cell = Row.createCell(++col);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Service Begin"));
            sheet.setColumnWidth(col, 3500);

            Cell = Row.createCell(++col);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Service End"));
            sheet.setColumnWidth(col, 3500);

            Cell = Row.createCell(++col);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Send via Mail"));
            sheet.setColumnWidth(col, 3500);
         }

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Promo Address"));
         sheet.setColumnWidth(col, 3500);

         if ( m_IsMailingList ) {
            Cell = Row.createCell(++col);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Contact"));
            sheet.setColumnWidth(col, 6000);
         }

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Street/PO Box"));
         sheet.setColumnWidth(col, 7000);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("City"));
         sheet.setColumnWidth(col, 4000);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("State"));

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Zip"));

         if ( m_IsMailingList ) {
            Cell = Row.createCell(++col);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Phone"));
            sheet.setColumnWidth(col, 3300);

            Cell = Row.createCell(++col);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Fax"));
            sheet.setColumnWidth(col, 3300);

            Cell = Row.createCell(++col);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Send Email Notice"));
            sheet.setColumnWidth(col, 6000);

            Cell = Row.createCell(++col);
            Cell.setCellType(CellType.STRING);
            Cell.setCellValue(new XSSFRichTextString("Send Paper Version"));
            sheet.setColumnWidth(col, 6000);
         }

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Fascor"));
         sheet.setColumnWidth(col, 3500);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Stop Number"));
         sheet.setColumnWidth(col, 3500);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Day"));
         sheet.setColumnWidth(col, 2000);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Setup Date"));
         sheet.setColumnWidth(col, 3300);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Tm"));
         sheet.setColumnWidth(col, 4000);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Glidden Exclusive"));
         sheet.setColumnWidth(col, 3900);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("Glidden Func Discount"));
         sheet.setColumnWidth(col, 3300);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("R12 Sales"));
         sheet.setColumnWidth(col, 4500);

         Cell = Row.createCell(++col);
         Cell.setCellType(CellType.STRING);
         Cell.setCellValue(new XSSFRichTextString("All Locations R12 Sales"));
         sheet.setColumnWidth(col, 6000);
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
   private boolean prepareStatements() throws Exception
   {
      StringBuffer sql = new StringBuffer(256);
      boolean prepared = false;

      if ( m_OraConn != null ) {
         //
         // SQL below gets list of appropriate active customers and their promotional contact data.
         // NOTE: there could be multiple promotional contacts for the same customer.
         sql.append("select distinct ");
         sql.append("   customer.customer_id, customer.name, to_char(customer.setup_date, 'mm/dd/yyyy') as setup_date, ");
         sql.append("   emery_rep.first || ' ' || emery_rep.last as rep_name, ");
         sql.append("   customer_status.description as status, ");
         sql.append("   warehouse.name as whse_name, ");
         sql.append("   customer.parent_id, promo_contact.ec_id, ");
         sql.append("   promo_contact.first, ");
         sql.append("   promo_contact.last, ");
         sql.append("   promo_contact.email1, ");
         sql.append("   promo_contact.department, ");
         sql.append("   promo_contact.street, ");
         sql.append("   promo_contact.city, ");
         sql.append("   promo_contact.province, ");
         sql.append("   promo_contact.postal_code, ");
         sql.append("   promo_contact.phone_number, ");
         sql.append("   promo_contact.extension, ");
         sql.append("   promo_contact.fax_number, ");
         sql.append("   promo_contact.email_books, ");
         sql.append("   promo_contact.paper_books, ");
         sql.append("   decode(glid_exclusive.description, 'GLIDDEN', 'Y', '') as glid_excl, ");
         sql.append("   decode(glid_func_disc.description, 'GLIDDEN', 'Y', '') as glid_disc, ");
         sql.append("   mailing_addr.street as street_mailing, ");
         sql.append("   mailing_addr.city as city_mailing, ");
         sql.append("   mailing_addr.state as state_mailing, ");
         sql.append("   mailing_addr.postal_code as zip_mailing ");
         sql.append("from ");
         sql.append("   customer ");
         sql.append("   inner join customer_status on customer.cust_status_id = customer_status.cust_status_id and customer_status.description in ('ACTIVE', 'CREDITHOLD') ");
         sql.append("   inner join cust_market_view on customer.customer_id = cust_market_view.customer_id and cust_market_view.market = 'CUSTOMER TYPE' and cust_market_view.class in ('PAINT', 'NAT ACCT', 'PRO', 'HDW', 'HOME CTR', 'MISC') ");
         sql.append("   inner join cust_rep on customer.customer_id = cust_rep.customer_id ");
         sql.append("   inner join emery_rep on cust_rep.er_id = emery_rep.er_id ");
         sql.append("   inner join emery_rep_type on cust_rep.rep_type_id = emery_rep_type.rep_type_id and emery_rep_type.description = 'SALES REP' ");
         sql.append("   inner join cust_warehouse on cust_warehouse.customer_id = customer.customer_id ");
         sql.append("   inner join warehouse on cust_warehouse.warehouse_id = warehouse.warehouse_id " + (m_WhseName != null? " and warehouse.name = '" + m_WhseName + "' ": ""));
         sql.append("   left outer join ( ");
         sql.append("     select ");
         sql.append("        cust_contact.customer_id, ");
         sql.append("        emery_contact.ec_id, ");
         sql.append("        emery_contact.first, emery_contact.last, ");
         sql.append("        emery_contact.email1, ");
         sql.append("        emery_contact.department, ");
         sql.append("        ec_address.street, ");
         sql.append("        ec_address.city, ");
         sql.append("        ec_address.province, ");
         sql.append("        ec_address.postal_code, ");
         sql.append("        emery_utils.format_phone(ec_phone.phone_number) as phone_number, ");
         sql.append("        ec_phone.extension, ");
         sql.append("        (case when ec_fax.phone_number is null then '' else emery_utils.format_phone(ec_fax.phone_number) end) as fax_number, ");
         sql.append("        ec_email_books.email_books, ");
         sql.append("        ec_paper_books.paper_books ");
         sql.append("     from ");
         sql.append("        emery_contact ");
         sql.append("        inner join cust_contact on emery_contact.ec_id = cust_contact.ec_id ");
         sql.append("        inner join cust_contact_type on cust_contact.cct_id = cust_contact_type.cct_id and cust_contact_type.description = 'PROMOTION' ");
         sql.append("        left outer join ( ");
         sql.append("           select ");
         sql.append("             emery_contact_address.ec_id, ");
         sql.append("             contact_address.street, ");
         sql.append("             contact_address.city, ");
         sql.append("             contact_address.province, ");
         sql.append("             contact_address.postal_code ");
         sql.append("          from ");
         sql.append("             emery_contact_address ");
         sql.append("             inner join contact_address on emery_contact_address.cont_addr_id = contact_address.cont_addr_id ");
         sql.append("             inner join address_type on contact_address.addr_type_id = address_type.addr_type_id and addr_type = 'BUSINESS' ");
         sql.append("        ) ec_address on ec_address.ec_id = emery_contact.ec_id ");
         sql.append("        left outer join ( ");
         sql.append("           select ");
         sql.append("             emery_contact_phone.ec_id, ");
         sql.append("             contact_phone.phone_number, ");
         sql.append("             contact_phone.extension ");
         sql.append("           from ");
         sql.append("             emery_contact_phone ");
         sql.append("             inner join contact_phone on emery_contact_phone.cont_phone_id = contact_phone.cont_phone_id ");
         sql.append("             inner join phone_type on contact_phone.phone_type_id = phone_type.phone_type_id and phone_type = 'BUSINESS' ");
         sql.append("        ) ec_phone on ec_phone.ec_id = emery_contact.ec_id ");
         sql.append("        left outer join ( ");
         sql.append("           select ");
         sql.append("             emery_contact_phone.ec_id, ");
         sql.append("             contact_phone.phone_number ");
         sql.append("           from ");
         sql.append("             emery_contact_phone ");
         sql.append("             inner join contact_phone on emery_contact_phone.cont_phone_id = contact_phone.cont_phone_id ");
         sql.append("             inner join phone_type on contact_phone.phone_type_id = phone_type.phone_type_id and phone_type = 'BUSINESS FAX' ");
         sql.append("        ) ec_fax on ec_fax.ec_id = emery_contact.ec_id ");
         sql.append("        left outer join ( ");
         sql.append("           select ");
         sql.append("              contact_promobook_opt.ec_id, ");
         sql.append("              wmsys.wm_concat(promo_book_type.name) as email_books ");
         sql.append("           from ");
         sql.append("              contact_promobook_opt ");
         sql.append("              inner join promo_book_type on contact_promobook_opt.book_type_id = promo_book_type.book_type_id ");
         sql.append("           where ");
         sql.append("              email_promo = 1 ");
         sql.append("           group by contact_promobook_opt.ec_id ");
         sql.append("        ) ec_email_books on ec_email_books.ec_id = emery_contact.ec_id ");
         sql.append("        left outer join ( ");
         sql.append("           select ");
         sql.append("              contact_promobook_opt.ec_id, ");
         sql.append("              wmsys.wm_concat(promo_book_type.name) as paper_books ");
         sql.append("           from ");
         sql.append("              contact_promobook_opt ");
         sql.append("              inner join promo_book_type on contact_promobook_opt.book_type_id = promo_book_type.book_type_id ");
         sql.append("           where ");
         sql.append("              send_paper_ver = 1 ");
         sql.append("           group by contact_promobook_opt.ec_id ");
         sql.append("        ) ec_paper_books on ec_paper_books.ec_id = emery_contact.ec_id ");
         sql.append("   ) promo_contact on promo_contact.customer_id = customer.customer_id ");
         sql.append("   left outer join ( "); // Check if customer gest the Glidden exclusive line
         sql.append("      select ");
         sql.append("         customer_id, item_grant.description ");
         sql.append("      from ");
         sql.append("         cust_grant, item_grant ");
         sql.append("      where ");
         sql.append("         item_grant.description = 'GLIDDEN' and ");
         sql.append("         item_grant.grant_id = cust_grant.grant_id and ");
         sql.append("         item_grant.is_restriction = 0 ");
         sql.append("   ) glid_exclusive on glid_exclusive.customer_id = customer.customer_id ");
         sql.append("   left outer join ( "); // Check if customers gets the Glidden functional discount
         sql.append("      select ");
         sql.append("         customer_id, discount.description ");
         sql.append("      from ");
         sql.append("         cust_discount, discount ");
         sql.append("      where ");
         sql.append("         discount.description = 'GLIDDEN' and ");
         sql.append("         discount.discount_id = cust_discount.discount_id and ");
         sql.append("         (trunc(cust_discount.beg_date) <= trunc(sysdate) and (trunc(sysdate) <= trunc(cust_discount.end_date) or cust_discount.end_date is null)) ");
         sql.append("   ) glid_func_disc on glid_func_disc.customer_id = customer.customer_id ");
         sql.append("   left outer join ( "); // Get regular mailing address in case there is no promotional one
         sql.append("      select ");
         sql.append("         cust_address.customer_id, ");
         sql.append("         addr1 || ' ' || addr2 as street, ");
         sql.append("         city, ");
         sql.append("         state, ");
         sql.append("         postal_code ");
         sql.append("      from ");
         sql.append("         cust_address ");
         sql.append("         inner join cust_addr_link on cust_address.cust_addr_id = cust_addr_link.cust_addr_id ");
         sql.append("         inner join addr_link_type on cust_addr_link.addr_link_type_id = addr_link_type.addr_link_type_id and addr_link_type.description = 'MAILING' ");
         sql.append("   ) mailing_addr on mailing_addr.customer_id = customer.customer_id ");
         sql.append("order by customer.name");
         m_CustData = m_OraConn.prepareStatement(sql.toString());

         //
         // Gets the FascorId-Stop#-Day string for each stop for this customer
         // Will pick the first one, per request of Marketing Manager
         sql.setLength(0);
         sql.append("select ");
         sql.append("   cust_trip.fascor_id, ");
         sql.append("   cust_trip.stop_num, ");
         sql.append("   decode(cust_trip.monday, 'Pick', 'Mon', '') || ");
         sql.append("   decode(cust_trip.tuesday, 'Pick', 'Tue', '') || ");
         sql.append("   decode(cust_trip.wednesday, 'Pick', 'Wed', '') || ");
         sql.append("   decode(cust_trip.thursday, 'Pick', 'Thu', '') || ");
         sql.append("   decode(cust_trip.friday, 'Pick', 'Fri', '') as day, ");
         sql.append("   decode(cust_trip.monday, 'Pick', 1, '') || ");
         sql.append("   decode(cust_trip.tuesday, 'Pick', 2, '') || ");
         sql.append("   decode(cust_trip.wednesday, 'Pick', 3, '') || ");
         sql.append("   decode(cust_trip.thursday, 'Pick', 4, '') || ");
         sql.append("   decode(cust_trip.friday, 'Pick', 5, '') as day_num ");
         sql.append("from cust_trip ");
         sql.append("where customer_id = ? ");
         sql.append("order by day_num");
         m_CustTrip = m_OraConn.prepareStatement(sql.toString());

         sql.setLength(0);
         sql.append("select ");
         sql.append("   cust_service.cust_serv_id, cust_service.comments, ");
         sql.append("   to_char(cust_service.beg_date, 'mm/dd/yyyy') as beg_date, ");
         sql.append("   to_char(cust_service.end_date, 'mm/dd/yyyy') as end_date, ");
         sql.append("   decode(mail_option.description, null, '', 'Y') as send_via_mail ");
         sql.append("from ");
         sql.append("   cust_service ");
         sql.append("   inner join service on cust_service.service_id = service.service_id and service.name = 'PROMO MAILING' ");
         sql.append("   left outer join ( ");
         sql.append("      select ");
         sql.append("         cust_serv_option.cust_serv_id, ");
         sql.append("         service_option.description ");
         sql.append("      from ");
         sql.append("         cust_serv_option ");
         sql.append("         inner join service_option on service_option.serv_opt_id = cust_serv_option.serv_opt_id and service_option.description = 'MUST GO ON MAIL' ");
         sql.append("   ) mail_option on mail_option.cust_serv_id = cust_service.cust_serv_id ");
         sql.append("where ");
         sql.append("   customer_id = ? and ");
         sql.append("   (trunc(cust_service.beg_date) <= trunc(sysdate) and (cust_service.end_date is null or trunc(cust_service.end_date) >= trunc(sysdate)))");
         m_MailService = m_OraConn.prepareStatement(sql.toString());

         m_CustSalesAllLocs = m_OraConn.prepareStatement(buildSalesSQLAllLocs(m_Month, m_Year));
         m_CustSales = m_OraConn.prepareStatement(buildSalesSQL(m_Month, m_Year));

         prepared = true;
      }

      return prepared;
   }

   /**
    * Sets the parameters for the report and also builds the file name.
    * 
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fname = new StringBuffer();
      String rptType;
      String tm = Long.toString(System.currentTimeMillis()).substring(3);

      //
      // We are for the time being, just going to grab whats in the first two elements.
      // This will change later to add more protection.  A 0 month or year means get the current
      // date from the system
      m_Month = Integer.parseInt(params.get(0).value);
      m_Year = Integer.parseInt(params.get(1).value);
      rptType = params.get(2).value;

      if ( params.size() > 3)
         m_WhseName = params.get(3).value;
      else
         m_WhseName = null;

      m_IsMailingList = rptType != null && rptType.equals("MAILINGLIST");

      //
      // Build the file name.
      if ( m_IsMailingList )
         fname.append("promo-mailinglist-");
      else
         fname.append("NOT-on-mailinglist-");

      fname.append(tm);
      fname.append("-");
      fname.append(Integer.toString(m_Month));
      fname.append(Integer.toString(m_Year));
      fname.append(".xlsx");
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
