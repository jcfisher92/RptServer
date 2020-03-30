/**
 * File: ShowRegExport.java
 * Description: Builds the show data export report. 
 *
 * @author Paul Davidson
 *
 * Create Date: 10/19/2006
 * Last Update: $Id: ShowRegExport.java,v 1.29 2013/01/09 17:16:15 npasnur Exp $
 * 
 * History:
 */
package com.emerywaterhouse.rpt.spreadsheet;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;
import com.emerywaterhouse.utils.Crypto;

public class ShowRegExport extends Report
{
   private HSSFCellStyle m_CellStyle;
   private HSSFRow m_CaptionRow;              // Reference to caption row object
   private HSSFSheet m_Sheet;
   private String m_ShowDesc;                 // Description (name) of show 
   private PreparedStatement m_StmtShowCusts; // Gets list of show customers
   private PreparedStatement m_StmtCustAtten; // Gets list of attendees for a specific account
   private PreparedStatement m_StmtCustMeals; // Gets meal info for a specific account
   private PreparedStatement m_StmtCustNumbs; // Gets customer#s of parent and children
   private PreparedStatement m_StmtCustRooms; // Gets hotel room info for a specific account
   private PreparedStatement m_StmtCustSeminars; // Gets seminar info for a specific account
   private HSSFWorkbook m_Wrkbk;

   private static final short MAX_COLS  = 200; // Max number of columns in spreadsheet
   private static final short MAX_ATTEN = 6;   // Max number of attendee name columns
   
   /**
    * Default constructor. Initialize report variables.
    */
   public ShowRegExport()
   {
      super();

      m_Wrkbk = new HSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      m_CaptionRow = null;
      setupWorkbook();

      m_ShowDesc = "";
      m_StmtShowCusts = null;
      m_StmtCustAtten = null;
      m_StmtCustMeals = null;
      m_StmtCustNumbs = null;
      m_StmtCustRooms = null;
      m_StmtCustSeminars = null;
   }
   
   /**
    * Cleanup any allocated resources.
    */
   @Override
   public void finalize() throws Throwable
   {      
      m_CaptionRow = null;
      m_Sheet = null;
      m_Wrkbk = null;      
      m_CellStyle = null;
      m_ShowDesc = null;
      
      super.finalize();
   }
   
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      String cardNum;
      int colNum = 0;
      int count;
      Crypto crypto = new Crypto();
      String custId = null;
      StringBuffer custs = new StringBuffer(); 
      HSSFRow row = null;
      int rowNum = 1;
      FileOutputStream outFile = null;
      boolean prntHotel;
      ResultSet rsetCust = null;
      ResultSet rsetNums =  null;
      ResultSet rsetAttd = null;
      ResultSet rsetMeal = null;
      ResultSet rsetRoom = null;
      ResultSet rsetSeminar = null;
      boolean result = false;
      long showCustId;
      boolean wantBook; 
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         rowNum = createCaptions();
                  
         m_StmtShowCusts.setString(1, m_ShowDesc);
         rsetCust = m_StmtShowCusts.executeQuery();

         while ( rsetCust.next() && m_Status == RptServer.RUNNING ) {            
            showCustId = rsetCust.getLong("show_cust_id");
            custId = rsetCust.getString("customer_id");
            
            setCurAction("Currently at customer# (parent): " + custId);
            
            row = createRow(rowNum++, MAX_COLS);
            colNum = 0;
            wantBook = false;
            
            if ( row != null ) {               
               //
               // Get customer#s of parent and all its children
               custs.setLength(0);
               m_StmtCustNumbs.setLong(1, showCustId);
               rsetNums = m_StmtCustNumbs.executeQuery();
               try {
                  while ( rsetNums.next() ) {
                     custs.append(rsetNums.getString("customer_id"));
                     custs.append(" ");
                  }
                  
                  if ( custs.length() > 0 )
                     custs.setLength(custs.length()-1);
               }
               finally {
                  closeRSet(rsetNums);
               }
               
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(custs.toString()));
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetCust.getString("name")));
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetCust.getString("email")));
               
               //
               // Print out attendee names with each name in a separate cell
               count = 0;
               m_StmtCustAtten.setLong(1, showCustId);
               rsetAttd = m_StmtCustAtten.executeQuery();
               try {
                  while ( rsetAttd.next() ) {
                     count++;
                     
                     if ( count <= MAX_ATTEN )
                        row.getCell(colNum++).setCellValue(
                              new HSSFRichTextString(rsetAttd.getString("first") + " " + rsetAttd.getString("last")));
                     
                     if ( !wantBook )
                        wantBook = ( rsetAttd.getInt("show_book") == 1 );
                  }
               }
               finally {
                  closeRSet(rsetAttd);
               }
               
               if ( count > MAX_ATTEN )
                  count = MAX_ATTEN;
               
               //
               // If there were less attendees than the max, increment by the
               // difference, so columns can line up correctly
               colNum = (colNum + (MAX_ATTEN-count));
               
               //
               // Print if this show customer wants a show book
               if ( wantBook )
                  row.getCell(colNum++).setCellValue(new HSSFRichTextString("Yes"));
               else
                  row.getCell(colNum++).setCellValue(new HSSFRichTextString("No"));
               
               //
               // Print out which attendee wants a show book, and the contact
               // name and customer# who wants the book.  This re-executes the
               // attendee list statement above, which is not the most efficient 
               // way of doing this.
               count = 0;
               m_StmtCustAtten.setLong(1, showCustId);
               rsetAttd = m_StmtCustAtten.executeQuery();
               try {
                  while ( rsetAttd.next() ) {
                     wantBook = ( rsetAttd.getInt("show_book") == 1 );
                     
                     if ( wantBook ) {
                        count++;
                        
                        if ( count <= MAX_ATTEN ) {
                           row.getCell(colNum++).setCellValue(
                                 new HSSFRichTextString(rsetAttd.getString("first") + " " + rsetAttd.getString("last")));
                           row.getCell(colNum++).setCellValue(
                                 new HSSFRichTextString(rsetAttd.getString("attendee_cust")));
                        }
                     }
                  }
               }
               finally {
                  closeRSet(rsetAttd);
               }
               
               if ( count > MAX_ATTEN )
                  count = MAX_ATTEN;
               
               //
               // If there were less show book requestors than the max, increment
               // by the difference, so columns can line up correctly
               colNum = (colNum + ((MAX_ATTEN-count) * 2));
               
               //
               // Print meal info for current show customer
               m_StmtCustMeals.setLong(1, showCustId);
               rsetMeal = m_StmtCustMeals.executeQuery();
               try {
                  if ( rsetMeal.next() ) {
                	 row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetMeal.getString("thurs_night_fare")));
                	 row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetMeal.getString("fri_breakfast")));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetMeal.getString("fri_lunch")));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetMeal.getString("fri_dinner")));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetMeal.getString("sat_breakfast")));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetMeal.getString("sat_lunch")));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetCust.getString("meal_comments")));
                  }
               }
               finally {
                  closeRSet(rsetMeal);
               }
               
               //
               // Print credit card info for hotel payment
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetCust.getString("card_name")));
               cardNum = rsetCust.getString("card_number");
               if ( cardNum != null && cardNum.length() > 0 )
                  row.getCell(colNum++).setCellValue(new HSSFRichTextString(crypto.decrypt(cardNum)));
               else
                  colNum++;
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetCust.getString("expiration")));

               //
               // Print hotel room info
               prntHotel = true;
               count = 0;
               m_StmtCustRooms.setLong(1, showCustId);
               rsetRoom = m_StmtCustRooms.executeQuery();
               try {
                  while ( rsetRoom.next() ) {
                     count++;
                     
                     if ( prntHotel ) {
                        m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString("Hotel"));
                        row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetRoom.getString("hotel_name")));
                        
                        prntHotel = false;
                     }
                     
                     m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString("Rm"+count+" Occupant1"));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetRoom.getString("occupant_1")));
                     m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString("Rm"+count+" Occupant2"));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetRoom.getString("occupant_2")));
                     
                     m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString("Rm"+count+" Bed Type"));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetRoom.getString("bed_type")));
                     
                     m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString("Rm"+count+" Smoking"));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetRoom.getString("smoking")));
   
                     m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString("Rm"+count+" Handicapped"));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetRoom.getString("handicapped")));
                     
                     m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString("Rm"+count+" Special Requests"));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetRoom.getString("room_comments")));
                     
                     m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString("Rm"+count+" Check-In"));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetRoom.getString("arrival")));
                     
                     m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString("Rm"+count+" Check-Out"));
                     row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetRoom.getString("departure")));
                  }
               }
               finally {
                  closeRSet(rsetRoom);
               }
               
               //
               // Print seminar info for current show customer
               m_StmtCustSeminars.setLong(1, showCustId);
               rsetSeminar = m_StmtCustSeminars.executeQuery();
               try {
                  while ( rsetSeminar.next() ) {
                	 m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString(rsetSeminar.getString("presented_by")+"("+rsetSeminar.getString("seminar_session")+")"));
                	 row.getCell(colNum++).setCellValue(new HSSFRichTextString(rsetSeminar.getString("seminar_attendee")));
                  }
               }
               catch(Exception e){
                  log.fatal("Error while adding show seminar information to the report:", e);
               }
               finally {
                  closeRSet(rsetSeminar);
               }
            }
         }
         
         m_Wrkbk.write(outFile);

         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {   
         closeRSet(rsetCust);
         
         rsetCust = null;
         rsetNums =  null;
         rsetAttd = null;
         rsetMeal = null;
         rsetRoom = null;
         rsetSeminar = null;
         custs = null;
         row = null;
         crypto = null;
         
         try {
            outFile.close();
         }
         catch( Exception e ) {
            log.error(e);
         }

         outFile = null;
      }

      return result;
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_StmtShowCusts);
      closeStmt(m_StmtCustAtten);
      closeStmt(m_StmtCustMeals);
      closeStmt(m_StmtCustNumbs);
      closeStmt(m_StmtCustRooms);
      closeStmt(m_StmtCustSeminars);
   }
   
   /**
    * Sets up a portion of the captions on the report.  The rest will
    * be added dynamically.
    */
   private int createCaptions()
   {
      int rowNum = 0;
      int colNum = 0;
            
      if ( m_Sheet == null )
         return 0;
      
      //
      // Create the row for the captions.
      m_CaptionRow = m_Sheet.createRow(rowNum);
      
      if ( m_CaptionRow != null ) {
         for ( int i = 0; i < MAX_COLS; i++ ) {
            m_CaptionRow.createCell(i);            
         }
      }
            
      m_CaptionRow.getCell(colNum).setCellValue(new HSSFRichTextString("Customer#(s)"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Cust Name"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Email (Primary Contact)"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Attendees"));
      
      // Allow for max number of attendees
      colNum = (colNum + (MAX_ATTEN-1));
      
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Show Account Stamp"));
      
      // Allow max_atten*2 columns for max_atten show book contact+cust pairs
      for ( int i = 1; i <= MAX_ATTEN; i++ ) {
         m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Show Account Stamp Contact" + i));
         m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Show Account Stamp Cust" + i));
      }
      
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Thurs Night Fare"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Fri Breakfast"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Fri Lunch"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Fri Recept Dinner"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Sat Breakfast"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Sat Lunch"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Special Meal Requests"));
      
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Card Type"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Card Number"));
      m_CaptionRow.getCell(++colNum).setCellValue(new HSSFRichTextString("Expiration Date"));
      
      return ++rowNum;
   }
   
   /**
    * Creates the show data export report.
    * 
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    * 
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
         log.fatal("ShowRegExport error:", ex);
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
    * @param rowNum The row number.
    * @param colCnt The number of columns in the row.
    * 
    * @return The formatted row of the spreadsheet.
    */
   private HSSFRow createRow(int rowNum, int colCnt)
   {
      HSSFRow row = null;
      HSSFCell cell = null;
      
      if ( m_Sheet == null )
         return row;

      row = m_Sheet.createRow(rowNum);

      //
      // set the type and style of the cell.
      if ( row != null ) {
         for ( int i = 0; i < colCnt; i++ ) {            
            cell = row.createCell(i);
            cell.setCellStyle(m_CellStyle);
         }
      }

      return row;
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
            //
            // Gets all show customer and credit card data.  Make sure that
            // only the topmost parent show customers are retrieved here, by 
            // checking if the show_cust.parent field is null.
            sql.append("select ");
            sql.append("   show_cust_id, ");
            sql.append("   customer.customer_id, ");
            sql.append("   customer.name, ");
            sql.append("   ( select min(sca.email) "); 
            sql.append("     from show_cust_attendee sca "); 
            sql.append("     where sca.primary_contact = 1 and "); 
            sql.append("        sca.show_cust_id = show_cust.show_cust_id ");
            sql.append("   ) as email, ");
            sql.append("   comments as meal_comments, ");
            sql.append("   card_name, ");
            sql.append("   card_number, ");
            sql.append("   to_char(expiration, 'mm/dd/yyyy') as expiration ");
            sql.append("from "); 
            sql.append("   show, ");
            sql.append("   show_cust, ");
            sql.append("   customer, ");
            sql.append("   credit_card ");
            sql.append("where ");
            sql.append("   show.name = ? and ");
            sql.append("   parent is null and ");
            sql.append("   show.show_id = show_cust.show_id and ");
            sql.append("   show_cust.customer_id = customer.customer_id and ");
            sql.append("   show_cust.cc_id = credit_card.cc_id(+) ");
            sql.append("order by customer_id");
            m_StmtShowCusts = m_OraConn.prepareStatement(sql.toString());
            
            //
            // Gets customer#s of parent show cust and all its children
            sql.setLength(0);
            sql.append("select customer_id "); 
            sql.append("from show_cust ");
            sql.append("start with show_cust_id = ? ");
            sql.append("connect by parent = prior show_cust_id ");
            m_StmtCustNumbs = m_OraConn.prepareStatement(sql.toString());
               
            //
            // Gets list of attendees for a show customer and all its children.
            // This returns duplicate rows for some reason, so added a distinct 
            // clause to the sql.
            sql.setLength(0);
            sql.append("select distinct "); 
            sql.append("   show_attendee_id, ");
            sql.append("   show_book, "); 
            sql.append("   first, "); 
            sql.append("   last, "); 
            sql.append("   customer_id as attendee_cust "); 
            sql.append("from "); 
            sql.append("   show_cust, "); 
            sql.append("   show_cust_attendee "); 
            sql.append("where "); 
            sql.append("   show_cust.show_cust_id = show_cust_attendee.show_cust_id ");
            sql.append("start with show_cust.show_cust_id = ? ");
            sql.append("connect by parent = prior show_cust.show_cust_id ");
            sql.append("order by last");            
            m_StmtCustAtten = m_OraConn.prepareStatement(sql.toString());

            //
            // Gets total numbers of meal types for a specific show customer
            // This is for the MarketPlace 2011
            sql.setLength(0);
            sql.append("select ");
            sql.append("   sum(decode(description, 'Thursday Night Light Fare', quantity, 0)) as thurs_night_fare, ");
            sql.append("   sum(decode(description, 'Friday Breakfast', quantity, 0)) as fri_breakfast, ");
            sql.append("   sum(decode(description, 'Friday Lunch', quantity, 0)) as fri_lunch, ");
            sql.append("   sum(decode(description, 'Friday Reception & Dinner', quantity, 0)) as fri_dinner, ");
            sql.append("   sum(decode(description, 'Saturday Breakfast', quantity, 0)) as sat_breakfast, ");
            sql.append("   sum(decode(description, 'Saturday Lunch', quantity, 0)) as sat_lunch ");
            sql.append("from ");
            sql.append("   show_cust, ");
            sql.append("   show_cust_meal, ");
            sql.append("   show_meal ");
            sql.append("where ");
            sql.append("   show_cust.show_cust_id = ? and ");
            sql.append("   show_cust.show_cust_id = show_cust_meal.show_cust_id and ");
            sql.append("   show_cust_meal.meal_id = show_meal.meal_id ");
            m_StmtCustMeals = m_OraConn.prepareStatement(sql.toString());
            
            //
            // Get all hotel room information for a specific show customer
            sql.setLength(0);
            sql.append("select ");
            sql.append("   room_id, ");
            sql.append("   show_hotel.name as hotel_name, ");
            sql.append("   occupant_1, occupant_2, bed_type, ");
            sql.append("   decode(smoking, 1, 'Yes', 'No') as smoking, ");
            sql.append("   decode(handicapped, 1, 'Yes', 'No') as handicapped, ");
            sql.append("   show_cust_hotel_room.comments as room_comments, ");
            sql.append("   to_char(arrival, 'mm/dd/yyyy') as arrival, ");
            sql.append("   to_char(departure, 'mm/dd/yyyy') as departure ");
            sql.append("from ");
            sql.append("   show_cust, ");
            sql.append("   show_cust_hotel_room, ");
            sql.append("   show_hotel ");
            sql.append("where ");
            sql.append("   show_cust.show_cust_id = ? and ");
            sql.append("   show_cust.show_cust_id = show_cust_hotel_room.show_cust_id and ");
            sql.append("   show_cust_hotel_room.hotel_id = show_hotel.hotel_id ");
            m_StmtCustRooms = m_OraConn.prepareStatement(sql.toString());
            
            //
            //Get the seminar information
            sql.setLength(0);
            sql.append("select ");
            sql.append("   presented_by,seminar_session,nvl(quantity,'0') as seminar_attendee ");
            sql.append("from ");
            sql.append("   show_cust_seminar ");
           	sql.append("   right outer join show_seminar on show_cust_seminar.seminar_id = show_seminar.seminar_id ");
            sql.append("   and show_cust_id = ? ");
            sql.append("order by seminar_session ");
            m_StmtCustSeminars = m_OraConn.prepareStatement(sql.toString());

            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("ShowRegExport.prepareStatements error: ", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("ShowRegExport.prepareStatements - null oracle connection");
      
      return isPrepared;
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();
      Param param = null;
      int pcount = params.size();
      String tmp = null;
               
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                           
         if ( param.name.equals("showdesc") )
            m_ShowDesc = param.value;
      }
      
      //
      // Build the report file name
      tmp = Long.toString(System.currentTimeMillis());
      fileName.append("show_export");      
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xls");
      m_FileNames.add(fileName.toString());
   }
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other inititialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {      
      HSSFCellStyle styleText;
            
      styleText = m_Wrkbk.createCellStyle();      
      styleText.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      
      m_CellStyle = styleText;
      
      styleText = null;
   }
}