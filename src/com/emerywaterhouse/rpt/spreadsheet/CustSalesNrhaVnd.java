/**
 * File: CustSalesNrhaVnd.java.java
 * Description: Rewrite of the Customer sales by class and by nrha dept and by venddor.  
 *    The report creates a cross tab report pivoting on the nrha dept so that cost and sales are 
 *    displayed across the row for each nrha dept and vendor.<p>
 *    
 *    Original author was Jeffrey Fisher
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 05/11/2005
 * Last Update: $Id: CustSalesNrhaVnd.java,v 1.8 2011/04/23 04:23:38 npasnur Exp $
 * 
 * History
 *    $Log: CustSalesNrhaVnd.java,v $
 *    Revision 1.8  2011/04/23 04:23:38  npasnur
 *    Changed the params order for method CellRangeAddress after POI upgrade
 *
 *    Revision 1.7  2009/02/18 14:50:46  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    03/28/2005 - Added log4j logging - jcf.
 * 
 *    03/17/2005 - JDK 1.5 type safety fixes.
 * 
 *    10/25/2004 - Switched ftp server address to name (addressed changed) jbh
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class CustSalesNrhaVnd extends Report
{
   private final int SALES_COL_START = 1;
   private final int MAX_NRHA = 12;

   private String m_BegDate;
   private int m_ColCount;
   private HashMap<String, Short> m_CustClass;
   private PreparedStatement m_CustSales;
   private String m_EndDate;
   private String[] m_NrhaList;
   private ArrayList<String> m_ClassHeadings;
   
   /**
    * default constructor
    */
   public CustSalesNrhaVnd()
   {
      super();
      
      m_NrhaList = new String[MAX_NRHA];
      m_ClassHeadings = new ArrayList<String>(10);
      m_CustClass = new HashMap<String, Short>(10);
   }

   /**
    * Cleanup when were done.
    * @throws Throwable
    */
   public void finalize() throws Throwable
   {
      for ( short i = 0; i < MAX_NRHA; i++ )
         m_NrhaList[i] = null;

      m_CustClass.clear();
      m_CustClass = null;

      m_ClassHeadings.clear();
      m_ClassHeadings = null;
      
      m_EndDate = null;
      m_NrhaList = null;
      m_BegDate = null;
   }
   
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {      
      XSSFWorkbook wrkBk = null;
      XSSFSheet sheet = null;
      XSSFRow row = null;
      FileOutputStream outFile = null;
      ResultSet custSales = null;
      int colNum = 0;
      int rowNum = 0;
      String curNrha = null;
      String curVnd = null;
      String nextVnd = null;
      boolean result = false;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      wrkBk = new XSSFWorkbook();
      sheet = wrkBk.createSheet();

      //
      // Loop through each of the NRHA departments that were listed in the input params.  This could be all the depts
      // or a subset.
      try {
         for ( int j = 0; j < MAX_NRHA; j++ ) {
            curNrha = m_NrhaList[j];

            if ( m_Status != RptServer.RUNNING || curNrha == null )
               break;

            //
            // Create the captions for each nrha department.
            rowNum = createCaptions(sheet, rowNum, curNrha);

            m_CustSales.setString(1, m_BegDate);
            m_CustSales.setString(2, m_EndDate);
            m_CustSales.setString(3, curNrha);
            custSales = m_CustSales.executeQuery();

            while ( custSales.next() && m_Status != RptServer.STOPPED ) {
               colNum = getColNum(custSales.getString("class"));
               nextVnd = custSales.getString("vendor_name");

               //
               // Check to see if we have the same vendor or if it's new.  If it's new then we need to move down
               // to the next row and start a new line item.
               if ( curVnd == null || nextVnd.compareTo(curVnd) != 0 ) {
                  row = createRow(sheet, rowNum);
                  curVnd = nextVnd;
                  rowNum++;
               }

               if ( row != null ) {
                  row.getCell(0).setCellValue(new XSSFRichTextString(curVnd));
                  row.getCell(colNum).setCellValue(custSales.getDouble("sales"));
                  row.getCell(++colNum).setCellValue(custSales.getDouble("cost"));
               }
            }

            custSales.close();

            //
            // Put a line break between nrha department listings.
            rowNum++;
         }

         wrkBk.write(outFile);
         wrkBk.close();
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.error("exception:", ex);
      }

      finally {
         sheet = null;
         row = null;
         wrkBk = null;
         
         try {
            outFile.close();
            outFile = null;
         }

         catch( Exception e ) {
            log.error(e);
         }         
      }

      return result;
   }

   /**
    * Closes a resultset and a statement.
    *
    * @param stmt The statement to close
    * @param set The resultset to close
    */
   private void closeStatement(Statement stmt, ResultSet set)
   {
      if ( set != null ) {
         try {
            set.close();
         }
         catch ( Exception e ) {
            ;
         }
      }

      if ( stmt != null ) {
         try {
            stmt.close();
         }
         catch ( Exception e ) {
            ;
         }
      }
   }

   /**
    * Sets the captions on the report.  This is just an ugly function, but I couldn't think of any better way to do it.
    * I suppose we could check for and odd/even situation wrt the nrha columns, but it just didn't seem worth it.  A
    * case statement seemed better than a gigantic if else boondogle.
    *
    * @param sheet HSSFSheet - the current sheet in the workbook.
    */
   private int createCaptions(XSSFSheet sheet, int startRow, String nrha)
   {
      XSSFRow row = null;
      XSSFCell cell = null;      
      String caption = "";
      int i = 0;

      if ( sheet != null ) {
         //
         // Create the report title
         if ( startRow == 0 ) {
            row = sheet.createRow(startRow);

            if ( row != null ) {
               caption = "NRHA, Customer Class, Vendor Sales " + m_BegDate + " - " + m_EndDate;
               cell = row.createCell(0);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(caption));

               //
               // Set all of the column widths
               for ( i = 0; i < m_ColCount; i++ ) {
                  if ( i == 0 )
                     sheet.setColumnWidth(0, 10000);
                  else
                     sheet.setColumnWidth(i, 3500);
               }
            }

            startRow+=2;
         }

         //
         // Create the NRHA caption
         caption = "NRHA " + nrha;
         row = sheet.createRow(startRow);
         cell = row.createCell(SALES_COL_START);
         cell.setCellType(CellType.STRING);
         cell.setCellValue(new XSSFRichTextString(caption));

         //
         // Create the cust class headings
         startRow++;
         row = sheet.createRow(startRow);
         if ( row != null ) {
            i = SALES_COL_START;

            for (int j = 0; j < m_ClassHeadings.size(); j++ ) {
               cell = row.createCell(i);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(new XSSFRichTextString(m_ClassHeadings.get(j)));

               //
               // Merge the cells for the caption so that it spans the sales and cost columns
               CellRangeAddress region = new CellRangeAddress(startRow, startRow, i, (i+1));
               sheet.addMergedRegion(region);
               region = null;

               //
               // skip over to the start of the next caption.
               i+=2;
            }
         }

         //
         // Create the Sales/Cost captions
         startRow++;
         row = sheet.createRow(startRow);
         if ( row != null ) {
            for ( i = SALES_COL_START; i < m_ColCount; i++ ) {
               cell = row.createCell(i);
               cell.setCellType(CellType.STRING);

               if ( i % 2 != 0 )
                  cell.setCellValue(new XSSFRichTextString("Sales"));
               else
                  cell.setCellValue(new XSSFRichTextString("Cost"));
            }
         }
      }

      return ++startRow;
   }
   
   /**
    * Creates the report.
    * 
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {      
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         m_OraConn = m_RptProc.getOraConn();
         
         if ( prepareStatements() ) {
            //
            // Loads the customer class data into an array and sets the number of columns
            // for the spreadsheet.
            getCustClasses();
            created = buildOutputFile();
         }
                        
      }
      
      catch ( Exception ex ) {
         log.fatal("exception:", ex);
      }
      
      finally {         
         setCurAction("");
         
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }
      
      return created;
   }
   
   /**
    * Creates a row in the worksheet.
    *
    * @param sheet HSSFSheet the current sheet in the workbook.
    * @param rowNum short the current row number in the sheet.
    *
    * @return HSSFRow a reference to a new row with formatted cells.
    */
   private XSSFRow createRow(XSSFSheet sheet, int rowNum)
   {
      XSSFRow Row = null;
      XSSFCell Cell = null;
      XSSFCellStyle style = null;

      if ( sheet == null )
         return Row;

      Row = sheet.createRow(rowNum);

      //
      // iterate through the columns and set the style for each.  The non caption rows are numeric/money types
      // that contain cost and sales amounts.
      if ( Row != null ) {
         for ( int i = 0; i < m_ColCount; i++ ) {
            if ( i < SALES_COL_START ) {
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
               style = null;
            }
         }
      }

      return Row;
   }

   /**
    * Returns the column number that the cust class starts in.
    *
    * @param custClass The customer class to look up
    * @return The column number that custClass starts in.
    */
   private short getColNum(String custClass)
   {
      return (m_CustClass.get(custClass)).shortValue();
   }

   /**
    * Creates a list of customer classes and puts them in a hash map.  This allows the lookup of the correct
    * column when displaying the data for a specific class.  Also sets the sorted list of headings for the classes.
    * For whatever reason, the hashmap does not return a sorted alpha list when getting the key set.  This causes the
    * headings to be incorrect.
    *
    * @throws SQLException
    */
   private void getCustClasses() throws SQLException
   {
      ResultSet set = null;
      Statement stmt = null;
      short col = 1;
      String custClass;

      try {
         stmt = m_OraConn.createStatement();
         set = stmt.executeQuery(
            "select distinct(class) from cust_market_view where market = 'CUSTOMER TYPE' order by class"
         );

         while ( set.next() && m_Status == RptServer.RUNNING ) {
            custClass = set.getString(1);
            m_CustClass.put(custClass, new Short(col));
            m_ClassHeadings.add(custClass);
            col += 2;
         }

         //
         // double the number of classes becuase we have sales and cost, then add one for the vendor
         // column.
         m_ColCount = (m_CustClass.size() * 2) + 1;
      }

      finally {
         closeStatement(stmt, set);
         stmt = null;
         set = null;
      }
   }

   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(256);      
      boolean isPrepared = false;

      if ( m_OraConn != null ) {
         try {
            sql.append("select ");
            sql.append("   vendor_name, class, sum(qty_shipped * unit_sell) Sales, sum(qty_shipped * unit_cost) Cost ");
            sql.append("from sale_item, cust_market_view cmv ");
            sql.append("where ");
            sql.append("   tran_type in('SALE', 'CREDIT') and sale_type = 'WAREHOUSE' and ");
            sql.append("   invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') and ");
            sql.append("   nrha = ? and market = 'CUSTOMER TYPE' and ");
            sql.append("   sale_item.customer_id = cmv.customer_id ");
            sql.append("group by vendor_name, class ");
            sql.append("order by vendor_name, class");
   
            m_CustSales = m_OraConn.prepareStatement(sql.toString());
            isPrepared = true;
         }
   
         catch (SQLException ex) {
            log.error(ex);
         }
      }
      
      return isPrepared;
   }

   /**
    * Sets the parameters of this report.  Handles processing the parameters and also
    * creates the file name for the report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();
      String tmp = Long.toString(System.currentTimeMillis());
      int i = 0;
      int j = 0;
      int k = 0;
      
      m_BegDate = params.get(0).value;
      m_EndDate = params.get(1).value;
      tmp = params.get(2).value;
      
      //
      // Load the nrha list parameter into an array.  The list comes in as a comma delimeted string.
      while ( i != -1 ) {
         i = tmp.indexOf(',', i);

         if ( i != -1 ) {
            m_NrhaList[k] = tmp.substring(j, i);

            i++;
            j = i;
            k++;
         }
         else {
            //
            // Handle the case of one nrha or the last one in the list.
            i = tmp.length();

            if ( j < i ) {
               m_NrhaList[k] = tmp.substring(j, i);
               i = -1;
            }
         }
      }
      
      tmp = Long.toString(System.currentTimeMillis());
      fileName.append("csnrhav");      
      fileName.append(m_BegDate.replaceAll("/", "-"));
      fileName.append(m_EndDate.replaceAll("/", "-"));
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
   }
}
