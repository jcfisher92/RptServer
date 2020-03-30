/**
 * File: TopVendors.java
 * Description: Report that shows the sales of Emery's top customers with the top vendors.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 01/08/2007
 * Last Update: $Id: TopVendors.java,v 1.5 2011/04/25 07:00:16 npasnur Exp $
 * 
 * History
 *    $Log: TopVendors.java,v $
 *    Revision 1.5  2011/04/25 07:00:16  npasnur
 *    Changed the params order for method CellRangeAddress after POI upgrade
 *
 *    Revision 1.4  2009/02/18 16:53:10  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.helper.VendorData;
import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class TopVendors extends Report
{   
   private PreparedStatement m_VndData;
   private PreparedStatement m_CustData;
   private PreparedStatement m_CustSales;
   private PreparedStatement m_CustName;
   private PreparedStatement m_VndSales;
   
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;
   
   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;
   
   private int m_ColCount;
   private int m_CustCount;
   private int m_VndCount;
   private String m_BegDate;
   private String m_EndDate;
   private VendorData m_Vendors[];
   
   /**
    * 
    */
   public TopVendors()
   {
      super();
      
      m_VndCount = 20;
      m_CustCount = 100;
            
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      setupWorkbook();      
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable 
    */
   public void finalize() throws Throwable
   {      
      if ( m_CellStyles != null ) {
         for ( int i = 0; i < m_CellStyles.length; i++ )
            m_CellStyles[i] = null;
      }
      
      if ( m_Vendors != null ) {
         for ( int i = 0; i < m_Vendors.length; i++ )
            m_Vendors[i] = null;
      }
      
      m_Sheet = null;
      m_Wrkbk = null;      
      m_CellStyles = null;
      
      super.finalize();
   }
   

   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      XSSFRow row = null;
      int rowNum = 0;
      int colNum = 0;
      FileOutputStream outFile = null;
      ResultSet vndData = null;
      ResultSet custData = null;
      ResultSet custSales = null;
      boolean result = false;
      VendorData vnd = null;
      int i = 0;
      String custId = null;
      String custName = null;
      double pct = 0;
      double custVndSales = 0;
      double vndSales = 0;
      int vndId;
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         m_VndData.setString(1, m_BegDate);
         m_VndData.setString(2, m_EndDate);
         vndData = m_VndData.executeQuery();
         
         m_CustData.setString(1, m_BegDate);
         m_CustData.setString(2, m_EndDate);
         custData = m_CustData.executeQuery();
         
         //
         // Buld the vendor list.  Fill up the array with the vendor data
         while ( vndData.next() && i < m_VndCount && m_Status == RptServer.RUNNING ) {
            vnd = new VendorData();            
            
            vnd.m_VndId = vndData.getInt(1);
            vnd.m_VndName = vndData.getString(2);
            vnd.m_Sales = getVendSales(vnd.m_VndId);
            
            m_Vendors[i++] = vnd;
         }
         
         rowNum = createCaptions();         
         i = 0;
         m_CustSales.setString(3, m_BegDate);
         m_CustSales.setString(4, m_EndDate);
         
         //
         // Go through each of the top customers and get their pct sales of each vendor
         while ( custData.next() && i < m_CustCount && m_Status == RptServer.RUNNING ) {
            colNum = 0;
            custId = custData.getString(1);
            custName = getCustName(custId);
            
            if ( custName == null )
               custName = "";            
            
            m_CustSales.setString(1, custId);            
            row = createRow(rowNum++, m_ColCount);
                        
            if ( row != null ) {               
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(custId + " " + custName));
               
               if ( m_Vendors != null ) {
                  for ( int j = 0; j < m_Vendors.length && m_Status == RptServer.RUNNING; j++ ) {
                     vndId = m_Vendors[j].m_VndId;
                     vndSales = m_Vendors[j].m_Sales;
                     
                     setCurAction(
                        "processing cust acct: " + custId + 
                        String.format(" (%d of %d)", i+1, m_CustCount) + 
                        " vendor " + Integer.toString(vndId) 
                     );
                     
                     m_CustSales.setString(2, Integer.toString(vndId));                     
                     custSales = m_CustSales.executeQuery();
                                          
                     if ( custSales.next() ) {
                        custVndSales = custSales.getDouble(1);
                        pct = custVndSales / vndSales;
                     }
                     
                     row.getCell(colNum++).setCellValue(custVndSales);
                     row.getCell(colNum++).setCellValue(pct);
                     
                     closeRSet(custSales);         
                     custVndSales = 0;
                     pct = 0;
                  }
               }
            }
            
            i++;
         }
         
         m_Sheet.createFreezePane(1, 3);
         m_Wrkbk.write(outFile);
         
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("TopVendors:", ex);
      }

      finally {         
         closeRSet(vndData);
         closeRSet(custData);
         closeRSet(custSales);
         
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error(e);
         }

         outFile = null;
         row = null;
         vnd = null;
         vndData = null;
         custData = null;
      }

      return result;
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_VndData);
      closeStmt(m_CustData);
      closeStmt(m_CustName);
      closeStmt(m_CustSales);
      closeStmt(m_VndSales);
   }
   
   /**
    * Sets the captions on the report.
    */
   private int createCaptions()
   {
      XSSFRow row = null;
      XSSFCell cell = null;
      XSSFFont font = null;
      XSSFCellStyle styleTitle;
      CellRangeAddress region = null;      
      int rowNum = 0;
      int colNum = 0;
      int vndNum = 0;
            
      if ( m_Sheet == null )
         return 0;
      
      font = m_Wrkbk.createFont();
      font.setFontHeightInPoints((short)10);
      font.setFontName("Arial");
      font.setBold(true);
      
      styleTitle = m_Wrkbk.createCellStyle();
      styleTitle.setAlignment(HorizontalAlignment.LEFT);
      styleTitle.setFont(font);
      
      row = m_Sheet.createRow(rowNum++);
      cell = row.createCell(colNum);
      cell.setCellStyle(styleTitle);
      cell.setCellValue(new XSSFRichTextString("Top Vendors: " + m_BegDate + " to " + m_EndDate));
      
      //
      // Merge the title cells 
      region = new CellRangeAddress(0, 0, colNum, (colNum+1));
      m_Sheet.addMergedRegion(region);
            
      //
      // Create the first row of captions.  These are the vendor names and will span
      // two cells below.
      row = m_Sheet.createRow(rowNum);
      if ( row != null ) {
         cell = row.createCell(colNum);
         cell.setCellStyle(styleTitle);
         cell.setCellValue(new XSSFRichTextString("Accts"));
                  
         try {
            for ( colNum = 1; colNum < m_ColCount; colNum+=2 ) {
               cell = row.createCell(colNum);
               cell.setCellStyle(styleTitle);
               cell.getCellStyle().setWrapText(true);
               
               //
               // Merge the title cells.  Gives a better look to the report.
               region = new CellRangeAddress(rowNum, rowNum, colNum, (colNum+1));
               m_Sheet.addMergedRegion(region);               
               cell.setCellValue(
                     new XSSFRichTextString(m_Vendors[vndNum].m_VndId + " " + m_Vendors[vndNum++].m_VndName)
               );
            }
            
            //
            // Set the height of the vendor names row so they'll break
            row.setHeight((short)(row.getHeight()*4));
            rowNum++;
         }
         
         catch( Exception ex ) {
            ;
         }
      }
      
      //
      // Create the two columns below the vendor names.  One for dollars, one for pct
      row = m_Sheet.createRow(rowNum++);
      colNum = 0;
      
      if ( row != null ) {
         //
         // Set the acct column width
         m_Sheet.setColumnWidth(colNum, 7000);
         
         //
         // Set the dollars and pct columns
         for ( int i = 1; i < m_ColCount; i++ ) {
            cell = row.createCell(i);
                        
            if ( i == 1 || i % 2  != 0 ) {
               cell.setCellValue(new XSSFRichTextString("Dollars"));
               m_Sheet.setColumnWidth(i, 3000);
            }
            else
               cell.setCellValue(new XSSFRichTextString("% Total"));
         }
      }

      cell = null;
      row = null;
      font = null;
      styleTitle = null;
      
      return rowNum;
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
         log.fatal("TopVendors:", ex);
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
   private XSSFRow createRow(int rowNum, int colCnt)
   {
      XSSFRow row = null;
      XSSFCell cell = null;
      
      if ( m_Sheet != null ) {
         row = m_Sheet.createRow(rowNum);

         //
         // set the type and style of the cell.
         if ( row != null ) {
            for ( int i = 0; i < colCnt; i++ ) {            
               cell = row.createCell(i);
               
               if ( i == 0 )
                  cell.setCellStyle(m_CellStyles[0]);
               else {            
                  if ( i % 2 != 0 )
                     cell.setCellStyle(m_CellStyles[1]);
                  else
                     cell.setCellStyle(m_CellStyles[2]);
               }
            }
         }
      }

      return row;
   }
   
   /**
    * Gets the customer name based on a customer account number
    * 
    * @param custId The customer number to get the customer name for.
    * @return The customer name
    */
   private String getCustName(String custId)
   {
      String name = "";
      ResultSet custName = null;
      
      try {
         setCurAction("getting acct name for acct: " + custId);
         m_CustName.setString(1, custId);         
         custName = m_CustName.executeQuery();
         
         if ( custName.next() ) {
            name = custName.getString(1);
         }
      }
      
      catch ( Exception ex ) {
         log.error("TopVendors.getCustName", ex);
      }
      
      finally {
         closeRSet(custName);
         custName = null;
      }
      
      return name;
   }
   
   /**
    * Gets the total vendor sales for a specific time period
    * 
    * @param vndId The vendor number to get sales for.
    * @return The vendor sales for the vendor in vndId
    */
   private double getVendSales(int vndId)
   {
      double sales = 0.0;
      ResultSet vndSales = null;
      String vnd = Integer.toString(vndId); 
      
      try {
         setCurAction("getting sales for vendor: " + vnd);
         m_VndSales.setString(1, vnd);
         m_VndSales.setString(2, m_BegDate);
         m_VndSales.setString(3, m_EndDate);
         vndSales = m_VndSales.executeQuery();
         
         if ( vndSales.next() ) {
            sales = vndSales.getDouble(1);
         }
      }
      
      catch ( Exception ex ) {
         log.error("TopVendors.getVendorSales", ex);
      }
      
      finally {
         closeRSet(vndSales);
         vndSales = null;
         vnd = null;
      }
      
      return sales;
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
            sql.append("select /*+rule*/ \r\n");
            sql.append("vendor.vendor_id, name, sum(qty_put_away * emery_cost) total_dollars \r\n");
            sql.append("from po_hdr, po_dtl, vendor \r\n");
            sql.append("where \r\n");
            sql.append("   po_date >= to_date(?, 'mm/dd/yyyy') and \r\n");
            sql.append("   po_date <= to_date(?, 'mm/dd/yyyy') and \r\n");
            sql.append("   po_hdr.po_hdr_id = po_dtl.po_hdr_id and \r\n");
            sql.append("   vendor.vendor_id = po_hdr.vendor_id \r\n");
            sql.append("group by vendor.vendor_id, name \r\n");
            sql.append("order by total_dollars desc");
            
            m_VndData = m_OraConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select \r\n");
            sql.append("   cust_acct, sum(dollars_shipped) total  \r\n");
            sql.append("from sale \r\n");
            sql.append("where \r\n");
            sql.append("   invoice_date >= to_date(?, 'mm/dd/yyyy') and \r\n");
            sql.append("   invoice_date <= to_date(?, 'mm/dd/yyyy') \r\n");
            sql.append("group by cust_acct \r\n");
            sql.append("order by total desc \r\n");            
            m_CustData = m_OraConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select /*+ rule*/ \r\n");
            sql.append("   sum(ext_sell) vnd_total \r\n");
            sql.append("from \r\n");
            sql.append("   inv_dtl \r\n");
            sql.append("where \r\n");
            sql.append("   cust_acct = ? and \r\n");
            sql.append("   vendor_nbr = ? and \r\n");            
            sql.append("   invoice_date >= to_date(?, 'mm/dd/yyyy') and \r\n");
            sql.append("   invoice_date <= to_date(?, 'mm/dd/yyyy') \r\n");
            m_CustSales = m_OraConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select sum(dollars_shipped) sales \r\n");
            sql.append("from vendorsales \r\n");
            sql.append("where \r\n");
            sql.append("   vendor_nbr = ? and \r\n");
            sql.append("   invoice_date >= to_date(?, 'mm/dd/yyyy') and \r\n");
            sql.append("   invoice_date <= to_date(?, 'mm/dd/yyyy')");
            m_VndSales = m_OraConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select name \r\n");
            sql.append("from customer \r\n");
            sql.append("where \r\n");
            sql.append("   customer_id = ?");
            m_CustName = m_OraConn.prepareStatement(sql.toString());
            
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
         log.error("TopVendors.prepareStatements - null oracle connection");
      
      return isPrepared;
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();      
      String tmp = Long.toString(System.currentTimeMillis());      
      int pcount = params.size();
      Param param = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                  
         if ( param.name.equals("custcount") )
            m_CustCount = Integer.parseInt(param.value);
         
         if ( param.name.equals("vndcount") )
            m_VndCount = Integer.parseInt(param.value);
         
         if ( param.name.equals("begdate") )
            m_BegDate = param.value;
         
         if ( param.name.equals("enddate") )
            m_EndDate = param.value;
      }
      
      //
      // Set the column count based on the number of vendors.
      // This will account for the pair of cells under the vendor and the accts column.
      m_ColCount = ((m_VndCount * 2) + 1);
      m_Vendors = new VendorData[m_VndCount];
      
      fileName.append("topvnd");      
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
   }
      
   /**
    * Sets up the styles for the cells based on the column data.  Does any other inititialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {      
      XSSFCellStyle styleText;      // Text left justified
      XSSFCellStyle styleInt;       // Style with 0 decimals
      XSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00)
      XSSFCellStyle stylePct;       // Style with 0 decimals + %
            
      styleText = m_Wrkbk.createCellStyle();      
      styleText.setAlignment(HorizontalAlignment.LEFT);
      
      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);

      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short)8);
      
      stylePct = m_Wrkbk.createCellStyle();
      stylePct.setAlignment(HorizontalAlignment.RIGHT);
      stylePct.setDataFormat((short)9);
      
      //
      // We only have three styles.  The money and pct styles are in the paired cells.
      m_CellStyles = new XSSFCellStyle[] {
         styleText,    // All text cells         
         styleMoney,   // Dollar value cells
         stylePct      // percent value cells
      };
      
      styleText = null;
      styleInt = null;
      styleMoney = null;
   }
}
