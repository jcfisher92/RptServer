/**
* File: BuyGroupInvoiceCheck.java
* Description: Buy Group Invoice Check
*
* @author Jeff Fisher
* @author Eric Verge
*
* Create Date: 07/03/2014
* Last Update: $Id: BuyGroupInvoiceCheck.java,v 1.8 2014/07/17 14:49:53 everge Exp $
*
* History:
*    $Log: BuyGroupInvoiceCheck.java,v $
*    Revision 1.8  2014/07/17 14:49:53  everge
*    Report now prints in landscape mode by default. Wow.
*
*    Revision 1.7  2014/07/10 20:23:18  everge
*    Fixed "sent gentran" column to be blank.
*
*    Revision 1.6  2014/07/10 16:45:37  everge
*    Added footer w/ page numbers
*
*    Revision 1.5  2014/07/10 16:10:15  everge
*    Removed title row in favor of a print header that includes report title and date, adjusted column widths to better fit data
*
*    Revision 1.4  2014/07/09 22:41:55  everge
*    Fixed the creation of the output file and formatting of the SQL query. Report runs successfully now.
*
*    Revision 1.3  2014/07/09 14:01:40  everge
*    Added this header and made exception logging more explicit
*
*/

package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
//import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;


public class BuyGroupInvoiceCheck extends Report
{
   private static final short MAX_COLS = 7;
   
   private PreparedStatement m_BuyGrpInvData;
   
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private HSSFCellStyle[] m_CellStyles;
   
   //
   // workbook entries.
   private HSSFWorkbook m_Wrkbk;
   private HSSFSheet m_Sheet;
   private HSSFFont m_FontBold;
   private HSSFFont m_FontNormal;
   
   /**
    * Default constructor 
    */
   public BuyGroupInvoiceCheck()
   {
      super();
      m_Wrkbk = new HSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      m_Sheet.getPrintSetup().setLandscape(true);
      setupWorkbook();      
   }

   /**
    * Cleanup any allocated resources.
    */
   public void finalize() throws Throwable
   {      
      if ( m_CellStyles != null ) {
         for ( int i = 0; i < m_CellStyles.length; i++ )
            m_CellStyles[i] = null;
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
      HSSFRow row = null;
      int rowNum = 0;
      int colNum = 0;
      FileOutputStream outFile = null;
      ResultSet buyGrpInvData = null;
      boolean result = false;
      
      StringBuffer fileName = new StringBuffer();      
      String tmp = Long.toString(System.currentTimeMillis());
                  
      fileName.append("buygp-invchk");      
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xls");
      m_FileNames.add(fileName.toString());
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         rowNum = createCaptions();    
         buyGrpInvData = m_BuyGrpInvData.executeQuery();
         m_CurAction = "Building output file";

         while (buyGrpInvData.next() && m_Status == RptServer.RUNNING) {
            row = createRow(rowNum++, MAX_COLS);
            colNum = 0;
            
            if ( row != null ) {               
               row.getCell(colNum++).setCellValue(buyGrpInvData.getString("Invoice#"));
               row.getCell(colNum++).setCellValue(buyGrpInvData.getString("Cust ID"));
               row.getCell(colNum++).setCellValue(buyGrpInvData.getString("Cust Name"));
               row.getCell(colNum++).setCellValue(buyGrpInvData.getString("Buying Group"));
               row.getCell(colNum++).setCellValue(buyGrpInvData.getString("Date Processed"));
               row.getCell(colNum++).setCellValue(buyGrpInvData.getString("Sent DB"));
               //"Sent Gentran" field left blank
            }
         }
         
         m_Wrkbk.write(outFile);
         buyGrpInvData.close();

         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your Buy Group Invoice Check had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {         
         row = null;
                  
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
      closeStmt(m_BuyGrpInvData);
   }
   
   /**
    * Sets the captions on the report.
    */
   private int createCaptions()
   {
	  if ( m_Sheet == null )
	     return 0;
	  
      HSSFRow row = null;
      HSSFCell cell = null;
      HSSFCellStyle styleTitle;// Bold, center-aligned style
      int rowNum = 0;
      int colNum = 0;
      
      styleTitle = m_Wrkbk.createCellStyle();
      styleTitle.setFont(m_FontBold);
      styleTitle.setAlignment(HorizontalAlignment.CENTER);
           
      // Create document header & footer for printing
      HSSFHeader header = m_Sheet.getHeader();
      header.setLeft(HSSFHeader.font("Helvetica", "Bold") +
          HSSFHeader.fontSize((short)16) + "Buy Group Invoice Check");
      header.setRight(HSSFHeader.date() + " ");
      
      HSSFFooter footer = m_Sheet.getFooter();
      footer.setRight("Pg " + HSSFFooter.page() + " of " + HSSFFooter.numPages());
      
      // Build column titles
      row = m_Sheet.createRow(rowNum++);
      
      if ( row != null ) {
         for ( int i = 0; i < MAX_COLS; i++ ) {
            cell = row.createCell(i);
            cell.setCellType(CellType.STRING);
            cell.setCellStyle(styleTitle);
         }
         
         m_Sheet.setColumnWidth(colNum, 2100);
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Invoice#"));
         
         m_Sheet.setColumnWidth(colNum, 1800);
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Cust ID"));
         
         m_Sheet.setColumnWidth(colNum, 9000);
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Cust Name"));
         
         m_Sheet.setColumnWidth(colNum, 3500);
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Buying Group"));
         
         m_Sheet.setColumnWidth(colNum, 3800);
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Date Processed"));
         
         m_Sheet.setColumnWidth(colNum, 1800);
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Sent DB"));
         
         m_Sheet.setColumnWidth(colNum, 3100);
         row.getCell(colNum).setCellValue(new HSSFRichTextString("Sent Gentran"));
      }
      
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
         m_EdbConn = m_RptProc.getEdbConn();         
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
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    * @param colCnt The number of columns in the row.
    * 
    * @return The fromatted row of the spreadsheet.
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
            cell.setCellStyle(m_CellStyles[i]);
         }
      }

      return row;
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;
      
      if ( m_EdbConn != null ) {
         try {
            sql.append("select ");
            sql.append("   invoice_num \"Invoice#\", invoice.customer_id \"Cust ID\", customer.name \"Cust Name\", ");
            sql.append("   buying_group.name \"Buying Group\", date_processed \"Date Processed\", ");
            sql.append("   sent_x12_inv \"Sent DB\", '' \"Sent Gentran\" ");
            sql.append("from invoice ");
            sql.append("join customer on customer.customer_id = invoice.customer_id ");
            sql.append("join cust_buy_group on cust_buy_group.customer_id = invoice.customer_id ");
            sql.append("   and cust_buy_group.end_date is null or cust_buy_group.end_date >= trunc(sysdate) ");
            sql.append("join buying_group on buying_group.buy_group_id = cust_buy_group.buy_group_id ");
            sql.append("   and buying_group.name in ('LBMA', 'LMC') ");
            sql.append("join sale_type on sale_type.sale_type_id = invoice.sale_type_id ");
            sql.append("   and sale_type.name in ('WAREHOUSE') ");
            sql.append("where invoice_date = trunc(sysdate)-1 and invoice_amt <> 0 ");
            sql.append("order by buying_group.name, invoice_num, customer.customer_id ");
            
            m_BuyGrpInvData = m_EdbConn.prepareStatement(sql.toString());
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
         log.error("BuyGroupInvoiceCheck.prepareStatements - null oracle connection");
      
      return isPrepared;
   }
      
   /**
    * Sets up the styles for the cells based on the column data. Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {      
      HSSFCellStyle 
         styleTextL = m_Wrkbk.createCellStyle(), // Text left justified
         styleTextR = m_Wrkbk.createCellStyle(), // Text right justified
         styleTextC = m_Wrkbk.createCellStyle(); // Text center justified
      
      // Create a font that is normal size & bold weight
      m_FontBold = m_Wrkbk.createFont();
      m_FontBold.setFontHeightInPoints((short)8);
      m_FontBold.setFontName("Helvetica");
      m_FontBold.setBold(true);
      
      // Create a font that is normal size & normal weight
      m_FontNormal = m_Wrkbk.createFont();
      m_FontNormal.setFontHeightInPoints((short)8);
      m_FontNormal.setFontName("Helvetica");
                  
      styleTextL.setAlignment(HorizontalAlignment.LEFT);
      styleTextL.setFont(m_FontNormal);
      styleTextR.setAlignment(HorizontalAlignment.RIGHT);
      styleTextR.setFont(m_FontNormal);
      styleTextC.setAlignment(HorizontalAlignment.CENTER);
      styleTextC.setFont(m_FontNormal);
      
      m_CellStyles = new HSSFCellStyle[] {
         styleTextL,    // col 0 inv#
         styleTextL,    // col 1 cust id
         styleTextL,    // col 2 cust name
         styleTextC,    // col 3 buy grp
         styleTextR,    // col 4 date proc
         styleTextR,    // col 5 sent db
         styleTextR,    // col 6 sent gntrn        
      };
      
      styleTextL = null;
      styleTextR = null;
      styleTextC = null;
   }
}
