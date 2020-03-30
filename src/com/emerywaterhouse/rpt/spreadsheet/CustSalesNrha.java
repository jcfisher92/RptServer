/**
 * File: CustSalesNrha.java
 * Description: 
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 5/10/2005
 * Last Update: $Id: CustSalesNrha.java,v 1.8 2012/07/31 17:06:03 npasnur Exp $
 * 
 * History
 *    $Log: CustSalesNrha.java,v $
 *    Revision 1.8  2012/07/31 17:06:03  npasnur
 *    Fixed the params order for method CellRangeAddress after POI upgrade
 *
 *    Revision 1.7  2009/02/18 14:30:39  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    10/25/2004 - Switched ftp server address to name (addressed changed) jbh
 *    
 *    06/14/2004 - Added an is not null to the main query to remove data that has null nrha data. - jcf
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

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


public class CustSalesNrha extends Report
{
   private final short COL_COUNT       = 27;
   private final short SALES_COL_START = 3;
   private final short DATA_ROW_START  = 2;
   private final short MAX_NRHA        = 12;

   private String m_BegDate;
   private PreparedStatement m_CustSales;
   private String m_EndDate;
   
   /**
    * default constructor
    */
   public CustSalesNrha()
   {
      super();
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
      String custId = "";
      String nextCustId = "";
      int nrha = 0;
      int i = SALES_COL_START;
      int rowNum = DATA_ROW_START;
      boolean result = false;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      wrkBk = new XSSFWorkbook();
      sheet = wrkBk.createSheet();

      try {
         createCaptions(sheet);

         m_CustSales.setString(1, m_BegDate);
         m_CustSales.setString(2, m_EndDate);
         custSales = m_CustSales.executeQuery();

         while ( custSales.next() && m_Status == RptServer.RUNNING ) {
            nextCustId = custSales.getString("customer_id");
            setCurAction("processing customer: " + nextCustId);

            //
            // We have to break on each customer and then create a cross tab report based on the nrha dept.  We can do
            // this by comparing first the customer and then the nrha for each iteration through the resultset
            //
            // Note - According to Karen Jorgensen, a customer will only have one class so we won't break on the
            //    class.
            if ( !nextCustId.equals(custId) ) {
               custId = nextCustId;
               row = createRow(sheet, rowNum);

               row.getCell(0).setCellValue(new XSSFRichTextString(custId));
               row.getCell(1).setCellValue(new XSSFRichTextString(custSales.getString("name")));
               row.getCell(2).setCellValue(new XSSFRichTextString(custSales.getString("class")));
               rowNum++;
            }

            //
            // Convert the nrha to an integer so we can use (n-1)*2 + colstart as a formula for determining the
            // column starting position based on the nrha.
            nrha = Short.parseShort(custSales.getString("nrha"));

            //
            // There are some bogus nrha departments in the nrha table.  We only want the first 12.
            if ( nrha <= MAX_NRHA ) {
               i = ((nrha-1) * 2 + SALES_COL_START);

               if ( row != null ) {
                  row.getCell(i).setCellValue(custSales.getDouble("sales"));
                  row.getCell((i+1)).setCellValue(custSales.getDouble("cost"));
               }
            }
         }

         wrkBk.write(outFile);
         wrkBk.close();
         custSales.close();

         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {
         sheet = null;
         row = null;
         wrkBk = null;

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
    * Sets the captions on the report.  This is just an ugly function, but I couldn't think of any better way to do it.
    * I suppose we could check for and odd/even situation wrt the nrha columns, but it just didn't seem worth it.  A
    * case statement seemed better than a gigantic if else boondogle.
    *
    * @param sheet HSSFSheet - the current sheet in the workbook.
    */
   private void createCaptions(XSSFSheet sheet)
   {
      XSSFRow row = null;
      XSSFCell cell = null;
      String caption = "";
      int j = 0;
      int colWidth = 4000;
      boolean merge = false;

      if ( sheet != null ) {
         for ( j = 0; j < 2; j++ ) {
            row = sheet.createRow(j);

            if ( row != null ) {
               for ( int i = 0; i < COL_COUNT; i++ ) {
                  cell = row.createCell(i);
                  cell.setCellType(CellType.STRING);

                  switch ( i ) {
                     case 0:
                        if ( j == 1 ) {
                           caption = "Customer ID";
                           merge = false;
                        }
                        else {
                           caption = "Customer NRHA Sales " + m_BegDate + " to " + m_EndDate;
                           merge = true;
                        }

                        colWidth = 3000;
                     break;

                     case 1:
                        if ( j == 1 )
                           caption = "Customer Name";
                        else
                           caption = "";

                        colWidth = 13000;
                        merge = false;
                     break;

                     case 2:
                        if ( j == 1 )
                           caption = "Class";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 3:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 01";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 4:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        merge = false;
                        colWidth = 4000;
                     break;

                     case 5:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 02";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 6:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 7:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 03";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 8:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 9:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 04";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 10:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 11:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 05";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 12:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 13:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 06";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 14:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 15:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 07";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 16:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 17:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 08";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 18:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 19:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 09";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 20:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 21:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 10";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 22:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 23:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 11";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 24:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     case 25:
                        if ( j == 1 ) {
                           caption = "Sales";
                           merge = false;
                        }
                        else {
                           caption = "NRHA 12";
                           merge = true;
                        }

                        colWidth = 4000;
                     break;

                     case 26:
                        if ( j == 1 )
                           caption = "Cost";
                        else
                           caption = "";

                        colWidth = 4000;
                        merge = false;
                     break;

                     default:
                        colWidth = 4000;
                        caption = "";
                        merge = false;
                     break;
                  }

                  sheet.setColumnWidth(i, colWidth);
                  cell.setCellValue(new XSSFRichTextString(caption));

                  if ( merge ) {
                     CellRangeAddress region = new CellRangeAddress(j, j, i, (i+1));
                     sheet.addMergedRegion(region);
                  }
               }
            }
         }
      }
   }

   /**
    * Cleanup allocated resources.    
    */
   private void cleanup()
   {
      m_BegDate = null;
      m_EndDate = null;
      
      if ( m_CustSales != null ) {
         try {
            m_CustSales.close();
            m_CustSales = null;
         }
         
         catch ( Exception ex ) {
            
         }
      }
   }
   
   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
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
         cleanup();
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
      // iterate through the columns and set the style for each.  The last 24 columns are numeric/money types
      // that contain cost and sales amounts.
      if ( Row != null ) {
         for ( int i = 0; i < COL_COUNT; i++ ) {
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
               Cell.setCellValue(0.0);

               style = null;
            }
         }
      }

      return Row;
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
            sql.append("select sale_item.customer_id, c.name, class, nrha, ");
            sql.append("sum(qty_shipped * unit_sell) Sales, sum(qty_shipped * unit_cost) Cost ");
            sql.append("from sale_item, cust_market_view cmv, customer c ");
            sql.append("where tran_type in('SALE', 'CREDIT') and sale_type = 'WAREHOUSE' and ");
            sql.append("invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') and ");
            sql.append("nrha is not null and ");
            sql.append("market = 'CUSTOMER TYPE' and ");
            sql.append("c.customer_id = sale_item.customer_id and ");
            sql.append("cmv.customer_id = sale_item.customer_id ");
            sql.append("group by sale_item.customer_id, c.name, class, nrha ");
            sql.append("order by class, sale_item.customer_id, nrha");
   
            m_CustSales = m_OraConn.prepareStatement(sql.toString());
            isPrepared = true;
         }
   
         catch ( SQLException ex ) {
            isPrepared = false;
            log.error("exception:", ex);
         }
      }
      
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
            
      m_BegDate = params.get(0).value;
      m_EndDate = params.get(1).value;
      String tmp1 = m_BegDate.replaceAll("/", "-");
      String tmp2 = m_EndDate.replaceAll("/", "-");;

      fileName.append("csnrha");      
      fileName.append(tmp1);
      fileName.append(tmp2);
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
   }
}
