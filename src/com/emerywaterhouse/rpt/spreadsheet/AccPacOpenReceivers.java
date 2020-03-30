/**
 * File: AccPacOpenReceivers.java
 * Description: Accpac report that shows the open receivers.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 09/28/2006
 * Last Update: $Id: AccPacOpenReceivers.java,v 1.8 2012/12/03 17:04:22 prichter Exp $
 * 
 * History
 *    $Log: AccPacOpenReceivers.java,v $
 *    Revision 1.8  2012/12/03 17:04:22  prichter
 *    Reworked the query to handle non-numeric termcodes in accpac
 *
 *    Revision 1.7  2009/03/04 20:47:54  jfisher
 *    Added fields per jason w
 *
 *    Revision 1.6  2009/02/17 22:38:58  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.5  2008/10/29 20:33:37  jfisher
 *    Fixed potential null variable warning
 *
 *    Revision 1.4  2008/08/29 15:59:15  jfisher
 *    Added the warehouse to the report.
 *
 *    Revision 1.3  2006/11/02 19:29:26  jfisher
 *    Added a new column to the spreadsheet.
 *
 *    Revision 1.2  2006/10/27 13:15:08  jfisher
 *    New query from systems
 *
 *    Revision 1.1  2006/10/17 14:00:17  jfisher
 *    added an outer join to the eis table.
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class AccPacOpenReceivers extends Report
{
   private static final short MAX_COLS = 14;
   
   private PreparedStatement m_RcvrData;
   private PreparedStatement m_RcvrPoHdr;
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;
   
   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;
   
   /**
    * Initialize report variables.
    */
   public AccPacOpenReceivers()
   {
      super();
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      setupWorkbook();
   }
   
   /**
    * Cleanup any allocated resources.
    * 
    * @throws Throwable
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
      FileOutputStream outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
      
      XSSFRow row = null;
      int rowNum = 0;
      int colNum = 0;      
      ResultSet rcvrData = null;
      ResultSet rs = null;
      boolean result = false;
      int vndId = 0;
      int termsCode = 0;
      int rowCount = 0;
      Date rcptDate = null;
      Date dueDate = null;
      Date discDate = null;
      String poNbr = null;
      String whsName = null;
      String reviewPurch = null;
      
      try {
         rowNum = createCaptions();
         rcvrData = m_RcvrData.executeQuery();

         while ( rcvrData.next() && m_Status == RptServer.RUNNING ) {            
            rowCount++;            
            vndId = rcvrData.getInt("VDCODE");            
            poNbr = rcvrData.getString("ponumber");
            termsCode = rcvrData.getInt("TERMSCODE");
            setCurAction(String.format("processing eis recipt data: %d row %d", vndId, rowCount));
            
            m_RcvrPoHdr.setInt(1, termsCode);
            m_RcvrPoHdr.setInt(2, termsCode);
            m_RcvrPoHdr.setString(3, poNbr);
            
            setCurAction(String.format("processing eis terms data: row %d vendor %d term %d ", vndId, rowCount, termsCode));
            rs = m_RcvrPoHdr.executeQuery();
            
            if ( rs.next() && m_Status == RptServer.RUNNING ) {
               rcptDate = rs.getDate("receipt_date");
               whsName = rs.getString("name");
               reviewPurch = rs.getString("review_purch");
               dueDate = rs.getDate("due_date");
               discDate = rs.getDate("discount_date");
               setCurAction(String.format("processed eis terms data: row %d term %d disc date %s", rowCount, termsCode, rs.getString("discount_date")));
            }
            else {
               rcptDate = null;
               dueDate = null;
               rcptDate = null;
               whsName = "";
               reviewPurch = "";
            }
                                    
            row = createRow(rowNum++, MAX_COLS);
            colNum = 0;
            
            if ( row != null ) {               
               row.getCell(colNum++).setCellValue(vndId);
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(rcvrData.getString("VDNAME")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(poNbr));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(rcvrData.getString("RCPNUMBER")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(formatDate(rcvrData.getString("DATE"))));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(rcvrData.getString("terms")));
               row.getCell(colNum++).setCellValue(rcvrData.getDouble("EXTENDED"));
               row.getCell(colNum++).setCellValue(rcvrData.getDouble("opnamt"));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(reviewPurch));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(whsName));
               
               row.getCell(colNum++).setCellValue(
                  rcptDate != null ? new XSSFRichTextString(sdf.format(rcptDate)) : new XSSFRichTextString("")
               );
               
               row.getCell(colNum++).setCellValue(
                  dueDate != null ? new XSSFRichTextString(sdf.format(dueDate)) : new XSSFRichTextString("")
               );
                              
               row.getCell(colNum++).setCellValue(
                  discDate != null ? new XSSFRichTextString(sdf.format(discDate)) : new XSSFRichTextString("")
               );
               
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(rcvrData.getString("group_code")));
            }
         }
         
         m_Wrkbk.write(outFile);
         rcvrData.close();

         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[AccPacOpenReceivers]", ex);
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
         rcptDate = null;
         dueDate = null;
      }

      return result;
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_RcvrData);
      closeStmt(m_RcvrPoHdr);      
   }
   
   /**
    * Sets the captions on the report.
    */
   private int createCaptions()
   {
      XSSFRow row = null;      
      int rowNum = 0;
      int colNum = 0;
      
      if ( m_Sheet == null )
         return 0;
      
      //
      // Create the row for the captions.
      row = m_Sheet.createRow(rowNum);
      
      if ( row != null ) {
         for ( int i = 0; i < MAX_COLS; i++ ) {
            row.createCell(i);            
         }
      }
      
      if ( row != null ) {
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Vnd ID"));
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Vnd Name"));
         m_Sheet.setColumnWidth(colNum++, 8000);      
         row.getCell(colNum).setCellValue(new XSSFRichTextString("PO Nbr"));
         m_Sheet.setColumnWidth(colNum++, 4000);      
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Rcpt Nbr"));
         m_Sheet.setColumnWidth(colNum++, 4000);      
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Date"));
         m_Sheet.setColumnWidth(colNum++, 3000);      
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Terms"));
         m_Sheet.setColumnWidth(colNum++, 8000);      
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Extended"));
         m_Sheet.setColumnWidth(colNum++, 3000);      
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Open Amt"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Purch Rev"));
         m_Sheet.setColumnWidth(colNum++, 2500);      
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Warehouse"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Rcpt Date"));         
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Due Date"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Disc Date"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Grp Code"));
      }
      else
         log.error("[AccPacOpenReceivers] accpac open recievers createCaptions - null row");
                  
      return ++rowNum;
   }
   
   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {      
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         m_SageConn = m_RptProc.getSageConn();
         m_EdbConn = m_RptProc.getEdbConn();
         
         if ( prepareStatements() )
            created = buildOutputFile();            
      }
      
      catch ( Exception ex ) {
         log.fatal("[AccPacOpenReceivers]", ex);
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
    * Formats the date from bcd format to normal short date format.  All dates in Accpac
    * are bcd and are in the form of yyyymmdd.
    * 
    * @param bcdDate The bcd date from accpac
    * @return A date formatted as yyyy/mm/dd
    */
   private String formatDate(String bcdDate)
   {
      StringBuffer date = new StringBuffer();
      
      if ( bcdDate != null ) {
         if ( bcdDate.length() == 8 ) {
            date.append(bcdDate.substring(0, 4));
            date.append("/");
            date.append(bcdDate.substring(4, 6));
            date.append("/");
            date.append(bcdDate.substring(6, 8));
         }
         else
            date.append(bcdDate);
      }
            
      return date.toString();
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;
      
      if ( m_SageConn != null ) {
         try {
            sql.append("select  ");
            sql.append("   VDCODE, VDNAME, rtrim(porcph1.PONUMBER) as ponumber, RCPNUMBER, "); 
            sql.append("   porcph1.DATE, ");
            sql.append("   rtrim(aprta.CODEDESC) as terms, EXTENDED, ");
            sql.append("   coalesce(porcph1.EXTENDED -  ");
            sql.append("      (select sum(poinvl.RQRECEIVED * porcpl.UNITCOST) as invqty "); 
            sql.append("      from EMEDAT.dbo.POINVL poinvl, EMEDAT.dbo.PORCPL porcpl ");
            sql.append("      where ");
            sql.append("         poinvl.RCPLSEQ = porcpl.RCPLSEQ and ");
            sql.append("         poinvl.RCPHSEQ = porcpl.RCPHSEQ and ");
            sql.append("         porcph1.RCPHSEQ = poinvl.RCPHSEQ ");
            sql.append("      group by poinvl.RCPHSEQ), EXTENDED "); 
            sql.append("   ) as opnamt, ");
            sql.append("   coalesce(len(rtrim(replace(rtrim(aprta.CODEDESC), ' +-.0123456789',' '))), 0) as est_due_date, ");
            sql.append("   coalesce(len(rtrim(replace(rtrim(aprta.CODEDESC), ' +-.0123456789',' '))), 0) as discount_date, ");            
            sql.append("   IDGRP as group_code, porcph1.TERMSCODE ");
            sql.append("from ");
            sql.append("   EMEDAT.dbo.PORCPH1 porcph1 ");            
            sql.append("join EMEDAT.dbo.APVEN apven on apven.VENDORID = porcph1.VDCODE ");
            sql.append("join EMEDAT.dbo.APRTA aprta on aprta.TERMSCODE = porcph1.TERMSCODE ");   
            sql.append("where ");
            sql.append("   ISCOMPLETE = 0 and ");
            sql.append("   porcph1.RCPHSEQ = ( ");
            sql.append("      select distinct PORCPL.RCPHSEQ "); 
            sql.append("      from EMEDAT.dbo.PORCPL ");
            sql.append("      where porcph1.RCPHSEQ = PORCPL.RCPHSEQ and PORCPL.COMPLETION = 1 "); 
            sql.append("   ) ");
            sql.append("order by VDCODE");                        
            m_RcvrData = m_SageConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select review_purch, warehouse.name, trunc(receipt_date) as receipt_date, ");
            sql.append("ejd.terms_procs.get_date(?, receipt_date) as due_date, ejd.terms_procs.discount_date(?, receipt_date) as discount_date ");
            sql.append("from rcvr_po_hdr ");
            sql.append("join warehouse on warehouse.fas_facility_id = rcvr_po_hdr.warehouse ");
            sql.append("where (po_nbr || '-' || to_char(emery_rcvr_nbr - 1000000000)) = ? ");
            m_RcvrPoHdr = m_EdbConn.prepareStatement(sql.toString());
                        
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[AccPacOpenReceivers]", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("[AccPacOpenReceievers] prepareStatements - null db connection");
      
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
            
      fileName.append("acp_openrcvr");      
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
            
      styleText = m_Wrkbk.createCellStyle();      
      styleText.setAlignment(HorizontalAlignment.LEFT);
      
      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);

      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short)8);
      
      m_CellStyles = new XSSFCellStyle[] {
         styleText,    // col 0 vnd id
         styleText,    // col 1 vnd name
         styleText,    // col 2 po#
         styleText,    // col 3 rcpt#
         styleText,    // col 4 date
         styleText,    // col 5 terms
         styleMoney,   // col 6 ext amount
         styleMoney,   // col 7 open amount
         styleText,    // col 8 ap rev
         styleText,    // col 9 warehouse
         styleText,    // col 10 receipt date
         styleText,    // col 11 est due date
         styleText,    // col 12 disc date
         styleText,    // col 13 group code
      };
      
      styleText = null;
      styleInt = null;
      styleMoney = null;
   }
}
