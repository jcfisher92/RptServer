/**
 * File: ClosedReceivers.java
 * Description: Create the closed receiver Excel spreadsheet report
 *
 * @author Jeff Fisher
 *
 * Create Date: 08/10/2011
 * Last Update: $Id: ClosedReceivers.java,v 1.1 2011/08/12 02:08:35 jfisher Exp $
 * $Revision: 1.1 $
 * 
 * History:
 *    $Log: ClosedReceivers.java,v $
 *    Revision 1.1  2011/08/12 02:08:35  jfisher
 *    initial add
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Calendar;

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


public class ClosedReceivers extends Report
{
   private static final short maxCols = 5;

   private String m_RptDate;
   private PreparedStatement m_RcvrData;

   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;

   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;

   public ClosedReceivers()
   {
      super();
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      setupWorkbook();

      //
      // set the default date parameter, format it to the previous day to avoid
      // any errors.
      Calendar cal = Calendar.getInstance();
      cal.roll(Calendar.DAY_OF_YEAR, false);
      m_RptDate = String.format("%td-%tb-%tY", cal, cal, cal);
   }

   /**
    * Cleanup any allocated resources.
    */
   @Override
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
    * Builds the file name for the report.  Done outside of the setParams method because
    * it's possible no params come in.
    */
   private void buildFileName()
   {
      String tmp = Long.toString(System.currentTimeMillis());
      tmp = tmp.substring(tmp.length()-5, tmp.length());
      m_FileNames.add(String.format("closed_rec_%s_%s.xlsx", m_RptDate, tmp));
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
      ResultSet rcvrData = null;
      boolean result = false;
      int vndId = 0;

      buildFileName();
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         rowNum = createCaptions();

         m_RcvrData.setString(1, m_RptDate);
         rcvrData = m_RcvrData.executeQuery();

         while ( rcvrData.next() && m_Status == RptServer.RUNNING ) {
            vndId = rcvrData.getInt("vendor_id");
            setCurAction("processing vend: " + vndId);

            row = createRow(rowNum++, maxCols);
            colNum = 0;

            if ( row != null ) {
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(rcvrData.getString("facility")));
               row.getCell(colNum++).setCellValue(vndId);
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(rcvrData.getString("name")));
               row.getCell(colNum++).setCellValue(rcvrData.getInt("lines"));
               row.getCell(colNum++).setCellValue(rcvrData.getInt("units"));
            }
         }

         m_Wrkbk.write(outFile);
         DbUtils.closeDbConn(null, null, rcvrData);

         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
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
      closeStmt(m_RcvrData);
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
      // Create the row for the header.
      row = m_Sheet.createRow(rowNum++);
      row.createCell(0);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Closed Receiver Report: " + m_RptDate));

      //
      // Create the row for the captions.
      row = m_Sheet.createRow(rowNum);

      if ( row != null ) {
         for ( int i = 0; i < maxCols; i++ ) {
            row.createCell(i);
         }
      }

      row.getCell(colNum).setCellValue(new XSSFRichTextString("Facility"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Vendor#"));
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Vendor Name"));
      m_Sheet.setColumnWidth(colNum++, 15000);
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Lines"));
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Units"));

      return ++rowNum;
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
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;

      if ( m_EdbConn != null ) {
         try {
            sql.append("select warehouse.name facility, vendor.vendor_id, vendor.name, ");
            sql.append("count(*) lines, sum(qty_received) units ");
            sql.append("from rcvr_po_hdr ");
            sql.append("join rcvr_dtl on rcvr_dtl.rcvr_po_hdr_id = rcvr_po_hdr.rcvr_po_hdr_id ");
            sql.append("join vendor on vendor.vendor_id = rcvr_po_hdr.vendor_id ");
            sql.append("join warehouse on warehouse.fas_facility_id = rcvr_po_hdr.warehouse ");
            sql.append("where date_closed = ? ");
            sql.append("group by warehouse.name, vendor.vendor_id, vendor.name ");
            sql.append("order by warehouse.name, vendor.name");

            m_RcvrData = m_EdbConn.prepareStatement(sql.toString());
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
         log.error("ClosedReceivers.prepareStatements - null oracle connection");

      return isPrepared;
   }

   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      for (Param p : params) {
         if ( p.name.equalsIgnoreCase("rptdate"))
            m_RptDate = p.value.trim();
      }
   }

   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {
      XSSFCellStyle styleText;      // Text left justified
      XSSFCellStyle styleInt;       // Style with 0 decimals

      styleText = m_Wrkbk.createCellStyle();
      styleText.setAlignment(HorizontalAlignment.LEFT);

      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);

      m_CellStyles = new XSSFCellStyle[] {
            styleText,    // col 0 facility
            styleText,    // col 1 vnd id
            styleText,    // col 2 vnd name
            styleInt,     // col 3 lines
            styleInt      // col 4 units
      };

      styleText = null;
      styleInt = null;
   }
}
