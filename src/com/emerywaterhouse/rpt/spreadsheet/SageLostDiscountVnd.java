/**
 * File: SageLostDiscountVnd.java
 * Description: Sage300 Lost discount by vendor report
 *
 * @author Jeff Fisher
 * 
 * Create Data: 04/15/2016
 * Last Update: 04/15/2016
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
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

/**
 * 
 */
public class SageLostDiscountVnd extends Report 
{
   private int m_Cutoff;
   private ArrayList<String> m_VndTypes;
   private XSSFWorkbook m_Wrkbk;
   private Sheet m_Sheet;
   private Row m_Row;
   private Font m_FontNorm;
   private Font m_FontBold;
   private XSSFCellStyle m_StyleHdrLeft = null;
   private XSSFCellStyle m_StyleTxtC = null;      // Text centered
   private XSSFCellStyle m_StyleTxtL = null;      // Text left justified
   private XSSFCellStyle m_StyleInt = null;       // Style with 0 decimals
   private XSSFCellStyle m_StyleDouble = null;    // numeric #,##0.00
   private XSSFCellStyle m_StyleDate = null;      // mm/dd/yyyy
   
   private PreparedStatement m_DiscData = null;
   
   public SageLostDiscountVnd()
   {
      super();
      
      m_Cutoff = 0;
      m_VndTypes = new ArrayList<String>();
      
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
   }
      
   /**
    * adds a numeric type cell to current row at col p_Col in current sheet
    *
    * @param col 0-based column number of spreadsheet cell
    * @param value numeric value to be stored in cell
    * @param style Excel style to be used to display cell
    */
   private void addCell(int col, double value, CellStyle style)
   {
      Cell cell = m_Row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellStyle(style);
      cell.setCellValue(value);
      
      cell = null;
   }
   
   /**
    * adds a numeric type cell to current row at col p_Col in current sheet
    *
    * @param col 0-based column number of spreadsheet cell
    * @param value integer value to be stored in cell
    * @param style Excel style to be used to display cell
    */
   private void addCell(int col, int value, CellStyle style)
   {
      Cell cell = m_Row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellStyle(style);
      cell.setCellValue(value);
      
      cell = null;
   }
   
   /**
    * adds a text type cell to current row at col p_Col in current sheet
    *
    * @param col 0-based column number of spreadsheet cell
    * @param value text value to be stored in cell
    * @param style Excel style to be used to display cell
    */
   private void addCell(int col, String value, CellStyle style)
   {
      Cell cell = m_Row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(value));
      cell.setCellStyle(style);
      cell = null;
   }
   
   /**
    * adds row to the current sheet
    *
    * @param row  0-based row number of row to be added
    */
   private void addRow(int row) 
   {
      m_Row = m_Sheet.createRow(row);
   }
   
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      boolean result = false;
      ResultSet discData = null;
      int curRow = 4;
                  
      try {
         //
         // Need to call this here after the parameters have been set.  The title contains 
         // the sequence numbers.
         setupWorkbook();
         
         m_DiscData.setInt(1, m_Cutoff);
         m_DiscData.setInt(2, m_Cutoff);         
         discData = m_DiscData.executeQuery();
         
         while ( discData.next() && m_Status == RptServer.RUNNING ) {            
            addRow(curRow);
            addCell(0, discData.getInt("IDVEND"), m_StyleTxtL);
            addCell(1, discData.getString("VENDNAME"), m_StyleTxtL);
            addCell(2, discData.getString("IDVENDGRP"), m_StyleTxtL);            
            addCell(3, discData.getString("IDINVC"), m_StyleTxtL);
            addCell(4, discData.getString("doc_type"), m_StyleTxtL);
            addCell(5, toDateStr(discData.getInt("DATEINVC")), m_StyleDate);
            addCell(6, toDateStr(discData.getInt("DATEDISC")), m_StyleDate);
            addCell(7, toDateStr(discData.getInt("DATEINVCDU")), m_StyleDate);
            addCell(8, discData.getDouble("AMTINVCHC"), m_StyleDouble);
            addCell(9, discData.getDouble("AMTDISCHC"), m_StyleDouble);
            addCell(10, discData.getDouble("InvAmtDue"), m_StyleDouble);
            addCell(11, discData.getString("lost_disc"), m_StyleTxtL);
                        
            curRow++;
         }

         m_Wrkbk.write(outFile);         
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[SageLostDiscountVnd]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, m_DiscData, discData);
      }
      
      return result;
   }
   
   /**
    * Closes prepared statements and cleans up member variables
    */
   protected void cleanup()
   {
      m_DiscData = null;

      m_Sheet = null;
      m_Wrkbk = null;
      
      m_FontNorm = null;
      m_FontBold = null;
      
      m_StyleHdrLeft = null;
      m_StyleTxtC = null;
      m_StyleTxtL = null;      
      m_StyleInt = null;
      m_StyleDouble = null;      
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
         m_SageConn = m_RptProc.getSageConn();
         prepareStatements();
         created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[SageLostDiscountVnd]", ex);
      }

      finally {
        cleanup();

        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * defines Excel fonts and styles
    */
   private void defineStyles()
   {
      short fontHeight = 11;
      XSSFDataFormat format = m_Wrkbk.createDataFormat();
      
      //
      // Normal Font
      m_FontNorm = m_Wrkbk.createFont();
      m_FontNorm.setFontName("Arial");
      m_FontNorm.setFontHeightInPoints(fontHeight);
      
      //
      // defines bold font
      m_FontBold = m_Wrkbk.createFont();
      m_FontBold.setFontName("Arial");
      m_FontBold.setFontHeightInPoints(fontHeight);
      m_FontBold.setBold(true);
      
      //
      // defines style column header, left-justified
      m_StyleHdrLeft = m_Wrkbk.createCellStyle();
      m_StyleHdrLeft.setFont(m_FontBold);
      m_StyleHdrLeft.setAlignment(HorizontalAlignment.LEFT);
      m_StyleHdrLeft.setVerticalAlignment(VerticalAlignment.TOP);
      
      m_StyleTxtL = m_Wrkbk.createCellStyle();
      m_StyleTxtL.setAlignment(HorizontalAlignment.LEFT);
      
      m_StyleTxtC = m_Wrkbk.createCellStyle();
      m_StyleTxtC.setAlignment(HorizontalAlignment.CENTER);
      
      m_StyleInt = m_Wrkbk.createCellStyle();
      m_StyleInt.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleInt.setDataFormat(format.getFormat("0"));
      
      m_StyleDouble = m_Wrkbk.createCellStyle();
      m_StyleDouble.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDouble.setDataFormat(format.getFormat("#,##0.00"));
      
      m_StyleDate = m_Wrkbk.createCellStyle();
      m_StyleDate.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDate.setDataFormat(format.getFormat("mm/dd/yyyy"));
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(256);      
      boolean isPrepared = false;
      Iterator<String> iter = null;
      int pcount = 1;
      
      if ( m_SageConn != null ) {
         try {
            sql.append("select ");
            sql.append("   apobl.IDVEND, apven.VENDNAME, apobl.IDVENDGRP, apobl.IDINVC, ");
            sql.append("   (CASE IDTRXTYPE WHEN '12' THEN 'IN' WHEN '13' THEN 'IN' WHEN '22' THEN 'CR' WHEN '32' THEN 'CR' END) as doc_type, ");
            sql.append("   apobl.DATEINVC, apobl.DATEDISC, apobl.DATEINVCDU, ");
            sql.append("   apobl.AMTINVCHC, apobl.AMTDISCHC, (AMTINVCHC - AMTDISCHC) as InvAmtDue, ");
            sql.append("   (CASE WHEN DATEDISC < ? then 'L' else '' END) as lost_disc ");
            sql.append("from EMEDAT.dbo.APOBL apobl ");
            sql.append("join EMEDAT.dbo.APVEN apven on apven.VENDORID = apobl.IDVEND and apven.AMTBALDUEH > 0.00 ");
            sql.append("where ");
            sql.append("   apobl.AMTDISCHC > 0 and DATEDISC <= ? and IDTRXTYPE <> 51 and apobl.SWPAID = 0 ");
            
            //
            // We may or may not have some vendor filters and JDBC seems notoriously bad about formatting an "in" with
            // parameters.  We'll just build the sql on the fly.
            if ( m_VndTypes.size() > 0 ) {
               sql.append(" and apobl.IDVENDGRP in (");      
               iter = m_VndTypes.iterator();
               
               while ( iter.hasNext() ) {
                  if ( pcount < m_VndTypes.size() )
                     sql.append(String.format("'%s', ", iter.next()));
                  else
                     sql.append(String.format("'%s')", iter.next()));
                  
                  pcount++;
               }
            }
            
            sql.append(" order by ");
            sql.append("   apobl.IDVEND, apobl.IDINVC");
                     
            m_DiscData = m_SageConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[SageLostDiscountVnd]", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("[SageLostDiscountVnd] prepareStatements - null sqlserver connection");
      
      return isPrepared;
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {      
      int pcount = params.size();
      Param param = null;
      StringBuffer tmp = new StringBuffer();
      Iterator<String> iter = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                  
         if ( param.name.equals("cutoff") )
            m_Cutoff = Integer.parseInt(param.value);
         
         //
         // 'EM', 'DA', 'EXPENS', 'WIRE', 'MG'
         if ( param.name.equals("vndtype") )
            m_VndTypes.add(param.value);
      }
      
      //
      // Build the file name.  Formatted as the cutoff date plus any vendor type filters.
      // LostDisc 20160414 [EM DA].xlsx or LostDisc 20160414 [ALL].xlsx
      tmp.append(" [");
      
      if ( !m_VndTypes.isEmpty() ) {         
         iter = m_VndTypes.iterator();
         pcount = 1;
                  
         while ( iter.hasNext() ) { 
            tmp.append(iter.next());
            
            if ( pcount < m_VndTypes.size() )
               tmp.append(" ");
            
            pcount++;
         }
         
         iter = null;
      }
      else
         tmp.append("ALL");
      
      tmp.append("]");
      
      m_FileNames.add(String.format("LostVndDisc %d%s.xlsx", m_Cutoff, tmp));      
   }
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {
      int col = 0;
      int m_CharWidth = 295;
      
      defineStyles();
      
      //
      // creates Excel title
      addRow(0);
      addCell(col, "Sage Lost Discount By Vendor Report", m_StyleHdrLeft);
      addRow(1);
      addCell(col,  String.format("Cutoff period: %s", toDateStr(m_Cutoff)), m_StyleHdrLeft);
                
      //
      // Add the captions
      addRow(3);
      m_Sheet.setColumnWidth(col, (9 * m_CharWidth));
      addCell(col, "Vendor #", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (50 * m_CharWidth));
      addCell(col, "Vendor Name", m_StyleTxtC);
            
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Vendor Grp", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Invoice #", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Doc Type", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Invoice Date", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Disc Date", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Due Date", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Invoice Amt", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Disc Amt", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Amt Due", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Lost Disc", m_StyleTxtC);      
   }
   
   /**
    * Converts the Accpac numeric date representation to an actual date value.
    * @param accDate the numeric accpac date value.  yyyymmdd as a number.  ex 20160131
    * @return
    */
   private String toDateStr(long accDate)
   {
      int year = 0;
      int month = 0;
      int day = 0;      
            
      day = (int)(accDate % 100);
      accDate = (long)(accDate / 100);
      month = (int)(accDate % 100);
      year = (int)(accDate / 100);
            
      
      return String.format("%02d/%02d/%d", month, day, year);
   }
   
}