/**
 * File: SageEarlyInterest.java
 * Description: Sage300 early interest report for customer invoices
 *
 * @author Jeff Fisher
 * 
 * Create Data: 04/27/2016
 * Last Update: 04/27/2016
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

public class SageEarlyInterest extends Report 
{
   private String m_BegCust;
   private String m_EndCust;   
   private long m_BegDate;
   private long m_EndDate;
   
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
   
   private PreparedStatement m_PaymentData = null;
   
   /**
    * 
    */
   public SageEarlyInterest() 
   {
      super();
      
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      
      m_BegCust = "000001";
      m_EndCust = "999999";
      m_BegDate = 20160101;
      m_EndDate = 20161231;
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
      ResultSet pmntData = null;
      int curRow = 5;
                  
      try {
         //
         // Need to call this here after the parameters have been set.  The title contains 
         // the sequence numbers.
         
         setupWorkbook();
         
         m_PaymentData.setString(1, m_BegCust);
         m_PaymentData.setString(2, m_EndCust);
         m_PaymentData.setLong(3, m_BegDate);
         m_PaymentData.setLong(4, m_EndDate);         
         pmntData = m_PaymentData.executeQuery();
         
         while ( pmntData.next() && m_Status == RptServer.RUNNING ) {            
            addRow(curRow);
            addCell(0, pmntData.getString("IDCUST"), m_StyleTxtL);
            addCell(1, pmntData.getString("NAMECUST"), m_StyleTxtL);
            addCell(2, toDateStr(pmntData.getInt("doc_date")), m_StyleDate);
            addCell(3, toDateStr(pmntData.getInt("due_date")), m_StyleDate);
            addCell(4, pmntData.getString("inv_nbr"), m_StyleTxtL);
            addCell(5, pmntData.getString("pmt_nbr"), m_StyleTxtL);
            addCell(6, pmntData.getString("check_nbr"), m_StyleTxtL);
            addCell(7, toDateStr(pmntData.getInt("pmt_date")), m_StyleDate);
            addCell(8, pmntData.getString("date_var"), m_StyleInt);
            addCell(9, pmntData.getDouble("AMTDUEHC"), m_StyleDouble);
            addCell(10, pmntData.getDouble("AMTPAYMHC"), m_StyleDouble);            
            addCell(11, pmntData.getString("remit_var"), m_StyleInt);
                        
            curRow++;
         }

         m_Wrkbk.write(outFile);         
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[SageEarlyInterest]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, m_PaymentData, pmntData);
      }
      
      return result;
   }
   
   /**
    * Closes prepared statements and cleans up member variables
    */
   protected void cleanup()
   {
      m_PaymentData = null;

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
         log.fatal("[SageEarlyInterest]", ex);
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
      m_StyleTxtC.setWrapText(true);
      
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
            
      if ( m_SageConn != null ) {
         try {
            sql.append("select ");
            sql.append("   arcus.IDCUST, arcus.NAMECUST, arobs.DATEINVC as doc_date, arobs.DATEDUE as due_date, arobs.IDINVC as inv_nbr, ");
            sql.append("   arobl.IDINVC as pmt_nbr, arobl.IDRMIT as check_nbr, arobl.DATEDUE as pmt_date, (arobs.DATEDUE - arobl.DATEDUE) as date_var, ");
            sql.append("   arobs.AMTDUEHC, arobp.AMTPAYMHC, (arobs.AMTDUEHC - abs(arobp.AMTPAYMHC)) as remit_var ");
            sql.append("from EMEDAT.dbo.ARCUS arcus ");
            sql.append("join EMEDAT.dbo.AROBL arobl on arobl.IDCUST = arcus.IDCUST ");
            sql.append("join EMEDAT.dbo.AROBP arobp on arobp.IDCUST = arobl.IDCUST and arobp.IDMEMOXREF = arobl.IDINVC ");
            sql.append("join EMEDAT.dbo.AROBS arobs on arobs.IDCUST = arobp.IDCUST and arobs.IDINVC = arobp.IDINVC and arobs.CNTPAYM = arobp.CNTPAYMNBR ");
            sql.append("where ");
            sql.append("   arcus.IDCUST between ? and ? and abs(arobp.AMTPAYMHC) > 0.00 and ");
            sql.append("   arobl.TRXTYPETXT = 11 and arobl.DATEDUE between ? and ? and ");
            sql.append("   (arobs.DATEDUE - arobl.DATEDUE) >= 7 and arobs.AMTDUEHC > 0.00 ");
            sql.append(" order by ");
            sql.append("   arcus.IDCUST, arobp.IDINVC, arobp.CNTPAYMNBR, arobs.DATEDUE ");
                     
            m_PaymentData = m_SageConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[SageEarlyInterest]", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("[SageEarlyInterest] prepareStatements - null sqlserver connection");
      
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
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                  
         if ( param.name.equals("begdate") )
            m_BegDate = Long.parseLong(param.value);
         
         if ( param.name.equals("enddate") )
            m_EndDate = Long.parseLong(param.value);
                  
         if ( param.name.equals("begcust") )
            m_BegCust = param.value;
         
         if ( param.name.equals("endcust") )
            m_EndCust = param.value;                  
      }
      
      m_FileNames.add(String.format("EarlyInterest [%d-%d %s-%s].xlsx", m_BegDate, m_EndDate, m_BegCust, m_EndCust));
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
      addCell(col, "Sage A/R Early Payment Customer Activity", m_StyleHdrLeft);
      addRow(1);
      addCell(col,  String.format("From Customer Number: %s to %s", m_BegCust, m_EndCust), m_StyleHdrLeft);
      addRow(2);
      addCell(col,  String.format("Show Early Pay Transactions for: %s to %s", toDateStr(m_BegDate), toDateStr(m_EndDate)), m_StyleHdrLeft);

      //
      // Add the captions
      addRow(4);
      m_Row.setHeightInPoints((2 * m_Sheet.getDefaultRowHeightInPoints()));
      m_Sheet.setColumnWidth(col, (9 * m_CharWidth));
      addCell(col, "Cust #", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (50 * m_CharWidth));
      addCell(col, "Customer Name", m_StyleTxtC);
            
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Doc\nDate", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Due\nDate", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (8 * m_CharWidth));
      addCell(col, "Invoice\nNbr", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Payment\nNbr", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Check\nNbr", m_StyleTxtC);
            
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Remit\nDate", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Date\nVariance", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (12 * m_CharWidth));
      addCell(col, "Amount\nDue", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Amount\nRemit", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (10 * m_CharWidth));
      addCell(col, "Remit\nVariance", m_StyleTxtC);      
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
