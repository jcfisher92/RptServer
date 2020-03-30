/**
 * File: SageCheckRegister.java
 * Description: Sage300 check listing for specific periods.
 *
 * @author Jeff Fisher
 * 
 * Create Data: 03/15/2016
 * Last Update: 03/15/2016
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

/**
 * 
 */
public class SageCheckRegister extends Report 
{
   private int m_StartSeq;
   private int m_EndSeq;
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
   
   private PreparedStatement m_CheckData = null;
   
   public SageCheckRegister()
   {
      super();
      
      m_StartSeq = 0;
      m_EndSeq = 0;
      
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
      ResultSet checkData = null;
      int curRow = 3;
      int tmpDate = 0;
            
      try {
         //
         // Need to call this here after the parameters have been set.  The title contains 
         // the sequence numbers.
         setupWorkbook();
         
         m_CheckData.setInt(1, m_StartSeq);
         m_CheckData.setInt(2, m_EndSeq);
         checkData = m_CheckData.executeQuery();
         
         while ( checkData.next() && m_Status == RptServer.RUNNING ) {            
            addRow(curRow);
            addCell(0, checkData.getInt("IDVEND"), m_StyleTxtL);
            addCell(1, checkData.getString("VENDNAME"), m_StyleTxtL);
            addCell(2, checkData.getString("docnum"), m_StyleTxtL);
            addCell(3, toDateStr(checkData.getInt("doc_date")), m_StyleDate);
            
            tmpDate = checkData.getInt("disc_date");
            if ( tmpDate > 0 )
               addCell(4, toDateStr(tmpDate), m_StyleDate);
            
            addCell(5, toDateStr(checkData.getInt("due_date")), m_StyleDate);
            addCell(6, checkData.getDouble("payable_amt"), m_StyleDouble);
            addCell(7, checkData.getDouble("discount"), m_StyleDouble);
            addCell(8, checkData.getDouble("adjustment"), m_StyleDouble);
            addCell(9, checkData.getDouble("net_payment"), m_StyleDouble);
            addCell(10, toDateStr(checkData.getInt("check_date")), m_StyleDate);
            addCell(11, checkData.getInt("check_num"), m_StyleInt);
            addCell(12, checkData.getDouble("AMTPAYM"), m_StyleDouble);
            
            curRow++;
         }

         m_Wrkbk.write(outFile);         
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[SageCheckRegister]", ex);
      }

      finally {
         DbUtils.closeDbConn(null, m_CheckData, checkData);
      }
      
      return result;
   }
   
   /**
    * Closes prepared statements and cleans up member variables
    */
   protected void cleanup()
   {
      m_CheckData = null;

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
         log.fatal("[SageCheckRegister]", ex);
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
      
      if ( m_SageConn != null ) {
         try {
            sql.append("select ");
            sql.append("   appjh.IDVEND, apven.VENDNAME,  ");
            sql.append("   rtrim(appjd.IDINVC) + convert(varchar(6), (CASE appjd.CNTPAYM WHEN 0 THEN '' ELSE ('-' + convert(varchar(5), appjd.CNTPAYM)) END)) as docnum, ");
            sql.append("   appjd.DATEINVC as doc_date, appjd.DATEDISC as disc_date, appjd.DATEDUE as due_date, ");
            sql.append("   (appjd.AMTDSCHCUR + appjd.AMTADJHCUR + appjd.AMTEXTNHDC) as payable_amt, ");
            sql.append("   appjd.AMTDSCHCUR as discount, appjd.AMTADJHCUR as adjustment, appjd.AMTEXTNDHC as net_payment, ");
            sql.append("   appjh.DATEINVC as check_date, appjh.IDRMIT as check_num, appym.AMTPAYM, ");
            sql.append("   appym.IDRMIT, appjh.POSTSEQNCE, appjd.CNTPAYM, appjh.TYPEBTCH ");
            sql.append("from ");
            sql.append("   EMEDAT.dbo.APPJH ");
            sql.append("join EMEDAT.dbo.APPJD appjd on appjh.TYPEBTCH = appjd.TYPEBTCH and appjh.POSTSEQNCE = appjd.POSTSEQNCE and ");
            sql.append("    appjh.CNTBTCH = appjd.CNTBTCH and appjh.CNTITEM = appjd.CNTITEM ");
            sql.append("left outer join EMEDAT.dbo.APVEN apven on apven.VENDORID = appjh.IDVEND ");
            sql.append("join EMEDAT.dbo.APPYM appym on appym.IDBANK = appjh.IDBANK and appym.IDVEND = appjh.IDVEND and ");
            sql.append("    appym.IDRMIT = appjh.IDRMIT and appym.LONGSERIAL = appjh.LONGSERIAL ");
            sql.append("where ");
            sql.append("   appjh.TYPEBTCH = 'PY' and appjh.PAYMTYPE = 2 and appjd.ACCTTYPE = 1 and ");
            sql.append("   (appjd.IDTRANS = 51 or appjd.IDTRANS = 57) and appjh.POSTSEQNCE between ? and ? ");
            sql.append(" order by appjh.IDVEND");
                     
            m_CheckData = m_SageConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[SageCheckRegister]", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("[SageCheckRegister] prepareStatements - null sqlserver connection");
      
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
                  
         if ( param.name.equals("startseq") )
            m_StartSeq = Integer.parseInt(param.value);
         
         if ( param.name.equals("endseq") )
            m_EndSeq = Integer.parseInt(param.value);
      }
      
      m_FileNames.add(String.format("checkreg_%d-%d.xlsx", m_StartSeq, m_EndSeq));      
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
      addCell(col, String.format("Sage Check Register Report [%d] to [%d]", m_StartSeq, m_EndSeq), m_StyleHdrLeft);
      
      //
      // Add the captions
      addRow(2);
      m_Sheet.setColumnWidth(col, (15 * m_CharWidth));
      addCell(col, "Vendor #", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (50 * m_CharWidth));
      addCell(col, "Vendor Name", m_StyleTxtC);
            
      m_Sheet.setColumnWidth(++col, (25 * m_CharWidth));
      addCell(col, "Doc#", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "Doc Date", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "Disc Date", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "Due Date", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "Payable", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "Discount", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "Adjustment", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "Net Payment", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "Check Date", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "Check #", m_StyleTxtC);
      
      m_Sheet.setColumnWidth(++col, (15 * m_CharWidth));
      addCell(col, "Check Amount", m_StyleTxtC);
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
