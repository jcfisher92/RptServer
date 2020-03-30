/**
 * File: PPVAccountAnalysis
 * Description: Analysis report of accpac ppv.
 *    Note - No header comments.  Not sure who wrote this or when.  02/18/2009 jcf
 * 
 * Author: ?
 * 
 * Create Date: ?
 * Last Update: $Id: PPVAccountAnalysis.java,v 1.3 2009/02/18 17:17:50 jfisher Exp $
 * 
 * History:
 *    $Log: PPVAccountAnalysis.java,v $
 *    Revision 1.3  2009/02/18 17:17:50  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
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


public class PPVAccountAnalysis extends Report
{
   private static final int MAX_COLS = 12;
   
   private String m_BegAcct;
   private String m_EndAcct;
   private String m_Year;
   private String m_Period;
   
   private PreparedStatement m_AcctData;
   
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;
   
   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;
   
   /**
    * 
    */
   public PPVAccountAnalysis()
   {
      super();
      
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
      ResultSet acctData = null;
      boolean result = false;
      int vndId = 0;
      String acct = "";

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         rowNum = createCaptions();
         
         //
         // See prepareStatements for the reason not all vars are bind vars
         m_AcctData.setString(1, m_Year);
         m_AcctData.setString(2, m_Period);
         
         acctData = m_AcctData.executeQuery();

         while ( acctData.next() && m_Status == RptServer.RUNNING ) {
            acct = acctData.getString("IDACCT");
            vndId = acctData.getInt("VENDORID");
            setCurAction("processing acct: " + acct + "vend: " + vndId);
            
            row = createRow(rowNum++, MAX_COLS);
            colNum = 0;
            
            if ( row != null ) {
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(acct));
               row.getCell(colNum++).setCellValue(vndId);
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(acctData.getString("VENDNAME")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(acctData.getString("FISCYR")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(acctData.getString("FISCPER")));
               row.getCell(colNum++).setCellValue(
                     new XSSFRichTextString(formatDate(acctData.getString("DATEINVC")))
               );
               row.getCell(colNum++).setCellValue(acctData.getDouble("AMTEXTNDHC"));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(acctData.getString("IDINVC")));
               
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(
                  getSrcCode(acctData.getInt("TRANSTYPE"), acctData.getString("SRCETYPE"))
               ));
               
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(acctData.getString("PONBR")));
               row.getCell(colNum++).setCellValue(acctData.getInt("POSTSEQNCE"));
               row.getCell(colNum++).setCellValue(acctData.getInt("CNTBTCH"));
            }
         }
         
         m_Wrkbk.write(outFile);
         acctData.close();

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
      closeStmt(m_AcctData);
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
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Acct"));
         m_Sheet.setColumnWidth(colNum++, 4000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Vnd ID"));
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Vnd Name"));
         m_Sheet.setColumnWidth(colNum++, 8000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Year"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Period"));
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Date Invc"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Amt Extnd HC"));
         m_Sheet.setColumnWidth(colNum++, 4000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("ID Invc"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Src Code"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("PO#"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Post Seq"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Batch No"));
      }
            
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
    * Formats the date from bcd format to normal int date format.  All dates in Accpac
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
    * Determines a code based on the transaction type.  Not sure what this means or why it's done.  This
    * was pulled from the crystal report that had this same data.  This was the function that was used.
    * 
    * @param tranType The transaction type number
    * @param srcType The source type.
    * 
    * @return The code.
    */
   private String getSrcCode(int tranType, String srcType)
   {
      String code = "";
      
      switch( tranType ) {
         case 1:
            code = "IN";
         break;
         
         case 2:
            code = "DB";
         break;
         
         case 3:
            code = "CR";
         break;
         
         case 4:
            code = "IT";
         break;
         
         case 5:
            code = "UC";
         break;
         
         case 6:
            code = "DT";
         break;
         
         case 7:
            code = "DF";
         break;
         
         case 8:
            code = "CT";
         break;
         
         case 9:
            code = "CF";
         break;
         
         case 10:
            code = "PI";
         break;
         
         case 11:
            code = "PY";
         break;
         
         case 12:
            code = "ED";
         break;
         
         case 14:
            code = "AD";
         break;
         
         case 16:
            code = "GL";
         break;
         
         case 18:
            if ( srcType != null )
               code = srcType;             
         break;
      }
      
      return code; 
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
            sql.append("select apven.VENDORID, apven.VENDNAME, appjd.IDACCT, appjd.FISCYR, appjd.FISCPER, ");
            sql.append("appjh.DATEINVC, appjd.AMTEXTNDHC, appjh.IDINVC, appjd.CNTBTCH, ");
            sql.append("appjd.TYPEBTCH, appjd.TRANSTYPE, appjd.SRCETYPE, appjd.POSTSEQNCE, apibh.PONBR ");
            sql.append("from EMEDAT.dbo.APPJD appjd ");
            sql.append("join EMEDAT.dbo.APPJH appjh on appjh.TYPEBTCH = appjd.TYPEBTCH and appjh.POSTSEQNCE = appjd.POSTSEQNCE and ");
            sql.append("                               appjh.CNTBTCH = appjd.CNTBTCH and appjh.CNTITEM = appjd.CNTITEM ");
            sql.append("join EMEDAT.dbo.APIBH apibh on apibh.CNTBTCH = appjh.CNTBTCH and apibh.CNTITEM = appjh.CNTITEM ");
            sql.append("join EMEDAT.dbo.APVEN apven on apven.VENDORID = apibh.IDVEND ");
            
            //
            // Note - we can't use bind variable here because Oracle is exclusive of the last operator's
            // value.  This means we are missing the last set of data that normally is returned.  This is 
            // different when the query has the literal value.
            sql.append(String.format("where IDACCT between '%s' and '%s' and ", m_BegAcct, m_EndAcct));
            
            sql.append("appjd.FISCYR = ? and ");
            sql.append("appjd.FISCPER = ? ");
            sql.append("order by appjd.IDACCT, apven.VENDNAME, appjh.DATEINVC");
            
            m_AcctData = m_SageConn.prepareStatement(sql.toString());
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
         log.error("PPVAcctAnalysis.prepareStatements - null sqlserver connection");
      
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
         
         if ( param.name.equals("begacct") )
            m_BegAcct = param.value;
         
         if ( param.name.equals("endacct") )
            m_EndAcct = param.value;
         
         if ( param.name.equals("year") )
            m_Year = param.value;
         
         if ( param.name.equals("period") )
            m_Period = param.value;
      }
            
      fileName.append("ppv");      
      fileName.append(m_BegAcct);
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
         styleText,    // col 0 acct id
         styleText,    // col 1 vnd id
         styleText,    // col 2 vnd name
         styleText,    // col 3 year
         styleText,    // col 4 period
         styleText,    // col 5 inv date
         styleMoney,   // col 6 ext amount
         styleText,    // col 7 invoice id
         styleText,    // col 8 source
         styleText,    // col 9 po
         styleInt,     // col 10 post seq
         styleInt,     // col 11 batch num         
      };
      
      styleText = null;
      styleInt = null;
      styleMoney = null;
   }
}
