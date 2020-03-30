/**
 * Title:         PositivePayWells.java
 * Description:   Positive Pay report designed to Wells Fargo's specifications
 * Company:       Emery-Waterhouse
 * @author        Stephen Martel
 */

package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class PositivePayWells extends Report {

   // Workbook and style stuff
   private XSSFWorkbook m_WrkBk;
   private XSSFSheet m_Sheet;
   private XSSFRow m_Row = null;

   private XSSFFont m_FontData;
   private XSSFFont m_FontBold;
   private XSSFCellStyle m_StyleText;
   private XSSFCellStyle m_StyleBold;

   private short m_RowNum = 0;
   private final int NUM_ALL_FIELDS = 7;
   
   // Finals supplied to us by Wells Fargo
   private final String ABA_NUM = "241253823";
   private final String ACCT_NUM = "9612000811";
   private final String NEW_CODE = "320";
   
   // Query variables:   
   private PreparedStatement m_PositivePayData;
   private String m_batch_start;
   private String m_batch_end;

   // Finals for retrieving fields from the Positive Pay data ResultSet
   //private final String BATCH_NUM = "batch_num";
   private final String CHK_NUM = "chk_num";
   private final String CHK_DATE = "chk_date";
   private final String CHK_AMT = "chk_amt";
   private final String PAYEE = "payee";
   
   private boolean m_displayHeader = false; // by default, we do not want a header line. we only bring it in for testing.
   
   /**
    * default constructor
    */
   public PositivePayWells()
   {
      super();
     
      m_MaxRunTime = RptServer.HOUR * 2;
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    */
   @Override
   public void finalize() throws Throwable
   {      
      m_PositivePayData = null;
      
      super.finalize();
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_PositivePayData);
   }
   
   /**
    * Creates the report file.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
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
         log.fatal("[PositivePayWells#createReport] ", ex);
      }

      finally {
        closeStatements();

        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }

      return created;
   }
   
   
   private boolean prepareStatements() throws SQLException
   {
      StringBuffer sql = new StringBuffer();
      
      if ( m_SageConn == null )
         return false;
      
      sql.append("select ");
      sql.append("   CNTBTCH as batch_num, ");      
      sql.append("   substring(IDRMIT, PATINDEX('%[^0]%', IDRMIT+'.'), LEN(IDRMIT)) as chk_num, ");
      sql.append("   convert(varchar(10), convert(DATETIME, convert(varchar(10),DATERMIT)), 110) as chk_date, ");
      sql.append("   right(replicate('0', 11) + (convert(varchar(11), cast(round(AMTRMIT * 100, 0) as int))), 11) as chk_amt, ");
      sql.append("   rtrim(NAMERMIT) as payee, ");
      sql.append("   (select count(distinct CNTBTCH) from EMEDAT.dbo.APTCR where CNTBTCH >= ? and CNTBTCH <= ?) as num_batches ");
      sql.append("from ");
      sql.append("   EMEDAT.dbo.APTCR ");
      sql.append("where ");
      sql.append("   CNTBTCH >= ? and ");
      sql.append("   CNTBTCH <=  ? ");
      sql.append("order by CNTBTCH, chk_num ");
      
      m_PositivePayData = m_SageConn.prepareStatement(sql.toString());
      
      return true;
   }
   
   /**
    * Executes the queries and builds the output file
    * 
    * @return true if the report was successfully built
    * @throws FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      Row row = null;
      int colNum = 0;
      FileOutputStream outFile = null;
      ResultSet positivePayData = null;
      boolean result = false;
      DecimalFormat df = new DecimalFormat("#.00"); // ensuring we have 2 digits in the cents' place
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      StringBuffer EmailText = new StringBuffer(1024);
      double checkTotal = 0;
      
      int batchCount = 0;
      
      initReport();
      
      try {
         m_PositivePayData.setString(1,m_batch_start);        
         m_PositivePayData.setString(2,m_batch_end);        
         m_PositivePayData.setString(3,m_batch_start);        
         m_PositivePayData.setString(4,m_batch_end);        
         positivePayData = m_PositivePayData.executeQuery();
         
         m_CurAction = "Building output file - info records";
         XSSFCellStyle currentStyle = m_StyleText;
         
         while ( positivePayData.next() && getStatus() != RptServer.STOPPED ) {
            row = m_Sheet.createRow(m_RowNum++);
            colNum = 0;
            
            // The check-total amount pulled in by the query does not have a decimal point. Put it in.
            double checkAmount = positivePayData.getDouble(CHK_AMT);
            checkAmount /= 100; // putting in the decimal places
            
            createCell(row, colNum++, ABA_NUM, currentStyle);
            createCell(row, colNum++, ACCT_NUM, currentStyle);
            createCell(row, colNum++, positivePayData.getString(CHK_NUM), currentStyle);
            createCell(row, colNum++, positivePayData.getString(CHK_DATE), currentStyle);
            createCell(row, colNum++, df.format(checkAmount), currentStyle);
            createCell(row, colNum++, NEW_CODE, currentStyle);
            createCell(row, colNum++, positivePayData.getString(PAYEE), currentStyle);
            
            if (batchCount == 0) {
            	batchCount = positivePayData.getInt("num_batches");
            }
            
            checkTotal += checkAmount;
         }

         m_WrkBk.write(outFile);
         
         
         EmailText.setLength(0);

         //we override the email text message here, mostly so whoever is running the report will get the check total
         EmailText.append("The Positive Pay report has finished running.\r\n");
         EmailText.append("The report file has been attached.\r\n");
         EmailText.append("\r\n");
         EmailText.append("Batches: ");
         EmailText.append(batchCount);
         EmailText.append("\r\n");
         EmailText.append("Total: ");
         
         DecimalFormat fmt = new DecimalFormat();
         EmailText.append(fmt.format(checkTotal));

         m_RptProc.setEmailMsg(EmailText.toString()); 

         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[PositivePayWells#buildOutputFile] ", ex);
      }

      finally {         
         
         closeRSet(positivePayData);
         
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[PositivePayWells#buildOutputFile] " , e);
         }

         outFile = null;
      }

      return result;
   }
   
   
   /**
    * Creates a cell of type String
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return HSSFCell newly created cell
    */
   private Cell createCell(Row row, int col, String val, CellStyle style)
   {
      Cell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(val));
      cell.setCellStyle(style);

      return cell;
   }
   
   
   /**
    * Creates the workbook and worksheet.  Creates any fonts and styles that
    * will be used.
    */
   private void initReport()
   {
      short col = 0;
      m_RowNum = 0;

      try {
         m_WrkBk = new XSSFWorkbook();
         
         //
         // Create a font that is normal size
         m_FontData = m_WrkBk.createFont();
         m_FontData.setFontHeightInPoints((short)11);
         m_FontData.setFontName("Calibri");

         //
         // Create a font that is normal size & bold
         m_FontBold = m_WrkBk.createFont();
         m_FontBold.setFontHeightInPoints((short)11);
         m_FontBold.setFontName("Calibri");
         m_FontBold.setBold(true);
         
         //
         // Setup the cell styles used in this report
         m_StyleText = m_WrkBk.createCellStyle();
         m_StyleText.setFont(m_FontData);
         m_StyleText.setAlignment(HorizontalAlignment.LEFT);
         
         // Style 8pt, left aligned, bold
         m_StyleBold = m_WrkBk.createCellStyle();
         m_StyleBold.setFont(m_FontBold);
         m_StyleBold.setAlignment(HorizontalAlignment.LEFT);
         
         m_Sheet = m_WrkBk.createSheet();
         m_Sheet.setMargin(XSSFSheet.BottomMargin, .25);
         m_Sheet.getPrintSetup().setLandscape(true);
         m_Sheet.getPrintSetup().setPaperSize((short)5);

         col = 0;

         // Initialize the default column widths
         for ( short i = 0; i < NUM_ALL_FIELDS; i++ )
            m_Sheet.setColumnWidth(i, 2000);

         // Wells Fargo does not want a header line. We only put it in if desired for testing.
         if (m_displayHeader) {
	         // Create the column headings
	         m_Row = m_Sheet.createRow(m_RowNum);
	         
	         m_Sheet.setColumnWidth(col, 3500);
	         createCell(m_Row, col++, "ABA Number", m_StyleBold);
	         m_Sheet.setColumnWidth(col, 4300);
	         createCell(m_Row, col++, "Account Number", m_StyleBold);
	         m_Sheet.setColumnWidth(col, 3900);
	         createCell(m_Row, col++, "Check Number", m_StyleBold);
	         m_Sheet.setColumnWidth(col, 5000);
	         createCell(m_Row, col++, "Date (MM-DD-YYYY)", m_StyleBold);
	         m_Sheet.setColumnWidth(col, 3800);
	         createCell(m_Row, col++, "Check Amount", m_StyleBold);
	         m_Sheet.setColumnWidth(col, 2000);
	         createCell(m_Row, col++, "Code", m_StyleBold);
	         m_Sheet.setColumnWidth(col, 10000);
	         createCell(m_Row, col++, "Payee", m_StyleBold);
	
	         m_RowNum++;
         } else {
        	 // set up column widths when there's no header line present
	         m_Sheet.setColumnWidth(col++, 3200);
	         m_Sheet.setColumnWidth(col++, 3300);
	         m_Sheet.setColumnWidth(col++, 3700);
	         m_Sheet.setColumnWidth(col++, 3200);
	         m_Sheet.setColumnWidth(col++, 3100);
	         m_Sheet.setColumnWidth(col++, 1700);
	         m_Sheet.setColumnWidth(col++, 10000);
         }
      }

      catch ( Exception e ) {
         log.error("[PositivePayWells#initReport] ", e );
      }
   }

   /**
    * Sets the parameters for the report.
    *    param(0) = batch number
    *    
    * @param params ArrayList<Param> - list of report parameters.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fname = new StringBuffer();
      int pcount = params.size();                     
      Param param = null;                             
      SimpleDateFormat formatter = new SimpleDateFormat ("yyyyMMddHHmmss");
      Date day = new Date();
       
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
         if ( param.name.equals("batchstart") )
            m_batch_start  = param.value;
         
         if ( param.name.equals("batchend") )
            m_batch_end  = param.value;
      }
      
      //
      // Build the file name.
      fname.append(formatter.format( day ));
      fname.append("-");
      fname.append("PositivePayWells.xlsx");
      m_FileNames.add(fname.toString());
   }
   
   
   /**
    * Main method for testing the PositivePayWells' output.
    * Can supply a LogDate here if desired for testing the queries on a specific date.
    * @param args
    *
   public static void main(String args[]) {
      PositivePayWells ppw = new PositivePayWells();
      
      // NOTE: PositivePayWells#createReport will need to swap how the connection settings
      //       are setup if you try to run this main - see comments in that method body.
      
      // if a date is not specified here, the system date will be used
      Param p1 = new Param();
      p1.name = "batchstart";
      p1.value = "59";
      Param p2 = new Param();
      p2.name = "batchend";
      p2.value = "59";
      ArrayList<Param> params = new ArrayList<Param>();
      params.add(p1);
      params.add(p2);
      
      ppw.m_FilePath = "C:\\PPW\\";
      //ppw.m_displayHeader = true;
      
      ppw.setParams(params);
      ppw.createReport();
   }*/
}
