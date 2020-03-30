/**
 * Title:         AceItemImport.java
 * Description:   Report for when items are added via the ACE feed.
 * Company:       Emery-Waterhouse
 * @author        Stephen Martel
 */

package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class AceItemImport extends Report {

   // Workbook and style stuff
   private XSSFWorkbook m_WrkBk;
   private XSSFSheet m_Sheet;
   private XSSFRow m_Row = null;

   private XSSFFont m_FontTitle;
   private XSSFFont m_FontBold;
   private XSSFFont m_FontData;
   private XSSFFont m_FontWarn;
   private XSSFFont m_FontError;

   private XSSFCellStyle m_StyleText;          // Text left justified, base
   private XSSFCellStyle m_StyleTextOdd;       // Text left justified, odd row fill color
   private XSSFCellStyle m_StyleTextOddWarn;   // Text left justified, odd row fill color, warning text color
   private XSSFCellStyle m_StyleTextOddError;  // Text left justified, odd row fill color, error text color
   private XSSFCellStyle m_StyleTextEven;      // Text left justified, even row fill color
   private XSSFCellStyle m_StyleTextEvenWarn;  // Text left justified, even row fill color, warning text color
   private XSSFCellStyle m_StyleTextEvenError; // Text left justified, even row fill color, error text color
   private XSSFCellStyle m_StyleTitle;         // Bold, larger size, centered, title color
   private XSSFCellStyle m_StyleTitleRight;    // Bold, larger size, right justified, title color
   private XSSFCellStyle m_StyleBold;          // Normal but bold, column header color

   private short m_RowNum = 0;
   
   // date to match in the process_log query, formatted like "dd-MMM-yyyy" ie "08-AUG-2014"
   // if no date is set as a param, the current system date will be used.
   private String m_LogDate = "";
   private String m_SystemDate = "";
   
   // Finals for retrieving the Report Parameters from the ResultSets
   private final String LOG_TIME = "log_time";
   private final String MSG_TYPE = "msg_type";
   private final String MSG1 = "msg1";
   private final String MSG2 = "msg2";
   private final String FIELD1 = "field1";
   private final String FIELD2 = "field2";
   private final String FIELD3 = "field3";
   private final String FIELD4 = "field4";
   
   private PreparedStatement m_InfoRecords;
   private PreparedStatement m_WarnRecords; // 'warn' and 'warning' // TODO is only 'warn' used, and never 'warning' ?
   private PreparedStatement m_ErrorRecords; // 'error' and 'abort' // TODO should 'abort' be separate, is it even used?
      
   
   /**
    * Creates the report file.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      // if no log date was specified as a parameter, then this date will be used.
      m_SystemDate = new SimpleDateFormat("dd-MMM-yyyy").format(Calendar.getInstance().getTime());
      
      try {
         
         // TODO: If testing with AceItemImport#main, then use this code chunk to set m_EdbConn;
         //java.util.Properties connProps = new java.util.Properties();
         //connProps.put("user", "eis_emery");
         //connProps.put("password", "mugwump");
         //m_EdbConn = java.sql.DriverManager.getConnection(
         //      "jdbc:oracle:thin:@10.128.0.127:1521:DANA",connProps);
           
         // TODO: If not testing with AceItemImport#main, use this line for the m_EdbConn!
         m_EdbConn = m_RptProc.getEdbConn();

         if ( prepareStatements() )
            created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[AceItemImport#createReport] ", ex);
      }

      finally {
        cleanup();

        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }

      return created;
   }
   
   
   /**
    * Prepares the sql queries for execution.
    * 
    * @return true if the statements were successfully prepared
    */
   private boolean prepareStatements() {
      StringBuilder sql = new StringBuilder(256);
      boolean isPrepared = false;
      
      // if no date is provided as a param (as dd-MMM-yyyy, ie 08-AUG-2014),
      // then the current system date will be used for this report.
      String date;
      if (m_LogDate.isEmpty()) {
         date = m_SystemDate;
      } else{
         date = m_LogDate;
      }
      
      if ( m_EdbConn != null ) {
         try {
            sql.setLength(0);
            sql.append("select * ");
            sql.append("from process_log ");
            sql.append("where proc_name = 'aceitems' and trunc(log_time) = '" + date + "' and (MSG_TYPE='abort' or MSG_TYPE='error')");
            // makes it grab item adds only, no item changes
            sql.append("and MSG1 = 'item add'");
            sql.append("order by 1");
            m_ErrorRecords = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select * ");
            sql.append("from process_log ");
            sql.append("where proc_name = 'aceitems' and trunc(log_time) = '" + date + "' and (MSG_TYPE='warn' or MSG_TYPE='warning')");
            // makes it grab item adds only, no item changes
            sql.append("and MSG1 = 'item add'");
            sql.append("order by 1");
            m_WarnRecords = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select * ");
            sql.append("from process_log ");
            sql.append("where proc_name = 'aceitems' and trunc(log_time) = '" + date + "' and MSG_TYPE='info'");
            // makes it grab item adds only, no item changes
            sql.append("and MSG1 = 'item add'");
            sql.append("order by 1");
            m_InfoRecords = m_EdbConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[AceItemImport#prepareStatements] ", ex);
         }
         finally {
            sql = null;
         }         
      }
      else {
         log.error("AceItemImport.prepareStatements - null edb connection");
      }
      
      return isPrepared;
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
      ResultSet errorData = null;
      ResultSet warnData = null;
      ResultSet infoData = null;
      boolean result = false;
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      initReport();
      
      try {
         
         // We want errors / aborts to be at the top of the document, so do those first
         
         errorData = m_ErrorRecords.executeQuery();
         m_CurAction = "Building output file - error records";
         XSSFCellStyle currentStyle = m_StyleTextOdd;
         XSSFCellStyle currentWarn = m_StyleTextOddWarn;
         XSSFCellStyle currentError = m_StyleTextOddError;
         
         while ( errorData.next() && getStatus() != RptServer.STOPPED ) {
            row = m_Sheet.createRow(m_RowNum++);
            colNum = 0;

            createCell(row, colNum++, errorData.getString(LOG_TIME), currentStyle);
            
            // we know all of these rows are msg_type 'error', so use the error style
            createCell(row, colNum++, errorData.getString(MSG_TYPE), currentError);
            
            createCell(row, colNum++, errorData.getString(MSG1), currentStyle);
            createCell(row, colNum++, errorData.getString(MSG2), currentStyle);
            createCell(row, colNum++, errorData.getString(FIELD1), currentStyle);
            createCell(row, colNum++, errorData.getString(FIELD2), currentStyle);
            createCell(row, colNum++, errorData.getString(FIELD3), currentStyle);
            createCell(row, colNum++, errorData.getString(FIELD4), currentStyle);
            
            // alternate rows between the odd/even colors to improve readability of the report
            if (currentStyle == m_StyleTextOdd){
               currentStyle = m_StyleTextEven;
               currentWarn = m_StyleTextEvenWarn;
               currentError = m_StyleTextEvenError;
            }
            else {
               currentStyle = m_StyleTextOdd;
               currentWarn = m_StyleTextOddWarn;
               currentError = m_StyleTextOddError;
            }
         }
         
         
         // Next up, we want the warnings
         
         warnData = m_WarnRecords.executeQuery();
         m_CurAction = "Building output file - warning records";
         
         while ( warnData.next() && getStatus() != RptServer.STOPPED ) {
            row = m_Sheet.createRow(m_RowNum++);
            colNum = 0;
            
            createCell(row, colNum++, warnData.getString(LOG_TIME), currentStyle);
            
            // we know all of these rows are msg_type 'warn', so use the warn style
            createCell(row, colNum++, warnData.getString(MSG_TYPE), currentWarn);
            
            createCell(row, colNum++, warnData.getString(MSG1), currentStyle);
            createCell(row, colNum++, warnData.getString(MSG2), currentStyle);
            createCell(row, colNum++, warnData.getString(FIELD1), currentStyle);
            createCell(row, colNum++, warnData.getString(FIELD2), currentStyle);
            createCell(row, colNum++, warnData.getString(FIELD3), currentStyle);
            createCell(row, colNum++, warnData.getString(FIELD4), currentStyle);
            
            // alternate rows between the odd/even colors to improve readability of the report
            if (currentStyle == m_StyleTextOdd){
               currentStyle = m_StyleTextEven;
               currentWarn = m_StyleTextEvenWarn;
               currentError = m_StyleTextEvenError;
            }
            else {
               currentStyle = m_StyleTextOdd;
               currentWarn = m_StyleTextOddWarn;
               currentError = m_StyleTextOddError;
            }
         }
         
         
         // Now for the info messages

         infoData = m_InfoRecords.executeQuery();
         m_CurAction = "Building output file - info records";
         
         while ( infoData.next() && getStatus() != RptServer.STOPPED ) {
            row = m_Sheet.createRow(m_RowNum++);
            colNum = 0;

            createCell(row, colNum++, infoData.getString(LOG_TIME), currentStyle);
            createCell(row, colNum++, infoData.getString(MSG_TYPE), currentStyle);
            createCell(row, colNum++, infoData.getString(MSG1), currentStyle);
            createCell(row, colNum++, infoData.getString(MSG2), currentStyle);
            createCell(row, colNum++, infoData.getString(FIELD1), currentStyle);
            createCell(row, colNum++, infoData.getString(FIELD2), currentStyle);
            createCell(row, colNum++, infoData.getString(FIELD3), currentStyle);
            createCell(row, colNum++, infoData.getString(FIELD4), currentStyle);
            
            // alternate rows between the odd/even colors to improve readability of the report
            if (currentStyle == m_StyleTextOdd){
               currentStyle = m_StyleTextEven;
               currentWarn = m_StyleTextEvenWarn;
               currentError = m_StyleTextEvenError;
            }
            else {
               currentStyle = m_StyleTextOdd;
               currentWarn = m_StyleTextOddWarn;
               currentError = m_StyleTextOddError;
            }
         }
         
         
         m_WrkBk.write(outFile);
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[AceItemImport#buildOutputFile] ", ex);
      }

      finally {         
         
         closeRSet(errorData);
         closeRSet(warnData);
         closeRSet(infoData);
         
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[AceItemImport#buildOutputFile] ", e);
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
      // if no date is provided as a param (as dd-MMM-yyyy, ie 08-AUG-2014),
      // then the current system date will be used for this report.
      String date;
      if (m_LogDate.isEmpty()) {
         date = m_SystemDate;
      } else{
         date = m_LogDate;
      }
         
      short col = 0;
      m_RowNum = 0;

      try {
         m_WrkBk = new XSSFWorkbook();

         // colors for the report
         XSSFColor headerColor = new XSSFColor(new java.awt.Color(238,236,225));  // for fill
         XSSFColor columnColor = new XSSFColor(new java.awt.Color(221,217,196));  // for fill
         XSSFColor oddRowColor = new XSSFColor(new java.awt.Color(242,242,242));  // for fill
         XSSFColor evenRowColor = new XSSFColor(new java.awt.Color(217,217,217)); // for fill
         XSSFColor warningColor = new XSSFColor(new java.awt.Color(154,0,0));  // for text
         XSSFColor errorColor = new XSSFColor(new java.awt.Color(255,0,0));    // for text
         
         //
         // Create a font for titles
         m_FontTitle = m_WrkBk.createFont();
         m_FontTitle.setFontHeightInPoints((short)10);
         m_FontTitle.setFontName("Arial");
         m_FontTitle.setBold(true);

         //
         // Create a font that is normal size & bold
         m_FontBold = m_WrkBk.createFont();
         m_FontBold.setFontHeightInPoints((short)8);
         m_FontBold.setFontName("Arial");
         m_FontBold.setBold(true);

         //
         // Create a font that is normal size & bold
         m_FontData = m_WrkBk.createFont();
         m_FontData.setFontHeightInPoints((short)8);
         m_FontData.setFontName("Arial");

         //
         // Create a font that is normal size & bold, with warning color
         m_FontWarn = m_WrkBk.createFont();
         m_FontWarn.setFontHeightInPoints((short)8);
         m_FontWarn.setFontName("Arial");
         m_FontWarn.setColor(warningColor);
         
         //
         // Create a font that is normal size & bold, with error color
         m_FontError = m_WrkBk.createFont();
         m_FontError.setFontHeightInPoints((short)8);
         m_FontError.setFontName("Arial");
         m_FontError.setColor(errorColor);
         
         //
         // Setup the cell styles used in this report
         m_StyleText = m_WrkBk.createCellStyle();
         m_StyleText.setFont(m_FontData);
         m_StyleText.setAlignment(HorizontalAlignment.LEFT);
         
         //
         // Style 8pth, left aligned, color for odd rows
         m_StyleTextOdd = m_WrkBk.createCellStyle();
         m_StyleTextOdd.setFont(m_FontData);
         m_StyleTextOdd.setAlignment(HorizontalAlignment.LEFT);
         m_StyleTextOdd.setFillForegroundColor(oddRowColor);
         m_StyleTextOdd.setFillPattern(FillPatternType.SOLID_FOREGROUND);

         //
         // Style 8pth, left aligned, color for odd rows, error text color
         m_StyleTextOddWarn = m_WrkBk.createCellStyle();
         m_StyleTextOddWarn.setFont(m_FontWarn);
         m_StyleTextOddWarn.setAlignment(HorizontalAlignment.LEFT);
         m_StyleTextOddWarn.setFillForegroundColor(oddRowColor);
         m_StyleTextOddWarn.setFillPattern(FillPatternType.SOLID_FOREGROUND);

         //
         // Style 8pth, left aligned, color for odd rows, warning text color
         m_StyleTextOddError = m_WrkBk.createCellStyle();
         m_StyleTextOddError.setFont(m_FontError);
         m_StyleTextOddError.setAlignment(HorizontalAlignment.LEFT);
         m_StyleTextOddError.setFillForegroundColor(oddRowColor);
         m_StyleTextOddError.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         
         //
         // Style 8pth, left aligned, color for even rows
         m_StyleTextEven = m_WrkBk.createCellStyle();
         m_StyleTextEven.setFont(m_FontData);
         m_StyleTextEven.setAlignment(HorizontalAlignment.LEFT);
         m_StyleTextEven.setFillForegroundColor(evenRowColor);
         m_StyleTextEven.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         
         //
         // Style 8pth, left aligned, color for odd rows, error text color
         m_StyleTextEvenWarn = m_WrkBk.createCellStyle();
         m_StyleTextEvenWarn.setFont(m_FontWarn);
         m_StyleTextEvenWarn.setAlignment(HorizontalAlignment.LEFT);
         m_StyleTextEvenWarn.setFillForegroundColor(evenRowColor);
         m_StyleTextEvenWarn.setFillPattern(FillPatternType.SOLID_FOREGROUND);

         //
         // Style 8pth, left aligned, color for odd rows, warning text color
         m_StyleTextEvenError = m_WrkBk.createCellStyle();
         m_StyleTextEvenError.setFont(m_FontError);
         m_StyleTextEvenError.setAlignment(HorizontalAlignment.LEFT);
         m_StyleTextEvenError.setFillForegroundColor(evenRowColor);
         m_StyleTextEvenError.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         
         // Style 8pt, left aligned, bold
         m_StyleBold = m_WrkBk.createCellStyle();
         m_StyleBold.setFont(m_FontBold);
         m_StyleBold.setAlignment(HorizontalAlignment.LEFT);
         m_StyleBold.setFillForegroundColor(columnColor);
         m_StyleBold.setFillPattern(FillPatternType.SOLID_FOREGROUND);

         m_StyleTitle = m_WrkBk.createCellStyle();
         m_StyleTitle.setFont(m_FontTitle);
         m_StyleTitle.setAlignment(HorizontalAlignment.LEFT);
         m_StyleTitle.setFillForegroundColor(headerColor);
         m_StyleTitle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

         m_StyleTitleRight = m_WrkBk.createCellStyle();
         m_StyleTitleRight.setFont(m_FontTitle);
         m_StyleTitleRight.setAlignment(HorizontalAlignment.LEFT);
         m_StyleTitleRight.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleTitleRight.setFillForegroundColor(headerColor);
         m_StyleTitleRight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         
         m_Sheet = m_WrkBk.createSheet();
         m_Sheet.setMargin(XSSFSheet.BottomMargin, .25);
         m_Sheet.getPrintSetup().setLandscape(true);
         m_Sheet.getPrintSetup().setPaperSize((short)5);

         
         
         // Create a top heading
         m_Row = m_Sheet.createRow(m_RowNum);
         m_Sheet.setColumnWidth(1, 4100);
         createCell(m_Row, col++, "Ace Items Imported", m_StyleTitle);
         createCell(m_Row, col++, "", m_StyleTitle); // just coloring in the top row...
         createCell(m_Row, col++, "", m_StyleTitle); // ""
         createCell(m_Row, col++, "", m_StyleTitle);
         createCell(m_Row, col++, "", m_StyleTitle);
         createCell(m_Row, col++, "", m_StyleTitle);
         createCell(m_Row, col++, "", m_StyleTitle);
         createCell(m_Row, col++, "On: " + date, m_StyleTitleRight);
         m_RowNum++;
         col = 0;

         // Initialize the default column widths
         for ( short i = 0; i < 20; i++ )
            m_Sheet.setColumnWidth(i, 2000);

         // Create the column headings
         m_Row = m_Sheet.createRow(m_RowNum);
         
         m_Sheet.setColumnWidth(col, 4700);
         createCell(m_Row, col++, "Log Time", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3300);
         createCell(m_Row, col++, "Message Type", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3000);
         createCell(m_Row, col++, "Change Type", m_StyleBold);
         m_Sheet.setColumnWidth(col, 7000);
         createCell(m_Row, col++, "Message", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Item Id", m_StyleBold); // TODO looks like ITEM_ID. is it always? or a message description?
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "UPC", m_StyleBold); // TODO looks like UPC. is it always UPC? or message description 2?
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Extra Details 1", m_StyleBold); // TODO what is field3?  message description 3?
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Extra Details 2", m_StyleBold); // TODO what is field4?  message description 4?

         m_RowNum++;
      }

      catch ( Exception e ) {
         log.error("[AceItemImport#initReport] ", e );
      }
   }
   
   /**
    * Resource cleanup
    */
   public void cleanup()
   {
      closeStmt(m_InfoRecords);
      closeStmt(m_WarnRecords);
      closeStmt(m_ErrorRecords);
   }
   
   /**
    * Sets the parameters of this report.
    * Param 0 - the log date to use in the process_log query. If not specified, the system date will be used.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();      
      String tmp = Long.toString(System.currentTimeMillis());
                  
      fileName.append("ace_items_report");      
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
      
      for (int i = 0; i < params.size(); i++) {
         Param param = params.get(i);
         
         if (param.name.equalsIgnoreCase("logdate"))
            m_LogDate = param.value;
      }
      
      if (m_LogDate == null || m_LogDate.length() == 0)
         m_LogDate = new SimpleDateFormat("dd-MMM-yyyy").format(Calendar.getInstance().getTime());
   }
   
   /**
    * Main method for testing the AceItemImport's output.
    * Can supply a LogDate here if desired for testing the queries on a specific date.
    * @param args
    */
   public static void main(String args[]) {
      AceItemImport ace = new AceItemImport();
      
      // NOTE: AceItemImport#createReport will need to swap how the connection settings
      //       are setup if you try to run this main - see comments in that method body.
      
      // if a date is not specified here, the system date will be used
      ace.m_LogDate = "27-Aug-2014";
      
      StringBuffer fileName = new StringBuffer();      
      String tmp = Long.toString(System.currentTimeMillis());
      fileName.append("ace_items_imported");
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      ace.m_FileNames.add(fileName.toString());
      
      ace.m_FilePath = "C:\\exp\\";
      
      ace.createReport();
   }
}
