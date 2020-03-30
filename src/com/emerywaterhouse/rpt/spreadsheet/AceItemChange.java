/**
 * Title:         AceItemChange.java
 * Description:   Report for when items are changed via the ACE feed.
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

public class AceItemChange extends Report {

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
   
   public static int NUM_ALL_FIELDS = 52; // count of all fields used from the item table + process log table
   public static int NUM_ITEM_FIELDS = 44; // count of all fields used from the item table
   
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
   private PreparedStatement m_ItemRecords;
   private PreparedStatement m_ItemEaRecords; //Execute this only when there's an item_ea_id
   
   
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
         
         // TODO: If testing with AceItemChange#main, then use this code chunk to set m_OraConn;
         //java.util.Properties connProps = new java.util.Properties();
         //connProps.put("user", "ejd");
         //FOR GROK
         //connProps.put("password", "boxer");
         //m_OraConn = java.sql.DriverManager.getConnection(
          //    "jdbc:oracle:thin:@10.128.0.9:1521:GROK",connProps);
              
         
         //FOR DANA
         //connProps.put("password", "mugwump");
         //m_OraConn = java.sql.DriverManager.getConnection(
         //      "jdbc:oracle:thin:@10.128.0.127:1521:DANA",connProps);
           
         // TODO: If not testing with AceItemChange#main, use this line for the m_OraConn!
         //m_OraConn = m_RptProc.getOraConn();
         m_EdbConn = m_RptProc.getEdbConn();

         if ( prepareStatements() )
            created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[AceItemChange#createReport] ", ex);
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
      StringBuffer sql = new StringBuffer(256);
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
            // makes it grab item changes only
            sql.append("and MSG1 = 'item change'");            
            sql.append("order by 1");
            m_ErrorRecords = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select * ");
            sql.append("from process_log ");
            sql.append("where proc_name = 'aceitems' and trunc(log_time) = '" + date + "' and (MSG_TYPE='warn' or MSG_TYPE='warning')");
            // makes it grab item changes only
            sql.append("and MSG1 = 'item change'");
            sql.append("order by 1");
            m_WarnRecords = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select * ");
            sql.append("from process_log ");
            sql.append("where proc_name = 'aceitems' and trunc(log_time) = '" + date + "' and MSG_TYPE='info'");
            // makes it grab item changes only
            sql.append("and MSG1 = 'item change'");
            sql.append("order by 1");
            m_InfoRecords = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select iea.item_id, iea.description, ejd.setup_date, eiw.status_date, ejd.pallet_qty, ");
            sql.append("ejd.WEIGHT, ejd.SOQ_COMMENT, ejd.STICKERS, ejd.EQUIVALENT, ejd.FLC_ID, ejd.DEPT_ID, ");
            sql.append("ejd.BROKEN_CASE_ID, ejd.HAZARD_ID, ejd.REGULATED_QTY, ejd.MARINE_POLLUTANT, ejd.FLASH_POINT, ");
            sql.append("ejd.AEROSOL, ejd.FLAMMABLE, ejd.OIL, ejd.FLAMMABLE_PLASTIC, ejd.LAST_HAZ_REVIEW, eiw.SEASONAL, ");
            sql.append("eiw.NEVER_OUT, eiw.STOCK_PACK, eiw.MIN_ORDER, eiw.VELOCITY_ID, eiw.RESTRICT_RESERVE_BEGIN, ");
            sql.append("eiw.RESTRICT_RESERVE_END, eiw.FORECAST_NO_DEMAND, eiw.DISP_ID, iea.ITEM_TYPE_ID, iea.DISPLAY_BRKDWN_ID, ");
            sql.append("iea.VDH_ID, iea.CONVENIENCE_PACK_1, iea.CONVENIENCE_PACK_2, iea.CONVENIENCE_PACK_3, iea.BUY_MULT, ");
            sql.append("iea.VENDOR_ID, iea.SHIP_UNIT_ID, iea.RET_UNIT_ID, iea.RETAIL_PACK, ");
            sql.append("(select item_id from item_entity_attr where item_ea_id = (select subitem_ea_id from item_ea_sub where item_ea_id = iea.item_ea_id)) as SUG_SUB_ITEM, ");         
            sql.append("NULL as MANUAL_FCST_SPLIT, NULL as PLANNING_VENDOR_ID, NULL as VIRTUAL ");
            sql.append("from item_entity_attr iea ");
            sql.append("inner join ejd.ejd_item ejd on iea.ejd_item_id = ejd.ejd_item_id ");
            sql.append("inner join ejd.ejd_item_warehouse eiw on eiw.ejd_item_id = iea.ejd_item_id ");                    
            sql.append("where iea.item_id = ? ");              
            sql.append("limit 1 "); 
            m_ItemRecords = m_EdbConn.prepareStatement(sql.toString());             

            sql.setLength(0);
            sql.append("select iea.item_id, iea.description, ejd.setup_date, eiw.status_date, ejd.pallet_qty, ");
            sql.append("ejd.WEIGHT, ejd.SOQ_COMMENT, ejd.STICKERS, ejd.EQUIVALENT, ejd.FLC_ID, ejd.DEPT_ID, ");
            sql.append("ejd.BROKEN_CASE_ID, ejd.HAZARD_ID, ejd.REGULATED_QTY, ejd.MARINE_POLLUTANT, ejd.FLASH_POINT, ");
            sql.append("ejd.AEROSOL, ejd.FLAMMABLE, ejd.OIL, ejd.FLAMMABLE_PLASTIC, ejd.LAST_HAZ_REVIEW, eiw.SEASONAL, ");
            sql.append("eiw.NEVER_OUT, eiw.STOCK_PACK, eiw.MIN_ORDER, eiw.VELOCITY_ID, eiw.RESTRICT_RESERVE_BEGIN, ");
            sql.append("eiw.RESTRICT_RESERVE_END, eiw.FORECAST_NO_DEMAND, eiw.DISP_ID, iea.ITEM_TYPE_ID, iea.DISPLAY_BRKDWN_ID, ");
            sql.append("iea.VDH_ID, iea.CONVENIENCE_PACK_1, iea.CONVENIENCE_PACK_2, iea.CONVENIENCE_PACK_3, iea.BUY_MULT, ");
            sql.append("iea.VENDOR_ID, iea.SHIP_UNIT_ID, iea.RET_UNIT_ID, iea.RETAIL_PACK, ");
            sql.append("(select item_id from item_entity_attr where item_ea_id = (select subitem_ea_id from item_ea_sub where item_ea_id = iea.item_ea_id)) as SUG_SUB_ITEM, ");         
            sql.append("NULL as MANUAL_FCST_SPLIT, NULL as PLANNING_VENDOR_ID, NULL as VIRTUAL ");
            sql.append("from item_entity_attr iea ");
            sql.append("inner join ejd.ejd_item ejd on iea.ejd_item_id = ejd.ejd_item_id ");
            sql.append("inner join ejd.ejd_item_warehouse eiw on eiw.ejd_item_id = iea.ejd_item_id ");                    
            sql.append("where iea.item_ea_id = ? ");              
            sql.append("limit 1 ");        
            m_ItemEaRecords = m_EdbConn.prepareStatement(sql.toString());            
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[AceItemChange#prepareStatements] ", ex);
         }
         finally {
            sql = null;
         }         
      }
      else {
         log.error("[AceItemChange#prepareStatements] null oracle connection");
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

            // now fill in the information of the changed item
            grabItemDetails(row, errorData.getString(FIELD1), colNum, currentStyle);
            
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

            // now fill in the information of the changed item
            grabItemDetails(row, warnData.getString(FIELD1), colNum, currentStyle);
            
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

            // now fill in the information of the changed item
            grabItemDetails(row, infoData.getString(FIELD1), colNum, currentStyle);
            
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

         log.fatal("[AceItemChange#buildOutputFile] ", ex);
      }

      finally {         
         
         closeRSet(errorData);
         closeRSet(warnData);
         closeRSet(infoData);
         
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[AceItemChange#buildOutputFile] " , e);
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
         createCell(m_Row, col++, "Ace Items Changed", m_StyleTitle);
         // just coloring in the top row...
         for (short x = 0; x < 6; x++) {
            createCell(m_Row, col++, "", m_StyleTitle);
         }
         createCell(m_Row, col++, "On: " + date, m_StyleTitleRight);
         
         // coloring the top row above the item_id details... if not using details, remove this loop
         createCell(m_Row, col++, "Item Details:", m_StyleTitle);
         for (short x = 0; x < NUM_ITEM_FIELDS-1; x++) {
            createCell(m_Row, col++, "", m_StyleTitle);
         }
         
         m_RowNum++;
         col = 0;

         // Initialize the default column widths
         for ( short i = 0; i < NUM_ALL_FIELDS; i++ )
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

         // write all of the latest item details for the changed item.
         m_Sheet.setColumnWidth(col, 20000);
         createCell(m_Row, col++, "Description", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4000);
         createCell(m_Row, col++, "Setup Date", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Status Date", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Weight", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Pallet Quantity", m_StyleBold);
         m_Sheet.setColumnWidth(col, 12000);
         createCell(m_Row, col++, "SOQ Comment", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Stock Pack", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Equivalent", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Min Order", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Buy Multi", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Stickers", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Sug Sub Item", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Seasonal", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Vendor ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "FLC ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Dept ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Ship Unit ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Return Unit ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4000);
         createCell(m_Row, col++, "Broken Case ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Velocity ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Hazard ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Retail Pack", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Disp ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4200);
         createCell(m_Row, col++, "Regulated Quantity", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4200);
         createCell(m_Row, col++, "Marine Pollutant", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4000);
         createCell(m_Row, col++, "Flash Point", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Aerosol", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Flammable", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Oil", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4200);
         createCell(m_Row, col++, "Flammable Plastic", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4200);
         createCell(m_Row, col++, "Last Hazard Review", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Item Type ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4800);
         createCell(m_Row, col++, "Restrict Reserve Begin", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4800);
         createCell(m_Row, col++, "Restrict Reserve End", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4700);
         createCell(m_Row, col++, "Display Breakdown ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "VDH ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4400);
         createCell(m_Row, col++, "Forecast No Demand", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4500);
         createCell(m_Row, col++, "Convenience Pack 1", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4500);
         createCell(m_Row, col++, "Convenience Pack 2", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4500);
         createCell(m_Row, col++, "Convenience Pack 3", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4700);
         createCell(m_Row, col++, "Manual Forecast Split", m_StyleBold);
         m_Sheet.setColumnWidth(col, 4500);
         createCell(m_Row, col++, "Planning Vendor ID", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Virtual", m_StyleBold);
         m_Sheet.setColumnWidth(col, 3800);
         createCell(m_Row, col++, "Never Out", m_StyleBold);
         
         m_RowNum++;
      }

      catch ( Exception e ) {
         log.error("[AceItemChange#initReport] ", e );
      }
   }
   
   public void grabItemDetails(Row row, String strItemId, int nColNum, XSSFCellStyle currentStyle) {
     
	   try { 	  
		   
    	  ResultSet itemData = null;
    	       
			if (strItemId.contains("item ea id:")) {
				String parsedEaId = ParseItemEaId(strItemId);
				m_ItemEaRecords.setString(1, parsedEaId);
				itemData = m_ItemEaRecords.executeQuery();
			} 
			else {
				String parsedId = getItemId(strItemId);
				m_ItemRecords.setString(1, parsedId);
				itemData = m_ItemRecords.executeQuery();
			}
 	  	        
         // there should only be 1 record returned with any given item_id, so just grab the first one
         if  (itemData.next()) {            
            createCell(row, nColNum++, itemData.getString("DESCRIPTION"), currentStyle);
            createCell(row, nColNum++, itemData.getString("SETUP_DATE"), currentStyle);
            createCell(row, nColNum++, itemData.getString("STATUS_DATE"), currentStyle);
            createCell(row, nColNum++, itemData.getString("WEIGHT"), currentStyle);
            createCell(row, nColNum++, itemData.getString("PALLET_QTY"), currentStyle);
            createCell(row, nColNum++, itemData.getString("SOQ_COMMENT"), currentStyle);
            createCell(row, nColNum++, itemData.getString("STOCK_PACK"), currentStyle);
            createCell(row, nColNum++, itemData.getString("EQUIVALENT"), currentStyle);
            createCell(row, nColNum++, itemData.getString("MIN_ORDER"), currentStyle);
            createCell(row, nColNum++, itemData.getString("BUY_MULT"), currentStyle);
            createCell(row, nColNum++, itemData.getString("STICKERS"), currentStyle);
            createCell(row, nColNum++, itemData.getString("SUG_SUB_ITEM"), currentStyle);
            createCell(row, nColNum++, itemData.getString("SEASONAL"), currentStyle);
            createCell(row, nColNum++, itemData.getString("VENDOR_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("FLC_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("DEPT_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("SHIP_UNIT_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("RET_UNIT_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("BROKEN_CASE_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("VELOCITY_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("HAZARD_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("RETAIL_PACK"), currentStyle);
            createCell(row, nColNum++, itemData.getString("DISP_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("REGULATED_QTY"), currentStyle);
            createCell(row, nColNum++, itemData.getString("MARINE_POLLUTANT"), currentStyle);
            createCell(row, nColNum++, itemData.getString("FLASH_POINT"), currentStyle);
            createCell(row, nColNum++, itemData.getString("AEROSOL"), currentStyle);
            createCell(row, nColNum++, itemData.getString("FLAMMABLE"), currentStyle);
            createCell(row, nColNum++, itemData.getString("OIL"), currentStyle);
            createCell(row, nColNum++, itemData.getString("FLAMMABLE_PLASTIC"), currentStyle);
            createCell(row, nColNum++, itemData.getString("LAST_HAZ_REVIEW"), currentStyle);
            createCell(row, nColNum++, itemData.getString("ITEM_TYPE_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("RESTRICT_RESERVE_BEGIN"), currentStyle);
            createCell(row, nColNum++, itemData.getString("RESTRICT_RESERVE_END"), currentStyle);
            createCell(row, nColNum++, itemData.getString("DISPLAY_BRKDWN_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("VDH_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("FORECAST_NO_DEMAND"), currentStyle);
            createCell(row, nColNum++, itemData.getString("CONVENIENCE_PACK_1"), currentStyle);
            createCell(row, nColNum++, itemData.getString("CONVENIENCE_PACK_2"), currentStyle);
            createCell(row, nColNum++, itemData.getString("CONVENIENCE_PACK_3"), currentStyle);
            createCell(row, nColNum++, itemData.getString("MANUAL_FCST_SPLIT"), currentStyle);
            createCell(row, nColNum++, itemData.getString("PLANNING_VENDOR_ID"), currentStyle);
            createCell(row, nColNum++, itemData.getString("VIRTUAL"), currentStyle);
            createCell(row, nColNum++, itemData.getString("NEVER_OUT"), currentStyle);
         }
         else {
            // use red error text, match the odd/even fill coloring
            if (currentStyle == m_StyleTextOdd){
               createCell(row, nColNum++, "No corresponding item was found.", m_StyleTextOddError);
            } else {
               createCell(row, nColNum++, "No corresponding item was found.", m_StyleTextEvenError);
            }
            
            // color the remaining columns
            for (short x = 0; x < NUM_ITEM_FIELDS-1; x++) {
               createCell(row, nColNum++, "", currentStyle);
            }
         }
         
      } catch (Exception ex) {
         log.error("[AceItemChange#grabItemDetails]", ex);
      }
   }
   
   /**
    * Parse item ea id: and retrieve the just the id value
    * @param id the value of the Item Ea Id column
    * @return the parsed out Item ea Id
    */
   public String ParseItemEaId(String strItemId) {
	   return strItemId.replaceAll("[^0-9]", "");	
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
    * Given the full value of the Item Id column, parses the 7 digit Item Id from it
    * @param id the value of the Item Id column
    * @return the parsed out Item Id
    */
   public String getItemId(String id) {
      String finalId = "";
      
      // If the length is 7, it is the item id, grab it.
      if (id.length() == 7) {
         finalId = id;
      } else if (id.length() > 7) {
         // if the length is > 7, then it is probably: "item id: xxxxxxx", so grab the last 7 chars.
         finalId = id.substring(id.length()-7, id.length());
      } else {
         // unfortunately, the id is smaller than expected and probably not recoverable.
         // however, maybe we got the id as (for example) "6306" instead of "0006306", so let's try adding zeroes.
         finalId = id;
         while (finalId.length() < 7) {
            finalId = "0" + finalId;
         }
      }
      
      return finalId;
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
    * Main method for testing the AceItemChange's output.
    * Can supply a LogDate here if desired for testing the queries on a specific date.
    * @param args
    */
   public static void main(String args[]) {
      AceItemChange ace = new AceItemChange();
      
      // NOTE: AceItemChange#createReport will need to swap how the connection settings
      //       are setup if you try to run this main - see comments in that method body.
      
      // if a date is not specified here, the system date will be used
      ace.m_LogDate = "08-Nov-2015";
      
      StringBuffer fileName = new StringBuffer();      
      String tmp = Long.toString(System.currentTimeMillis());
      fileName.append("ace_items_changed");
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      ace.m_FileNames.add(fileName.toString());
      
      ace.m_FilePath = "C:\\exp\\";
      
      ace.createReport();
   }
}
