/**
 * Title:         POBackhaulExt.java
 * Company:       Emery-Waterhouse
 * @author        Stephen Martel
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
import java.util.Calendar;

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
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class POBackhaulExt extends Report {

   // Workbook and style stuff
   private XSSFWorkbook m_WrkBk;
   private XSSFSheet m_Sheet;
   private XSSFRow m_Row = null;

   private XSSFFont m_FontTitle;
   private XSSFFont m_FontBold;
   private XSSFFont m_FontData;

   private XSSFCellStyle m_StyleText;  		  // Text left justified
   private XSSFCellStyle m_StyleTextRight;  	  // Text right justified
   private XSSFCellStyle m_StyleTitleCenter;   // Bold, larger size, centered, title color
   private XSSFCellStyle m_StyleTitle;         // Bold, larger size, centered, title color
   private XSSFCellStyle m_StyleTitleRight;    // Bold, larger size, right justified, title color
   private XSSFCellStyle m_StyleBold;          // Normal but bold, column header color

   private short m_RowNum = 0;
   
   // field names for retrieving from the query
   private static int NUM_FIELDS = 10;
   private final String WAREHOUSE = "warehouse";
   private final String EIS_VND_NBR = "eis_vnd_nbr";
   private final String VND_NAME = "vnd_name";
   private final String PO_NBR = "po_nbr";
   private final String PO_DATE = "open_date";
   private final String CARRIER = "carrier_name";
   private final String WEIGHT = "po_weight";
   private final String ORDER_COST = "po_cost";
   private final String DATE_NEEDED = "due_in_date";
   private final String COMMENTS = "comments";

   private PreparedStatement m_POStmt;
   
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
      	m_EdbConn = m_RptProc.getEdbConn();

         if ( prepareStatements() )
            created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[POBackhaulExt#createReport] ", ex);
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
      
      if ( m_EdbConn != null ) {
         try {
            sql.setLength(0);
            sql.append("SELECT PO_BACKHAUL_VIEW.EIS_VND_NBR, PO_BACKHAUL_VIEW.VND_NAME, PO_BACKHAUL_VIEW.PO_NBR, PO_BACKHAUL_VIEW.OPEN_DATE, ");
            sql.append("       PO_BACKHAUL_VIEW.PO_WEIGHT, PO_BACKHAUL_VIEW.PO_COST, PO_BACKHAUL_VIEW.COMMENTS, PO_BACKHAUL_VIEW.CARRIER_NAME, ");
            sql.append("       PO_BACKHAUL_VIEW.DUE_IN_DATE, PO_BACKHAUL_VIEW.VENDOR_CITY, PO_BACKHAUL_VIEW.VENDOR_STATE, PO_BACKHAUL_VIEW.WAREHOUSE ");
            sql.append("FROM   EJD.PO_BACKHAULEXT_VIEW PO_BACKHAUL_VIEW ");            
            sql.append("ORDER BY PO_BACKHAUL_VIEW.WAREHOUSE, PO_BACKHAUL_VIEW.VND_NAME, PO_BACKHAUL_VIEW.OPEN_DATE DESC ");
            m_POStmt = m_EdbConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[POBackhaulExt#prepareStatements] ", ex);
         }
         finally {
            sql = null;
         }         
      }
      else {
         log.error("[POBackhaulExt#prepareStatements] null oracle connection");
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
      ResultSet rs = null;
      boolean result = false;

      m_FileNames.add(m_RptProc.getUid() + "_pobackhaul_ext_" + getStartTime() + ".xlsx");
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      initReport();
      
      try {
         
      	rs = m_POStmt.executeQuery();
         m_CurAction = "Building output file - error records";
         
         while ( rs.next() && getStatus() != RptServer.STOPPED ) {
            row = m_Sheet.createRow(m_RowNum++);
            colNum = 0;

            createCell(row, colNum++, rs.getString(WAREHOUSE), m_StyleText);
            createCell(row, colNum++, rs.getString(EIS_VND_NBR), m_StyleTextRight);
            createCell(row, colNum++, rs.getString(VND_NAME), m_StyleText);
            createCell(row, colNum++, rs.getString(PO_NBR), m_StyleText);

            // Let's format the PO date to look like how it did in the Crystal Report
            Date date = rs.getDate(PO_DATE);
            createCell(row, colNum++, new SimpleDateFormat("MM-dd-yyyy").format(date), m_StyleTextRight);

            createCell(row, colNum++, rs.getString(CARRIER), m_StyleText);
            createCell(row, colNum++, rs.getString(WEIGHT), m_StyleTextRight);
            createCell(row, colNum++, rs.getString(ORDER_COST), m_StyleTextRight);
            
            // Let's format the due date to look like how it did in the Crystal Report
            date = rs.getDate(DATE_NEEDED);
            createCell(row, colNum++, new SimpleDateFormat("MM-dd-yyyy").format(date), m_StyleTextRight);
            
            createCell(row, colNum++, rs.getString(COMMENTS), m_StyleText);

         }

         m_WrkBk.write(outFile);
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[POBackhaulExt#buildOutputFile] ", ex);
      }

      finally {         
         
         closeRSet(rs);
         
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[POBackhaulExt#buildOutputFile] " , e);
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
         // Setup the cell styles used in this report
         m_StyleText = m_WrkBk.createCellStyle();
         m_StyleText.setFont(m_FontData);
         m_StyleText.setAlignment(HorizontalAlignment.LEFT);

         m_StyleTextRight = m_WrkBk.createCellStyle();
         m_StyleTextRight.setFont(m_FontData);
         m_StyleTextRight.setAlignment(HorizontalAlignment.RIGHT);
         
         m_StyleTitleCenter = m_WrkBk.createCellStyle();
         m_StyleTitleCenter.setFont(m_FontTitle);
         m_StyleTitleCenter.setAlignment(HorizontalAlignment.CENTER);
         
         // Style 8pt, left aligned, bold
         m_StyleBold = m_WrkBk.createCellStyle();
         m_StyleBold.setFont(m_FontBold);
         m_StyleBold.setAlignment(HorizontalAlignment.LEFT);

         m_StyleTitle = m_WrkBk.createCellStyle();
         m_StyleTitle.setFont(m_FontTitle);
         m_StyleTitle.setAlignment(HorizontalAlignment.LEFT);

         m_StyleTitleRight = m_WrkBk.createCellStyle();
         m_StyleTitleRight.setFont(m_FontTitle);
         m_StyleTitleRight.setAlignment(HorizontalAlignment.LEFT);
         m_StyleTitleRight.setAlignment(HorizontalAlignment.RIGHT);
         
         m_Sheet = m_WrkBk.createSheet();
         m_Sheet.setMargin(XSSFSheet.BottomMargin, .25);
         m_Sheet.getPrintSetup().setLandscape(true);
         m_Sheet.getPrintSetup().setPaperSize((short)5);

         
         
         // Create a top heading
         m_Row = m_Sheet.createRow(m_RowNum);
         String date = new SimpleDateFormat("MM-dd-yyyy").format(Calendar.getInstance().getTime());
         createCell(m_Row, col++, "Date: ", m_StyleTitle);
         createCell(m_Row, col++, date, m_StyleTitle);
         createCell(m_Row, col++, "Purchase Order Back Haul Extended Report", m_StyleTitle);

         // just coloring in the rest of the top row...
         for (short x = 0; x < NUM_FIELDS-3; x++) {
            createCell(m_Row, col++, "", m_StyleTitle);
         }
         
         m_RowNum++;
         col = 0;

         // Initialize the default column widths
         for ( short i = 0; i < NUM_FIELDS; i++ )
            m_Sheet.setColumnWidth(i, 3000);

         // Create the column headings
         m_Row = m_Sheet.createRow(m_RowNum);
         
         m_Sheet.setColumnWidth(col, 3000);
         createCell(m_Row, col++, "Warehouse", m_StyleBold);
         m_Sheet.setColumnWidth(col, 2800);
         createCell(m_Row, col++, "Eis Vnd #", m_StyleBold);
         m_Sheet.setColumnWidth(col, 12000);
         createCell(m_Row, col++, "Vendor Name", m_StyleBold);
         createCell(m_Row, col++, "PO #", m_StyleBold);
         createCell(m_Row, col++, "PO Date", m_StyleBold);
         createCell(m_Row, col++, "Carrier", m_StyleBold);
         createCell(m_Row, col++, "Weight", m_StyleBold);
         createCell(m_Row, col++, "Order $", m_StyleBold);
         createCell(m_Row, col++, "Date Needed", m_StyleBold);
         m_Sheet.setColumnWidth(col, 20000);
         createCell(m_Row, col++, "Comments", m_StyleBold);
         
         m_RowNum++;
      }

      catch ( Exception e ) {
         log.error("[POBackhaulExt#initReport] ", e );
      }
   }
   
   
   /**
    * Resource cleanup
    */
   public void cleanup()
   {
   	DbUtils.closeDbConn(null, m_POStmt, null);
   }
   

   /**
    * Sets the parameters of this report.
    * Param 0 - the log date to use in the process_log query. If not specified, the system date will be used.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
   	// none
   }

}
