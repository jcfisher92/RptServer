/**
 * File: InventoryAging.java
 * Description: Inventory aging report for items in fascor.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 11/21/2006
 * Last Update: $Id: InventoryAging.java,v 1.33 2013/12/20 14:35:34 sgillis Exp $
 * 
 * History
 *    $Log: InventoryAging.java,v $
 *    Revision 1.33  2013/12/20 14:35:34  sgillis
 *    also removed cost > 0 check
 *
 *    Revision 1.32  2013/12/19 20:07:03  sgillis
 *    removed filter check on (qty > buy_mult)
 *
 *    Revision 1.31  2013/12/18 19:44:25  sgillis
 *    added status
 *
 *    Revision 1.30  2013/12/18 18:42:00  sgillis
 *    Modified to be customer-based, added test main(), removed inert FasConn checks
 *
 *    Revision 1.29  2011/04/05 17:56:20  jfisher
 *    1. Added the cube and location back to the report using the fascor db but in a single query.
 *    2. Changed the query so it only pulls back items with a quantity on hand and then remvoed the loop check.
 *    3. Removed all the unused and commented out stuff left behind by others.
 *
 *    Revision 1.28  2011/03/01 14:15:39  smurdock
 *    changed >= to <= for m_FilterDate
 *
 *    Revision 1.27  2010/10/19 04:46:56  smurdock
 *    frigged around with Wos's to get them to add up correctly
 *
 *    Revision 1.26  2010/10/18 06:03:36  smurdock
 *    includes functioning filter dates for item setup, fixed cost and avg cost bug
 *
 *    Revision 1.25  2010/09/23 10:10:16  jfisher
 *    Fixed warnings and removed unused methods.
 *
 *    Revision 1.24  2010/07/14 13:23:48  epearson
 *    Modified report logic to accept user defined date ranges per Fred Arsenault's request.
 *
 *    Revision 1.23  2010/05/10 14:45:56  smurdock
 *    Fascor calls were causing big perfomrance issues for warehouse.
 *
 *    Qtyonyhand now comes from Accpac.
 *
 *    Cube and Home Loc have been dropped (though the columns remain in hopes of a future restoration).  Fred agreed to lose these.
 *
 *    No more Fascor connection at all -- runs in 3 mintues instead of 4 hours too.
 *
 *    Revision 1.22  2010/04/26 04:24:33  smurdock
 *    now using average cost instead of current cost for calculations per Fred Arsenault
 *
 *    Revision 1.21  2010/04/19 08:54:09  smurdock
 *    set fascor connection to read uncomiitted to reduce locking
 *
 *    Revision 1.20  2010/03/26 22:34:54  smurdock
 *    added average cost
 *
 *    Revision 1.19  2010/03/23 16:26:41  jfisher
 *    Fixed a bug where the sales weren't displayed per DC.
 *
 *    Revision 1.18  2009/04/10 09:35:53  smurdock
 *    fixed bug looking for length of null m_dept
 *
 *    Revision 1.17  2009/03/06 21:01:58  jfisher
 *    Fixed an issue with the description field in the query.
 *
 *    Revision 1.16  2009/03/05 22:12:38  jfisher
 *    Added velocity code and filtering by department per Fred Arsenault.
 *
 *    Revision 1.15  2009/03/04 20:48:23  jfisher
 *    Fixed problem with running out of rows on large reports and fixed a region bug.
 *
 *    Revision 1.14  2009/02/18 15:11:50  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.13  2008/12/09 20:56:21  prichter
 *    Fixed another bug comparing strings
 *
 *    Revision 1.12  2008/12/09 20:54:34  prichter
 *    Fixed a bug in getQtyOnHand
 *
 *    Revision 1.11  2008/11/20 19:36:22  smurdock
 *    yet more code to handle Pittston or Portland
 *
 *    Revision 1.10  2008/10/29 21:24:42  jfisher
 *    Fixed some warnings
 *
 *    Revision 1.9  2008/10/29 21:14:27  jfisher
 *    Fixed some warnings
 *
 *    
 *    Revision 1.7  2008/08/01 14:42:30  smurdock
 *    added filter by Distribution Center
 *
 *    Revision 1.6  2008/03/06 14:01:48  jfisher
 *    Added additional columns per Adam B.
 *
 *    Revision 1.5  2008/02/20 16:20:31  jfisher
 *    Made a change on the output for Darrin
 *
 *    Revision 1.4  2008/01/23 19:27:58  jfisher
 *    Changed from wos 156 to wos 104 per Darrin's request
 *
 *    Revision 1.3  2006/12/07 16:29:40  jfisher
 *    added member vars to the finaization method.
 *
 *    Revision 1.2  2006/12/06 18:46:47  jfisher
 *    production version
 *   
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class InventoryAging extends Report
{
   private static final short MAX_COLS = 28;
   private static final double FIFTYTWO_WEEK_MILLIS = 31449600; 
   //private static final double YEAR_MILLIS = FIFTYTWO_WEEK_MILLIS +  86400000;  // 365 days
   
   
   private static final int MAX_VALUE  = 52000;
  
   private PreparedStatement m_SaleData;
   private PreparedStatement m_GetDCName;
   
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;
   
   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;
   
   private String m_BegDate;
   private String m_Dept;
   private String m_EndDate;
   private boolean m_Filtered;
   private String m_FilterDate;
   private String m_Warehouse;  //FASCOR id , sez Jeff.  From Delphi.
   private String m_Warehouse_Name;
           
   /**
    * 
    */
   public InventoryAging()
   {
      super();
      
      m_Filtered = false;      
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet("SLOB Page 1");
      
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
      m_BegDate = null;
      m_EndDate = null;
      m_FilterDate = null;
      m_SaleData = null;
      m_Dept = null;
            
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
      int pageNum = 1;
      String msg = "processing page %d, row %d, item %s";
      FileOutputStream outFile = null;
      ResultSet saleData = null;
      ResultSet DCData = null;
      long dateDiff = 0;
      long dateDiffdecimated = 0;
      SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
      boolean result = false;
      boolean show = true;
      String itemId = null;
      String setupDate = null;
      String convPck1 = null;
      String homeLoc = null;
      int buyMult = 0;
      int qtyOnHand = 0;
      int totUnits = 0;
      int wos52Units = 0;
      double cube = 0.0;
      double wosOnHand = 0.0;
      double wos1 = 0.0;
      double cost = 0.0;
      double wos0Cost = 0.0;
      double wos13Cost = 0.0;
      double wos26Cost = 0.0;
      double wos39Cost = 0.0;
      double wos52Cost = 0.0;
      double wos104Cost = 0.0;
          
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){
            m_GetDCName.setString(1,m_Warehouse);
            DCData = m_GetDCName.executeQuery();
            
            while ( DCData.next() && rowNum < 10  && m_Status == RptServer.RUNNING ) {
               m_Warehouse_Name = DCData.getString("name");
            }
         }

         rowNum = createCaptions();
         
         if ( m_Filtered && (m_FilterDate != null && m_FilterDate.length() > 0) ) {
            m_SaleData.setString(1, m_BegDate);
            m_SaleData.setString(2, m_EndDate);
            m_SaleData.setString(3, m_FilterDate);
            m_SaleData.setString(4, m_BegDate);
            m_SaleData.setString(5, m_EndDate);
            m_SaleData.setString(6, m_FilterDate);
         }
         else {
            m_SaleData.setString(1, m_BegDate);
            m_SaleData.setString(2, m_EndDate);
            m_SaleData.setString(3, m_BegDate);
            m_SaleData.setString(4, m_EndDate);
         }
         
         saleData = m_SaleData.executeQuery();

         // Calculate the difference between m_BegDate and m_EndDate
         dateDiff = getDateDiff(m_BegDate, m_EndDate, "MM/dd/yyyy");
         dateDiffdecimated = dateDiff/1000;  //JAVA HAVING TROUBLE WITH BIG INTEGERSD< EVERYTHING DIVIDEWD BY `1000
          
         while ( saleData.next() && m_Status == RptServer.RUNNING ) {
            itemId = saleData.getString("item_id");
            setCurAction(String.format(msg, pageNum, rowNum, itemId));            
            
            //lets' try getting on hand from accpac -- Fascor dirty reads may be causing problems
            qtyOnHand = saleData.getInt("qtyonhand");
            cube = saleData.getDouble("cube");
            homeLoc = saleData.getString("locs");
            buyMult = saleData.getInt("buy_mult");
            totUnits = saleData.getInt("tot_units");
            
            //changed cur_cost to avg_cost per Fred Arsemnault 4/21/2010 but i have doubts so leaving this in
            //cost = saleData.getDouble("cur_cost");
            cost = saleData.getDouble("avg_cost");
            setupDate = sdf.format(saleData.getDate("setup_date"));
            convPck1 = saleData.getString("convenience_pack_1");
            
            if ( m_Filtered )
               show = qtyOnHand > 0;
                        
            if ( show ) {
               //  torshnit
               // Calculate the weeks of supply, can't be a negative value.
               //wos1 = ((double)totUnits)/YR_WEEKS;
               wos1 = (((double)totUnits) * (dateDiffdecimated/FIFTYTWO_WEEK_MILLIS)) / 52;
               
               wosOnHand = wos1 == 0 ? MAX_VALUE : (qtyOnHand/wos1);
               
               wos0Cost = qtyOnHand * cost;                     // current cost
               
               wos13Cost = (qtyOnHand - (wos1 * 13)) * cost;    // for the quarter
               wos13Cost = wos13Cost > 0 ? wos13Cost : 0;
               
               wos26Cost = (qtyOnHand - (wos1 * 26)) * cost;    // for half the year
               wos26Cost = wos26Cost > 0 ? wos26Cost : 0;
               
               wos39Cost = (qtyOnHand - (wos1 * 39)) * cost;    // for 3/4 the year
               wos39Cost = wos39Cost > 0 ? wos39Cost : 0;
               
               wos52Cost = (qtyOnHand - (wos1 * 52)) * cost;    // for the year
               wos52Cost = wos52Cost > 0 ? wos52Cost : 0;
               
               wos52Units = (int) (wos52Cost / saleData.getDouble("avg_cost"));
               
               wos104Cost = (qtyOnHand - (wos1 * 104)) * cost;  // for two years
               wos104Cost = wos104Cost > 0 ? wos104Cost : 0;
                       
               row = createRow(rowNum++, MAX_COLS);
               colNum = 0;
               
               if ( row != null ) {               
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(itemId));
                  row.getCell(colNum++).setCellValue(saleData.getString("disposition"));
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(setupDate));
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(saleData.getString("description")));
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(saleData.getString("name")));
                  row.getCell(colNum++).setCellValue(saleData.getInt("stock_pack"));
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(saleData.getInt("broken_case_id") == 1 ? "N" : "Y"));
                  row.getCell(colNum++).setCellValue(buyMult);
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(saleData.getString("dept_num")));
                  row.getCell(colNum++).setCellValue(saleData.getDouble("tot_sales"));
                  row.getCell(colNum++).setCellValue(saleData.getInt("tot_lines"));
                  row.getCell(colNum++).setCellValue(totUnits);
                  row.getCell(colNum++).setCellValue(qtyOnHand);
                  //row.getCell(colNum++).setCellValue(cost > 0 ? cost : 0.000001);
                  row.getCell(colNum++).setCellValue(saleData.getDouble("cur_cost"));
                  row.getCell(colNum++).setCellValue(saleData.getDouble("avg_cost"));                  
                  row.getCell(colNum++).setCellValue(wos1);
                  row.getCell(colNum++).setCellValue(wosOnHand);
                  row.getCell(colNum++).setCellValue(wos0Cost);
                  row.getCell(colNum++).setCellValue(wos13Cost);
                  row.getCell(colNum++).setCellValue(wos26Cost);                  
                  row.getCell(colNum++).setCellValue(wos39Cost);
                  row.getCell(colNum++).setCellValue(wos52Cost);
                  row.getCell(colNum++).setCellValue(wos52Units);
                  row.getCell(colNum++).setCellValue(wos104Cost);                  
                  row.getCell(colNum++).setCellValue(cube);
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(homeLoc));
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(convPck1 == null ? "" : convPck1));
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(saleData.getString("velocity")));
               }
            }
            
            if ( rowNum > 65000 ) {
               m_Sheet.createFreezePane(1, 3);
               pageNum++;
               m_Sheet = m_Wrkbk.createSheet("SLOB Page " + pageNum);
               rowNum = createCaptions();               
            }
         }
         
         m_Wrkbk.write(outFile);
         saleData.close();

         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {
         closeStmt(m_SaleData);
         
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error(e);
         }

         outFile = null;
         row = null;
         convPck1 = null;
         itemId = null;
         setupDate = null;
         sdf = null;
         saleData = null;         
      }

      return result;
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_SaleData);
   }
   
   /**
    * Sets the captions on the report.
    */
   private int createCaptions()
   {
      XSSFCellStyle styleCaption;
      XSSFCellStyle styleTitle;
      XSSFFont font;
      XSSFCell cell = null;
      XSSFRow row = null;
      CellRangeAddress region = null;      
      int rowNum = 0;
      int colNum = 0;
      short rowHeight = 1000;
      StringBuffer title = new StringBuffer("SLOB Report: ");
            
      if ( m_Sheet == null )
         return 0;
      
      font = m_Wrkbk.createFont();
      font.setFontHeightInPoints((short)10);
      font.setFontName("Arial");
      font.setBold(true);
      
      styleTitle = m_Wrkbk.createCellStyle();
      styleTitle.setAlignment(HorizontalAlignment.LEFT);
      styleTitle.setFont(font);
      
      styleCaption = m_Wrkbk.createCellStyle();      
      styleCaption.setFont(font);
      styleCaption.setAlignment(HorizontalAlignment.CENTER);
      styleCaption.setWrapText(true);
      
      //
      // Create the title
      title.append(m_BegDate);
      title.append(" - ");
      title.append(m_EndDate);
         
      //
      // set the report title
      row = m_Sheet.createRow(rowNum++);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitle);
      cell.setCellValue(new XSSFRichTextString(title.toString()));
      
      //
      // Set the filter date if there is one
      title.setLength(0);
      if ( m_Filtered ) {
         if ( m_FilterDate != null ) {         
            title.append("Filter: On    Date: ");
            title.append(m_FilterDate);
         }
      }
      else
         title.append("Filter: Off");
      
      if ((m_Warehouse_Name != null) && (m_Warehouse_Name.length() > 0)) {
         title.append(" DC = ");
         title.append(m_Warehouse_Name);
      }
      
      //
      // Merge the title cells.  Gives a better look to the report.
      region = new CellRangeAddress(0, 0, 0, 2);
      m_Sheet.addMergedRegion(region);
      
      region = new CellRangeAddress(1, 1, 0, 2);
      m_Sheet.addMergedRegion(region);
      
      row = m_Sheet.createRow(rowNum++);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitle);
      cell.setCellValue(new XSSFRichTextString(title.toString()));
      
      //
      // Create the row for the captions.
      row = m_Sheet.createRow(rowNum++);
      row.setHeight(rowHeight);
      
      for ( int i = 0; i < MAX_COLS; i++ ) {
         cell = row.createCell(i);
         cell.setCellStyle(styleCaption);
      }
      
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Item #"));
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Status"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Add Date"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Description"));
      m_Sheet.setColumnWidth(colNum++, 4000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Vendor"));
      m_Sheet.setColumnWidth(colNum++, 4000);
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Vendor\nStock Pack"));      
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Broken\nCase"));
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Buy\nMult"));
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Dept"));
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Total\nSales($)"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Total\nLines Shipped"));
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Total\nUnits Shipped"));      
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("On-Hand Qty"));
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Current\nCost"));
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Average\nCost"));
      row.getCell(colNum).setCellValue(new XSSFRichTextString("1 Week of\nSupply\n(1-WOS Units)"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Weeks\nof Supply\nOn-hand"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Est\nWOS-0"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Est\nWOS-13"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Est\nWOS-26"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Est\nWOS-39"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Est\nWOS-52"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Est WOS-52\n / Avg Cost"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Est\nWOS-104"));
      m_Sheet.setColumnWidth(colNum++, 3000);      
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Cube"));      
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Locations"));
      m_Sheet.setColumnWidth(colNum++, 4000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Conv\nPack1"));      
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Vel"));      
            
      font = null;
      styleTitle = null;
      styleCaption = null;
      title = null;
      region = null;
      
      return rowNum;
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
         m_EdbConn = m_RptProc.getEdbConn();
         
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
    * Returns the difference in milliseconds between the two given dates.
    * 
    * @param beginDate
    * @param endDate
    * @param dateFormat
    * @return The difference in time between the dates in milliseconds
    * @throws ParseException
    */
   private long getDateDiff(String beginDate, String endDate, String pattern) throws ParseException{
      SimpleDateFormat sdf = new SimpleDateFormat(pattern);
      
      sdf.getCalendar().setLenient(false);
      
      Date begin = sdf.parse(beginDate);
       Date end = sdf.parse(endDate);
       
       return Math.abs(end.getTime() - begin.getTime());
   }
      
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(256);
      StringBuffer sqlDCName = new StringBuffer(256);

      boolean isPrepared = false;
      
      if ( m_EdbConn != null ) {
         try {
            sql.append("select ");
            sql.append("item_entity_attr.item_id, nvl(qty.available_qty, 0) as qtyonhand, ejd_item.setup_date, item_entity_attr.description,  \r\n");
            sql.append("vendor.name, ejd_item_warehouse.stock_pack, item_disp.disposition, "); 
            sql.append("ejd_item.broken_case_id, item_entity_attr.buy_mult, emery_dept.dept_num, ");
            sql.append("tot_sales, ");
            sql.append("tot_lines, ");
            sql.append("tot_units, ");
            sql.append("buy as cur_cost, ");
            sql.append("item_entity_attr.convenience_pack_1, item_velocity.velocity, ");
            sql.append("round(decode(qty.available_qty, 0, buy, " +
            		"fi_cost.totalcost / qty.available_qty),2) as avg_cost, ");
            sql.append("nvl(sku1.\"cube\"::varchar, sku4.\"cube\"::varchar) as cube, faslocs.locs ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id ");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0) {
               if (m_Warehouse.equals("01"))
                 sql.append(" and ejd_item_warehouse.warehouse_id = 1 ");
               if (m_Warehouse.equals("04"))
                 sql.append(" and ejd_item_warehouse.warehouse_id = 2 ");
            }
            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id ");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0) {
               if (m_Warehouse.equals("01"))
                 sql.append("     and ejd_item_price.warehouse_id = 1 ");
               if (m_Warehouse.equals("04"))
                 sql.append("     and ejd_item_price.warehouse_id = 2 ");
            }
            //join in totals
            sql.append("join (select item_id, sum(unit_sell*qty_shipped) as tot_sales, " +
            		"count(item_id) as tot_lines, sum(qty_shipped) as tot_units ");
            sql.append("from sale_item ");
            sql.append("join cust_warehouse on cust_warehouse.customer_id = sale_item.customer_id ");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0){
               if (m_Warehouse.equals("01"))
                 sql.append(" and cust_warehouse.warehouse_id = 1 ");
               if (m_Warehouse.equals("04"))
                 sql.append(" and cust_warehouse.warehouse_id = 2 ");
             }
            sql.append(" where invoice_date >= to_date(?,'mm/dd/yyyy') and invoice_date <= to_date(?,'mm/dd/yyyy') ");
            sql.append(" and sale_type in ('WAREHOUSE','TOOLREPAIR','APG WHS') ");
            sql.append(" group by item_id) totals ");
            sql.append(" on totals.item_id = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9) ");
            //end join on totals
            //join in cost from emeryd
            sql.append("join \r\n");
            sql.append("(select item_entity_attr.item_id, sum(fi.totalcost) as totalcost ");
            sql.append("   from item_entity_attr ");
            sql.append("   join ejd.sage300_iciloc_mv fi on  fi.itemno = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9)  ");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0){
               if (m_Warehouse.equals("01"))
                 sql.append(" and fi.location = '01' ");
               if (m_Warehouse.equals("04"))
                 sql.append(" and fi.location = '02' ");
             }
            sql.append("group by \r\n");
            sql.append("item_entity_attr.item_id) fi_cost on fi_cost.item_id = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9) ");
            //yay done emeryd
            sql.append("join vendor on item_entity_attr.vendor_id = vendor.vendor_id ");
            sql.append("join item_disp on ejd_item_warehouse.disp_id = item_disp.disp_id ");
            sql.append("join emery_dept on ejd_item.dept_id = emery_dept.dept_id ");
            
            if(( m_Dept != null) &&( m_Dept.length() > 0 ))
               sql.append(String.format(" and dept_num = %s ", m_Dept) );
            
            sql.append("join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id  ");
            sql.append("left outer join (select item_id, avail_qty as available_qty from ejd_item_warehouse ");
            sql.append("join item_entity_attr on item_entity_attr.ejd_item_id = ejd_item_warehouse.ejd_item_id ");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0){
               if (m_Warehouse.equals("01"))
                 sql.append(" where warehouse_id = 1 ");
               if (m_Warehouse.equals("04"))
                 sql.append(" where warehouse_id = 2 ");
             }
            sql.append(") ");
            sql.append(" qty on item_entity_attr.item_id = qty.item_id and item_entity_attr.item_type_id not in (8,9) ");
            
            //
            // Pull the sku_master and location information from fascor for now.  It seems to work ok.
            // Dean needs all locations and Fred doesn't want repeating groups.  The wm_concat will
            // give us all the locations in a comma separated field.
            sql.append("left outer join sku_master sku1 on sku1.sku = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9) and sku1.warehouse = 'PORTLAND' ");
            sql.append("left outer join sku_master sku4 on sku4.sku = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9) and sku4.warehouse = 'PITTSTON' ");
            sql.append("left outer join ( ");
            sql.append("   select sku, string_agg(distinct location.\"loc_id\", ',') as locs ");
            sql.append("   from loc_allocation ");																	
            sql.append("   join location on location.\"loc_id\" = loc_allocation.\"loc_id\" ");
            sql.append("   group by sku ");
            sql.append(" ) faslocs on faslocs.sku = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9)   ");
            
            if ( m_Filtered && (m_FilterDate != null && m_FilterDate.length() > 0) ) {
               sql.append("where \r\n");
               sql.append("   ejd_item.setup_date <= to_date(?, 'mm/dd/yyyy') \r\n");
            }

            //we also require items that weren't sold, but that we have quantity for, 
            //so put on your union hat cause #sqlgotreal
            sql.append("union ");
            sql.append("select ");
            sql.append("item_entity_attr.item_id, nvl(qty.available_qty, 0) as qtyonhand, ejd_item.setup_date, item_entity_attr.description,  \r\n");
            sql.append("vendor.name, ejd_item_warehouse.stock_pack, item_disp.disposition, ");
            sql.append("ejd_item.broken_case_id, item_entity_attr.buy_mult, emery_dept.dept_num, ");
            sql.append("0 tot_sales, ");
            sql.append("0 tot_lines, "); //we specifically don't have sales for these items
            sql.append("0 tot_units, ");
            sql.append("buy as cur_cost, ");
            sql.append("item_entity_attr.convenience_pack_1, item_velocity.velocity, ");
            sql.append("round(decode(qty.available_qty, 0, buy, " +
                    "fi_cost.totalcost / qty.available_qty),2) as avg_cost, ");
            sql.append("nvl(sku1.\"cube\"::varchar, sku4.\"cube\"::varchar) as cube, faslocs.locs ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id ");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0) {
               if (m_Warehouse.equals("01"))
                 sql.append(" and ejd_item_warehouse.warehouse_id = 1 ");
               if (m_Warehouse.equals("04"))
                 sql.append(" and ejd_item_warehouse.warehouse_id = 2 ");
            }
            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id ");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0) {
               if (m_Warehouse.equals("01"))
                 sql.append("     and ejd_item_price.warehouse_id = 1 ");
               if (m_Warehouse.equals("04"))
                 sql.append("     and ejd_item_price.warehouse_id = 2 ");
            }
            //no totals to join
            //join in cost from emeryd
            sql.append("join \r\n");
            sql.append("(select item_entity_attr.item_id, sum(fi.totalcost) as totalcost ");
            sql.append("   from item_entity_attr ");
            sql.append("   join ejd.sage300_iciloc_mv fi on fi.itemno = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9) ");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0){
               if (m_Warehouse.equals("01"))
                 sql.append(" and fi.location = '01' ");
               if (m_Warehouse.equals("04"))
                 sql.append(" and fi.location = '02' ");
             }
            sql.append("group by \r\n");
            sql.append("item_entity_attr.item_id) fi_cost on fi_cost.item_id = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9) ");
            //done emeryd
            sql.append("join vendor on item_entity_attr.vendor_id = vendor.vendor_id ");
            sql.append("join item_disp on ejd_item_warehouse.disp_id = item_disp.disp_id ");
            sql.append("join emery_dept on ejd_item.dept_id = emery_dept.dept_id ");
            
            if(( m_Dept != null) &&( m_Dept.length() > 0 ))
               sql.append(String.format(" and dept_num = %s ", m_Dept) );
            
            sql.append("join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id  ");
            sql.append("left outer join (select item_id, avail_qty as available_qty from ejd_item_warehouse ");
            sql.append("join item_entity_attr on item_entity_attr.ejd_item_id = ejd_item_warehouse.ejd_item_id ");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0){
               if (m_Warehouse.equals("01"))
                 sql.append(" where warehouse_id = 1 ");
               if (m_Warehouse.equals("04"))
                 sql.append(" where warehouse_id = 2 ");
             }
            sql.append(") ");
            sql.append(" qty on item_entity_attr.item_id = qty.item_id and item_entity_attr.item_type_id not in (8,9) ");
            
            //
            // Pull the sku_master and location information from fascor for now.  It seems to work ok.
            // Dean needs all locations and Fred doesn't want repeating groups.  The wm_concat will
            // give us all the locations in a comma separated field.
            sql.append("left outer join sku_master sku1 on sku1.sku = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9) and sku1.warehouse = 'PORTLAND' ");
            sql.append("left outer join sku_master sku4 on sku4.sku = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9) and sku4.warehouse = 'PITTSTON' ");
            sql.append("left outer join ( ");
            sql.append("   select sku, string_agg(distinct location.\"loc_id\", ',') as locs ");
            sql.append("   from loc_allocation ");																	
            sql.append("   join location on location.\"loc_id\" = loc_allocation.\"loc_id\" ");
            sql.append("   group by sku ");
            sql.append(" ) faslocs on faslocs.sku = item_entity_attr.item_id and item_entity_attr.item_type_id not in (8,9)   ");
            sql.append(" where item_entity_attr.item_id in ");
            sql.append(" (select item_id from ejd_item_warehouse join item_entity_attr using (ejd_item_id) where avail_qty > 0 ");
            sql.append(" minus ");
            sql.append(" (select distinct item_id from sale_item  ");
            sql.append(" join cust_warehouse cw on cw.customer_id = sale_item.customer_id ");
            sql.append(" where invoice_date >= to_date(?,'mm/dd/yyyy') and invoice_date <= to_date(?,'mm/dd/yyyy') ");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0){
               if (m_Warehouse.equals("01"))
                 sql.append(" and cw.warehouse_id = 1 ");
               if (m_Warehouse.equals("04"))
                 sql.append(" and cw.warehouse_id = 2 ");
             }
                        
            sql.append(" and sale_type in ('WAREHOUSE','TOOLREPAIR','APG WHS') ");
            sql.append(" and cw.customer_id not in ('199796','037940'))) ");
            if ( m_Filtered && (m_FilterDate != null && m_FilterDate.length() > 0) ) {
               //this side always has a where clause
               sql.append(" and ejd_item.setup_date <= to_date(?, 'mm/dd/yyyy') \r\n");
            }            
            
            sql.append("order by item_id");          
            m_SaleData = m_EdbConn.prepareStatement(sql.toString());
                                    
            sqlDCName.setLength(0);
            sqlDCName.append("select name from warehouse where fas_facility_id = ?");
            m_GetDCName = m_EdbConn.prepareStatement(sqlDCName.toString());
            
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
         log.error("InventoryAging.prepareStatements - null edb connection");
      
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
                  
         if ( param.name.equals("begdate") )
            m_BegDate = param.value;
         
         if ( param.name.equals("enddate") )
            m_EndDate = param.value;
         
         if ( param.name.equals("filtered") )
            m_Filtered = param.value.equalsIgnoreCase("true");
         
         if ( param.name.equals("filterdate") )
            m_FilterDate = param.value;
         
         if ( param.name.equals("dc") )
            m_Warehouse = param.value;
         
         if ( param.name.equalsIgnoreCase("dept") )
            m_Dept = param.value;
      }
      
      fileName.append("invtaging");
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
   }
      
   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {      
      XSSFCellStyle styleTxtC = null;      // Text centered
      XSSFCellStyle styleTxtL = null;      // Text left justified
      XSSFCellStyle styleInt = null;       // Style with 0 decimals
      XSSFCellStyle styleDouble = null;    // numeric #,##0.00
      XSSFCellStyle styleMoney = null;     // Money ($#,##0.00_);[Red]($#,##0.00)
      XSSFDataFormat format = null;
      
      format = m_Wrkbk.createDataFormat();
            
      styleTxtL = m_Wrkbk.createCellStyle();
      styleTxtL.setAlignment(HorizontalAlignment.LEFT);
      
      styleTxtC = m_Wrkbk.createCellStyle();
      styleTxtC.setAlignment(HorizontalAlignment.CENTER);
      
      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);
      
      styleDouble = m_Wrkbk.createCellStyle();
      styleDouble.setAlignment(HorizontalAlignment.RIGHT);
      styleDouble.setDataFormat(format.getFormat("#,##0.000"));   
      
      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short)8);
      
      m_CellStyles = new XSSFCellStyle[] {
         styleTxtC,    // col 0 item
         styleTxtC,    // col 1 status
         styleTxtC,    // col 2 date
         styleTxtL,    // col 3 descripton
         styleTxtL,    // col 4 vendor name         
         styleTxtC,    // col 5 stock pack
         styleTxtC,    // col 6 broken case
         styleTxtC,    // col 7 buy mult
         styleTxtC,    // col 8 dept
         styleMoney,   // col 9 total sales
         styleInt,     // col 10 total lines shipped
         styleInt,     // col 11 tot units shipped
         styleInt,     // col 12 On hand qty
         styleMoney,   // col 13 Cost
         styleMoney,   // col 14 Avg Cost
         styleDouble,  // col 15 1-wos
         styleDouble,  // col 16 wos on hand
         styleMoney,   // col 17 wos 0
         styleMoney,   // col 18 wos 13
         styleMoney,   // col 19 wos 26
         styleMoney,   // col 20 wos 39
         styleMoney,   // col 21 wos 52
         styleInt,     // col 22 wos 52 / avg cost
         styleMoney,   // col 23 wos 156
         styleDouble,  // col 24 cube
         styleTxtL,    // col 25 home loc
         styleTxtC,    // col 26 conv pack1
         styleTxtC,    // col 27 velocity code
      };
      
      styleTxtC = null;
      styleTxtL = null;
      styleInt = null;
      styleDouble = null;
      styleMoney = null;
      format = null;
   }
}
