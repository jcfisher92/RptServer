/**
 * File: DailyBrief.java
 * Description: The daily brief report.  This class replaces the class that runs in the middle tier.
 *    Peggy Richter was the original author.
 *
 * @author Peggy Richter
 * @author Jeffrey Fisher
 *
 * Create Data: 04/06/2005
 * Last Update: $Id: DailyBrief.java,v 1.29 2013/12/23 15:08:01 jfisher Exp $
 *
 * History:
 *    $Log: DailyBrief.java,v $
 *    Revision 1.29  2013/12/23 15:08:01  jfisher
 *    Added new columns per Dan G.
 *
 *    Revision 1.28  2013/02/22 19:58:47  prichter
 *    Fixed a query with an ambiguous column reference.  Oracle upgrade.
 *
 *    Revision 1.27  2009/09/23 17:16:26  prichter
 *    Never out query fix.  It was picking up cuts from the other warehouse.
 *
 *    Revision 1.26  2009/08/24 15:24:43  prichter
 *    Fixed the calculation of m_yesterday, which is used to determine which sales date to use.
 *
 *    Revision 1.25  2009/08/12 15:50:25  prichter
 *    Fixed the never out query
 *
 *    Revision 1.24  2009/07/23 19:53:22  prichter
 *    Rewrote the never out query.  Moved never_out from item to item_warehouse.
 *
 *    Revision 1.23  2009/04/29 20:12:54  jfisher
 *    removed unused import
 *
 *    Revision 1.22  2009/04/29 18:41:01  smurdock
 *    added detail felds  units ordered, units shippied, lines ordered, lines shipped, dollars ordered, dollars shipped pct dollars ordered shipped
 *
 *    Revision 1.21  2009/03/26 20:37:03  prichter
 *    Changed the default never out fill rate from 0% to 100%.  If no never out items were shipped, the report displayed 0% which is misleading.  Also fixed a bug in the addRegion method.
 *
 *    Revision 1.20  2009/02/18 14:55:05  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    08/03/2005 - Merged previous change with the changes needed for the new report server.  pjr
 *
 *    07/05/2005 - Added total selected vendor service level. Mantis# 674
 *               - Removed separate Villager's entries.  pjr
 *
 *    03/25/2005 - Added log4j logging. jcf
 *
 *    03/23/2005 - Added the close() call for the m_OraCon var.  jcf
 *
 *    03/10/2005 - Added unit service levels.  Removed pro focus vendor section.  CR# 611.  pjr.
 *
 *    07/13/2004 - Added PRO Focus Vendors service level section. Mantis# 487  pjr
 *
 *    05/03/2004 - Changed the email notification so that the m_DistList member variable was not used.  This is causing
 *       problems because the list gets cleaned up before it's used in the email service. - jcf
 *
 *    04/07/2004 - Applied Email class changes. - jcf
 *
 *    03/05/2004 - Changed the # of ship days calc again to count all M-W as 1 day and Thur as .5 days.
 *               - Added MTD totals line. pjr
 *
 *    02/10/2004 - Changed the # of ship days calculation to not count Friday.  pjr
 *
 *    12/17/2003 - Changed the email distribution list to parse the email addresses out of the last param which is
 *       a semicolon seperated list of addresses.  This works with the change of going to xml and web services. - jcf
 *
 *    03/03/2003 - Changed to report the prior month if the current date is the first business day of the month
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.sql.Types;
import java.util.ArrayList;
import java.util.GregorianCalendar;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
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


public class DailyBrief extends Report
{
   //Parameters
   private java.sql.Date m_Date;
   private java.sql.Date m_Yesterday;
   private String m_RelLines;
   private String m_ShiftEnd;
   private String m_Warehouse;

   private double m_Sales;
   private double m_FillRate;
   private double m_UnitFillRate;   //pjr 03/09/2005
   private String m_MnthName;
   private int m_WhsBudget;
   private double m_Days = -1;
   private double m_MtdDays = -1;
   private double m_MtdSalesAmt;
   private double m_LineNeverOut = 0;
   private double m_DollarNeverOut = 0;
   private double m_UnitNeverOut = 0;   //pjr 03/09/2005
   private int m_WhsId;

   XSSFWorkbook m_Wrkbk = new XSSFWorkbook();
   XSSFSheet m_Sheet = m_Wrkbk.createSheet();
   private XSSFFont m_Font;
   private XSSFFont m_FontTitle;
   private XSSFFont m_FontRed;
   private XSSFFont m_FontBold;
   private XSSFFont m_FontData;
   private XSSFFont m_FontBoldData;
   private XSSFFont m_FontUl;

   private XSSFCellStyle m_StyleDolla;  // Formatted 0.00
   private XSSFCellStyle m_StyleText;  // Text right justified
   private XSSFCellStyle m_StyleBoldText;  //size 10 bold right-alligned
   private XSSFCellStyle m_StyleTitle; // Bold, centered
   private XSSFCellStyle m_StyleTtlWrap; // Bold, centered, word wrap
   private XSSFCellStyle m_StyleBold;  // Normal but bold
   private XSSFCellStyle m_StyleDec;   // Style with 2 decimals
   private XSSFCellStyle m_StyleInt;   // Style with 0 decimals
   private XSSFCellStyle m_StylePct;   // Style with 0 decimals + %
   private XSSFCellStyle m_StyleRed;   // Highlighted/red fields & 1 decimals
   private XSSFCellStyle m_StyleLabel; // Text labels, right justify, 8pt
   private XSSFCellStyle m_StyleUl;    // Underlined, numeric, 3 decimals

   private PreparedStatement m_ProFocusVendors = null;
   private PreparedStatement m_VendorSvcLvl = null;  //pjr 07/05/2005
   private PreparedStatement m_VipVendors = null;
   private PreparedStatement m_WhsSales = null;
   private PreparedStatement m_NeverOut = null;

   private CallableStatement m_Budgets = null;
   private CallableStatement m_MtdSales = null;
   private CallableStatement m_ShipDays = null;

   //private HSSFDataFormat df = null;

   /**
    * default constructor
    */
   public DailyBrief() {
      super();

      //
      // Create the default font for this workbook
      m_Font = m_Wrkbk.createFont();
      m_Font.setFontHeightInPoints((short) 8);
      m_Font.setFontName("Arial");

      //
      // Create a font for titles
      m_FontTitle = m_Wrkbk.createFont();
      m_FontTitle.setFontHeightInPoints((short)10);
      m_FontTitle.setFontName("Arial");
      m_FontTitle.setBold(true);

      //
      // Create a font for red fields
      m_FontRed = m_Wrkbk.createFont();
      m_FontRed.setFontHeightInPoints((short)10);
      m_FontRed.setFontName("Arial");
      m_FontRed.setBold(true);
      m_FontRed.setColor(IndexedColors.RED.index);

      //
      // Create a font that is normal size & bold
      m_FontBold = m_Wrkbk.createFont();
      m_FontBold.setFontHeightInPoints((short)8);
      m_FontBold.setFontName("Arial");
      m_FontBold.setBold(true);

      //
      // Create a font that is normal size & bold
      m_FontData = m_Wrkbk.createFont();
      m_FontData.setFontHeightInPoints((short)10);
      m_FontData.setFontName("Arial");

      //
      // Create a font that is size 10 & bold
      m_FontBoldData = m_Wrkbk.createFont();
      m_FontBoldData.setFontHeightInPoints((short)10);
      m_FontBoldData.setFontName("Arial");

      //
      // Create a font that is normal size & underlined
      m_FontUl = m_Wrkbk.createFont();
      m_FontUl.setFontHeightInPoints((short)10);
      m_FontUl.setFontName("Ariel");
      m_FontUl.setUnderline(FontUnderline.SINGLE_ACCOUNTING);

      //
      // Setup the cell styles used in this report
      m_StyleText = m_Wrkbk.createCellStyle();
      m_StyleText.setFont(m_FontData);
      m_StyleText.setAlignment(HorizontalAlignment.RIGHT);


      m_StyleDolla = m_Wrkbk.createCellStyle();
      m_StyleDolla.setFont(m_FontData);
      m_StyleDolla.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDolla.setDataFormat((short)0x2a); //charlie likes this format for currency, says dan guimond
      //m_StyleDolla.setDataFormat(df.getFormat("#,##0.00"));
      //m_StyleDolla.setDataFormat(df.getFormat("$* #,##0"));


      m_StyleBoldText = m_Wrkbk.createCellStyle();
      m_StyleBoldText.setFont(m_FontData);
      m_StyleBoldText.setAlignment(HorizontalAlignment.RIGHT);

      m_StyleLabel = m_Wrkbk.createCellStyle();
      m_StyleLabel.setFont(m_Font);
      m_StyleLabel.setAlignment(HorizontalAlignment.RIGHT);

      m_StyleTitle = m_Wrkbk.createCellStyle();
      m_StyleTitle.setFont(m_FontTitle);
      m_StyleTitle.setAlignment(HorizontalAlignment.RIGHT);

      m_StyleTtlWrap = m_Wrkbk.createCellStyle();
      m_StyleTtlWrap.setFont(m_FontTitle);
      m_StyleTtlWrap.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleTtlWrap.setWrapText(true);




      m_StyleRed = m_Wrkbk.createCellStyle();
      m_StyleRed.setFont(m_FontRed);
      m_StyleRed.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleRed.setDataFormat((short)4);

      m_StyleDec = m_Wrkbk.createCellStyle();
      m_StyleDec.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleDec.setDataFormat((short)4);

      m_StyleInt = m_Wrkbk.createCellStyle();
      m_StyleInt.setAlignment(HorizontalAlignment.RIGHT);
      m_StyleInt.setDataFormat((short)3);

      m_StylePct = m_Wrkbk.createCellStyle();
      m_StylePct.setAlignment(HorizontalAlignment.RIGHT);
      m_StylePct.setDataFormat((short)9);

      m_StyleBold = m_Wrkbk.createCellStyle();
      m_StyleBold.setFont(m_FontBold);
      m_StyleBold.setAlignment(HorizontalAlignment.RIGHT);

      m_StyleUl = m_Wrkbk.createCellStyle();
      m_StyleUl.setFont(m_FontUl);
      m_StyleUl.setDataFormat((short)3);
      m_StyleUl.setAlignment(HorizontalAlignment.RIGHT);

      m_Sheet.setColumnWidth(0, 5000);
      m_Sheet.setColumnWidth(1, 2500);
      m_Sheet.setColumnWidth(2, 2500);
      m_Sheet.setColumnWidth(3, 2500);
      m_Sheet.setColumnWidth(4, 2000);
      m_Sheet.setColumnWidth(5, 6500);
      m_Sheet.setColumnWidth(6, 2000);
      m_Sheet.setColumnWidth(7, 2000);
      m_Sheet.setColumnWidth(8, 2000);
      m_Sheet.setColumnWidth(9, 2000);
      m_Sheet.setColumnWidth(10, 2000);
      m_Sheet.setColumnWidth(11, 2000);
      m_Sheet.setColumnWidth(12, 2000);
      m_Sheet.setColumnWidth(13, 2000);
      m_Sheet.setColumnWidth(14, 2600);
      m_Sheet.setColumnWidth(15, 2600);
      m_Sheet.setColumnWidth(16, 2600);
   }

   /**
    * Convenience method that adds a completely bordered cell with the specified alignment.
    *
    * @param rowNum - the row index.
    * @param colNum - the column index.
    * @return HSSFCell - the newly added cell, or a reference to the existing one.
    */
   private XSSFCell addBordCell(int rowNum, int colNum)
   {
      return addCell(rowNum, colNum);
   }

   /**
    * Convenience method that adds a new String type cell with no borders and the specified alignment.
    *
    * @param rowNum int - the row index.
    * @param colNum short - the column index.
    * @param val String - the cell value.
    *
    * @return HSSFCell - the newly added String type cell, or a reference to the existing one.
    */
   private XSSFCell addCell(int rowNum, int colNum, String val, XSSFCellStyle style)
   {
      XSSFCell cell = addCell(rowNum, colNum);

      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(val));
      cell.setCellStyle(style);

      return cell;
   }

   /**
    * Convenience method that adds a new numeric type cell with no borders and the specified alignment.
    *
    * @param rowNum - the row index.
    * @param colNum short - the column index.
    * @param val double - the cell value.
    * @param style HSSFCellStyle - the cell style and format
    *
    * @return HSSFCell - the newly added numeric type cell, or a reference to the existing one.
    */
   private XSSFCell addCell(int rowNum, int colNum, double val, XSSFCellStyle style)
   {
      XSSFCell cell = addCell(rowNum, colNum);

      cell.setCellType(CellType.NUMERIC);
      cell.setCellStyle(style);
      cell.setCellValue(val);

      return cell;
   }

   /**
    * Adds a new cell with the specified borders and horizontal alignment.
    *
    * @param rowNum - the row index.
    * @param colNum short - the column index.
    *
    * @return HSSFCell - the newly added cell, or a reference to the existing one.
    */
   private XSSFCell addCell(int rowNum, int colNum)
   {
      XSSFRow row = addRow(rowNum);
      XSSFCell cell = row.getCell(colNum);

      if ( cell == null )
         cell = row.createCell(colNum);

      row = null;

      return cell;
   }

   /**
    * Convenience method that adds a merged region with the given value and border style.  This method assumes
    * that the region will only have 1 row, and 1 or more cells.  Also assumes a string type value.
    *
    * @param rowNum - start from this row.
    * @param colFrom short - start from this column.
    * @param colTo short - to this column.
    * @param value String - region cell value.
    * @param merge boolean - if true then don't merge it.  This is here because of POI bug#16362. See bugzilla.
    *
    * Note:  The bug reference in the merge parameter is only an issue if the number of merged cells exceeds 1000.
    *        This isn't an issue in the
    */
   private void addRegion(int rowNum, int colFrom, int colTo, String value, boolean merge, XSSFCellStyle style)
   {
      XSSFCell cell = null;

      addRow(rowNum);

      cell = addBordCell(rowNum, colFrom);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(value));
      cell.setCellStyle(style);

      if ( merge ) {
         CellRangeAddress region = new CellRangeAddress(rowNum, rowNum, colFrom, colTo);
         m_Sheet.addMergedRegion(region);
      }
   }

   /**
    * Adds a new row or returns the existing one.
    *
    * @param rowNum int - the row index.
    * @return HSSFRow - the row object added, or a reference to the existing one.
    */
   private XSSFRow addRow(int rowNum)
   {
      XSSFRow row = m_Sheet.getRow(rowNum);

      if ( row == null )
         row = m_Sheet.createRow(rowNum);

      return row;
   }

   /**
    * Builds the output file
    * @return true if the file was created, false if not.
    * @throws FileNotFoundException
    */
   public boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream OutFile = null;
      boolean result = true;
      short rownum = 0;
      double whsSales = -1;
      double whsFillRate = -1;
      double whsUnitFillRate = -1; //pjr 03/09/2005
      double totSales = -1;
      XSSFCell cell = null;
      double diff;
      double pct;

      StringBuffer fileName = new StringBuffer();
      String tmp = Long.toString(System.currentTimeMillis());
      fileName.append(m_Warehouse);
      fileName.append("-dailybrief");
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
      OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         m_CurAction = "building output file";
         m_WhsId = getWhsId();

         getShippingDays();
         getWhsSales();
         totSales = m_Sales;
         whsFillRate = m_FillRate;
         whsUnitFillRate = m_UnitFillRate; //pjr 03/09/2005

         whsSales = totSales;
         getBudget();
         getMtdSales();
         neverOutFillRates();

         m_Sheet.setFitToPage(true);

         //Create the title as a single cell that spans the width of the page
         String dateStr = m_Date.toString();
         dateStr =  dateStr.substring(5,7) + "-" +  dateStr.substring(8) + "-" + dateStr.substring(0,4);
         addRegion(0, 0, 7, m_Warehouse + " Daily Brief " + dateStr, true, m_StyleTitle);

         addRegion(2, 0, 1, "Today", true, m_StyleTitle);

         addCell(3, 0, "# of Lines for the day", m_StyleLabel);
         addCell(3, 1, m_RelLines, m_StyleText);

         dateStr = m_Yesterday.toString();
         dateStr =  dateStr.substring(5,7) + "-" +  dateStr.substring(8) + "-" + dateStr.substring(0,4);
         addRegion(5, 0, 4, "Results for " + dateStr, true, m_StyleTitle);

         addCell(6, 0, "Warehouse sales", m_StyleLabel);
         addCell(6, 1, whsSales, m_StyleInt);

         addCell(7, 0, "Time shift ended", m_StyleLabel);
         addCell(7, 1, m_ShiftEnd, m_StyleText);

         addCell(9, 1, "Line", m_StyleTitle); //pjr 03/09/2005
         addCell(9, 2, "Unit", m_StyleTitle); //pjr 03/09/2005

         addCell(10, 0, "Fill rate", m_StyleLabel);
         if ( totSales > 0 ) {
            if ( whsFillRate < 97 )
               addCell(10, 1, svcLvlFmt(whsFillRate), m_StyleRed);
            else
               addCell(10, 1, svcLvlFmt(whsFillRate), m_StyleText);

            if ( whsUnitFillRate < 97 ) //pjr 03/09/2005
               addCell(10, 2, svcLvlFmt(whsUnitFillRate), m_StyleRed);
            else
               addCell(10, 2, svcLvlFmt(whsUnitFillRate), m_StyleText);
         }

         //this row comes after row 10 because getShippingDays() sets m_MnthName
         addCell(13, 0, "Shipping Days in " + m_MnthName.trim(), m_StyleLabel);
         addCell(13, 1, m_Days, m_StyleText);
         addCell(15, 0, "Warehouse Sales Budget", m_StyleLabel);
         addCell(15, 1, m_WhsBudget, m_StyleInt);

         addCell(6, 5, "Select Vendor Fill Rates", m_StyleTitle);
         addCell(6, 6, "Line Rate", m_StyleTtlWrap);
         addCell(6, 7, "Unit Rate", m_StyleTtlWrap);
         addCell(6, 8, "Total Rate", m_StyleTtlWrap);
         addCell(6, 9, "Lines Ord", m_StyleTtlWrap);
         addCell(6, 10, "Lines Ship", m_StyleTtlWrap);
         addCell(6, 11, "Lines Cut", m_StyleTtlWrap);
         addCell(6, 12, "Units Ord", m_StyleTtlWrap);
         addCell(6, 13, "Units Ship", m_StyleTtlWrap);
         addCell(6, 14, "Dollars Ord", m_StyleTtlWrap);
         addCell(6, 15, "Dollars Ship", m_StyleTtlWrap);
         addCell(6, 16, "Dollars Lost", m_StyleTtlWrap);
         rownum = displayVendors();
         rownum++;

         cell = addCell(rownum, 1, "Budget", m_StyleLabel);
         cell.setCellStyle(m_StyleBold);
         cell = addCell(rownum, 2, "Actual", m_StyleLabel);
         cell.setCellStyle(m_StyleBold);
         cell = addCell(rownum, 3, "Variance", m_StyleLabel);
         cell.setCellStyle(m_StyleBold);
         rownum++;

         addCell(rownum, 0, "Per Day Average Sales", m_StyleLabel);
         addCell(rownum, 1, m_WhsBudget / m_Days, m_StyleInt);
         addCell(rownum, 2, m_MtdSalesAmt / m_MtdDays, m_StyleInt);
         diff = (m_MtdSalesAmt / m_MtdDays) - (m_WhsBudget / m_Days);
         addCell(rownum, 3, diff, m_StyleInt);
         if ( m_WhsBudget != 0 ) {
            pct = diff / ((m_WhsBudget ) / m_Days);
            addCell(rownum, 4, pct, m_StylePct);
         }
         rownum++;

         addCell(rownum, 0, "Month to Date Sales", m_StyleLabel);
         addCell(rownum, 1, m_WhsBudget / m_Days * m_MtdDays, m_StyleInt);
         addCell(rownum, 2, m_MtdSalesAmt, m_StyleInt);
         diff = m_MtdSalesAmt - (m_WhsBudget / m_Days * m_MtdDays);
         addCell(rownum, 3, diff, m_StyleInt);
         if ( m_WhsBudget != 0 ) {
            pct = diff / (m_WhsBudget / m_Days * m_MtdDays);
            addCell(rownum, 4, pct, m_StylePct);
         }
         else
            pct = 0;

         rownum++;
         rownum++;

         addRegion(rownum++, 0, 1, "NEVER OUT FILL RATE", true, m_StyleTitle);

         addCell(rownum, 0, "Line %", m_StyleLabel);
         if ( m_LineNeverOut < 100 )
          addCell(rownum++, 1, svcLvlFmt(m_LineNeverOut), m_StyleRed);
 //        boop2 = svcLvlFmt(m_LineNeverOut);
 //        addCell(rownum++, 1, boop2, m_StyleRed);


         else
            addCell(rownum++, 1, svcLvlFmt(m_LineNeverOut), m_StyleText);

         addCell(rownum, 0, "Dollar %", m_StyleLabel);
         if ( m_DollarNeverOut < 100 )
          addCell(rownum++, 1, svcLvlFmt(m_DollarNeverOut), m_StyleRed);
 //         boop2 = svcLvlFmt(m_DollarNeverOut);
 //         addCell(rownum++, 1, boop2, m_StyleRed);

         else
            addCell(rownum++, 1, svcLvlFmt(m_DollarNeverOut), m_StyleText);

         addCell(rownum, 0, "Unit %", m_StyleLabel);
         if ( m_UnitNeverOut < 100 )
            addCell(rownum++, 1, svcLvlFmt(m_UnitNeverOut), m_StyleRed);
         else
            addCell(rownum++, 1, svcLvlFmt(m_UnitNeverOut), m_StyleText);

         m_Wrkbk.write(OutFile);
      }

      catch( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         result = false;
      }


      return result;
   }

   /**
    * Displays a column of vendor names and their fill rates (service level) for the prior day.
    * @return short
    */
   private short displayVendors() throws Exception
   {
      ResultSet rs = null;
      short rownum = 7;
      XSSFCell cell = null;
      double dollaship;
      double dollaord;
      double dollarate;

      try {
         m_VipVendors.setString(1, m_Warehouse);
         m_VipVendors.setDate(2, m_Yesterday);
         m_VipVendors.setDate(3, m_Yesterday);
         rs = m_VipVendors.executeQuery();

         while ( rs.next() ) {
            dollaship = rs.getDouble("dolla_ship");
            dollaord = rs.getDouble("dolla_ord");

            if (dollaord == 0){  //dealing with divide by zero and what we display as a fill rate for zero values
               dollarate = 100;  //(believe me, it's easier here than in the query)
            }
            else // dollaord is not zero
               dollarate = (dollaship/dollaord) * 100;

            addCell(rownum, 5, rs.getString("vendor_name"), m_StyleLabel);
            cell = addCell(rownum, 6, svcLvlFmt(rs.getDouble("fillrate")), m_StyleText);

            if ( rs.getDouble("fillrate") > 0 && rs.getDouble("fillrate") < 98 )
               cell.setCellStyle(m_StyleRed);

            if ((rs.getDouble("unit_ord") == 0) && (rs.getDouble("unit_ship") == 0))
                cell = addCell(rownum, 7, "100.0",m_StyleText);  // no units ordered or shipped = 100%, said somebody
            else
                cell = addCell(rownum, 7, svcLvlFmt(rs.getDouble("unitfillrate")), m_StyleText);

            if ( rs.getDouble("unitfillrate") > 0 && rs.getDouble("unitfillrate") < 98 )
               cell.setCellStyle(m_StyleRed);

            cell = addCell(rownum, 8, svcLvlFmt(dollarate), m_StyleText);
            if (dollarate > 0 && dollarate < 98 )
               cell.setCellStyle(m_StyleRed);

            cell = addCell(rownum, 9, rs.getDouble("line_ord"), m_StyleText);
            cell = addCell(rownum, 10, rs.getDouble("line_ship"), m_StyleText);
            cell = addCell(rownum, 11, rs.getDouble("line_cut"), m_StyleText);
            cell = addCell(rownum, 12, rs.getDouble("unit_ord"), m_StyleText);
            cell = addCell(rownum, 13, rs.getDouble("unit_ship"), m_StyleText);
            cell = addCell(rownum, 14, rs.getDouble("dolla_ord"), m_StyleDolla);
            cell = addCell(rownum, 15, rs.getDouble("dolla_ship"), m_StyleDolla);
            cell = addCell(rownum++, 16, rs.getDouble("dollalost"), m_StyleDolla);
         }
      }

      catch ( Exception e ) {
         log.error( e );
      }

      finally {
      	DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      //
      // Add the total service level for select vendors.  pjr 07/05/2005
      try {
         m_VendorSvcLvl.setString(1, m_Warehouse);
         m_VendorSvcLvl.setDate(2, m_Yesterday);
         m_VendorSvcLvl.setDate(3, m_Yesterday);
         rs = m_VendorSvcLvl.executeQuery();

         while ( rs.next() ) {
            dollaship = rs.getDouble("dolla_ship");
            dollaord = rs.getDouble("dolla_ord");
            if (dollaord == 0) {  //dealing with divide by zero and what we display as a fill rate for zero values
                  dollarate = 100.0;
            } //end of divide by zero stuff
            else  // dollaord is not sero
               dollarate = (dollaship/dollaord) * 100;

            addRegion(rownum, 4, 5, "Service Level of Select Vendors", true, m_StyleTitle);
            cell = addCell(rownum, 6, svcLvlFmt(rs.getDouble("fillrate")), m_StyleBoldText);

            if ( rs.getDouble("fillrate") > 0 && rs.getDouble("fillrate") < 98 )
               cell.setCellStyle(m_StyleRed);

            if ((rs.getDouble("unit_ord") == 0) && (rs.getDouble("unit_ship") == 0))
               cell = addCell(rownum, 7, "100.0",m_StyleText);  // no units ordered or shipped = 100%, said somebody
            else
               cell = addCell(rownum, 7, svcLvlFmt(rs.getDouble("unitfillrate")), m_StyleBoldText);

            if ( rs.getDouble("unitfillrate") > 0 && rs.getDouble("unitfillrate") < 98 )
               cell.setCellStyle(m_StyleRed);

            cell = addCell(rownum, 8, svcLvlFmt(dollarate), m_StyleBoldText);
            if ( dollarate > 0 && dollarate < 98 )
               cell.setCellStyle(m_StyleRed);

            cell = addCell(rownum, 9, rs.getDouble("line_ord"), m_StyleBoldText);
            cell = addCell(rownum, 10, rs.getDouble("line_ship"), m_StyleBoldText);
            cell = addCell(rownum, 11, rs.getDouble("line_cut"), m_StyleBoldText);
            cell = addCell(rownum, 12, rs.getDouble("unit_ord"), m_StyleText);
            cell = addCell(rownum, 13, rs.getDouble("unit_ship"), m_StyleText);
            cell = addCell(rownum, 14, rs.getDouble("dolla_ord"), m_StyleDolla);
            cell = addCell(rownum, 15, rs.getDouble("dolla_ship"), m_StyleDolla);
            cell = addCell(rownum, 16, rs.getDouble("dollalost"), m_StyleDolla);
         }
      }

      catch ( Exception e ) {
         log.error( "Exception", e );
      }

      finally {
      	DbUtils.closeDbConn(null, null, rs);
         rs = null;
      }

      return rownum;
   }


   /**
    * Returns the warehouse budget for the current month
    */
   private void getBudget()
   {
   	m_WhsBudget = 0;

      try {
         m_CurAction = "getting budget";
         m_Budgets.setDate(1, m_Yesterday);
         m_Budgets.setInt(2, m_WhsId);
         m_Budgets.execute();
         m_WhsBudget = m_Budgets.getInt(3);
      }
      catch ( Exception e ) {
         log.error( "exception", e );
      }
   }

   /**
    * Returns the month to date warehouse sales
    */
   private void getMtdSales()
   {
      try {
         m_CurAction = "getting mtd sales";
         m_MtdSales.setDate(1, m_Yesterday);
         m_MtdSales.setString(2, m_Warehouse);
         m_MtdSales.execute();
         m_MtdSalesAmt = m_MtdSales.getDouble(3);
      }
      catch ( Exception e ) {
         log.error( "exception", e );
      }
   }

   /**
    * PLSQL block that calculates the number of shipping days for the month and then number of shipping days
    * month to date
    */
   private void getShippingDays()
   {

      try {
         m_CurAction = "getting shipping days";
         m_ShipDays.setDate(1, m_Date);
         m_ShipDays.execute();
         m_Days = m_ShipDays.getDouble(2);
         m_MnthName = m_ShipDays.getString(3);
         m_MtdDays = m_ShipDays.getDouble(4);
         m_Yesterday = m_ShipDays.getDate(5);
      }
      catch ( Exception e ) {
         log.error( "exception", e );
      }
   }

   /**
    * Returns the warehouse id using the warehouse name passed as a parameter
    * @return int - the warehouse id
    */
   private int getWhsId()
   {
   	Statement stmt = null;
   	ResultSet rs = null;

   	try {
   		stmt = m_OraConn.createStatement();
   		rs = stmt.executeQuery("select warehouse_id from warehouse where name = '" + m_Warehouse + "'");

   		if ( rs.next() )
   			return rs.getInt("warehouse_id");
   	}

   	catch ( Exception e ) {
   		log.error("Exception", e);
   	}

   	finally {
   		DbUtils.closeDbConn(null, stmt, rs);
   		rs = null;
   		stmt = null;
   	}

   	return -1;
   }

   /**
    * Returns the total warehouse sales for the prior day
    */
    private void getWhsSales()
    {
       ResultSet rs = null;

       m_Sales = -1;
       m_FillRate = -1;
       m_CurAction = "getting whs sales";

       try {
          m_WhsSales.setDate(1, m_Yesterday);
          m_WhsSales.setString(2, m_Warehouse);
          rs = m_WhsSales.executeQuery();

          while ( rs.next() ) {
             m_Sales = rs.getDouble("sales");
             m_FillRate = rs.getDouble("fillrate");
             m_UnitFillRate = rs.getDouble("unitfillrate");
          }
       }

       catch ( Exception e ) {
          log.error( "exception", e );
       }

       finally {
          DbUtils.closeDbConn(null, null, rs);
          rs = null;
       }
    }

   /**
    * Closes prepared statements and cleans up member variables
    */
   protected void cleanup()
   {
      closeStatement(m_Budgets);
      m_Budgets = null;

      closeStatement(m_MtdSales);
      m_MtdSales = null;

      closeStatement(m_NeverOut);
      m_NeverOut = null;

      closeStatement(m_ShipDays);
      m_ShipDays = null;

      closeStatement(m_ProFocusVendors);
      m_ProFocusVendors = null;

      closeStatement(m_VendorSvcLvl);
      m_VendorSvcLvl = null;

      closeStatement(m_VipVendors);
      m_VipVendors = null;

      closeStatement(m_WhsSales);
      m_WhsSales = null;

      m_Sheet = null;
      m_Wrkbk = null;
      m_Font = null;
      m_FontTitle = null;
      m_FontRed = null;
      m_FontBold = null;
      m_FontData = null;
      m_StyleText = null;
      m_StyleTitle = null;
      m_StyleBold = null;
      m_StyleDec = null;
      m_StyleInt = null;
      m_StylePct = null;
      m_StyleRed = null;
      m_StyleLabel = null;
   }

   /**
    * Closes a single prepared statement
    * @param stmt
    */
   private void closeStatement(PreparedStatement stmt)
   {
      if ( stmt != null ) {
         try {
            stmt.close();
         }
         catch ( Exception e ) {
            ;
         }
      }
   }

   /**
    * Creates the report.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;

      try {
         m_OraConn = m_RptProc.getOraConn();
         prepareStatements();
         created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("exception:", ex);
      }

      finally {
        cleanup();

        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Calculates service levels for never out items
    */
   private void neverOutFillRates()
   {
   	ResultSet rs = null;

      m_CurAction = "getting fill rates";

      try {
         m_NeverOut.setDate(1, m_Yesterday);
         m_NeverOut.setDate(2, m_Yesterday);
         m_NeverOut.setDate(3, m_Yesterday);
         m_NeverOut.setString(4, m_Warehouse);

         rs = m_NeverOut.executeQuery();

         if ( rs.next() ) {
         	m_LineNeverOut = rs.getDouble("line_svclvl");
         	m_DollarNeverOut = rs.getDouble("dollar_svclvl");
         	m_UnitNeverOut = rs.getDouble("unit_svclvl");
         }
      }

      catch ( Exception e ) {
         log.error( "exception", e );
      }

      finally {
      	DbUtils.closeDbConn(null, null, rs);
      	rs = null;
      }
   }

   /**
    * Prepares the sql statements
    * @throws Exception
    */
   private void prepareStatements() throws Exception
   {
      StringBuffer sql = new StringBuffer();
      m_CurAction = "preparing statements";

      if ( m_OraConn != null ) {
         sql.setLength(0);
         sql.append("declare ");
         sql.append("   rundate    date := ?; ");
         sql.append("   whsid      integer := ?; ");
         sql.append("   yr         varchar2(4); ");
         sql.append("   mnth       varchar2(2); ");
         sql.append("   whsbudget  number := 0; ");
         sql.append("begin ");
         sql.append("   select to_char(rundate, 'yyyy') into yr from dual; ");
         sql.append("   select to_char(rundate, 'mm') into mnth from dual; ");
         sql.append("   if length(mnth) = 1 then mnth := '0' || mnth; end if; ");
         sql.append("   select sum(budget_amount) into whsbudget ");
         sql.append("   from salesyear, cust_warehouse ");
         sql.append("   where sales_year = yr and ");
         sql.append("         sales_month = mnth and ");
         sql.append("         sale_type = 'WAREHOUSE' and ");
         sql.append("         cust_warehouse.customer_id = salesyear.cust_nbr and ");
         sql.append("         cust_warehouse.warehouse_id = whsid; ");
         sql.append("   ? := whsbudget; ");
         sql.append("end; ");
         m_Budgets = m_OraConn.prepareCall(sql.toString());
         m_Budgets.registerOutParameter(3, Types.INTEGER);

         sql.setLength(0);
         sql.append("declare ");
         sql.append("   rundate    date := ?; ");
         sql.append("   whsname    varchar2(40) := ?; ");
         sql.append("   begdate    date; ");
         sql.append("   mtdsales   number; ");
         sql.append("   vsales     number; ");
         sql.append("begin ");
         sql.append("   begdate := last_day(add_months(rundate, -1)) + 1; ");
         sql.append("   select sum(dollars_shipped) into mtdsales ");
         sql.append("   from sale ");
         sql.append("   where invoice_date >= begdate and ");
         sql.append("         invoice_date <= rundate and ");
         sql.append("         warehouse = whsname and ");
         sql.append("         sale_type = 'WAREHOUSE'; ");
         sql.append("   ? := mtdsales; ");
         sql.append("end; ");
         m_MtdSales = m_OraConn.prepareCall(sql.toString());
         m_MtdSales.registerOutParameter(3, Types.DOUBLE);

         sql.setLength(0);
         sql.append("select lines.svclvl line_svclvl, units.svclvl unit_svclvl, dollars.svclvl dollar_svclvl ");
         sql.append("from warehouse ");
         sql.append("join ( ");
         sql.append("   select inv_dtl.warehouse, ");
         sql.append("          round(decode(count(*), 0, 100, sum(decode(qty_shipped, 0, 0, 1)) / count(*)) * 100, 3) svclvl ");
         sql.append("   from inv_dtl ");
         sql.append("   join item_warehouse on item_warehouse.item_id = inv_dtl.item_nbr and ");
         sql.append("                          item_warehouse.never_out = 1 ");
         sql.append("   join warehouse on warehouse.warehouse_id = item_warehouse.warehouse_id ");
         sql.append("   where inv_dtl.invoice_date = ? and ");
         sql.append("         inv_dtl.sale_type = 'WAREHOUSE' and ");
         sql.append("         inv_dtl.tran_type = 'SALE' and  ");
         sql.append("         inv_dtl.qty_ordered > 0 and ");
         sql.append("         inv_dtl.warehouse = warehouse.name ");
         sql.append("   group by warehouse   ");
         sql.append(") lines on lines.warehouse = warehouse.name ");
         sql.append("join ( ");
         sql.append("   select inv_dtl.warehouse, ");
         sql.append("          round(decode(sum(qty_ordered), 0, 100, sum(qty_shipped) / sum(qty_ordered)) * 100, 3) svclvl ");
         sql.append("   from inv_dtl ");
         sql.append("   join item_warehouse on item_warehouse.item_id = inv_dtl.item_nbr and ");
         sql.append("                          item_warehouse.never_out = 1 ");
         sql.append("   join warehouse on warehouse.warehouse_id = item_warehouse.warehouse_id ");
         sql.append("   where inv_dtl.invoice_date = ? and ");
         sql.append("         inv_dtl.sale_type = 'WAREHOUSE' and ");
         sql.append("         inv_dtl.tran_type = 'SALE' and  ");
         sql.append("         inv_dtl.qty_ordered > 0 and ");
         sql.append("         inv_dtl.warehouse = warehouse.name ");
         sql.append("   group by warehouse   ");
         sql.append(") units on units.warehouse = warehouse.name ");
         sql.append("join ( ");
         sql.append("   select inv_dtl.warehouse, ");
         sql.append("          round(decode(sum(qty_ordered * unit_sell), 0, 100, sum(qty_shipped * unit_sell) / sum(qty_ordered * unit_sell)) * 100, 3) svclvl ");
         sql.append("   from inv_dtl ");
         sql.append("   join item_warehouse on item_warehouse.item_id = inv_dtl.item_nbr and ");
         sql.append("                          item_warehouse.never_out = 1 ");
         sql.append("   join warehouse on warehouse.warehouse_id = item_warehouse.warehouse_id ");
         sql.append("   where inv_dtl.invoice_date = ? and ");
         sql.append("         inv_dtl.sale_type = 'WAREHOUSE' and ");
         sql.append("         inv_dtl.tran_type = 'SALE' and  ");
         sql.append("         inv_dtl.qty_ordered > 0 and ");
         sql.append("         inv_dtl.warehouse = warehouse.name ");
         sql.append("   group by warehouse   ");
         sql.append(") dollars on dollars.warehouse = warehouse.name ");
         sql.append("where warehouse.name = ? ");
         m_NeverOut = m_OraConn.prepareStatement(sql.toString());

         sql.setLength(0);
         sql.append("      declare  ");
         sql.append("         rundate    date := ?; ");
         sql.append("         begdate    date; ");
         sql.append("         enddate    date; ");
         sql.append("         day        varchar2(1); ");
         sql.append("         cnt        number(5,2) := 0; ");
         sql.append("         month      varchar2(10); ");
         sql.append("         mtdcnt     number(5,2) := 0;  ");
         sql.append("         yesterday  date := rundate - 1; ");
         sql.append("      begin ");
         sql.append("         select max(load_date) into enddate from trip where load_date < rundate; ");
         sql.append("         enddate := last_day(enddate); ");
         sql.append("         begdate := add_months(enddate, -1) + 1; ");
         sql.append("         select to_char(enddate, 'Month') into month from dual; ");
         sql.append("         while begdate <= enddate loop ");
         sql.append("            select to_char(begdate, 'd') into day from dual; ");
         sql.append("            if day in ('2','3','4') then ");
         sql.append("               cnt := cnt + 1; ");
         sql.append("               if begdate < rundate then ");
         sql.append("                  mtdcnt := mtdcnt + 1; ");
         sql.append("               end if; ");
         sql.append("            elsif day = '5' then ");
         sql.append("               cnt := cnt + .5; ");
         sql.append("               if begdate < rundate then ");
         sql.append("                  mtdcnt := mtdcnt + .5; ");
         sql.append("               end if; ");
         sql.append("            end if; ");
         sql.append("            begdate := begdate + 1; ");
         sql.append("         end loop; ");
         sql.append("         select max(load_date) into yesterday from trip ");
         sql.append("         where load_date < rundate and ");
         sql.append("            exists(select * from trip_stop, shipment ");
         sql.append("                   where trip_stop.trip_id = trip.trip_id and ");
         sql.append("                         shipment.trip_stop_id = trip_stop.trip_stop_id); ");
         sql.append("         ? := cnt; ");
         sql.append("         ? := month; ");
         sql.append("         ? := mtdcnt; ");
         sql.append("         ? := yesterday; ");
         sql.append("      end; ");
         m_ShipDays = m_OraConn.prepareCall(sql.toString());
         m_ShipDays.registerOutParameter(2, Types.DOUBLE);
         m_ShipDays.registerOutParameter(3, Types.VARCHAR);
         m_ShipDays.registerOutParameter(4, Types.DOUBLE);
         m_ShipDays.registerOutParameter(5, Types.DATE);

         sql.setLength(0);
         sql.append("select sum(dollars_shipped) sales, ");
         sql.append("round(sum(lines_shipped) / sum(lines_ordered) * 100,1) fillrate, ");
         sql.append("round(sum(units_shipped) / sum(units_ordered) * 100,1) unitfillrate " );
         sql.append("from sale ");
         sql.append("where invoice_date = ? and ");
         sql.append("      sale_type = 'WAREHOUSE' and ");
         sql.append("      warehouse = ? ");
         m_WhsSales = m_OraConn.prepareStatement(sql.toString());

         // pjr 07/05/2005 Added total of select vendor service levels
         sql.setLength(0);
         sql.append("select round((sum(decode(si.qty_shipped, 0, 0, 1)) / count(*)) * 100, 1) fillrate,  ");
         sql.append("       round(sum(nvl(si.qty_shipped,0))/decode(sum(nvl(si.qty_ordered,0)),0,1,sum(si.qty_ordered)) * 100,1) unitfillrate, ");
         sql.append("       nvl(sum(si.qty_ordered),0) unit_ord, ");
         sql.append("       nvl(sum(si.qty_shipped),0) unit_ship,   ");
         sql.append("       count(si.item_id) line_ord, ");
         sql.append("       count(si.item_id) - count(zerolines.item_id) line_ship, ");
         sql.append("       count(zerolines.item_id) line_cut, ");
         sql.append("       nvl(sum(si.qty_ordered * si.unit_sell),0) dolla_ord, ");
         sql.append("       nvl(sum(si.qty_shipped * si.unit_sell),0) dolla_ship, ");
         sql.append("       nvl(round(sum(si.qty_ordered * si.unit_sell) - sum(si.qty_shipped * si.unit_sell),2),0) dollalost ");
         sql.append("from vendor_class ");
         sql.append("join vnd_class_value on vnd_class_value.vnd_class_val_id = vendor_class.vnd_class_val_id and ");
         sql.append("                        vnd_class_value.description = ? ");
         sql.append("join vnd_class_type on vnd_class_type.vnd_class_id = vnd_class_value.vnd_class_id and ");
         sql.append("                       vnd_class_type.description = 'DAILY BRIEF' ");
         sql.append("join vendor on vendor.vendor_id = vendor_class.vendor_id ");
         sql.append("left outer join sale_item si on si.vendor_id = vendor.vendor_id and ");
         sql.append("                       si.tran_type = 'SALE' and");
         sql.append("                       si.invoice_date = ? and ");
         sql.append("                       si.warehouse = vnd_class_value.description ");
         sql.append("left outer join sale_item zerolines on zerolines.sale_item_id = si.sale_item_id and ");
         sql.append("                       zerolines.tran_type = 'SALE' and ");
         sql.append("                       zerolines.invoice_date = ? and ");
         sql.append("                       zerolines.warehouse = vnd_class_value.description and ");
         sql.append("                       zerolines.qty_shipped = 0 ");
         m_VendorSvcLvl = m_OraConn.prepareStatement(sql.toString());

         // pjr 03/10/2005 - Use Daily Brief list instead of VIP list
         // 11/14 - If sum qty ord is zero show 100% fill rate, based on Peggy's recommendation
         sql.setLength(0);
         sql.append("select vendor.name vendor_name,  ");
         sql.append("       round((sum(decode(si.qty_shipped, 0, 0, 1)) / count(*)) * 100, 1) fillrate,  ");
         sql.append("       round(sum(nvl(si.qty_shipped,0))/decode(sum(nvl(si.qty_ordered,0)),0,1,sum(si.qty_ordered)) * 100,1) unitfillrate, ");
         sql.append("       nvl(sum(si.qty_ordered),0) unit_ord, ");
         sql.append("       nvl(sum(si.qty_shipped),0) unit_ship,   ");
         sql.append("       count(si.item_id) line_ord, ");
         sql.append("       count(si.item_id) - count(zerolines.item_id) line_ship, ");
         sql.append("       count(zerolines.item_id) line_cut, ");
         sql.append("       nvl(sum(si.qty_ordered * si.unit_sell),0) dolla_ord, ");
         sql.append("       nvl(sum(si.qty_shipped * si.unit_sell),0) dolla_ship, ");
         sql.append("       nvl(round(sum(si.qty_ordered * si.unit_sell) - sum(si.qty_shipped * si.unit_sell),2),0) dollalost ");
         sql.append("from vendor_class ");
         sql.append("join vnd_class_value on vnd_class_value.vnd_class_val_id = vendor_class.vnd_class_val_id and ");
         sql.append("                        vnd_class_value.description = ? ");
         sql.append("join vnd_class_type on vnd_class_type.vnd_class_id = vnd_class_value.vnd_class_id and ");
         sql.append("                       vnd_class_type.description = 'DAILY BRIEF' ");
         sql.append("join vendor on vendor.vendor_id = vendor_class.vendor_id ");
         sql.append("left outer join sale_item si on si.vendor_id = vendor.vendor_id and ");
         sql.append("                       si.tran_type = 'SALE' and");
         sql.append("                       si.invoice_date = ? and ");
         sql.append("                       si.warehouse = vnd_class_value.description ");
         sql.append("left outer join sale_item zerolines on zerolines.sale_item_id = si.sale_item_id and ");
         sql.append("                       zerolines.tran_type = 'SALE' and ");
         sql.append("                       zerolines.invoice_date = ? and ");
         sql.append("                       zerolines.warehouse = vnd_class_value.description and ");
         sql.append("                       zerolines.qty_shipped = 0 ");
         sql.append(" group by vendor.name ");
         sql.append("order by vendor.name ");
         m_VipVendors = m_OraConn.prepareStatement(sql.toString());
      }

      sql = null;
   }

   /**
    * Sets the parameters for the report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      String tmp;
      int yr;
      int mo;
      int day;

      //
      // param(0) = date
      tmp = params.get(0).value;
      yr = new Integer(tmp.substring(6)).intValue();
      mo = new Integer(tmp.substring(0,2)).intValue();
      day = new Integer(tmp.substring(3,5)).intValue();
      m_Date = new java.sql.Date(new GregorianCalendar(yr, mo - 1, day).getTimeInMillis());

      //
      //param(1) = Number of lines released
      m_RelLines = params.get(1).value;

      //
      // param(2) = Shift End
      m_ShiftEnd = params.get(2).value;

      //
      // param(3) = Warehouse Name
      m_Warehouse = params.get(3).value;
      m_Warehouse.toUpperCase();
   }

   /**
    * Returns a string that represents a number with 1 decimal  This is necessary because POI does not yet
    * have a DataFormat with a single decimal, it only offers 2 or none.
    * @param val - Decimal number
    * @return String - Character representation.
    */
   private String svcLvlFmt(double val)
   {
      String tmp = new Double(val).toString();
      int i;

      i = tmp.indexOf(".");
      if ( i == -1 )
         tmp = tmp + ".0";
      else {
         if ( tmp.length() > i + 1 )
            tmp = tmp.substring( 0, i + 2 );
      }
       return tmp;
   }
}
