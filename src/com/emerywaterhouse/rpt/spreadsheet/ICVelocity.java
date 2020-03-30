/**
 * Title:			ICVelocity.java
 * Description:	Item velocity data
 * Company:			Emery-Waterhouse
 * @author			prichter
 *
 * Create Date:	Jun 18, 2008
 * Last Update:   $Id: ICVelocity.java,v 1.10 2012/07/12 00:32:31 jfisher Exp $
 * <p>
 * History:
 *		$Log: ICVelocity.java,v $
 *		Revision 1.10  2012/07/12 00:32:31  jfisher
 *		Moved the in_catalog check to the item_warehouse.active line.  Removed test code and commented out sections.  Fixed warnings.
 *
 *		Revision 1.9  2012/05/05 06:07:28  pberggren
 *		Removed redundant loading of system properties.
 *
 *		Revision 1.8  2012/05/04 01:17:18  pberggren
 *		Added "SkuQty/" to web.service call
 *
 *		Revision 1.7  2012/05/03 07:55:10  prichter
 *		Fix to web service ip address
 *
 *		Revision 1.6  2012/05/03 04:41:04  pberggren
 *		Alter getOnHand Method for different variable to call .57
 *
 *		Revision 1.5  2012/05/03 04:37:17  pberggren
 *		Added server.properties call to force report to .57
 *
 *		Revision 1.4  2012/05/03 04:22:11  pberggren
 *		Added server.properties call to force report to .57
 *
 *		Revision 1.3  2010/10/10 18:48:51  smurdock
 *		left out the last two periods
 *
 *		Revision 1.2  2010/10/10 18:44:46  smurdock
 *		Report for sales over 12 week periods for I and C velocity codes
 *
 *		Revision 1.1  2010/10/07 17:11:49  smurdock
 *		committing unworking version to see if anybody else can get it to work
 *
 *		Revision 1.5  2009/02/24 22:03:51  smurdock
 *		lotsa user request updates.  service level by buyer, nbc upgrade, on order as net of ordered - received, promo id for all customers, QOH centered
 *
 *		Revision 1.4  2008/10/30 15:56:57  jfisher
 *		Fixed some warnings
 *
 *		Revision 1.3  2008/07/04 20:05:35  prichter
 *		Added repeating column headings option
 *
 *		Revision 1.2  2008/06/29 17:55:26  prichter
 *		Bug fixes from testing.  Items with no outstanding PO's were dropped from the report.  Changed 'INACTIVE' to 'INACTIVE ITEM' when filtering inactive items.  Added filter for cancelled lines.  Center PO related data and suppress repeats.
 *
 *		Revision 1.1  2008/06/28 16:51:05  prichter
 *		Initial add
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class ICVelocity extends Report 
{
	private PreparedStatement m_Item;

   private XSSFWorkbook m_WrkBk;
   private XSSFSheet m_Sheet;
   private XSSFRow m_Row = null;
   private Header m_Header;

   private XSSFFont m_Font;
   private XSSFFont m_FontTitle;
   private XSSFFont m_FontBold;
   private XSSFFont m_FontData;

   private XSSFCellStyle m_StyleText;  		// Text left justified
   private XSSFCellStyle m_StyleTextRight;  	// Text right justified
   private XSSFCellStyle m_StyleTextCenter; 	// Text centered
   private XSSFCellStyle m_StyleTitle; 		// Bold, centered
   private XSSFCellStyle m_StyleBold;  		// Normal but bold
   private XSSFCellStyle m_StyleBoldRight; 	// Normal but bold & right aligned
   private XSSFCellStyle m_StyleBoldCenter; 	// Normal but bold & centered
   private XSSFCellStyle m_StyleDec;   		// Style with 2 decimals
   private XSSFCellStyle m_StyleDecBold;		// Style with 2 decimals, bold
   private XSSFCellStyle m_StyleHeader; 		// Bold, centered 12pt
   private XSSFCellStyle m_StyleInt;   		// Style with 0 decimals

   // Parameters
   private String m_BegDate= null;
   private String m_EndDate= null;
   private String m_Warehouse = null;
   private String m_Period = null;
   private String m_VendorID = null;
   private String m_DeptNum = null;

   private short m_RowNum = 0;
   private ArrayList<String> m_FacilityList = new ArrayList<String>();
   private ArrayList<Integer> m_WhsList = new ArrayList<Integer>();

   /**
    * Builds the output file
    * @return boolean.  True if the file was created, false if not.
    * @throws FileNotFoundException
    */
   public boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream outFile = null;
      boolean result = true;
      ResultSet rs = null;
      int col;

      String lastPeriod = "begin";
      int  PeriodQtyShip = 0;
      int PeriodItemCnt = 0;
      double PeriodDollarShip = 0;
      int PeriodOnHand = 0;
      double PeriodDollarInv = 0;

      m_FileNames.add(m_RptProc.getUid() + "ICVelo" + getStartTime() + ".xlsx");
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      initReport();

      try {
         setCurAction( "Running the ICVelocity query" );
         rs = m_Item.executeQuery();

         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            setCurAction( "Processing " + rs.getString("period"));

            // Check for the start of a new period
            if ( !rs.getString("period").equals(lastPeriod) ) {
            	if ( !lastPeriod.equals("begin"))
            		PeriodTrailer(m_Period,PeriodItemCnt, PeriodQtyShip, PeriodDollarShip, PeriodOnHand,PeriodDollarInv);

               PeriodQtyShip = 0;
               PeriodItemCnt = 0;
               PeriodDollarShip = 0;
               PeriodOnHand = 0;
               PeriodDollarInv = 0;
            	PeriodHeader(rs.getString("Period"));
            }

            col = (short)1;
            m_Row = m_Sheet.createRow(m_RowNum++);
            createCell(m_Row, col++, rs.getString("item_id"), m_StyleText);
            createCell(m_Row, col++, rs.getString("setup_date"), m_StyleText);
            createCell(m_Row, col++, rs.getString("description"), m_StyleText);
            createCell(m_Row, col++, rs.getString("vendor_name"), m_StyleText);
            createCell(m_Row, col++, rs.getString("vendor_id"), m_StyleText);
            createCell(m_Row, col++, rs.getString("dept_num"), m_StyleText);
            createCell(m_Row, col++, rs.getString("stock_pack"), m_StyleText);
            createCell(m_Row, col++, rs.getString("nbc"), m_StyleText);
            createCell(m_Row, col++, rs.getString("bm"), m_StyleText);
            createCell(m_Row, col++, rs.getInt("shipqty"), m_StyleInt);
            createCell(m_Row, col++, rs.getInt("shipdoll"), m_StyleDec);
            createCell(m_Row, col++, rs.getDouble("onhand"), m_StyleInt);
            createCell(m_Row, col++, rs.getString("inv_dollars"), m_StyleDec);
            createCell(m_Row, col++, rs.getString("todays_buy"), m_StyleDec);
            createCell(m_Row, col++, rs.getString("average_cost"), m_StyleDec);
            createCell(m_Row, col++, ("  "), m_StyleText); // this is "Reviewed" eventually

            PeriodQtyShip += rs.getInt("shipqty");
            PeriodItemCnt++;
            PeriodDollarShip += rs.getDouble("shipdoll");
            PeriodOnHand += rs.getInt("onhand");
            PeriodDollarInv  += rs.getDouble("inv_dollars");

            lastPeriod = rs.getString("period");
         }
         PeriodTrailer(m_Period,PeriodItemCnt, PeriodQtyShip, PeriodDollarShip, PeriodOnHand,PeriodDollarInv);
         m_WrkBk.write(outFile);
         setCurAction( "Complete");
      }

      catch( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("The report had the following Error: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         result = false;
      }

      finally {
      	DbUtils.closeDbConn(null, null, rs);

         try {
            outFile.close();
            outFile = null;
         }
         catch ( Exception e ) {
            log.error( e );
         }
      }

      return result;
   }

   /**
    * Creates the buyer level header row.  Includes all column headings.
    * @param buyerNbr String - the buyer number
    * @param name String - the buyer name
    */
   private void PeriodHeader(String PHeader)
   {
      short col = 0;
      m_Row = m_Sheet.createRow(m_RowNum++);
      createCell(m_Row, col++, PHeader, m_StyleBold );
      createCell(m_Row, col++, "Item #", m_StyleBold );
      createCell(m_Row, col++, "Setup Date", m_StyleBold );
      createCell(m_Row, col++, "Description", m_StyleBold );
      createCell(m_Row, col++, "Vendor Name", m_StyleBold );
      createCell(m_Row, col++, "Vendor #", m_StyleBold );
      createCell(m_Row, col++, "Dept #", m_StyleBold );
      createCell(m_Row, col++, "Stock Pack", m_StyleBold );
      createCell(m_Row, col++, "NBC", m_StyleBold );
      createCell(m_Row, col++, "BM", m_StyleBold );
      createCell(m_Row, col++, "Units Shipped", m_StyleBold );
      createCell(m_Row, col++, "Dollars Shipped", m_StyleBold );
      createCell(m_Row, col++, "On Hand", m_StyleBold );
      createCell(m_Row, col++, "Inv Dollars", m_StyleBold );
      createCell(m_Row, col++, "Cost", m_StyleBold );
      createCell(m_Row, col++, "Avg Cost", m_StyleBold );
      createCell(m_Row, col++, "Reviewed", m_StyleBold );
   }

   /**
    * Resource cleanup
    */
   public void cleanup()
   {
   	DbUtils.closeDbConn(null, m_Item, null);
   	m_Item = null;
   	m_Header = null;
   	m_Font = null;
   	m_FontTitle = null;
   	m_FontBold = null;
   	m_FontData = null;
   	m_StyleText = null;
   	m_StyleTextRight = null;
   	m_StyleTextCenter = null;
   	m_StyleBold = null;
   	m_StyleBoldRight = null;
   	m_StyleBoldCenter = null;;
   	m_StyleDec = null;
   	m_StyleDecBold = null;
   	m_StyleHeader = null;
   	m_StyleInt = null;
   	m_BegDate = null;
   	m_EndDate = null;
   	m_Warehouse = null;
   	m_VendorID = null;
      m_DeptNum = null;

   	if ( m_FacilityList != null )
   		m_FacilityList.clear();

   	if ( m_WhsList != null )
   		m_WhsList.clear();

   	m_FacilityList = null;
   	m_WhsList = null;

   	m_WrkBk = null;
   	m_Sheet = null;
   	m_Row = null;
   }

   /**
    * Creates a cell of type numeric
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return XSSFCell newly created cell
    */
   private XSSFCell createCell(XSSFRow row, int col, double val, XSSFCellStyle style)
   {
      XSSFCell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellValue(val);
      cell.setCellStyle(style);

      return cell;
   }

   /**
    * Creates a cell of type numeric
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return XSSFCell newly created cell
    */
   private XSSFCell createCell(XSSFRow row, int col, int val, XSSFCellStyle style)
   {
      XSSFCell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellValue(val);
      cell.setCellStyle(style);

      return cell;
   }

   /**
    * Creates a cell of type String
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return XSSFCell newly created cell
    */
   private XSSFCell createCell(XSSFRow row, int col, String val, XSSFCellStyle style)
   {
      XSSFCell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(val));
      cell.setCellStyle(style);

      return cell;
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
         m_OraConn = m_RptProc.getOraConn();

         if ( prepareStatements() )
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
    * Creates the workbook and worksheet.  Creates any fonts and styles that
    * will be used.
    */
   private void initReport()
   {
      XSSFDataFormat df;
      short col = 0;

      try {
         m_WrkBk = new XSSFWorkbook();

         df = m_WrkBk.createDataFormat();

         //
         // Create the default font for this workbook
         m_Font = m_WrkBk.createFont();
         m_Font.setFontHeightInPoints((short) 8);
         m_Font.setFontName("Arial");

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
         // Create a font that is 12 pt & bold
         m_FontBold = m_WrkBk.createFont();
         m_FontBold.setFontHeightInPoints((short)8);
         m_FontBold.setFontName("Arial");
         m_FontBold.setBold(true);

         //
         // Setup the cell styles used in this report
         m_StyleText = m_WrkBk.createCellStyle();
         m_StyleText.setFont(m_FontData);
         m_StyleText.setAlignment(HorizontalAlignment.LEFT);

         m_StyleTextRight = m_WrkBk.createCellStyle();
         m_StyleTextRight.setFont(m_FontData);
         m_StyleTextRight.setAlignment(HorizontalAlignment.RIGHT);

         m_StyleTextCenter = m_WrkBk.createCellStyle();
         m_StyleTextCenter.setFont(m_FontData);
         m_StyleTextCenter.setAlignment(HorizontalAlignment.CENTER);

         // Style 8pt, left aligned, bold
         m_StyleBold = m_WrkBk.createCellStyle();
         m_StyleBold.setFont(m_FontBold);
         m_StyleBold.setAlignment(HorizontalAlignment.LEFT);

         // Style 8pt, right aligned, bold
         m_StyleBoldRight = m_WrkBk.createCellStyle();
         m_StyleBoldRight.setFont(m_FontBold);
         m_StyleBoldRight.setAlignment(HorizontalAlignment.RIGHT);

         // Style 8pt, centered, bold
         m_StyleBoldCenter = m_WrkBk.createCellStyle();
         m_StyleBoldCenter.setFont(m_FontBold);
         m_StyleBoldCenter.setAlignment(HorizontalAlignment.CENTER);

         m_StyleDec = m_WrkBk.createCellStyle();
         m_StyleDec.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleDec.setFont(m_FontData);
         m_StyleDec.setDataFormat(df.getFormat("#,##0.00"));

         m_StyleDecBold = m_WrkBk.createCellStyle();
         m_StyleDecBold.setFont(m_FontBold);
         m_StyleDecBold.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleDecBold.setDataFormat(df.getFormat("#,##0.00"));

         m_StyleHeader = m_WrkBk.createCellStyle();
         m_StyleHeader.setFont(m_FontTitle);
         m_StyleHeader.setAlignment(HorizontalAlignment.CENTER);

         m_StyleInt = m_WrkBk.createCellStyle();
         m_StyleInt.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleInt.setFont(m_FontData);
         m_StyleInt.setDataFormat((short)3);

         m_StyleTitle = m_WrkBk.createCellStyle();
         m_StyleTitle.setFont(m_FontTitle);
         m_StyleTitle.setAlignment(HorizontalAlignment.LEFT);

         m_Sheet = m_WrkBk.createSheet();
         m_Sheet.setMargin(XSSFSheet.BottomMargin, .25);
         m_Sheet.getPrintSetup().setLandscape(true);
         m_Sheet.getPrintSetup().setPaperSize((short)5);

         m_Header = m_Sheet.getHeader();
         m_Header.setCenter(HSSFHeader.font("Arial", "Bold") + HSSFHeader.fontSize((short) 12) + "Daily Sales Order Cuts");
         m_Header.setLeft(HSSFHeader.font("Arial", "Bold") + HSSFHeader.fontSize((short) 12) + " " + m_BegDate + " thru " + m_EndDate);
         m_Header.setRight(HSSFHeader.font("Arial", "Bold") + HSSFHeader.fontSize((short) 12) + HSSFHeader.page());

         m_RowNum = 0;

         // Initialize the default column widths
         for ( short i = 0; i < 20; i++ )
         	m_Sheet.setColumnWidth(i, 2000);

         m_Sheet.setColumnWidth(1, 2000);
         m_Sheet.setColumnWidth(2, 7000);

         // Create the column headings
      	m_Row = m_Sheet.createRow(m_RowNum++);
      	col = (short)0;
         m_Sheet.setColumnWidth(col, 2200);
         createCell(m_Row, col++, "I & C Velocity Code Summary", m_StyleBold);

         m_Row = m_Sheet.createRow(m_RowNum++);
         col = (short)0;
         m_Sheet.setColumnWidth(col, 2200);
         createCell(m_Row, col++, "DC:", m_StyleBold);
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ) {
            if ( m_Warehouse.equals("1"))
               createCell(m_Row, col++," Portland" , m_StyleBold);
            else if ( m_Warehouse.equals("2"))
               createCell(m_Row, col++," Pittston" , m_StyleBold); // if it is not 1 or 2 leave blank!  no fall thorugh value
         }
         else
            createCell(m_Row, col++," Both" , m_StyleBold);

         m_Row = m_Sheet.createRow(m_RowNum++);
         col = (short)0;
         m_Sheet.setColumnWidth(col, 2200);
         createCell(m_Row, col++, "Dept:", m_StyleBold);
         if ( m_DeptNum != null && m_DeptNum.length() > 0 )
               createCell(m_Row, col++,m_DeptNum.toString(), m_StyleBold);
         else
            createCell(m_Row, col++," All" , m_StyleBold);

         m_Row = m_Sheet.createRow(m_RowNum++);
         col = (short)0;
         m_Sheet.setColumnWidth(col, 2200);
         createCell(m_Row, col++, "Vendor:", m_StyleBold);
         if ( m_VendorID != null && m_VendorID.length() > 0 )
            createCell(m_Row, col++,m_VendorID.toString(), m_StyleBold);
         else
            createCell(m_Row, col++," All" , m_StyleBold);

         m_RowNum++;
      }

      catch ( Exception e ) {
         log.error( e );
      }
   }

   /**
    * Creates a trailer line for an item
    *
    * @param itemQtyOrd int - the total qty ordered for the item
    * @param itemQtyShip - the total qty shipped for the item
    * @param itemLineCnt - the number of lines cut
    * @param itemAmtCut - the value of the items cut
    */
   private void PeriodTrailer(String Period, int PeriodItemCnt, int PeriodQtyShip, double PeriodDollarShip, int PeriodOnHand, double PeriodDollarInv)
   {
   	short col = 0;

   	// Only print a total line for the item if there are more than 1 line
   	if ( PeriodItemCnt > 1 ) {
         m_Row = m_Sheet.createRow(m_RowNum++);
	      createCell(m_Row, col++, "Summary", m_StyleBoldRight);
	      createCell(m_Row, col++, PeriodItemCnt, m_StyleBoldRight);
         createCell(m_Row, col++, "< #Items", m_StyleBold);

	      createCell(m_Row, col+7, PeriodQtyShip, m_StyleBoldRight);
	      createCell(m_Row, col+8, PeriodDollarShip, m_StyleDecBold);
         createCell(m_Row, col+9, PeriodOnHand, m_StyleBoldRight);
         createCell(m_Row, col+10, PeriodDollarInv, m_StyleDecBold);
         m_Row = m_Sheet.createRow(m_RowNum++);

      }
   }

   private boolean prepareStatements()
   {
   	StringBuffer sql = new StringBuffer();
      StringBuffer dcsql = new StringBuffer();
      StringBuffer vendorsql = new StringBuffer();
      StringBuffer deptsql = new StringBuffer();

   	try {

         vendorsql.setLength(0);
         if ( m_VendorID != null && m_VendorID.length() > 0 ) {
            vendorsql.append("join vendor v on v.vendor_id = item.vendor_id and v.vendor_id = ");
            vendorsql.append(m_VendorID);
            vendorsql.append(" \r\n");
         }
         else {
            vendorsql.append("join vendor v on v.vendor_id = item.vendor_id \r\n");
         }


         deptsql.setLength(0);
         if ( m_DeptNum != null && m_DeptNum.length() > 0 ) {
            deptsql.append("join emery_dept ed on ed.dept_id = item.dept_id and ed.dept_num = '");
            deptsql.append(m_DeptNum);
            deptsql.append("' \r\n");

         }
         else {
            deptsql.append("join emery_dept ed on ed.dept_id = item.dept_id \r\n");
         }

         dcsql.setLength(0);
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ) {
            //first we limit the item selection by warehouse
            dcsql.append("join item_warehouse iw on iw.item_id = item.item_id and iw.warehouse_id = ");
            dcsql.append(m_Warehouse);
            dcsql.append(" \r\n");
            //now we get totals only from the rquested warehouse
            dcsql.append("   left outer join  \r\n");
            dcsql.append("            (select iA.item_id, nvl(iA.qty_on_hand,0) onhand, nvl(iA.total_cost,0) totalcost,   \r\n");
            dcsql.append("            nvl(iA.average_cost,0) average_cost   \r\n");
            dcsql.append("            from item_inventory iA    \r\n");
            dcsql.append("            where iA.warehouse_id = ");
            dcsql.append(m_Warehouse);
            dcsql.append(" and iA.inventory_date in   \r\n");
            dcsql.append("              (select max(iAA.inventory_date) from item_inventory iAA   \r\n");
            dcsql.append("                 where iAA.warehouse_id = ");
            dcsql.append(m_Warehouse);
            dcsql.append(" and iAA.item_id = iA.item_id)) ii on ii.item_id = item.item_id   \r\n");

         }
         else {   /// we have to get totals from both of the damn warehouses
            dcsql.append("left outer join         \r\n");
            dcsql.append("   (select item_id, nvl(sum(onhand),0) onhand, nvl(sum(totalcost),0) totalcost, \r\n");
            dcsql.append("   decode(sum(onhand),0,0, round((sum(totalcost) / sum(onhand)),3)) average_cost \r\n");
            dcsql.append("   from  \r\n");
            dcsql.append("      (select iA.item_id, nvl(iA.qty_on_hand,0) onhand, nvl(iA.total_cost,0) totalcost from item_inventory iA  \r\n");
            dcsql.append("      where iA.warehouse_id = 1 and iA.inventory_date in  \r\n");
            dcsql.append("         (select max(iAA.inventory_date) from item_inventory iAA \r\n");
            dcsql.append("           where iAA.warehouse_id = 1 and iAA.item_id = iA.item_id) \r\n");
            dcsql.append("      union \r\n");
            dcsql.append("      select iB.item_id, nvl(iB.qty_on_hand,0) onhand, nvl(iB.total_cost,0) totalcost  from item_inventory iB  \r\n");
            dcsql.append("      where iB.warehouse_id = 2 and iB.inventory_date in  \r\n");
            dcsql.append("         (select max(iBB.inventory_date) from item_inventory iBB \r\n");
            dcsql.append("         where iBB.warehouse_id =2and iBB.item_id = iB.item_id) \r\n");
            dcsql.append("      )  \r\n");
            dcsql.append("      group by item_id) ii   on ii.item_id = item.item_id    \r\n");
         }

         sql.setLength(0);
         sql.append("select 'A 0-13' period,item.item_id, to_char(item.setup_date,'mm/dd/yyyy') setup_date, \r\n");
         sql.append("   item.description, v.name vendor_name, item.vendor_id, \r\n");
         sql.append("   ed.dept_num, item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC') nbc, item.buy_mult bm,\r\n");
         sql.append("   ii.onhand, ii.totalcost inv_dollars, ii.average_cost,  \r\n");
         sql.append("   item_price_procs.todays_buy(item.item_id) todays_buy, \r\n");
         sql.append("   sum(id.qty_shipped) shipqty,sum(id.dollars_shipped) shipdoll \r\n");
         sql.append("from item \r\n");
         sql.append(vendorsql);
         sql.append(deptsql);
         sql.append("left outer join itemsales id on id.item_nbr = item.item_id \r\n");
         sql.append("   and id.invoice_date > add_months(trunc(sysdate), - 3) \r\n");
         sql.append("join item_velocity iv on iv.velocity_Id = item.velocity_id \r\n");
         sql.append("     and iv.velocity_id in \r\n");
         sql.append("       (select iv2.velocity_id from item_velocity iv2 \r\n");
         sql.append("        where iv2.velocity in ('I','C')) \r\n");
         sql.append(dcsql);
         sql.append("where \r\n");
         sql.append("   item.setup_date >  add_months(trunc(sysdate), -3) and \r\n");
         sql.append("   item.item_id in  \r\n");
         sql.append("  (select distinct(iw.item_id) from item_warehouse iw  \r\n");
         sql.append("   where iw.active = 1 and in_catalog = 1)  \r\n");
         sql.append("   group by item.item_id, to_char(item.setup_date,'mm/dd/yyyy'),\r\n");
         sql.append("   item.description, ed.dept_num, v.name, \r\n");
         sql.append("   item.vendor_id,item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC'), item.buy_mult, \r\n");
         sql.append("   ii.onhand,ii.totalcost, ii.average_cost \r\n");
         sql.append("   union all \r\n");
         sql.append("       \r\n");
         sql.append(" select 'B 13-26' period,item.item_id, to_char(item.setup_date,'mm/dd/yyyy') setup_date,  \r\n");
         sql.append("   item.description, v.name vendor_name, item.vendor_id, \r\n");
         sql.append("   ed.dept_num, item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC') nbc, item.buy_mult bm,\r\n");
         sql.append("   ii.onhand, ii.totalcost inv_dollars, ii.average_cost, \r\n");
         sql.append("   item_price_procs.todays_buy(item.item_id) todays_buy, \r\n");
         sql.append("   sum(id.qty_shipped) shipqty,sum(id.dollars_shipped) shipdoll \r\n");
         sql.append("from item \r\n");
         sql.append(vendorsql);
         sql.append(deptsql);
         sql.append("left outer join itemsales id on id.item_nbr = item.item_id \r\n");
         sql.append("   and id.invoice_date >= add_months(trunc(sysdate), - 6) \r\n");
         sql.append("    \r\n");
         sql.append("join item_velocity iv on iv.velocity_Id = item.velocity_id \r\n");
         sql.append("     and iv.velocity_id in \r\n");
         sql.append("       (select iv2.velocity_id from item_velocity iv2 \r\n");
         sql.append("        where iv2.velocity in ('I','C')) \r\n");
         sql.append(dcsql);
         sql.append("where \r\n");
         sql.append("   item.in_catalog = 1 and \r\n");
         sql.append("   item.setup_date > add_months(trunc(sysdate), -6) and setup_date <= add_months(trunc(sysdate), -3) and \r\n");
         sql.append("   item.item_id in  \r\n");
         sql.append("  (select distinct(iw.item_id) from item_warehouse iw  \r\n");
         sql.append("   where iw.active = 1)  \r\n");
         sql.append("       \r\n");
         sql.append("   group by item.item_id, to_char(item.setup_date,'mm/dd/yyyy'),\r\n");
         sql.append("   item.description, ed.dept_num, v.name, \r\n");
         sql.append("   item.vendor_id,item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC'), item.buy_mult, \r\n");
         sql.append("   ii.onhand,ii.totalcost, ii.average_cost \r\n");

         sql.append("   union all \r\n");

         sql.append("select 'C 26-39' period,item.item_id, to_char(item.setup_date,'mm/dd/yyyy') setup_date,  \r\n");
         sql.append("   item.description, v.name vendor_name, item.vendor_id, \r\n");
         sql.append("   ed.dept_num, item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC') nbc, item.buy_mult bm,\r\n");
         sql.append("    ii.onhand, ii.totalcost inv_dollars, ii.average_cost, \r\n");
         sql.append("   item_price_procs.todays_buy(item.item_id) todays_buy, \r\n");
         sql.append("   sum(id.qty_shipped) shipqty,sum(id.dollars_shipped) shipdoll \r\n");
         sql.append("from item \r\n");
         sql.append(vendorsql);
         sql.append(deptsql);
         sql.append("left outer join itemsales id on id.item_nbr = item.item_id \r\n");
         sql.append("and id.invoice_date >= add_months(trunc(sysdate), - 9) \r\n");
         sql.append("    \r\n");
         sql.append("join item_velocity iv on iv.velocity_Id = item.velocity_id \r\n");
         sql.append("     and iv.velocity_id in \r\n");
         sql.append("       (select iv2.velocity_id from item_velocity iv2 \r\n");
         sql.append("        where iv2.velocity in ('I','C')) \r\n");
         sql.append(dcsql);
         sql.append("where \r\n");
         sql.append("   item.in_catalog = 1 and \r\n");
         sql.append("   item.setup_date > add_months(trunc(sysdate), -9) and item.setup_date <= add_months(trunc(sysdate), -6) and \r\n");
         sql.append("   item.item_id in  \r\n");
         sql.append("  (select distinct(iw.item_id) from item_warehouse iw  \r\n");
         sql.append("   where iw.active = 1)  \r\n");
         sql.append("          \r\n");
         sql.append("   group by item.item_id, to_char(item.setup_date,'mm/dd/yyyy'),\r\n");
         sql.append("   item.description, ed.dept_num, v.name, \r\n");
         sql.append("   item.vendor_id,item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC'), item.buy_mult, \r\n");
         sql.append("   ii.onhand,ii.totalcost, ii.average_cost \r\n");

         sql.append("   union all \r\n");

         sql.append("select 'D 39-52' period,item.item_id, to_char(item.setup_date,'mm/dd/yyyy') setup_date, \r\n");
         sql.append("   item.description, v.name vendor_name, item.vendor_id, \r\n");
         sql.append("   ed.dept_num, item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC') nbc, item.buy_mult bm,\r\n");
         sql.append("    ii.onhand, ii.totalcost inv_dollars, ii.average_cost,   \r\n");
         sql.append("   item_price_procs.todays_buy(item.item_id) todays_buy, \r\n");
         sql.append("   sum(id.units_shipped) shipqty,sum(id.dollars_shipped) shipdoll \r\n");
         sql.append("from item \r\n");
         sql.append(vendorsql);
         sql.append(deptsql);
         sql.append("left outer join monthlyitemsales id on id.item_nbr = item.item_id \r\n");
         sql.append("and id.year_month >= (to_char(add_months(sysdate, -9),'YYYYMM')) \r\n");
         sql.append("    \r\n");
         sql.append("join item_velocity iv on iv.velocity_Id = item.velocity_id \r\n");
         sql.append("     and iv.velocity_id in \r\n");
         sql.append("       (select iv2.velocity_id from item_velocity iv2 \r\n");
         sql.append("        where iv2.velocity in ('I','C')) \r\n");
         sql.append(dcsql);
         sql.append("where \r\n");
         sql.append("   item.in_catalog = 1 and \r\n");
         sql.append("   item.setup_date > add_months(trunc(sysdate), -12) and item.setup_date <= add_months(trunc(sysdate), -9) and \r\n");
         sql.append("   item.item_id in  \r\n");
         sql.append("  (select distinct(iw.item_id) from item_warehouse iw  \r\n");
         sql.append("   where iw.active = 1)  \r\n");
         sql.append("       \r\n");
         sql.append("   group by item.item_id, to_char(item.setup_date,'mm/dd/yyyy'),\r\n");
         sql.append("   item.description, ed.dept_num, v.name, \r\n");
         sql.append("   item.vendor_id,item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC'), item.buy_mult, \r\n");
         sql.append("   ii.onhand,ii.totalcost, ii.average_cost \r\n");

         sql.append("   union all \r\n");

         sql.append("select 'E 52-104' period,item.item_id, to_char(item.setup_date,'mm/dd/yyyy') setup_date, \r\n");
         sql.append("   item.description, v.name vendor_name, item.vendor_id, \r\n");
         sql.append("   ed.dept_num, item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC') nbc, item.buy_mult bm,\r\n");
         sql.append("    ii.onhand, ii.totalcost inv_dollars, ii.average_cost, \r\n");
         sql.append("   item_price_procs.todays_buy(item.item_id) todays_buy, \r\n");
         sql.append("   sum(id.units_shipped) shipqty,sum(id.dollars_shipped) shipdoll \r\n");
         sql.append("from item \r\n");
         sql.append(vendorsql);
         sql.append(deptsql);
         sql.append("left outer join monthlyitemsales id on id.item_nbr = item.item_id \r\n");
         sql.append("and id.year_month >= (to_char(add_months(sysdate, -12),'YYYYMM')) \r\n");
         sql.append(" \r\n");
         sql.append("join item_velocity iv on iv.velocity_Id = item.velocity_id \r\n");
         sql.append("     and iv.velocity_id in \r\n");
         sql.append("       (select iv2.velocity_id from item_velocity iv2 \r\n");
         sql.append("        where iv2.velocity in ('I','C')) \r\n");
         sql.append(dcsql);
         sql.append("where \r\n");
         sql.append("   item.in_catalog = 1 and \r\n");
         sql.append("   item.setup_date > add_months(trunc(sysdate), -24) and item.setup_date <= add_months(trunc(sysdate), -12) and \r\n");
         sql.append("   item.item_id in  \r\n");
         sql.append("  (select distinct(iw.item_id) from item_warehouse iw  \r\n");
         sql.append("   where iw.active = 1)  \r\n");
         sql.append("    \r\n");
         sql.append("   group by item.item_id, to_char(item.setup_date,'mm/dd/yyyy'),\r\n");
         sql.append("   item.description, ed.dept_num, v.name, \r\n");
         sql.append("   item.vendor_id,item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC'), item.buy_mult, \r\n");
         sql.append("   ii.onhand,ii.totalcost, ii.average_cost \r\n");

         sql.append("   union all \r\n");

         sql.append("select 'F 104+' period,item.item_id, to_char(item.setup_date,'mm/dd/yyyy') setup_date, \r\n");
         sql.append("   item.description, v.name vendor_name, item.vendor_id, \r\n");
         sql.append("   ed.dept_num, item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC') nbc, item.buy_mult bm,\r\n");
         sql.append("    ii.onhand, ii.totalcost inv_dollars, ii.average_cost, \r\n");
         sql.append("   item_price_procs.todays_buy(item.item_id) todays_buy, \r\n");
         sql.append("   sum(id.units_shipped) shipqty,sum(id.dollars_shipped) shipdoll \r\n");
         sql.append("from item \r\n");
         sql.append(vendorsql);
         sql.append(deptsql);
         sql.append("left outer join monthlyitemsales id on id.item_nbr = item.item_id \r\n");
         sql.append("and id.year_month >= (to_char(add_months(sysdate, -12),'YYYYMM')) \r\n");
         sql.append("    \r\n");
         sql.append(dcsql);
         sql.append("where \r\n");
         sql.append("   item.in_catalog = 1 and \r\n");
         sql.append("   item.setup_date <= add_months(trunc(sysdate), -24) and \r\n");
         sql.append("   item.item_id in  \r\n");
         sql.append("  (select distinct(iw.item_id) from item_warehouse iw  \r\n");
         sql.append("   where iw.active = 1)  \r\n");
         sql.append("       \r\n");
         sql.append("   group by item.item_id, to_char(item.setup_date,'mm/dd/yyyy'),\r\n");
         sql.append("   item.description, ed.dept_num, v.name, \r\n");
         sql.append("   item.vendor_id,item.stock_pack, \r\n");
         sql.append("   decode(item.broken_case_id,1,'  ','NBC'), item.buy_mult, \r\n");
         sql.append("   ii.onhand,ii.totalcost, ii.average_cost \r\n");
         sql.append("order by period, shipqty, setup_date ");
   		m_Item = m_OraConn.prepareStatement(sql.toString());

   		return true;
   	}

   	catch ( Exception e ) {
   		log.error("exception", e);
   		return false;
   	}

   	finally {
   		sql = null;
         dcsql = null;
         vendorsql = null;
         deptsql = null;
   	}
   }

   /**
    * Sets the parameters for the report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      int pcount = params.size();
      Param param = null;


      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);

         if ( param.name.equals("vendor") )
             m_VendorID = param.value;

         if ( param.name.equals("dept") )
             m_DeptNum = param.value;

         if ( param.name.equals("warehouse") && param.value.trim().length() > 0 ) {
            if (param.value.trim().equals("01"))
               m_Warehouse = "1";
            else if (param.value.trim().equals("04"))
               m_Warehouse = "2";
         }

         if ( param.name.equals("begdate") )
            m_BegDate = param.value;

         if ( param.name.equals("enddate") )
            m_EndDate = param.value;
      }
   }
}

