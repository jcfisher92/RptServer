/**
 * Title:			CutReport.java
 * Description:	Daily warehouse item cuts by buyer
 * Company:			Emery-Waterhouse
 * @author			prichter
 * @version			1.0
 * <p>
 * Create Date:	Jun 18, 2008
 * Last Update:   $Id: CutReport.java,v 1.14 2013/09/11 14:25:42 tli Exp $
 * <p>
 * History:
 *		$Log: CutReport.java,v $
 *		Revision 1.14  2013/09/11 14:25:42  tli
 *		Converted the facilityId to facilityName when needed
 *
 *		Revision 1.13  2013/09/09 18:33:38  tli
 *		Replace SkuQty web service call with item_qty_view
 *
 *		Revision 1.12  2012/10/05 14:06:00  jfisher
 *		Changes to deal with the timeout on the sku quantity web service.
 *
 *		Revision 1.11  2012/08/29 19:53:02  jfisher
 *		Switched web service calls from Wasp to Axis2
 *
 *		Revision 1.10  2012/05/04 01:56:09  jfisher
 *		Fixed the web service call issue and removed the extraneous property loading.
 *
 *		Revision 1.8  2012/05/03 07:55:10  prichter
 *		Fix to web service ip address
 *
 *		Revision 1.7  2012/05/03 04:39:10  pberggren
 *		*** empty log message ***
 *
 *		Revision 1.6  2012/05/03 04:19:15  pberggren
 *		Added server.properties call to force report to .57
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

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Header;
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
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class CutReport extends Report
{
   private PreparedStatement m_Item;
   private PreparedStatement m_ItemWhs;
   private PreparedStatement m_OnHand;
   private PreparedStatement m_Whs;
   private PreparedStatement m_DeptSvcLvl;
   private PreparedStatement m_OldestPO;

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
   private String m_BegDate;
   private String m_EndDate;
   private String m_Warehouse;
   private String m_Buyer;
   private String m_Vendor;

   private short m_RowNum = 0;
   private ArrayList<String> m_FacilityList = new ArrayList<String>();
   private ArrayList<Integer> m_WhsList = new ArrayList<Integer>();

   private PreparedStatement m_ItemDCQty;

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
      ResultSet rs2 = null;
      int col;

      String lastVendor = "begin";
      String lastItem = "begin";
      String lastBuyer = "begin";
      int itemQtyOrd = 0;
      int itemQtyShip = 0;
      int itemLineCnt = 0;
      double itemAmtCut = 0.0;
      int vndQtyOrd = 0;
      int vndQtyShip = 0;
      int vndLineCnt = 0;
      double vndAmtCut = 0.0;
      double deptSvcLvl = 0.0;

      int totOnHand = 0;
      int qty;

      int ohCol = 0;

      String oldestPoNbr = "";
      String oldestPoDate = "";
      String oldestPoQtyOrd = "";
      int oldestPoQtyRcvd = 0;
      String oldestPoDueIn = "";


      m_FileNames.add(m_RptProc.getUid() + "cutreport" + getStartTime() + ".xlsx");
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      initReport();

      try {
         setCurAction( "Running the cut report query" );
         m_Item.setString(1, m_BegDate);
         m_Item.setString(2, m_EndDate);
         m_Item.setString(3, m_BegDate);
         m_Item.setString(4, m_EndDate);
         rs = m_Item.executeQuery();

         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            setCurAction( "Processing " + rs.getString("vendor_nbr") + rs.getString("vendor_name"));

            // Check for the start of a new vendor
            if ( !rs.getString("vendor_nbr").equals(lastVendor) ) {
               // If this isn't the first vendor, display item and vendor trailers
               if ( !lastItem.equals("begin"))
                  itemTrailer(itemQtyOrd, itemQtyShip, itemLineCnt, itemAmtCut);

               if ( !lastVendor.equals("begin"))
                  vendorTrailer(vndQtyOrd, vndQtyShip, vndLineCnt, vndAmtCut);

               itemQtyOrd = 0;
               itemQtyShip = 0;
               itemLineCnt = 0;
               itemAmtCut = 0.0;
               vndQtyOrd = 0;
               vndQtyShip = 0;
               vndLineCnt = 0;
               vndAmtCut = 0.0;

               // If this is the start of a new buyer, force a page break
               // and print the column headings.  And get the service level.

               if ( !rs.getString("buying_dept").equals(lastBuyer) ) {

                  m_DeptSvcLvl.setString(1, rs.getString("buying_dept"));
                  m_DeptSvcLvl.setString(2, m_BegDate);
                  m_DeptSvcLvl.setString(3, m_EndDate);
                  rs2 = m_DeptSvcLvl.executeQuery();
                  while ( rs2.next() ) {
                     deptSvcLvl = rs2.getDouble("service_level");
                  }
                  if ( !lastBuyer.equals("begin") ){
                     m_Sheet.setRowBreak(m_RowNum++);
                  }

                  buyerHeader(rs.getString("buying_dept"), rs.getString("buyer_name"), deptSvcLvl);
               }

               vendorHeader(rs.getString("vendor_nbr"), rs.getString("vendor_Name"), rs.getDouble("service_level"));
               itemHeader(rs.getString("item_nbr"), rs.getString("item_descr"));
            }

            else {
               if ( !rs.getString("item_nbr").equals(lastItem)) {
                  if ( !lastItem.equals("begin") )
                     itemTrailer(itemQtyOrd, itemQtyShip, itemLineCnt, itemAmtCut);

                  itemQtyOrd = 0;
                  itemQtyShip = 0;
                  itemLineCnt = 0;
                  itemAmtCut = 0.0;
                  itemHeader(rs.getString("item_nbr"), rs.getString("item_descr"));
               }
            }

            col = (short)1;
            m_Row = m_Sheet.createRow(m_RowNum++);
            createCell(m_Row, col++, rs.getString("cust_nbr"), m_StyleText);
            createCell(m_Row, col++, rs.getString("cust_name"), m_StyleText);
            createCell(m_Row, col++, rs.getString("warehouse"), m_StyleText);
            createCell(m_Row, col++, rs.getInt("qty_ordered"), m_StyleInt);
            createCell(m_Row, col++, rs.getInt("qty_shipped"), m_StyleInt);
            col++;
            createCell(m_Row, col++, rs.getDouble("cut_amt"), m_StyleDec);
            createCell(m_Row, col++, rs.getString("promo_nbr"), m_StyleTextCenter);
            createCell(m_Row, col++, rs.getString("backorder_date") == null ? "N" : "Y", m_StyleTextCenter);

            // Only print this data once for each item
            if ( !rs.getString("item_nbr").equals(lastItem) ) {
               createCell(m_Row, col++, rs.getString("velocity"), m_StyleTextCenter);
               createCell(m_Row, col++, rs.getString("nbc"), m_StyleTextCenter);
               createCell(m_Row, col++, rs.getString("ship_unit"), m_StyleTextCenter);
               ohCol = col++;
               createCell(m_Row, col++, rs.getString("net_on_order"), m_StyleTextCenter);
               // createCell(m_Row, col++, rs.getString("po_nbr"), m_StyleTextCenter);
               // createCell(m_Row, col++, rs.getString("po_date"), m_StyleTextCenter);
               // createCell(m_Row, col++, rs.getString("qty_on_order"), m_StyleTextCenter);
               // createCell(m_Row, col++, rs.getInt("qty_put_away"), m_StyleTextCenter);
               // createCell(m_Row, col++, rs.getString("due_in"), m_StyleTextCenter);

               ResultSet rsPO = null;
               oldestPoNbr = "";
               oldestPoDate = "";
               oldestPoQtyOrd = "";
               oldestPoQtyRcvd = 0;
               oldestPoDueIn = "";

               try {
                  m_OldestPO.setInt(1, rs.getInt("item_ea_id"));                  
                  m_OldestPO.setString(2, rs.getString("fas_facility_id"));
                  rsPO = m_OldestPO.executeQuery();

                  while (rsPO.next()) {
                     oldestPoNbr = rsPO.getString("po_nbr");
                     oldestPoDate = rsPO.getString("po_date");
                     oldestPoQtyOrd = rsPO.getString("qty_ordered");
                     oldestPoQtyRcvd = rsPO.getInt("qty_put_away");
                     oldestPoDueIn = rsPO.getString("due_in");
                  }
               }
               
               finally {
                  closeRSet(rsPO);
                  rsPO = null;
               }

               createCell(m_Row, col++, oldestPoNbr, m_StyleTextCenter);
               createCell(m_Row, col++, oldestPoDate, m_StyleTextCenter);
               createCell(m_Row, col++, oldestPoQtyOrd, m_StyleTextCenter);
               createCell(m_Row, col++, oldestPoQtyRcvd, m_StyleTextCenter);
               createCell(m_Row, col++, oldestPoDueIn, m_StyleTextCenter);

               totOnHand = 0;
               for ( int i = 0; i < m_FacilityList.size(); i++ ) {
                  if ( itemAtFacility(rs.getString("item_nbr"), m_WhsList.get(i))) {
                     qty = getOnHand(m_FacilityList.get(i), rs.getString("item_nbr"));
                     totOnHand += qty;
                     createCell(m_Row, col++, qty, m_StyleTextCenter);
                  }
                  else
                     createCell(m_Row, col++, "n/a", m_StyleTextCenter);
               }

               createCell(m_Row, ohCol, totOnHand, m_StyleTextCenter);
            }

            itemQtyOrd += rs.getInt("qty_ordered");
            itemQtyShip += rs.getInt("qty_shipped");
            itemLineCnt++;
            itemAmtCut += rs.getDouble("cut_amt");

            vndQtyOrd += rs.getInt("qty_ordered");
            vndQtyShip += rs.getInt("qty_shipped");
            vndLineCnt++;
            vndAmtCut += rs.getDouble("cut_amt");

            lastItem = rs.getString("item_nbr");
            lastVendor = rs.getString("vendor_nbr");
            lastBuyer = rs.getString("buying_dept");
         }

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
   private void buyerHeader(String buyerNbr, String name, double SvcLvl)


   {

      m_Row = m_Sheet.createRow(m_RowNum++);
      createCell(m_Row, (short)0, buyerNbr + " " + name + " - " + SvcLvl +"%", m_StyleBold );
   }

   /**
    * Resource cleanup
    */
   public void cleanup()
   {
      DbUtils.closeDbConn(null, m_Item, null);
      DbUtils.closeDbConn(null, m_ItemWhs, null);
      DbUtils.closeDbConn(null, m_OnHand, null);
      DbUtils.closeDbConn(null, m_Whs, null);
      DbUtils.closeDbConn(null, m_ItemDCQty, null);
      DbUtils.closeDbConn(null, m_OldestPO, null);

      m_Item = null;
      m_ItemWhs = null;
      m_OnHand = null;
      m_Whs = null;
      m_OldestPO = null;

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
      m_Buyer = null;
      m_Vendor = null;

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
    * @return HSSFCell newly created cell
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
    * @return HSSFCell newly created cell
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
    * @return HSSFCell newly created cell
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
         m_EdbConn = m_RptProc.getEdbConn();

         if ( m_EdbConn != null ) {
            if ( prepareStatements() )
               created = buildOutputFile();
         }
         else
            throw new Exception("null database connection");
      }

      catch ( Exception ex ) {
         log.fatal("[CutReport]", ex);
      }

      finally {
         cleanup();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Returns an item's on hand quantity a the given facility.  The data is
    * taken from the Fascor database.
    *
    * @param facility String - the fascor facility id
    * @param itemId String - the item id
    * @return int - the quantity on hand.
    * @throws Exception
    */
   private int getOnHand(String facilityId, String itemId) throws Exception
   {
      int qty = 0;
      ResultSet rset = null;

      if ( itemId != null && itemId.length() == 7 ) {	      
         m_ItemDCQty.setString(1, itemId);

         if( facilityId.equals("01") || facilityId.equals("02") )
            m_ItemDCQty.setString(2, "PORTLAND");
         else {
            if( facilityId.equals("04") || facilityId.equals("05") )
               m_ItemDCQty.setString(2, "PITTSTON");
         }

         try {
            rset = m_ItemDCQty.executeQuery();

            if ( rset.next() )
               qty = rset.getInt("available_qty");
         }

         finally {
            closeRSet(rset);
            rset = null;
         }
      }

      return qty;
   }


   /**
    * Creates the workbook and worksheet.  Creates any fonts and styles that
    * will be used.
    */
   private void initReport()
   {
      XSSFDataFormat df;
      ResultSet rs = null;
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

         // Create a list if warehouses for later use
         rs = m_Whs.executeQuery();

         while ( rs.next() ) {
            m_WhsList.add(new Integer(rs.getInt("warehouse_id")));
            m_FacilityList.add(rs.getString("fas_facility_id"));
         }

         // Create the column headings
         m_Row = m_Sheet.createRow(m_RowNum);
         col = (short)3;
         m_Sheet.setColumnWidth(col, 2200);
         createCell(m_Row, col++, "Facility", m_StyleBold);
         m_Sheet.setColumnWidth(col, 1800);
         createCell(m_Row, col++, "Qty Ord", m_StyleBoldRight);
         m_Sheet.setColumnWidth(col, 1800);
         createCell(m_Row, col++, "Qty Ship", m_StyleBoldRight);
         m_Sheet.setColumnWidth(col, 1200);
         createCell(m_Row, col++, "Lines", m_StyleBoldRight);
         m_Sheet.setColumnWidth(col, 2000);
         createCell(m_Row, col++, "$ Cut", m_StyleBoldRight);
         m_Sheet.setColumnWidth(col, 1500);
         createCell(m_Row, col++, "Promo", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 1000);
         createCell(m_Row, col++, "BO", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 1000);
         createCell(m_Row, col++, "Vel", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 1500);
         createCell(m_Row, col++, "NBC", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 1200);
         createCell(m_Row, col++, "Unit", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 1200);
         createCell(m_Row, col++, "QOH", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 1200);
         createCell(m_Row, col++, "QOO", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2000);
         createCell(m_Row, col++, "PO", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2200);
         createCell(m_Row, col++, "PO Date", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 1200);
         createCell(m_Row, col++, "Qty", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 1200);
         createCell(m_Row, col++, "Rcvd", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2200);
         createCell(m_Row, col++, "Due In", m_StyleBoldCenter);

         for ( int i = 0; i < m_WhsList.size(); i++ ) {
            m_Sheet.setColumnWidth(col, 1200);
            createCell(m_Row, col++, "OH" + m_WhsList.get(i), m_StyleBoldCenter);
         }

         // Set the column heading row to repeat on each page      
         m_Sheet.setRepeatingRows(CellRangeAddress.valueOf("1:1"));
         m_Sheet.setRepeatingColumns(CellRangeAddress.valueOf("A:AM"));
         m_RowNum++;
      }

      catch ( Exception e ) {
         log.error( "[CutReport]", e );
      }
   }

   /**
    * Returns true if an item is actively purchased at the given warehouse
    * @param itemId String - the item id
    * @param whs int - the warehouse id
    * @return boolean - true if the item is stocked at this warehouse
    * @throws Exception
    */
   private boolean itemAtFacility(String itemId, int whs) throws Exception
   {
      ResultSet rs = null;

      try {
         m_ItemWhs.setString(1, itemId);
         m_ItemWhs.setInt(2, whs);
         rs = m_ItemWhs.executeQuery();

         return rs.next();
      }

      finally {
         closeRSet(rs);
         rs = null;
      }
   }

   /**
    * Creates a header line for an item
    * @param itemId String - the item id
    * @param descr String - the item description
    * @throws Exception
    */
   private void itemHeader(String itemId, String descr) throws Exception
   {
      short col = 0;

      m_Row = m_Sheet.createRow(m_RowNum++);
      createCell(m_Row, col++, itemId, m_StyleText);
      createCell(m_Row, col++, descr, m_StyleText);
   }

   /**
    * Creates a trailer line for an item
    *
    * @param itemQtyOrd int - the total qty ordered for the item
    * @param itemQtyShip - the total qty shipped for the item
    * @param itemLineCnt - the number of lines cut
    * @param itemAmtCut - the value of the items cut
    */
   private void itemTrailer(int itemQtyOrd, int itemQtyShip, int itemLineCnt, double itemAmtCut)
   {
      short col = 3;

      // Only print a total line for the item if there are more than 1 line
      if ( itemLineCnt > 1 ) {
         m_Row = m_Sheet.createRow(m_RowNum++);
         createCell(m_Row, col++, "Item Totals", m_StyleBoldRight);
         createCell(m_Row, col++, itemQtyOrd, m_StyleBoldRight);
         createCell(m_Row, col++, itemQtyShip, m_StyleBoldRight);
         createCell(m_Row, col++, itemLineCnt, m_StyleBoldRight);
         createCell(m_Row, col++, itemAmtCut, m_StyleDecBold);
      }
   }

   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer();

      try {
         sql.setLength(0);
         sql.append("select inv_dtl.vendor_nbr, inv_dtl.vendor_name, inv_dtl.item_nbr, inv_dtl.item_descr, inv_dtl.qty_ordered, ");
         sql.append("   inv_dtl.qty_shipped, (inv_dtl.qty_ordered - inv_dtl.qty_shipped) * inv_dtl.unit_sell cut_amt, ");
         sql.append("   inv_dtl.buying_dept, inv_dtl.buyer_name, item_velocity.velocity, inv_dtl.ship_unit, ");
         sql.append("   decode(inv_dtl.nbc, 'Y', 'N' || trim(to_char(inv_dtl.stock_pack)), ' ') as nbc, ");
         sql.append("   decode(inv_dtl.action, 'BACKORDER', 'Y', 'N') backordered, ");
         sql.append("   inv_dtl.warehouse, inv_dtl.promo_nbr, inv_hdr.cust_nbr, inv_hdr.cust_name,action, ");
         //sql.append("   po_hdr.po_nbr, to_char(po_hdr.po_date, 'mm/dd/yyyy') as po_date, ");
         sql.append("   inv_dtl.item_ea_id, warehouse.fas_facility_id, ");
         sql.append("   oo.net_on_order, ");  //Pam wants net so they get net
         //sql.append("   po_dtl.qty_ordered as qty_on_order, po_dtl.qty_put_away, ");
         //sql.append("   to_char(sched_in_date, 'mm/dd/yyyy') as due_in, ");
         sql.append("   svc_lvl.service_level, inv_dtl.backorder_date ");
         sql.append("from inv_dtl ");
         sql.append("join inv_hdr on inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id ");
         sql.append("join item_entity_attr on inv_dtl.item_ea_id = item_entity_attr.item_ea_id  ");
         sql.append("join warehouse on warehouse.name = inv_dtl.warehouse ");   		
         sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.warehouse_id = warehouse.warehouse_id ");
         sql.append("join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id ");
         sql.append("join item_type on item_type.item_type_id = item_entity_attr.item_type_id and item_type.itemtype = 'STOCK' ");
         //sql.append("left outer join ( ");
         //sql.append("   select po_dtl.warehouse, po_dtl.item_nbr, min(po_hdr.po_hdr_id) po_hdr_id ");
         //sql.append("   from po_hdr ");
         //sql.append("   join po_dtl on po_dtl.po_hdr_id = po_hdr.po_hdr_id and po_dtl.status <> 'CLOSED' and po_dtl.cancelled = 'N' ");
         //sql.append("   where po_hdr.status <> 'CLOSED' and po_hdr.cancelled = 'N' ");
         //sql.append("   group by po_dtl.warehouse, po_dtl.item_nbr ");
         //sql.append("   ) oldest_po on oldest_po.item_nbr = inv_dtl.item_nbr and oldest_po.warehouse = warehouse.fas_facility_id ");
         //sql.append("left outer join po_hdr on po_hdr.po_hdr_id = oldest_po.po_hdr_id ");
         //sql.append("left outer join po_dtl on po_dtl.po_hdr_id = oldest_po.po_hdr_id and po_dtl.item_nbr = inv_dtl.item_nbr ");
         sql.append("left outer join ( ");
         sql.append("   select po_hdr.warehouse, item_nbr, sum(qty_ordered) qty_ordered, sum(qty_put_away) qty_put_away, ");
         sql.append("   (sum(qty_ordered) - sum(qty_put_away)) net_on_order");
         sql.append("   from po_dtl ");
         sql.append("   join po_hdr on po_hdr.po_hdr_id = po_dtl.po_hdr_id and po_hdr.status <> 'CLOSED' and po_hdr.cancelled = 'N' ");
         sql.append("   where po_dtl.status <> 'CLOSED' and po_dtl.cancelled = 'N' ");
         sql.append("   group by po_hdr.warehouse, item_nbr ");
         sql.append("   ) oo on oo.item_nbr = inv_dtl.item_nbr and oo.warehouse = warehouse.fas_facility_id ");
         sql.append("left outer join ( ");
         sql.append("   select vendor_nbr, round(sum(decode(qty_shipped, 0, 0, 1)) / count(*) * 100, 1) service_level ");
         sql.append("   from inv_dtl  ");
         sql.append("   where invoice_date >= to_date(?, 'mm/dd/yyyy') and ");
         sql.append("         invoice_date <= to_date(?, 'mm/dd/yyyy') and ");
         sql.append("         sale_type = 'WAREHOUSE' and tran_type = 'SALE' and ");
         sql.append("         qty_ordered > 0 and ");
         sql.append("         (action is null or ( ");
         sql.append("            action not in ('ADJUST ORDER QTY','BACKORDER','CANCELLED','FUTURED','INACTIVE ITEM') AND ");
         sql.append("            action not like ('SUB ITEM%'))) ");
         sql.append("   group by vendor_nbr, vendor_name ");
         sql.append("   ) svc_lvl on svc_lvl.vendor_nbr = inv_dtl.vendor_nbr ");
         sql.append("where inv_dtl.invoice_date >= to_date(?, 'mm/dd/yyyy') and ");
         sql.append("      inv_dtl.invoice_date <= to_date(?, 'mm/dd/yyyy') and ");
         sql.append("		inv_dtl.sale_type = 'WAREHOUSE' and ");
         sql.append("		inv_dtl.tran_type = 'SALE' and ");

         if ( !m_Warehouse.equalsIgnoreCase("All") )
            sql.append("      inv_dtl.warehouse = '" + m_Warehouse.toUpperCase() + "' and ");

         if ( !m_Buyer.equalsIgnoreCase("All") )
            sql.append("      inv_dtl.buyer_name = '" + m_Buyer.toUpperCase() + "' and ");

         if ( m_Vendor.trim().length() == 6 )
            sql.append("      inv_dtl.vendor_nbr = '" + m_Vendor + "' and ");

         sql.append("   inv_dtl.qty_shipped < inv_dtl.qty_ordered and ");
         sql.append("   inv_dtl.action not in ('INACTIVE ITEM','CANCELLED') and ");
         sql.append("   inv_dtl.action not like 'SUB%' ");
         sql.append("		order by buying_dept, inv_dtl.vendor_name, warehouse.warehouse_id, item_velocity.velocity, item_entity_attr.item_id ");
         m_Item = m_EdbConn.prepareStatement(sql.toString());

         sql.setLength(0);

         sql.append("select emery_dept.dept_num, round(sum(decode(qty_shipped, 0, 0, 1)) / count(*) * 100, 1) service_level ");
         sql.append("from inv_dtl ");
         sql.append("join item_entity_attr on inv_dtl.item_ea_id = item_entity_attr.item_ea_id ");
         sql.append("join ejd_item on item_entity_attr.ejd_item_id = ejd_item.ejd_item_id ");
         sql.append("join emery_dept on ejd_item.dept_id = emery_dept.dept_id ");
         sql.append("and emery_dept.dept_num = ? ");
         sql.append("where invoice_date <= to_date(?, 'mm/dd/yyyy') and ");
         sql.append("invoice_date >= to_date(?, 'mm/dd/yyyy') and ");
         sql.append("sale_type = 'WAREHOUSE' and tran_type = 'SALE' and ");
         sql.append("qty_ordered > 0 and ");
         sql.append("(action is null or ( action not in ('ADJUST ORDER QTY','BACKORDER','CANCELLED','FUTURED','INACTIVE ITEM') AND action not like ('SUB ITEM%'))) ");
         sql.append("group by dept_num ");
         m_DeptSvcLvl = m_EdbConn.prepareStatement(sql.toString());

         m_ItemWhs = m_EdbConn.prepareStatement("select * from ejd_item_warehouse " +
               "join item_entity_attr on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id " +
               "and item_entity_attr.item_id = ? WHERE warehouse_id = ? and can_plan = 1");

         m_Whs = m_EdbConn.prepareStatement("select * from warehouse order by warehouse_id");

         m_ItemDCQty = m_EdbConn.prepareStatement("select avail_qty as available_qty from ejd_item_warehouse " +
               "join item_entity_attr on item_entity_attr.ejd_item_id = ejd_item_warehouse.ejd_item_id " +
               "where item_entity_attr.item_id = ? and warehouse_id = (select warehouse_id from warehouse where name = ?) ");

         sql.setLength(0);
         sql.append("select po_hdr.po_hdr_id, ");
         sql.append("  po_hdr.po_nbr, ");
         sql.append("  to_char(po_hdr.po_date, 'mm/dd/yyyy') as po_date, ");
         sql.append("  po_dtl.qty_ordered, ");
         sql.append("  po_dtl.qty_put_away, ");
         sql.append("  to_char(po_dtl.sched_in_date, 'mm/dd/yyyy') as due_in ");
         sql.append("from po_hdr ");
         sql.append("join po_dtl on po_dtl.po_hdr_id = po_hdr.po_hdr_id ");
         sql.append("where ");
         sql.append("   po_hdr.status <> 'CLOSED' and po_hdr.cancelled = 'N' and ");
         sql.append("   po_dtl. item_ea_id = ? and po_dtl.warehouse = ? and po_dtl.status <> 'CLOSED' and po_dtl.cancelled = 'N' ");
         sql.append("order by po_hdr.po_date asc ");
         sql.append("limit 1 ");
         m_OldestPO = m_EdbConn.prepareStatement(sql.toString());


         return true;
      }

      catch ( Exception e ) {
         log.error("[CutReport]", e);
         return false;
      }

      finally {
         sql = null;
      }
   }

   /**
    * Sets the parameters for the report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      m_BegDate = params.get(0).value;
      m_EndDate = params.get(1).value;
      m_Warehouse = params.get(2).value;
      m_Buyer = params.get(3).value;
      m_Vendor = params.get(4).value;
   }

   /**
    * Creates a vendor header row
    * @param vendorId String - the vendor id
    * @param name String - the vendor name
    * @param svcLvl double - the vendor's line service level on this date
    * @throws Exception
    */
   private void vendorHeader(String vendorId, String name, double svcLvl) throws Exception
   {
      short col = 0;

      m_Row = m_Sheet.createRow(m_RowNum++);
      createCell(m_Row, col++, vendorId, m_StyleTitle);
      createCell(m_Row, col++, name + " - " + svcLvl + "%", m_StyleTitle);
   }

   /**
    * Creates a trailer line for an vendor
    *
    * @param itemQtyOrd int - the total qty ordered for the item
    * @param itemQtyShip - the total qty shipped for the item
    * @param itemLineCnt - the number of lines cut
    * @param itemAmtCut - the value of the items cut
    */
   private void vendorTrailer(int qtyOrd, int qtyShip, int lineCnt, double amtCut)
   {
      short col = 3;

      m_Row = m_Sheet.createRow(m_RowNum++);
      createCell(m_Row, col++, "Vendor Totals", m_StyleBoldRight);
      createCell(m_Row, col++, qtyOrd, m_StyleBoldRight);
      createCell(m_Row, col++, qtyShip, m_StyleBoldRight);
      createCell(m_Row, col++, lineCnt, m_StyleBoldRight);
      createCell(m_Row, col++, amtCut, m_StyleDecBold);
      m_RowNum++;
   }

}

