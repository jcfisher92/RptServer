/**
 * Title: POExpedite.java
 * Description: POExpedite report used by buyers.
 * <p>
 * Input parameters to this report:
 *    1. Emery dept# (emery_dept.dept_id)
 *    2. Vendor#     (vendor.vendor_id)
 *    3. Warehouse#  (warehouse.warehouse_id or ALL)
 * <p>
 * NOTE:
 *    - If vendor input param = 0, then get all vendors within the input department
 *    - If warehouse input param = 0, get available item qty for all warehouses
 * <p>
 * Company: Emery-Waterhouse
 * @author Paul Davidson
 * @version 1.0
 * <p>
 * Create Date: June 30, 2008
 * Last Update: $Id: POExpedite.java,v 1.20 2014/02/14 14:54:41 jfisher Exp $
 * <p>
 * History:
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class POExpedite extends Report
{

   private PreparedStatement m_AllFacilities; // Gets all whse facilities
   private SimpleDateFormat m_DateFmt;        // Date formatter
   private String m_EmDeptId;                 // Emery dept# input paramter
   private PreparedStatement m_Facility;      // Gets whse facility number
   private ArrayList<String> m_FacilityList;  // List of all whse facilities
   private String m_FacilityName;             // Name of current facility
   private XSSFFont m_Font;                   //
   private XSSFFont m_FontTitle;              //
   private XSSFFont m_FontTitleUndrLn;        //
   private XSSFFont m_FontBold;               //
   private XSSFFont m_FontData;               //
   private Header m_Header;               //
   private PreparedStatement m_PODtl;         // Gets PO data requested
   private XSSFRow m_Row = null;              //
   private int m_RowNum = 0;                //
   private XSSFSheet m_Sheet;                 //
   private XSSFCellStyle m_StyleBorderBttm;   // Bottom border cell style
   private XSSFCellStyle m_StyleText;  		 // Text left justified
   private XSSFCellStyle m_StyleTextRight;  	 // Text right justified
   private XSSFCellStyle m_StyleTextCenter; 	 // Text centered
   private XSSFCellStyle m_StyleTitle; 		 // Bold, centered
   private XSSFCellStyle m_StyleTitleUndrLn;  // Bold, underlined
   private XSSFCellStyle m_StyleBold;  		 // Normal but bold
   private XSSFCellStyle m_StyleBoldRight; 	 // Normal but bold & right aligned
   private XSSFCellStyle m_StyleBoldCenter; 	 // Normal but bold & centered
   private XSSFCellStyle m_StyleDec;   		 // Style with 2 decimals
   private XSSFCellStyle m_StyleDecBold;		 // Style with 2 decimals, bold
   private XSSFCellStyle m_StyleHeader; 		 // Bold, centered 12pt
   private XSSFCellStyle m_StyleInt;   		 // Style with 0 decimals
   private String m_VendorId;                 // Vendor# input paramter
   private String m_VendPhone;                // Vendor contact phone#
   private String m_VendPhExt;                // Vendor phone extension
   private PreparedStatement m_VndContact;    // Gets vendor contact info
   private PreparedStatement m_ItemDCQty;
   private String m_WhseId;                   // Warehouse input parameter
   private XSSFWorkbook m_WrkBk;              //

   /**
    * Default constructor.
    */
   public POExpedite()
   {
      super();

      m_DateFmt = new SimpleDateFormat("MM/dd/yyyy");
      m_FacilityList = new ArrayList<String>();
      m_FacilityName = "";
      m_VendPhone = "";
      m_VendPhExt = "";
   }

   /**
    * Builds the spreadsheet file for the PO expedite report.
    *
    * @return boolean.  True if the file was created, false if not.
    * @throws FileNotFoundException
    */
   public boolean buildOutputFile() throws FileNotFoundException
   {
      int charWidth;
      int col;
      String facility;
      StringBuffer hdrStuff = new StringBuffer();
      String itemId;
      int idx = 0;
      int numCols = 15;
      FileOutputStream outFile = null;
      long poHdrId = -1;
      Date poDate = null;
      String poNum = null;
      long priorPoHdrId = -1;
      long priorVndId = -1;
      Date rcptDate = null;
      CellRangeAddress region = null;
      boolean result = true;
      int rowCount = 0;
      ResultSet rs = null;
      boolean showAllVendors; // If true then get all vendors within the input dept
      long vndId;

      charWidth = 256;
      m_FileNames.add(m_RptProc.getUid() + "poexpedite" + getStartTime() + ".xlsx");
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      initReport();

      try {
         setCurAction("Running the po expedite query");

         showAllVendors = m_VendorId.equals("0");

         if ( !showAllVendors )
            loadVndContact(Long.parseLong(m_VendorId));

         loadFacilityList();
         facility = getFacility();
         numCols = m_WhseId.equals("0") ? numCols+m_FacilityList.size() : numCols+1;

         //
         // Create headers for column data, and make it print on every page
         m_Row = m_Sheet.createRow(m_RowNum);
         col = 0;

         createCell(m_Row, col, "Vendor#", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (8*charWidth));
         col++;

         createCell(m_Row, col, "Vendor Name", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (19*charWidth));
         col++;

         createCell(m_Row, col, "Vendor Phone#", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (11*charWidth));
         col++;

         createCell(m_Row, col, "PO#", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (8*charWidth));
         col++;

         createCell(m_Row, col, "PO Date", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (9*charWidth));
         col++;

         createCell(m_Row, col, "Due In", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (9*charWidth));
         col++;

         createCell(m_Row, col, "Whse", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (9*charWidth));
         col++;

         createCell(m_Row, col, "Item#", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (7*charWidth));
         col++;

         createCell(m_Row, col, "Description", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (20*charWidth));
         col++;

         createCell(m_Row, col, "UOM", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (4*charWidth));
         col++;

         createCell(m_Row, col, "Qty Ord", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (7*charWidth));
         col++;

         createCell(m_Row, col, "Qty Recv", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (7*charWidth));
         col++;

         createCell(m_Row, col, "Outstanding", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (7*charWidth));
         col++;

         if ( m_WhseId.equals("0") ) {  // If whseId == 0, then user wants all warehouses
            for (String fac:  m_FacilityList) {
               createCell(m_Row, col, fac.substring(fac.indexOf("|")+1) + " OH", m_StyleTitleUndrLn);
               m_Sheet.setColumnWidth(col, (11*charWidth));
               col++;
            }
         }
         else {
            createCell(m_Row, col, m_FacilityName + " OH", m_StyleTitleUndrLn);
            m_Sheet.setColumnWidth(col, (11*charWidth));
            col++;
         }

         createCell(m_Row, col, "Item Level Comment", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (50 * charWidth));
         col++;

         createCell(m_Row, col, "Internal Comment", m_StyleTitleUndrLn);
         m_Sheet.setColumnWidth(col, (50 * charWidth));
         col++;

         
         // Set the column heading row to repeat on each page         
         m_Sheet.setRepeatingRows(CellRangeAddress.valueOf("1:1"));
         m_Sheet.setRepeatingColumns(CellRangeAddress.valueOf("A:A"));
         m_RowNum++;

         m_PODtl.setInt(++idx, Integer.parseInt(m_EmDeptId));

         if ( !showAllVendors ) {
            m_PODtl.setLong(++idx, Long.parseLong(m_VendorId));
         }

         if ( !m_WhseId.equals("0") ) {
            m_PODtl.setInt(++idx, Integer.parseInt(m_WhseId));
         }

         rs = m_PODtl.executeQuery();

         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            rowCount++;
            vndId = rs.getLong("vendor_id");

            if ( showAllVendors )
              loadVndContact(vndId);

            //
            // If at a new vendor, then add a page break
            if ( showAllVendors && (priorVndId != -1 && vndId != priorVndId) ) {
               m_Sheet.setRowBreak(m_RowNum);
            }

            //
            // Set document header info for print date and Emery dept
            if ( rowCount == 1 ) {
               hdrStuff.setLength(0);
               hdrStuff.append("\nDept# ");
               hdrStuff.append(rs.getString("dept_num"));
               hdrStuff.append(" (");
               hdrStuff.append(rs.getString("dept_name"));
               hdrStuff.append(")");
               m_Header.setLeft(HSSFHeader.font("Arial", "Bold") + HSSFHeader.date() + hdrStuff.toString());
            }

            itemId = rs.getString("item_nbr");
            poHdrId = rs.getLong("po_hdr_id");
            poDate = rs.getDate("po_date");
            poNum = rs.getString("po_nbr");
            rcptDate = rs.getDate("sched_in_date");

            setCurAction("Processing PO# " + poNum);

            //
            // If at a new PO, then add a blank row with line to enhance readability
            if ( priorPoHdrId != -1 && poHdrId != priorPoHdrId ) {
               m_Row = m_Sheet.createRow(m_RowNum);
               region = new CellRangeAddress(m_RowNum, m_RowNum, 0, numCols);

               for ( int i = 0; i < numCols; i++ ) {
                  createCell(m_Row, i, "", m_StyleBorderBttm);
               }

               m_Sheet.addMergedRegion(region);
               m_RowNum++;
               m_Row = m_Sheet.createRow(m_RowNum++);

               for ( int i = 0; i < numCols; i++ ) {
                  createCell(m_Row, i, "", m_StyleText);
               }
            }

            priorPoHdrId = rs.getLong("po_hdr_id");
            priorVndId = vndId;

            col = 0;
            m_Row = m_Sheet.createRow(m_RowNum++);

            createCell(m_Row, col++, rs.getString("vendor_id"), m_StyleText);
            createCell(m_Row, col++, rs.getString("vendor_name"), m_StyleText);
            createCell(m_Row, col++, m_VendPhone + " " + m_VendPhExt, m_StyleText);
            createCell(m_Row, col++, rs.getString("po_nbr"), m_StyleText);
            createCell(m_Row, col++, (poDate == null ? "" : m_DateFmt.format(poDate)), m_StyleText);
            createCell(m_Row, col++, (rcptDate == null ? "" : m_DateFmt.format(rcptDate)), m_StyleText);
            createCell(m_Row, col++, rs.getString("whse"), m_StyleText);
            createCell(m_Row, col++, itemId, m_StyleText);
            createCell(m_Row, col++, rs.getString("descr"), m_StyleText);
            createCell(m_Row, col++, rs.getString("ship_unit"), m_StyleTextCenter);
            createCell(m_Row, col++, rs.getInt("qty_ordered"), m_StyleInt);
            createCell(m_Row, col++, rs.getInt("qty_put_away"), m_StyleInt);
            createCell(m_Row, col++, rs.getInt("outstanding"), m_StyleInt);

            if ( m_WhseId.equals("0") ) {  // If whseId == 0, then user wants all warehouses
               for (String fac:  m_FacilityList) {
                  createCell(m_Row, col++, getOnHand(itemId, fac.substring(0, fac.indexOf("|"))), m_StyleInt);
               }
            }
            else {
               createCell(m_Row, col++, getOnHand(itemId, facility), m_StyleInt);
            }

            createCell(m_Row, col++, rs.getString("comments"), m_StyleText);
            createCell(m_Row, col++, rs.getString("internal_comments"), m_StyleText);
         }

         m_WrkBk.write(outFile);
         setCurAction("Complete");
      }

      catch( Exception ex ) {
         log.error("POExpedite exception", ex);
         m_ErrMsg.append("The report had the following Error: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         result = false;
      }

      finally {
      	closeRSet(rs);
      	rs = null;
      	hdrStuff = null;

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
    * Resource cleanup
    */
   public void cleanup()
   {
   	DbUtils.closeDbConn(null, m_PODtl, null);
   	DbUtils.closeDbConn(null, m_Facility, null);
   	DbUtils.closeDbConn(null, m_AllFacilities, null);
   	DbUtils.closeDbConn(null, m_VndContact, null);
  	DbUtils.closeDbConn(null, m_ItemDCQty, null);

   	m_Facility = null;
   	m_PODtl = null;
   	m_AllFacilities = null;
   	m_VndContact = null;

   	m_Header = null;
   	m_Font = null;
   	m_FontTitle = null;
   	m_FontTitleUndrLn = null;
   	m_FontBold = null;
   	m_FontData = null;
   	m_StyleBorderBttm = null;
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

   	m_WrkBk = null;
   	m_Sheet = null;
   	m_Row = null;
   	m_DateFmt = null;
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
         m_EdbConn = m_RptProc.getEdbConn();

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
    * Gets the Fascor facility number of the current warehouse.
    *
    * @return String - Facility number.
    * @throws SQLException
    */
   private String getFacility() throws SQLException
   {
      String facility = "01";
      ResultSet rs = null;

      try {
         m_Facility.setInt(1, Integer.parseInt(m_WhseId));
         rs = m_Facility.executeQuery();

         if ( rs.next() ) {
            facility = rs.getString("fas_facility_id");

            //
            // Since we're here just get the name as well
            m_FacilityName = rs.getString("name");
         }

         return facility;
      }
      finally {
         closeRSet(rs);
         rs = null;
      }
   }

   /**
    * Returns the on hand quantity of an item at a facility.
    *
    * @param item String - the item id
    * @param facility String - the fascor facility id
    * @return int - the quantity on hand.
    * @throws Exception
    */
   private int getOnHand(String itemId, String facilityId) throws Exception
   {
      int qty = 0;
      ResultSet rset = null;

      if ( itemId != null && itemId.length() == 7 ) {
         try {
            m_ItemDCQty.setString(1, itemId);

         	if( facilityId.equals("01") || facilityId.equals("02") )
               m_ItemDCQty.setString(2, "PORTLAND");
            else {
            	if( facilityId.equals("04") || facilityId.equals("05") )
                  m_ItemDCQty.setString(2, "PITTSTON");
            }

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

      try {
         m_WrkBk = new XSSFWorkbook();

         df = m_WrkBk.createDataFormat();

         //
         // Create the default font for this workbook
         m_Font = m_WrkBk.createFont();
         m_Font.setFontHeightInPoints((short)8);
         m_Font.setFontName("Arial");

         //
         // Create a font for titles
         m_FontTitle = m_WrkBk.createFont();
         m_FontTitle.setFontHeightInPoints((short)10);
         m_FontTitle.setFontName("Arial");
         m_FontTitle.setBold(true);;

         //
         // Create underlined title font
         m_FontTitleUndrLn = m_WrkBk.createFont();
         m_FontTitleUndrLn.setFontHeightInPoints((short)8);
         m_FontTitleUndrLn.setFontName("Arial");
         m_FontTitleUndrLn.setBold(true);;
         m_FontTitleUndrLn.setUnderline(XSSFFont.U_SINGLE);

         //
         // Create a font that is normal size & bold
         m_FontBold = m_WrkBk.createFont();
         m_FontBold.setFontHeightInPoints((short)8);
         m_FontBold.setFontName("Arial");
         m_FontBold.setBold(true);;

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

         m_StyleBorderBttm = m_WrkBk.createCellStyle();
         m_StyleBorderBttm.setBorderBottom(BorderStyle.THIN);

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
         m_StyleBoldCenter.setAlignment(HorizontalAlignment.RIGHT);

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
         m_StyleHeader.setAlignment(HorizontalAlignment.RIGHT);

         m_StyleInt = m_WrkBk.createCellStyle();
         m_StyleInt.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleInt.setFont(m_FontData);
         m_StyleInt.setDataFormat((short)3);

         m_StyleTitle = m_WrkBk.createCellStyle();
         m_StyleTitle.setFont(m_FontTitle);
         m_StyleTitle.setAlignment(HorizontalAlignment.LEFT);

         m_StyleTitleUndrLn = m_WrkBk.createCellStyle();
         m_StyleTitleUndrLn.setFont(m_FontTitleUndrLn);
         m_StyleTitleUndrLn.setAlignment(HorizontalAlignment.LEFT);

         m_Sheet = m_WrkBk.createSheet();
         m_Sheet.getPrintSetup().setLandscape(true);
         m_Sheet.getPrintSetup().setPaperSize(XSSFPrintSetup.LEGAL_PAPERSIZE);
         m_Sheet.getFooter().setRight("Page " + HSSFFooter.page() + " of " + HSSFFooter.numPages());

         m_Header = m_Sheet.getHeader();
         m_Header.setCenter(HSSFHeader.font("Arial", "Bold") + HSSFHeader.fontSize((short) 12) + "Purchase Order Expedite Report");

         m_RowNum = 0;
      }

      catch ( Exception e ) {
         log.error("POExpedite.initReport()", e);
      }
   }

   /**
    * Loads list of whse facility ids and names.
    *
    * @throws SQLException
    */
   private void loadFacilityList() throws SQLException
   {
      ResultSet rs = null;

      try {
         m_FacilityList.clear();

         rs = m_AllFacilities.executeQuery();

         while ( rs.next() ) {
            m_FacilityList.add(rs.getString("fas_facility_id") + "|" + rs.getString("name"));
         }
      }
      finally {
         closeRSet(rs);
         rs = null;
      }
   }

   /**
    * Loads contact info of current vendor.
    *
    * @throws SQLException
    */
   private void loadVndContact(long vendId) throws SQLException
   {
      ResultSet rs = null;

      try {
         m_VndContact.setLong(1, vendId);
         rs = m_VndContact.executeQuery();

         if ( rs.next() ) {
            m_VendPhone = rs.getString("phone_number");
            m_VendPhExt = rs.getString("extension");
         }

         m_VendPhone = m_VendPhone == null ? "" : m_VendPhone;
         m_VendPhExt = m_VendPhExt == null ? "" : m_VendPhExt;
      }
      finally {
         closeRSet(rs);
         rs = null;
      }
   }

   /**
    * Prepares any SQL statements used by this report.
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer();
      //String orderBy = "order by ";
      String orderBy = "order by po_hdr.vendor_name, po_hdr.warehouse, (po_hdr.po_date, po_hdr.po_nbr), po_dtl.descr ASC";

      try {
         sql.append("select ");
         sql.append("   po_hdr.po_hdr_id, ");
         sql.append("   po_hdr.po_nbr, ");
         sql.append("   po_hdr.po_date, ");
         sql.append("   po_hdr.vendor_id, ");
         sql.append("   po_hdr.vendor_name, ");
         sql.append("   po_hdr.warehouse, ");
         sql.append("   po_dtl.sched_in_date, ");
         sql.append("   po_dtl.item_nbr, ");
         sql.append("   po_dtl.descr, ");
         sql.append("   po_dtl.ship_unit, ");
         sql.append("   po_dtl.qty_ordered, ");
         sql.append("   po_dtl.qty_put_away, ");
         sql.append("   (po_dtl.qty_ordered - po_dtl.qty_put_away) as outstanding, ");
         sql.append("   po_dtl.comments, ");
         sql.append("   po_hdr.internal_comments, ");
         sql.append("   emery_dept.dept_num, emery_dept.name as dept_name, ");
         sql.append("   (select whs.name from warehouse whs where whs.fas_facility_id = po_hdr.warehouse) as whse ");
         sql.append("from ");
         sql.append("   po_hdr ");
         sql.append("   inner join po_dtl on (po_hdr.po_hdr_id = po_dtl.po_hdr_id and po_dtl.status = 'OPEN') ");
         sql.append("   inner join vendor_dept on (po_hdr.vendor_id = vendor_dept.vendor_id) ");
         sql.append("   inner join emery_dept on (vendor_dept.dept_id = emery_dept.dept_id) ");
         sql.append("where ");
         sql.append("   po_hdr.status = 'OPEN' and ");
         sql.append("   emery_dept.dept_id = ? ");

         if ( !m_VendorId.equals("0") ) {
            sql.append(" and po_hdr.vendor_id = ? ");
            //orderBy += "   po_hdr.po_date ASC ";
         }

         if ( !m_WhseId.equals("0") ) {
            sql.append(" and po_hdr.warehouse = (select fas_facility_id from warehouse where warehouse_id = ?) ");
            //if ( !m_VendorId.equals("0") ) {
            //	orderBy += "   ,po_hdr.vendor_id ";
           // }
          //  else {
          //  	orderBy += "   po_hdr.vendor_id ";
           // }
         }

         sql.append(orderBy);
         m_PODtl = m_EdbConn.prepareStatement(sql.toString());

         m_Facility = m_EdbConn.prepareStatement("select fas_facility_id, name from warehouse where warehouse_id = ?");

         m_AllFacilities = m_EdbConn.prepareStatement("select fas_facility_id, name from warehouse where warehouse_id in (1, 2) order by warehouse_id"); // remove where clause to include Ace warehouses

         sql.setLength(0);
         sql.append("select ");
         sql.append("   phone_number, ");
         sql.append("   extension ");
         sql.append("from ");
         sql.append("   vendor_contact_view ");
         sql.append("where ");
         sql.append("   vendor_id = ? and ");
         sql.append("   phone_type = 'BUSINESS' and ");
         sql.append("   is_default = 1");
         m_VndContact = m_EdbConn.prepareStatement(sql.toString());

         sql.setLength(0);
         sql.append("select qoh as available_qty ");
         sql.append("from ejd_item_warehouse ");
         sql.append("where ejd_item_id = ? and warehouse_id = (select warehouse_id from warehouse where name = ?) ");
         m_ItemDCQty = m_EdbConn.prepareStatement(sql.toString());

         return true;
      }

      catch ( Exception e ) {
         log.error("[PO Expedite]", e);
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
   @Override
   public void setParams(ArrayList<Param> params)
   {
      m_EmDeptId = params.get(0).value;
      m_VendorId = params.get(1).value;
      m_WhseId = params.get(2).value;
   }
}