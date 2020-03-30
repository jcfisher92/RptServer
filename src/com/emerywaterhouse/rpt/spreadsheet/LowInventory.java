/**
 * File: LowInventory.java
 * Description: Reports on low inventory based on demand streams.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 05/09/2007
 * Last Update: $Id: LowInventory.java,v 1.19 2014/04/16 18:05:46 tli Exp $
 *
 * History
 *    $Log: LowInventory.java,v $
 *    Revision 1.19  2014/04/16 18:05:46  tli
 *    Tweaks on reports
 *
 *    Revision 1.18  2014/04/15 20:28:46  tli
 *    Added ItemVelocityProject and PromoServiceLevel reports
 *
 *    Revision 1.17  2013/09/09 18:33:38  tli
 *    Replace SkuQty web service call with item_qty_view
 *
 *    Revision 1.16  2012/08/29 19:53:02  jfisher
 *    Switched web service calls from Wasp to Axis2
 *
 *    Revision 1.15  2012/05/05 06:07:17  pberggren
 *    Removed redundant loading of system properties.
 *
 *    Revision 1.14  2012/05/03 07:55:10  prichter
 *    Fix to web service ip address
 *
 *    Revision 1.13  2012/05/03 04:35:20  pberggren
 *    Added server.properties call to force report to .57
 *
 *    Revision 1.12  2012/05/03 04:25:01  pberggren
 *    Added server.properties call to force report to .57
 *
 *    Revision 1.11  2009/02/18 17:23:35  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.10  2008/10/29 21:24:26  jfisher
 *    Fixed some warnings
 *
 *    Revision 1.9  2008/08/01 23:28:09  smurdock
 *    added query by dc
 *
 *    Revision 1.8  2007/12/12 06:19:39  jfisher
 *    Removed unused vars
 *
 *    Revision 1.7  2007/07/25 22:43:24  smurdock
 *    put back exclusion of promo sales for regular demand queries oops
 *
 *    Revision 1.6  2007/07/17 18:33:58  smurdock
 *    added select by department
 *
 *    Revision 1.5  2007/05/31 19:21:05  jfisher
 *    Set the max run to 12 hours.
 *
 *    Revision 1.4  2007/05/31 19:17:47  jfisher
 *    Added the ATP back to the caption, fixed some formatting and removed the unused m_RegDemand var
 *
 *    Revision 1.3  2007/05/31 18:58:10  smurdock
 *    altered the regdemand query to run once inside the item query instead of separately for each item.
 *
 *    Revision 1.1  2007/05/23 17:54:00  jfisher
 *    initial add
 *
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class LowInventory extends Report
{
   private static final int MAX_COLS = 17;
   private static final int MAX_VALUE  = 52000;

   private PreparedStatement m_PoQty;
   private PreparedStatement m_SaleData;
   private PreparedStatement m_GetDeptId;
   private PreparedStatement m_GetDCName;

   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;

   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;

   //
   // Report params
   private int m_VndId;
   private int m_DmdWks;
   private int m_DeptId;
   private String m_DeptNum = "";
   private boolean m_UsePo;
   private boolean m_UseRegDmd;
   private double m_MaxWos;
   private String m_Warehouse;  //FASCOR id , sez Jeff.  From Delphi.
   private String m_Warehouse_Name;

   private PreparedStatement m_ItemDCQty;
   
   /**
    *
    */
   public LowInventory()
   {
      super();

      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      m_MaxRunTime = RptServer.HOUR * 12;

      setupWorkbook();
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable
    */
   @Override
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
      ResultSet saleData = null;
      ResultSet deptData = null;
      ResultSet DCData = null;
      SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
      boolean result = false;
      String itemId = null;
      String setupDate = null;
      int endDate = 0;
      int buyMult = 0;
      int qtyOnHand = 0;
      int totUnits = 0;
      double estUnitsShort = 0;
      double wosOnHand = 0.0;
      double wos1 = 0.0;
      double cost = 0.0;
      double estPctStock = 0.0;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         if (!(m_Warehouse == "")){
             m_GetDCName.setString(1,m_Warehouse);
             DCData = m_GetDCName.executeQuery();
             while ( DCData.next() && m_Status == RptServer.RUNNING ) {
                m_Warehouse_Name = DCData.getString("name");
             }
         }
         rowNum = createCaptions();

         if (!(m_DeptNum == "")){
             m_GetDeptId.setString(1,m_DeptNum);
             m_DeptId = 999;  //if a department has been entered, first default to a nonexistent dept_id
             deptData = m_GetDeptId.executeQuery();
             while ( deptData.next() && m_Status == RptServer.RUNNING ) {
                m_DeptId = deptData.getInt("dept_id");
             }
          }
          else
         	 m_DeptId = 0;  //no dept entered


         endDate = (m_DmdWks * 7) + 1;
         if (m_UseRegDmd) { // regular demand
            if ( m_DeptId == 0){ // regular demand, no dept
               if ( m_VndId == 0){//regular demand, no dept, no vendor
                  m_SaleData.setInt(1, endDate);
                  m_SaleData.setInt(2, endDate);
                  m_SaleData.setInt(3, endDate);
               }
               else {  //regular demand, no dept, yes vendor
                  m_SaleData.setInt(1, endDate);
                  m_SaleData.setInt(2, m_VndId);
                  m_SaleData.setInt(3, endDate);
                  m_SaleData.setInt(4, m_VndId);
                  m_SaleData.setInt(5, endDate);
                  m_SaleData.setInt(6, m_VndId);
               }
            }
            else {//regular demand, yes dept
               if ( m_VndId == 0){//regular demand, yes dept, no vendor
                  m_SaleData.setInt(1, endDate);
                  m_SaleData.setInt(2, m_DeptId);
                  m_SaleData.setInt(3, endDate);
                  m_SaleData.setInt(4, endDate);
                  m_SaleData.setInt(5, m_DeptId);
               }
               else {  //regular demand, yes dept, yes dept, yes vendor
                  m_SaleData.setInt(1, endDate);
                  m_SaleData.setInt(2, m_VndId);
                  m_SaleData.setInt(3, m_DeptId);
                  m_SaleData.setInt(4, endDate);
                  m_SaleData.setInt(5, m_VndId);
                  m_SaleData.setInt(6, endDate);
                  m_SaleData.setInt(7, m_VndId);
                  m_SaleData.setInt(8, endDate);

              }
            }
         }
         else{ // no regular demand
             if ( m_DeptId == 0){ // no regular demand, no dept
                 if ( m_VndId == 0){//regular demand, no dept, no vendor
                    m_SaleData.setInt(1, endDate);
                 }
                 else {  //no regular demand, no dept, yes vendor
                    m_SaleData.setInt(1, endDate);
                    m_SaleData.setInt(2, m_VndId);
                 }
              }
              else {//no regular demand, yes dept
                 if ( m_VndId == 0){//no regular demand, yes dept, no dept
                    m_SaleData.setInt(1, endDate);
                    m_SaleData.setInt(2, m_DeptId);
                 }
                 else {  //no regular demand, yes dept, yes dept, yes vendor
                    m_SaleData.setInt(1, endDate);
                    m_SaleData.setInt(2, m_VndId);
                    m_SaleData.setInt(3, m_DeptId);

                }
              }


         }

         saleData = m_SaleData.executeQuery();

         while ( saleData.next() && m_Status == RptServer.RUNNING ) {
            itemId = saleData.getString("item_id");
            setCurAction("processing item: " + itemId);

            qtyOnHand = getQtyOnHand(itemId);
            buyMult = saleData.getInt("buy_mult");
            totUnits = m_UseRegDmd ? saleData.getInt("reg_dmd_qty") : saleData.getInt("tot_units");
            cost = saleData.getDouble("cur_cost");
            setupDate = sdf.format(saleData.getDate("setup_date"));

            //
            // Do the calculations
            wos1 = ((double)totUnits)/m_DmdWks;
            wosOnHand = wos1 == 0 ? MAX_VALUE : (qtyOnHand/wos1);
            Math.round(estUnitsShort = (m_MaxWos - wosOnHand) * wos1);

            if ( qtyOnHand == 0 && estUnitsShort == 0 )
               estPctStock = 1;
            else
               estPctStock = qtyOnHand / (qtyOnHand + estUnitsShort);

            //FOLLOWING LINE TEST ONLY
            //m_MaxWos = 10000000;
            if ( wosOnHand <= m_MaxWos ) {
               row = createRow(rowNum++, MAX_COLS);
               colNum = 0;

               if ( row != null ) {
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(itemId));
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(setupDate));
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(saleData.getString("description")));
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(saleData.getString("name")));
                  row.getCell(colNum++).setCellValue(saleData.getInt("stock_pack"));
                  row.getCell(colNum++).setCellValue(
                     new XSSFRichTextString(saleData.getInt("broken_case_id") == 1 ? "N" : "Y")
                  );
                  row.getCell(colNum++).setCellValue(buyMult);
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString(saleData.getString("dept_num")));
                  row.getCell(colNum++).setCellValue(saleData.getDouble("tot_sales"));
                  row.getCell(colNum++).setCellValue(saleData.getInt("tot_lines"));
                  row.getCell(colNum++).setCellValue(totUnits);
                  row.getCell(colNum++).setCellValue(qtyOnHand);
                  row.getCell(colNum++).setCellValue(cost);
                  row.getCell(colNum++).setCellValue(wos1);
                  row.getCell(colNum++).setCellValue(wosOnHand);
                  row.getCell(colNum++).setCellValue(estUnitsShort);
                  row.getCell(colNum++).setCellValue(estPctStock);
               }
            }
         }

         m_Sheet.createFreezePane(1, 3);
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
      closeStmt(m_SaleData);
      closeStmt(m_PoQty);
      closeStmt(m_GetDeptId);
      closeStmt(m_GetDCName);
      closeStmt(m_ItemDCQty);
      
      m_SaleData = null;
      m_PoQty = null;
      m_GetDeptId = null;
      m_GetDCName = null;
      m_ItemDCQty = null; 
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
      StringBuffer title = new StringBuffer("Low Inventory Report: ");

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
      title.append(" Demand Weeks = ");
      title.append(m_DmdWks);
      title.append(" Regular Demand - ");
      title.append(m_UseRegDmd ? "Yes" : "No");


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
      if ( m_VndId > 0 ) {
         title.append("Filter: Vendor = ");
         title.append(m_VndId);
      }
      else
         title.append("Filter: All Vendors");
      if ( m_DeptNum.compareTo("00")> 0){
         title.append(" Dept = ");
         title.append(m_DeptNum);
      }
      title.append(" Max WOS = ");
      title.append(m_MaxWos);
      title.append(" Use PO Qty - ");
      title.append(m_UsePo ? "Yes" : "No");

      if ((m_Warehouse_Name != null) && (m_Warehouse_Name.length() > 0)) {
         title.append(" DC = ");
         title.append(m_Warehouse_Name);

      }

      //
      // Merge the title cells.  Gives a better look to the report.
      region = new CellRangeAddress(0, 0, 0, 4);
      m_Sheet.addMergedRegion(region);

      region = new CellRangeAddress(1, 1, 0, 4);
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
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("ATP Inventory"));
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Current\nCost"));
      row.getCell(colNum).setCellValue(new XSSFRichTextString("1 Week of\nSupply\n(1-WOS Units)"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("ATP\nWeeks\nof Supply\nOn-hand"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Est\nUnits\nShort"));
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Est\n%-Instock"));
      m_Sheet.setColumnWidth(colNum++, 3000);

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
    * Uses a web service to get the qty on hand in fascor for the specific item and
    * any open POs if that option is being used.
    *
    * @param item The item number to check the quantities
    * @return The quantity on hand in fascor
    * @throws Exception
    */
   private int getQtyOnHand(String item) throws Exception
   {
	   int qty = 0;
	   ResultSet rset = null;
	      
	   if ( item != null && item.length() == 7 ) {
	         try {
	        	 m_ItemDCQty.setString(1, item);	        	 
	             rset = m_ItemDCQty.executeQuery();

	            if ( rset.next() )
	            	qty = rset.getInt("qoh");
	         }
	         finally {
	            closeRSet(rset);
	            rset = null;
	         }
	      }
	   if ( m_UsePo ) {
	         m_PoQty.setString(1, item);

	         try {
	        	 rset = m_PoQty.executeQuery();

	            if ( rset.next() )
	               qty += rset.getInt(1);
	         }

	         finally {
	            closeRSet(rset);
	            rset = null;
	         }
	      }

	      return qty;
	    
   }
  

   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(256);
      StringBuffer sqld = new StringBuffer(256);
      StringBuffer sqlDeptId = new StringBuffer(256);
      StringBuffer sqlDCName = new StringBuffer(256);
      boolean isPrepared = false;

      if ( m_EdbConn != null ) {
         try {
            sqld.append(" join (select vendor_nbr, item_nbr, sum(qty_shipped) as shipd_qty, nvl(kit_qty::integer, 0) as kitd_qty \r\n");
            //,review_cycle
            //,dept_num
            sqld.append(" from inv_dtl \r\n");
            if ((m_Warehouse != null) && m_Warehouse.length() > 0) {
               sqld.append("   join warehouse on inv_dtl.warehouse = warehouse.name and warehouse.fas_facility_id = ");
               sqld.append(m_Warehouse);
               sqld.append(" \r\n");
            }
            sqld.append("      join inv_hdr \r\n");
            sqld.append("      ON inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id \r\n");
            sqld.append("      AND ((inv_hdr.order_type is null) or (inv_hdr.order_type <> 'STOCK')) \r\n");
            sqld.append("      AND inv_hdr.cust_nbr not in ('199818','199796') \r\n");
            sqld.append("   join item_entity_attr \r\n");
            sqld.append("     ON item_entity_attr.item_ea_id = inv_dtl.item_ea_id \r\n");
            if ( m_VndId > 0 )
               sqld.append("   AND item_entity_attr.vendor_id = ? \r\n");
            sqld.append(" join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            if ( m_DeptNum.compareTo("00")> 0)
               sqld.append("     AND  ejd_item.dept_id = ? \r\n");
            sqld.append(" join ejd_item_warehouse \r\n");
 		      sqld.append(" on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id \r\n");
 		      sqld.append(" AND ejd_item_warehouse.disp_id = 1 \r\n");
 		      sqld.append(" AND ejd_item_warehouse.warehouse_id = warehouse.warehouse_id \r\n");
            sqld.append("   join item_type on \r\n");
            sqld.append("     (item_type.item_type_id = item_entity_attr.item_type_id \r\n");
            sqld.append("      AND itemtype = 'STOCK' \r\n");
            sqld.append("      AND  flc_id < '9998') \r\n");
            sqld.append("      OR \r\n");
            sqld.append("      (item_type.item_type_id = item_entity_attr.item_type_id \r\n");
            sqld.append("     AND itemtype = 'TOOL PART') \r\n");
            sqld.append("   left outer join \r\n");
            sqld.append("      (select trim(itemno) as kit_item_id, \r\n");
            sqld.append("       sum(quantity) as kit_qty  \r\n");
            sqld.append("       FROM ejd.sage300_icaded_mv icaded \r\n");
            sqld.append("       join ejd.sage300_icadedo_mv icadedo \r\n");
            sqld.append("       ON icadedo.adjenseq = icaded.adjenseq \r\n");
            sqld.append("       AND  icadedo.lineno = icaded.lineno \r\n");
            sqld.append("       AND optfield = 'REASONCODE'  \r\n");
            sqld.append("       AND value in ('70','71') \r\n");
            sqld.append("       join ejd.sage300_icadeh_mv icadeh \r\n");
            sqld.append("       ON icadeh.adjenseq = icaded.adjenseq \r\n");
            sqld.append("       AND  icadeh.transdate > to_number(to_char(trunc(now()) - ?, 'yyyyMMdd')) \r\n");
            sqld.append("       group by trim(itemno)) kit_item \r\n");
            sqld.append("       on kit_item_id = item_nbr \r\n");
            sqld.append("    where \r\n");
            if ( m_VndId > 0 )
               sqld.append("       vendor_nbr = ? AND \r\n");
            sqld.append("      inv_dtl.invoice_date > trunc(now()) - ? \r\n");
            sqld.append("      AND inv_dtl.promo_nbr is null \r\n");
            sqld.append("      AND inv_dtl.sale_type = 'WAREHOUSE' \r\n");
            sqld.append("      AND inv_dtl.tran_type = 'SALE' \r\n");
            sqld.append("      AND inv_dtl.cust_nbr not in ('199818','199796') \r\n");
            sqld.append("      group by item_nbr, vendor_nbr, kit_qty) regdemand  \r\n");
            sqld.append("     on item_entity_attr.item_id = regdemand.item_nbr \r\n");

            sql.append("select \r\n");
            if ( m_UseRegDmd )
               sql.append("   (shipd_qty + kitd_qty) as reg_dmd_qty, \r\n");
            sql.append("   item_id, setup_date, description, vendor.name, ejd_item_warehouse.stock_pack, \r\n");
            sql.append("   broken_case_id, buy_mult, emery_dept.dept_num, \r\n");
            sql.append("   nvl(sum(ext_sell), 0) as tot_sales, \r\n");
            sql.append("   nvl(count(inv_dtl_id), 0) as tot_lines, \r\n");
            sql.append("   nvl(sum(qty_shipped), 0) as tot_units, \r\n");
            sql.append("   ejd_item_price.buy as cur_cost \r\n");
            sql.append("   from \r\n");
            sql.append("   item_entity_attr \r\n");
            sql.append("   join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id \r\n");
            sql.append("   join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id \r\n");
            sql.append("   join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id and ejd_item_price.warehouse_id = ejd_item_warehouse.warehouse_id \r\n");
            sql.append("   join item_type on \r\n");
            sql.append("     item_entity_attr.item_type_id = item_type.item_type_id \r\n");
            sql.append("      and item_type.itemtype = 'STOCK' \r\n");
            sql.append("   left outer join inv_dtl on \r\n");
            sql.append("      item_entity_attr.item_id = inv_dtl.item_nbr and \r\n");
            sql.append("	  inv_dtl.invoice_date > trunc(now()) - ? \r\n");
            sql.append("   join emery_dept on \r\n");
            sql.append("      ejd_item.dept_id = emery_dept.dept_id \r\n");
            sql.append("   join vendor on \r\n");
            sql.append("      item_entity_attr.vendor_id = vendor.vendor_id \r\n");
            if ( m_UseRegDmd ) {
              sql.append(sqld);
            }
            sql.append("where \r\n");
            if ( m_VndId > 0 )
               sql.append("   item_entity_attr.vendor_id = ? and \r\n");
            if ( m_DeptNum.compareTo("00")> 0)
               sql.append("  ejd_item.dept_id = ? and \r\n");
            sql.append("   ejd_item_warehouse.disp_id = 1 \r\n");
            sql.append(" group by  \r\n");
            sql.append("    item_id, setup_date, vendor.name, description, ejd_item_warehouse.stock_pack, broken_case_id, emery_dept.dept_num,  \r\n");
            sql.append("    buy_mult, ejd_item_price.buy \r\n");
            if ( m_UseRegDmd )
               sql.append(",shipd_qty, kitd_qty \r\n");
            sql.append("order by item_id \r\n");

            m_SaleData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select sum(qty_ordered) qty ");
            sql.append("from po_dtl ");
            sql.append("where item_nbr = ? and status = 'OPEN'");

            m_PoQty = m_EdbConn.prepareStatement(sql.toString());

            sqlDCName.setLength(0);
            sqlDCName.append("select name from warehouse where fas_facility_id = ?");
            m_GetDCName = m_EdbConn.prepareStatement(sqlDCName.toString());


            sqlDeptId.setLength(0);
            sqlDeptId.append("select dept_id from emery_dept where dept_num = ?");
            m_GetDeptId = m_EdbConn.prepareStatement(sqlDeptId.toString());
            
            
            
            m_ItemDCQty = m_EdbConn.prepareStatement("select sum(qoh) as qoh from ejd_item_warehouse where ejd_item_id = (select ejd_item_id from item_entity_attr where item_id = ?) and warehouse_id in (1, 2) ");

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
         log.error("InventoryAging.prepareStatements - null Edb connection");

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

         if ( param.name.equals("dmdwks") )
            m_DmdWks = Integer.parseInt(param.value);

         //
         // Incomming vendor ids may be blank.  A 0 means don't filter by vendor id.
         if ( param.name.equals("vendor") ) {
            if ( param.value.length() == 6 )
               m_VndId = Integer.parseInt(param.value);
            else
               m_VndId = 0;
         }
         //
         // User may choose to filter by dc
         if (param.name.equals("dc"))
            m_Warehouse = param.value;

         //
         // We deal with blank dept nums when we look up dept_ids
         if ( param.name.equals("dept") ) {
            if ( param.value.length() > 0 )
               m_DeptNum = (param.value);
            //else
               //m_DeptNum = "00";
         }


         if ( param.name.equals("maxwos") )
            m_MaxWos = Double.parseDouble(param.value);

         if ( param.name.equals("usepo") )
            m_UsePo = Boolean.parseBoolean(param.value);

         if ( param.name.equals("useregdmd") )
            m_UseRegDmd = Boolean.parseBoolean(param.value);
      }

      fileName.append("lowinv");
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
      XSSFCellStyle styleTxtC = null;      // Text centered
      XSSFCellStyle styleTxtL = null;      // Text left justified
      XSSFCellStyle styleInt = null;       // Style with 0 decimals
      XSSFCellStyle styleDouble = null;    // numeric #,##0.00
      XSSFCellStyle styleMoney = null;     // Money ($#,##0.00_);[Red]($#,##0.00)
      XSSFCellStyle stylePct = null;       // percentage
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

      stylePct = m_Wrkbk.createCellStyle();
      stylePct.setAlignment(HorizontalAlignment.RIGHT);
      stylePct.setDataFormat((short)9);

      m_CellStyles = new XSSFCellStyle[] {
         styleTxtC,    // col 0 item
         styleTxtC,    // col 1 date
         styleTxtL,    // col 2 descripton
         styleTxtL,    // col 3 vendor name
         styleTxtC,    // col 4 stock pack
         styleTxtC,    // col 5 broken case
         styleTxtC,    // col 6 buy mult
         styleTxtC,    // col 7 dept
         styleMoney,   // col 8 total sales
         styleInt,     // col 9 total lines shipped
         styleInt,     // col 10 tot units shipped
         styleInt,     // col 11 On hand qty
         styleMoney,   // col 12 Cost
         styleDouble,  // col 13 1-wos
         styleDouble,  // col 14 wos on hand
         styleInt,     // col 15 units short
         stylePct      // col 16 % in stock
      };

      styleTxtC = null;
      styleTxtL = null;
      styleInt = null;
      styleDouble = null;
      styleMoney = null;
      stylePct = null;
      format = null;
   }
}
