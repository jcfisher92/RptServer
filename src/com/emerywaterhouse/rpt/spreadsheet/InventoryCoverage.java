/**
 * File: InventoryCoverage.java
 * Description: Report that shows coverage for all active items
 *
 * @author Eric Verge, stolen from NeverOut.java by Jeff Fisher
 *
 * Create Date: 12/04/2014
 * Last Update: $Id: InventoryCoverage.java,v 1.3 2015/01/14 16:59:43 everge Exp $
 *
 * History: 
 *    $Log: InventoryCoverage.java,v $
 *    Revision 1.3  2015/01/14 16:59:43  everge
 *    Added columns for setup date and average sales per month, changed 4 week sales to monthly sales
 *
 *    Revision 1.2  2014/12/12 22:34:25  everge
 *    Fixed bug where coverage calculations weren't accurate when there were no customer orders. Changed Rolling 30 to count orders rather than shipments. Added some columns.
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;

public class InventoryCoverage extends Report
{
   private static final short maxCols = 24;
   private static final SpreadsheetVersion SS_VERSION = SpreadsheetVersion.EXCEL2007; // change to SpreadsheetVersion.EXCEL2007 to use XSSF
   
   //
   // indexes to the special cell styles array
   private static final short csIntRed        = 0;
   private static final short csPercentRed    = 1;
   private static final short csPercentOrange = 2;
   private static final short csPercentYellow = 3;
   private static final short csPercentGold   = 4;
   
   //
   // DB Data
   private PreparedStatement m_PoData;
   
   //
   // The cell styles for each of the columns in the spreadsheet.
   private CellStyle[] m_CellStyles;
   private CellStyle[] m_CellStylesEx;
   
   //
   // workbook entries.
   private Workbook m_Wrkbk;
   private Sheet m_Sheet;
   private CreationHelper m_CreateHelper;
      
   /**
    * Default constructor.
    */
   public InventoryCoverage()
   {
      super();
      
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet("Inventory Coverage Report");
      m_CreateHelper = m_Wrkbk.getCreationHelper();
      m_MaxRunTime = RptServer.HOUR * 12;      
            
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
      
      if ( m_CellStylesEx != null ) {
         for ( int i = 0; i < m_CellStylesEx.length; i++ )
            m_CellStylesEx[i] = null;
      }
      
      m_Sheet = null;     
      m_Wrkbk = null;      
      m_CellStyles = null;
      m_CellStylesEx = null;
      m_PoData = null;
      
      super.finalize();
   }
   
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      
      Row row = null;
      Cell cell = null;
      int rowNum = 0;
      int colNum = 0;
      int maxRows = SS_VERSION.getMaxRows();
      int sheetNum = 2;
      String msg = "processing dept %s, item %s";
      FileOutputStream outFile = null;      
      ResultSet poData = null;
      boolean result = false; 
      String item = null;
      String dept = null;
      int qoh = 0;
      int custOrders = 0;
      int fourWeekSales = 0;
      int fcst4 = 0;
      int fcst8 = 0;
      int adjFcst4 = 0;
      StringBuffer fileName = new StringBuffer();      
      SimpleDateFormat df = new SimpleDateFormat("MM dd yy");
      
      fileName.append(df.format(new Date()));
      fileName.append(" Inventory Coverage Report.xls");
      // set file extension to .xlsx if SS_VERSION is EXCEL2007
      if (SS_VERSION == SpreadsheetVersion.EXCEL2007) {
         fileName.append("x");
      }
      m_FileNames.add(fileName.toString());
            
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      rowNum = createCaptions();
      
      try {
         poData = m_PoData.executeQuery();

         while ( poData.next() && m_Status == RptServer.RUNNING ) {
            item = poData.getString("item_id");
            dept = poData.getString("dept_num");
            setCurAction(String.format(msg, dept, item));
            qoh = poData.getInt("qoh");
            custOrders = poData.getInt("cust_orders");
            fourWeekSales = poData.getInt("monthly_sales");
            fcst4 = poData.getInt("fcst_4");
            fcst8 = poData.getInt("fcst_8");
            adjFcst4 = poData.getInt("adj_fcst4");
                                    
            row = createRow(rowNum++, maxCols);
            colNum = 0;
            
            if ( row != null ) {
               row.getCell(colNum++).setCellValue(m_CreateHelper.createRichTextString(poData.getString("whsname")));
            	row.getCell(colNum++).setCellValue(m_CreateHelper.createRichTextString(dept));
               row.getCell(colNum++).setCellValue(m_CreateHelper.createRichTextString(poData.getString("bname")));
               row.getCell(colNum++).setCellValue(m_CreateHelper.createRichTextString(poData.getString("vname")));
               row.getCell(colNum++).setCellValue(m_CreateHelper.createRichTextString(item));
               row.getCell(colNum++).setCellValue(m_CreateHelper.createRichTextString(poData.getString("velocity")));
               row.getCell(colNum++).setCellValue(m_CreateHelper.createRichTextString(poData.getString("item_description")));
               row.getCell(colNum++).setCellValue(m_CreateHelper.createRichTextString(poData.getString("setup_date")));
               row.getCell(colNum++).setCellValue(m_CreateHelper.createRichTextString(poData.getString("next_po_date")));
               
               //
               // If qty on hand is 0, fill with a red background.
               cell = row.getCell(colNum++);
               cell.setCellValue(qoh);               
               if( qoh <= 0 )
                  cell.setCellStyle(m_CellStylesEx[csIntRed]);
                                 
               row.getCell(colNum++).setCellValue(poData.getInt("open_pos"));
               row.getCell(colNum++).setCellValue(m_CreateHelper.createRichTextString(poData.getString("on_order")));
               row.getCell(colNum++).setCellValue(poData.getInt("total_available"));
               row.getCell(colNum++).setCellValue(poData.getDouble("avg_per_month"));
               row.getCell(colNum++).setCellValue(fourWeekSales);
               row.getCell(colNum++).setCellValue(custOrders);
               row.getCell(colNum++).setCellValue(fcst4);
               row.getCell(colNum++).setCellValue(fcst8);
               
               // When coverage data is not available, leave cell blank 
               
               cell = row.getCell(colNum++);
               if ( custOrders > 0 ) {
                  // If on hand coverage < 100%, fill with red background
                  cell.setCellValue(poData.getDouble("oh_coverage"));               
                  if( poData.getInt("oh_coverage") < 1 )
                     cell.setCellStyle(m_CellStylesEx[csPercentRed]);
               }
               
               cell = row.getCell(colNum++);
               if ( (custOrders + fcst4) > 0) {
                  // If on hand + 4 week forecast coverage < 100%, fill with orange background
                  cell.setCellValue(poData.getDouble("oh_fcst4_coverage"));               
                  if( poData.getInt("oh_fcst4_coverage") < 1 )
                     cell.setCellStyle(m_CellStylesEx[csPercentOrange]);
               }
               
               cell = row.getCell(colNum++);
               if ( (custOrders + fcst4) > 0 ) {
                  // If on hand + 4 week forecast coverage < 100% (including open po's), 
                  // fill with gold background
                  cell.setCellValue(poData.getDouble("oh_po_fcst4_coverage"));               
                  if( poData.getInt("oh_po_fcst4_coverage") < 1 )
                     cell.setCellStyle(m_CellStylesEx[csPercentGold]);
               }
               
               row.getCell(colNum++).setCellValue(adjFcst4);
               
               if ( (custOrders + adjFcst4) > 0 ) {
                  cell = row.getCell(colNum++);
                  // If on hand + adjusted 4 week forecast coverage < 100%, fill with orange background
                  cell.setCellValue(poData.getDouble("oh_adj_fcst4_coverage"));
                  if( poData.getInt("oh_adj_fcst4_coverage") < 1)
                     cell.setCellStyle(m_CellStylesEx[csPercentOrange]);
               }
               
               cell = row.getCell(colNum++);
               if ( (custOrders + fcst8 > 0) ) {
                  // If on hand + 8 week forecast coverage < 100% (including open po's), 
                  // fill with yellow background
                  cell.setCellValue(poData.getDouble("oh_po_fcst8_coverage"));               
                  if( poData.getInt("oh_po_fcst8_coverage") < 1 )
                     cell.setCellStyle(m_CellStylesEx[csPercentYellow]);
               }
            }
            // if there are more rows than will fit on the current sheet, then create a new sheet for the excess
            if (rowNum >= maxRows) {
               m_Sheet = m_Wrkbk.createSheet(m_Sheet.getSheetName() + " - pg " + sheetNum++);
               rowNum = createCaptions();
            }
         }
         
         m_Wrkbk.write(outFile);
         poData.close();

         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[InventoryCoverage]", ex);
      }

      finally {         
         closeStatements();
         
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error(e);
         }

         outFile = null;
         row = null;
         cell = null;
         poData = null;
         item = null;
         dept = null;
      }

      return result;
   }
   
   /**
    * Close all the open queries
    */
   private void closeStatements()
   {     
      closeStmt(m_PoData);
   }
   
   /**
    * Sets the captions on the report.
    */
   private int createCaptions()
   {
      SimpleDateFormat df = new SimpleDateFormat("MM.dd.yy");
      CellStyle csCaption = null;      
      Font font = null;
      Cell cell = null;
      Row row = null;
      int rowNum = 0;
      int colNum = 0;
      short rowHeight = 1000;
      
      font = m_Wrkbk.createFont();
      font.setFontHeightInPoints((short)8);
      font.setFontName("Arial");
      font.setBold(true);
           
      csCaption = m_Wrkbk.createCellStyle();      
      csCaption.setFont(font);
      csCaption.setAlignment(HorizontalAlignment.CENTER);
      csCaption.setWrapText(true);
     
      try {
         if ( m_Sheet != null ) {
            //
            // Create the row for the captions.
            row = m_Sheet.createRow(rowNum++);
            row.setHeight(rowHeight);
            
            for ( int i = 0; i < maxCols; i++ ) {
               cell = row.createCell(i);
               cell.setCellStyle(csCaption);
            }
            
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Warehouse"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Dept#"));
            m_Sheet.setColumnWidth(colNum++, 1300);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Buyer Name"));
            m_Sheet.setColumnWidth(colNum++, 4000);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Vendor Name"));
            m_Sheet.setColumnWidth(colNum++, 6000);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Item#"));
            m_Sheet.setColumnWidth(colNum++, 2000);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Velocity"));
            m_Sheet.setColumnWidth(colNum++, 1800);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Item Description"));
            m_Sheet.setColumnWidth(colNum++, 6000);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Setup Date"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Next PO \nDue Date"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString(df.format(new Date())+ "\nQty On\nHand"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Open\nPOs"));
            m_Sheet.setColumnWidth(colNum++, 1500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("On\nOrder?"));
            m_Sheet.setColumnWidth(colNum++, 1600);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Total On\nHand or In\nPipeline"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Average Sales Per Month"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Rolling\n4 Week\nUnit Sales"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Open\nCustomer\nOrders"));
            m_Sheet.setColumnWidth(colNum++, 2250);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Forecasted\nSales from\nPrescient -\n4 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Forecasted\nSales from\nPrescient -\n8 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("QOH Coverage of Open Cust Orders"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("QOH Coverage of Open Cust Orders + 4 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Total Coverage of Open Cust Orders + 4 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Adjusted Forecast -\n4 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("QOH Coverage of Open Cust Orders + Adj 4 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            row.getCell(colNum).setCellValue(m_CreateHelper.createRichTextString("Total Coverage of Open Cust Orders + 8 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 2500);
            
            // freeze header row
            m_Sheet.createFreezePane(0, 1);
         }
      }
      
      finally {
         font = null;         
         csCaption = null;
         df = null;
      }
      
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
         if ( m_EdbConn == null )
            m_EdbConn = m_RptProc.getEdbConn();
         
         if ( prepareStatements() )
            created = buildOutputFile();
      }
      
      catch ( Exception ex ) {
         log.fatal("[InventoryCoverage]", ex);
      }
      
      finally {
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
   private Row createRow(int rowNum, int colCnt)
   {
      Row row = null;
      Cell cell = null;
      
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
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(4000);
      boolean isPrepared = false;
      
      if ( m_EdbConn != null ) {
         try {            
            //
            // customer information based on accounts or store number.
            // Gross sales are net of credits.
            sql.setLength(0);
            sql.append("select   ");
            sql.append("  warehouse.name as whsname, "); 
            sql.append("  emery_dept.dept_num, "); 
            sql.append("  buyer.name as bname, "); 
            sql.append("  vendor.name as vname, "); 
            sql.append("  item_entity_attr.item_id, "); 
            sql.append("  item_entity_attr.description as item_description, "); 
            sql.append("  to_char(ejd_item.setup_date, 'dd-MON-yy') as setup_date, "); 
            sql.append("  coalesce(item_velocity.velocity, 'N') as velocity, "); 
            sql.append("  qoh, "); 
            sql.append("  coalesce(po_qty.qty_ordered, 0) as open_pos, "); 
            sql.append("  decode(po_qty.qty_ordered, null, 'no', 'yes') as on_order, "); 
            sql.append("  coalesce(to_char(last_po.due_date, 'dd-MON-yy'), 'NONE') as next_po_date, "); 
            sql.append("  qoh + coalesce(po_qty.qty_ordered, 0) as total_available, "); 
            sql.append("  coalesce(ord_qty.qty_ordered, 0) as cust_orders, "); 
            sql.append("  coalesce(monthly_avg.avg_per_month, 0) as avg_per_month, "); 
            sql.append("  coalesce(wk4_sales.qty_shipped, 0) as monthly_sales, "); 
            sql.append("  coalesce(fcst_4_week.fcst_qty, 0) as fcst_4, "); 
            sql.append("  coalesce(fcst_8_week.fcst_qty, 0) as fcst_8, "); 

            sql.append("  case coalesce(ord_qty.qty_ordered, 0) "); 
            sql.append("    when 0 "); 
            sql.append("      then 1.0 "); 
            sql.append("    else (coalesce(qoh, 0) / coalesce(ord_qty.qty_ordered, 0)) "); 
            sql.append("  end as oh_coverage, ");
            
            sql.append("  case (coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_4_week.fcst_qty, 0)) "); 
            sql.append("    when 0 "); 
            sql.append("      then 1.0 "); 
            sql.append("    else (coalesce(qoh, 0)  / (coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_4_week.fcst_qty, 0))) "); 
            sql.append("  end oh_fcst4_coverage, ");
            
            sql.append("  case (coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_4_week.fcst_qty, 0)) "); 
            sql.append("    when 0 "); 
            sql.append("      then 1.0 "); 
            sql.append("    else (coalesce(qoh, 0) + coalesce(po_qty.qty_ordered, 0))  / (coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_4_week.fcst_qty, 0)) "); 
            sql.append("  end oh_po_fcst4_coverage, ");
            
            sql.append("  greatest(coalesce(wk4_sales.qty_shipped, 0), coalesce(fcst_4_week.fcst_qty, 0)) adj_fcst4, ");
            
            sql.append("  case (coalesce(ord_qty.qty_ordered, 0) + greatest(coalesce(wk4_sales.qty_shipped, 0), coalesce(fcst_4_week.fcst_qty, 0))) "); 
            sql.append("    when 0 "); 
            sql.append("      then 1.0 "); 
            sql.append("    else coalesce(qoh, 0) / (coalesce(ord_qty.qty_ordered, 0) + greatest(coalesce(wk4_sales.qty_shipped, 0), coalesce(fcst_4_week.fcst_qty, 0))) "); 
            sql.append("  end oh_adj_fcst4_coverage, ");
            
            sql.append("  case (coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_8_week.fcst_qty, 0)) "); 
            sql.append("    when 0 "); 
            sql.append("      then 1.0 "); 
            sql.append("    else (coalesce(qoh, 0) + coalesce(po_qty.qty_ordered, 0))  / (coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_8_week.fcst_qty, 0)) "); 
            sql.append("  end oh_po_fcst8_coverage ");

            sql.append("from item_entity_attr ");
            sql.append("inner join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("inner join warehouse on warehouse.warehouse_id in (1, 2) ");
            sql.append("inner join ejd_item_warehouse ");
            sql.append("  on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id ");
            sql.append("  and ejd_item_warehouse.warehouse_id = warehouse.warehouse_id ");
            sql.append("  and ejd_item_warehouse.disp_id = 1 ");
            sql.append("inner join item_type on item_type.item_type_id = item_entity_attr.item_type_id ");
            sql.append("inner join emery_dept emery_dept on emery_dept.dept_id = ejd_item.dept_id ");
            sql.append("inner join buyer on buyer.buyer_id = emery_dept.buyer_id ");
            sql.append("inner join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("left outer join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id ");
            
            sql.append("left outer join ( ");
            sql.append("  select max(due_in_date) due_date, pd.item_nbr item_id, pd.warehouse ");
            sql.append("  from po_hdr ph ");
            sql.append("  join po_dtl pd using(po_hdr_id) ");
            sql.append("  where due_in_date is not null and pd.status = 'OPEN' ");
            sql.append("  group by pd.item_nbr, pd.warehouse ");
            sql.append(") last_po on last_po.item_id = item_entity_attr.item_id and last_po.warehouse = warehouse.fas_facility_id ");
            
            sql.append("left outer join ( ");
            sql.append("  select po_dtl.warehouse, po_dtl.item_ea_id, sum(coalesce(qty_ordered, 0) - coalesce(qty_put_away, 0)) as qty_ordered ");
            sql.append("  from po_dtl ");
            sql.append("  join po_hdr on po_hdr.po_hdr_id = po_dtl.po_hdr_id and po_hdr.status = 'OPEN' ");
            sql.append("  where po_dtl.status = 'OPEN' ");
            sql.append("  group by po_dtl.warehouse, po_dtl.item_ea_id ");
            sql.append(") po_qty on po_qty.warehouse = warehouse.fas_facility_id and po_qty.item_ea_id = item_entity_attr.item_ea_id ");
            
            sql.append("left outer join ( ");
            sql.append("  select order_header.warehouse_id, order_line.item_ea_id, sum(order_line.qty_ordered) qty_ordered ");
            sql.append("  from order_line ");
            sql.append("  join order_header on order_line.order_id = order_header.order_id ");
            sql.append("  join order_status on order_status.order_status_id = order_header.order_status_id and ");
            sql.append("                       order_status.description in ('NEW','WAITING FOR INVENTORY','WAITING CREDIT APPROVAL') ");
            sql.append("  join order_status line_status on line_status.order_status_id = order_line.order_status_id and ");
            sql.append("                                   line_status.description in ('NEW') ");
            sql.append("  left outer join promotion on promotion.promo_id = order_line.promo_id ");
            sql.append("  group by order_header.warehouse_id, order_line.item_ea_id ");
            sql.append(") ord_qty on ord_qty.warehouse_id = warehouse.warehouse_id and ord_qty.item_ea_id = item_entity_attr.item_ea_id ");

            sql.append("left outer join ( ");
            sql.append("  select warehouse_id, item_nbr, avg(month_sales) avg_per_month ");
            sql.append("  from (select warehouse_id, item_nbr, trunc(invoice_date, 'MONTH') sale_month, sum(qty_ordered) month_sales ");
            sql.append("        from itemsales ");
            sql.append("        where invoice_date >= trunc(add_months(sysdate, -12), 'MONTH') and invoice_date < trunc(sysdate, 'MONTH') ");
            sql.append("        group by warehouse_id, item_nbr, trunc(invoice_date, 'MONTH')) ");
            sql.append("  group by warehouse_id, item_nbr ");
            sql.append(") monthly_avg on monthly_avg.warehouse_id = warehouse.warehouse_id and monthly_avg.item_nbr = item_entity_attr.item_id ");
            
            sql.append("left outer join ( ");
            sql.append("  select warehouse_id, item_nbr, sum(qty_shipped) as qty_shipped ");
            sql.append("  from itemsales ");
            sql.append("  where invoice_date > current_date - 28 ");
            sql.append("  group by warehouse_id, item_nbr ");
            sql.append(") wk4_sales on wk4_sales.warehouse_id = warehouse.warehouse_id and wk4_sales.item_nbr = item_entity_attr.item_id ");

            sql.append("left outer join fcst_4_week ");
            sql.append("  on fcst_4_week.whs_name = warehouse.name ");
            sql.append("  and fcst_4_week.item_id = item_entity_attr.item_id ");

            sql.append("left outer join fcst_8_week ");
            sql.append("  on fcst_8_week.whs_name = warehouse.name ");
            sql.append("  and fcst_8_week.item_id = item_entity_attr.item_id ");

            sql.append("where item_type.item_type_id < 8 ");
            
            sql.append("order by warehouse.name, ");
            sql.append("  case (nvl(ord_qty.qty_ordered, 0) + nvl(fcst_4_week.fcst_qty, 0)) ");
            sql.append("    when 0 "); 
            sql.append("      then 1.0 "); 
            sql.append("    else nvl(avail_qty, 0)  / (nvl(ord_qty.qty_ordered, 0) + nvl(fcst_4_week.fcst_qty, 0)) ");
            sql.append("  end, ");
            sql.append("  dept_num, vendor.name, item_entity_attr.item_id ");

            m_PoData = m_EdbConn.prepareStatement(sql.toString());
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[InventoryCoverage]", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("[InventoryCoverage] prepareStatements - null connection");
      
      return isPrepared;
   }
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    * Note - 
    *    The styles came from the original Excel worksheet.
    */
   private void setupWorkbook()
   {            
      CellStyle styleTxtC = null;       // Text centered
      CellStyle styleTxtL = null;       // Text left justified
      CellStyle styleInt = null;        // 0 decimals and a comma
      CellStyle csIntRed = null;        // 0 decimals and a comma, red foreground
      CellStyle styleDouble = null;     // numeric 1 decimal and a comma
      CellStyle stylePercent = null;    // Integer percent with comma
      CellStyle csPercentRed = null;	  // Integer percent with comma red
      CellStyle csPercentOrange = null; // Integer percent with comma orange
      CellStyle csPercentGold = null;	  // Integer percent with comma gold
      CellStyle csPercentYellow = null; // Integer percent with comma yellow
      DataFormat format = null;
      Font font = null;
            
      format = m_Wrkbk.createDataFormat();
      
      font = m_Wrkbk.createFont();
      font.setFontHeightInPoints((short)8);
      font.setFontName("Arial");
            
      styleTxtL = m_Wrkbk.createCellStyle();
      styleTxtL.setAlignment(HorizontalAlignment.LEFT);
      styleTxtL.setFont(font);
      
      styleTxtC = m_Wrkbk.createCellStyle();
      styleTxtC.setAlignment(HorizontalAlignment.CENTER);
      styleTxtC.setFont(font);
      
      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setFont(font);
      styleInt.setDataFormat(format.getFormat("_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)"));
      
      styleDouble = m_Wrkbk.createCellStyle();
      styleDouble.setAlignment(HorizontalAlignment.RIGHT);
      styleDouble.setFont(font);
      styleDouble.setDataFormat(format.getFormat("_(* #,##0.0_);_(* (#,##0.0);_(* \"-\"??_);_(@_)"));
      
      stylePercent = m_Wrkbk.createCellStyle();
      stylePercent.setAlignment(HorizontalAlignment.RIGHT);
      stylePercent.setFont(font);
      stylePercent.setDataFormat(format.getFormat("#,##0%"));

      //
      // These are used in the conditional formatting.
      csIntRed = m_Wrkbk.createCellStyle();
      csIntRed.setAlignment(HorizontalAlignment.RIGHT);
      csIntRed.setFont(font);
      csIntRed.setDataFormat(format.getFormat("_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)"));
      csIntRed.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      csIntRed.setFillForegroundColor(IndexedColors.RED.getIndex());
      
      csPercentRed = m_Wrkbk.createCellStyle();
      csPercentRed.setAlignment(HorizontalAlignment.RIGHT);
      csPercentRed.setFont(font);
      csPercentRed.setDataFormat(format.getFormat("#,##0%"));
      csPercentRed.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      csPercentRed.setFillForegroundColor(IndexedColors.RED.getIndex());
            
      csPercentOrange = m_Wrkbk.createCellStyle();
      csPercentOrange.setAlignment(HorizontalAlignment.RIGHT);
      csPercentOrange.setFont(font);
      csPercentOrange.setDataFormat(format.getFormat("#,##0%"));
      csPercentOrange.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      csPercentOrange.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
      
      csPercentYellow = m_Wrkbk.createCellStyle();
      csPercentYellow.setAlignment(HorizontalAlignment.RIGHT);
      csPercentYellow.setFont(font);
      csPercentYellow.setDataFormat(format.getFormat("#,##0%"));
      csPercentYellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      csPercentYellow.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
      
      csPercentGold = m_Wrkbk.createCellStyle();
      csPercentGold.setAlignment(HorizontalAlignment.RIGHT);
      csPercentGold.setFont(font);
      csPercentGold.setDataFormat(format.getFormat("#,##0%"));
      csPercentGold.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      csPercentGold.setFillForegroundColor(IndexedColors.GOLD.getIndex());
            
      m_CellStyles = new CellStyle[] {
      	styleTxtC,    // col 0 warehouse
         styleTxtC,    // col 1 dept
         styleTxtL,    // col 2 vendor name
         styleTxtL,    // col 3 buyer
         styleTxtC,    // col 4 item id
         styleTxtC,    // col 5 velocity code
         styleTxtL,    // col 6 item desc
         styleTxtC,    // col 7 setup date
         styleTxtC,    // col 8 next PO due
         styleInt,     // col 9 qoh
         styleInt,     // col 10 open pos
         styleTxtC,    // col 11 on order?
         styleInt,	  // col 12 qoh + pos
         styleDouble,  // col 13 average sales per month for last 12 months
         styleInt,     // col 14 4 week sales
         styleInt,     // col 15 open orders
         styleInt,	  // col 16 4 week forecast
         styleInt,	  // col 17 8 week forecast
         stylePercent, // col 18 qoh coverage
         stylePercent, // col 19 qoh + 4 week forecast coverage
         stylePercent, // col 20 qoh + 4 week forecast + open po coverage
         styleInt,     // col 21 adjusted 4 week forecast
         stylePercent, // col 22 qoh + adjusted 4 week forecast coverage
         stylePercent  // col 23 qoh + 8 week forecast + open po coverage
      };
      
      m_CellStylesEx = new CellStyle[] {
         csIntRed,        // special style #1
         csPercentRed,    // special style #2
         csPercentOrange, // special style #3
         csPercentYellow, // special style #4
         csPercentGold	  // special style #5
      };
      
      styleTxtC = null;
      styleTxtL = null;
      styleInt = null;
      styleDouble = null;
      csIntRed = null;
      csPercentRed = null;
      csPercentOrange = null;
      csPercentYellow = null;      
      format = null;
      font = null;   
   }

   /*
   public static void main(String... args) throws SQLException 
   {
        System.out.println(Calendar.getInstance().getTime());
        BasicConfigurator.configure();
        InventoryCoverage ic = new InventoryCoverage();
        ic.log = Logger.getLogger(Report.class);

        Connection conn;
        Properties connProps = new Properties();
        connProps.put("user", "ejd");
        connProps.put("password", "boxer");
        conn = DriverManager.getConnection("jdbc:edb://172.30.1.33/emery_jensen", connProps);
        ic.m_EdbConn = conn;

        ic.m_FilePath = "C:/Users/bcornwell/temp/";
        boolean res = ic.createReport();
        System.out.println(res);
        System.out.println(Calendar.getInstance().getTime());
   }
   */
}
