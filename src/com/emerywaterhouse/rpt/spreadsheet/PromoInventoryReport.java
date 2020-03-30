/**
 * File: PromoInventoryReport.java
 * Description: Report based off the never out report, to show promo inventory for promos or packets.
 *
 * @author Eric Brownewell
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

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
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

public class PromoInventoryReport extends Report 
{
   private static final short maxCols = 23;

   //
   // indexes to the special cell styles array
   private static final short csIntRed = 0;
   private static final short csPercentRed = 1;
   private static final short csPercentOrange = 2;
   private static final short csPercentYellow = 3;
   private static final short csPercentGold = 4;

   private static final String allWarehouses = "'PORTLAND','PITTSTON'";

   //promo/packet storage string and boolean
   boolean usePacket = false;
   String ids;

   //item/vendor/buyer limiters:
   String itemIds, vendorIds, buyerIds, warehouseList;

   java.sql.Date startDate, endDate;

   //
   // DB Data
   private PreparedStatement m_PoData;

   //
   // The cell styles for each of the columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;
   private XSSFCellStyle[] m_CellStylesEx;

   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;

   /**
    * Default constructor
    */
   public PromoInventoryReport() 
   {
      super();

      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet("Promo Inventory Report");

      //m_Wrkbk.setRepeatingRowsAndColumns(m_Wrkbk.getSheetIndex(m_Sheet), 0, maxCols, 1, 2);

      m_Sheet.createFreezePane(0, 2);
      m_Sheet.setAutoFilter(new CellRangeAddress(1, 1, 0, maxCols - 1));

      m_MaxRunTime = RptServer.HOUR * 12;

      setupWorkbook();
   }

   /**
    * Cleanup any allocated resources.
    *
    * @throws Throwable
    */
   public void finalize() throws Throwable 
   {
      if (m_CellStyles != null) {
         for (int i = 0; i < m_CellStyles.length; i++)
            m_CellStyles[i] = null;
      }

      if (m_CellStylesEx != null) {
         for (int i = 0; i < m_CellStylesEx.length; i++)
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
    * @throws FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException 
   {

      XSSFRow row = null;
      XSSFCell cell = null;
      int rowNum = 0;
      int colNum = 0;
      String msg = "processing dept %s, item %s";
      FileOutputStream outFile = null;
      ResultSet poData = null;
      boolean result = false;
      String item = null;
      String dept = null;
      double qoh = 0;
      double monthSales = 0;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      rowNum = createCaptions();

      try {
         m_PoData.setDate(1, startDate);
         m_PoData.setDate(2, endDate);

         poData = m_PoData.executeQuery();

         while ( poData.next() && m_Status == RptServer.RUNNING ) {
            item = poData.getString("item_id");
            dept = poData.getString("dept_num");
            setCurAction(String.format(msg, item, dept));
            qoh = poData.getInt("qoh");
            monthSales = poData.getDouble("monthly_sales");

            row = createRow(rowNum++, maxCols);
            colNum = 0;

            if ( row != null ) {
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(poData.getString("warehouse")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(dept));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(poData.getString("bname")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(poData.getString("vname")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(poData.getString("item_id")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(poData.getString("description")));

               //
               // If qty on hand is 0, fill with a red background.
               cell = row.getCell(colNum++);
               cell.setCellValue(qoh);
               
               if (qoh <= 0)
                  cell.setCellStyle(m_CellStylesEx[csIntRed]);

               row.getCell(colNum++).setCellValue(poData.getInt("open_pos"));
               row.getCell(colNum++).setCellValue(poData.getInt("total_available"));

               //new columns being added to report
               row.getCell(colNum++).setCellValue(poData.getInt("r7_promo_sales"));
               row.getCell(colNum++).setCellValue(poData.getInt("total_promo_order_sales"));
               row.getCell(colNum++).setCellValue(poData.getInt("prescient_promo_forecast"));
               row.getCell(colNum++).setCellValue(poData.getInt("r7_sales"));
               //end new columns

               row.getCell(colNum++).setCellValue(monthSales);
               row.getCell(colNum++).setCellValue(poData.getInt("cust_orders"));
               row.getCell(colNum++).setCellValue(poData.getInt("fcst_4"));
               row.getCell(colNum++).setCellValue(poData.getInt("fcst_8"));

               // If on hand coverage < 100%, fill with red background
               cell = row.getCell(colNum++);
               cell.setCellValue(poData.getDouble("oh_coverage"));
               
               if (poData.getInt("oh_coverage") < 1)
                  cell.setCellStyle(m_CellStylesEx[csPercentRed]);

               // If on hand + 4 week forecast coverage < 100%, fill with orange background
               cell = row.getCell(colNum++);
               cell.setCellValue(poData.getDouble("oh_fcst4_coverage"));
               
               if (poData.getInt("oh_fcst4_coverage") < 1)
                  cell.setCellStyle(m_CellStylesEx[csPercentOrange]);


               // If on hand + 8 week forecast coverage < 100% (including open po's),
               // fill with yellow background
               cell = row.getCell(colNum++);
               cell.setCellValue(poData.getDouble("oh_po_fcst4_coverage"));
               
               if (poData.getInt("oh_po_fcst4_coverage") < 1)
                  cell.setCellStyle(m_CellStylesEx[csPercentGold]);

               // If on hand + 8 week forecast coverage < 100% (including open po's),
               // fill with yellow background
               cell = row.getCell(colNum++);
               cell.setCellValue(poData.getDouble("oh_po_fcst8_coverage"));
               
               if (poData.getInt("oh_po_fcst8_coverage") < 1)
                  cell.setCellStyle(m_CellStylesEx[csPercentYellow]);

               cell = row.getCell(colNum++);
               cell.setCellValue(poData.getString("portland_active"));
               cell = row.getCell(colNum++);
               cell.setCellValue(poData.getString("pittston_active"));

            }
         }

         m_Wrkbk.write(outFile);
         poData.close();

         result = true;
      } 
      
      catch (Exception ex) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[PromoInventoryReport]", ex);
      } 
      
      finally {
         closeStatements();

         try {
            outFile.close();
         } 
         
         catch (Exception e) {
            ;
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
   private void closeStatements() {
      closeStmt(m_PoData);
   }

   /**
    * Sets the captions on the report.
    */
   private int createCaptions() 
   {
      SimpleDateFormat df = new SimpleDateFormat("MM.dd.yy");
      XSSFCellStyle csCaption = null;
      XSSFFont font = null;
      XSSFCell cell = null;
      XSSFRow row = null;
      int rowNum = 0;
      int colNum = 0;
      short rowHeight = 1000;

      font = m_Wrkbk.createFont();
      font.setFontHeightInPoints((short) 8);
      font.setFontName("Arial");
      font.setBold(true);

      csCaption = m_Wrkbk.createCellStyle();
      csCaption.setFont(font);
      csCaption.setAlignment(HorizontalAlignment.CENTER);
      csCaption.setWrapText(true);

      row = m_Sheet.createRow(rowNum++);
      row.createCell(0);

      row.getCell(0).setCellValue(buildTitle());

      try {
         if (m_Sheet != null) {
            //
            // Create the row for the captions.
            row = m_Sheet.createRow(rowNum++);
            row.setHeight(rowHeight);

            for (int i = 0; i < maxCols; i++) {
               cell = row.createCell(i);
               cell.setCellStyle(csCaption);
            }

            row.getCell(colNum).setCellValue(new XSSFRichTextString("Warehouse"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Dept#"));
            m_Sheet.setColumnWidth(colNum++, 1500);
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Buyer Name"));
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Vendor Name"));
            m_Sheet.setColumnWidth(colNum++, 6000);
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Item#"));
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Item Description"));
            m_Sheet.setColumnWidth(colNum++, 6000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString(df.format(new Date()) + "\nQty On\nHand"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Open\nPOs"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Total On\nHand or In\nPipeline"));
            m_Sheet.setColumnWidth(colNum++, 3000);

            //new columns, this is where we differ from the never-out report
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Rolling 7\nDays Promo\nSales"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Total Promo\nSales"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Prescient\nPromo\nForecast"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Rolling 7\nDays Unit\nSales"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            //end new columns

            row.getCell(colNum).setCellValue(new XSSFRichTextString("Rolling\n30 Days\nUnit Sales"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Open\nCustomer\nOrders"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Forecasted\nSales from\nPrescient -\n4 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Forecasted\nSales from\nPrescient -\n8 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("QOH Coverage of Open Customer Orders"));
            m_Sheet.setColumnWidth(colNum++, 3500);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("QOH Coverage of Open Cust Orders + 4 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 3500);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Total Coverage of Open Cust Orders + 4 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 3500);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Total Coverage of Open Cust Orders + 8 Weeks"));
            m_Sheet.setColumnWidth(colNum++, 3500);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Portland Active"));
            m_Sheet.setColumnWidth(colNum++, 3500);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Pittston Active"));
            m_Sheet.setColumnWidth(colNum++, 3500);
         }
      } 
      
      finally {
         font = null;
         csCaption = null;
         df = null;
      }

      return rowNum;
   }

   private String buildTitle() 
   {
      String promoOrPacket;

      if ( usePacket ) {
         promoOrPacket = "Packet";
      } 
      else {
         promoOrPacket = "Promo";
      }

      return String.format("Promo Inventory Report for %s #%s Prescient date range %s to %s", promoOrPacket, ids.replace("'", ""), startDate, endDate);
   }

   /**
    * @see Report#createReport()
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
         log.fatal("[PromoInventoryReport]", ex);
      } 
      
      finally {
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Creates a row in the worksheet.
    *
    * @param rowNum The row number.
    * @param colCnt The number of columns in the row.
    * @return The formatted row of the spreadsheet.
    */
   private XSSFRow createRow(int rowNum, int colCnt) 
   {
      XSSFRow row = null;
      XSSFCell cell = null;

      if ( m_Sheet != null ) {
         row = m_Sheet.createRow(rowNum);
   
         //
         // set the type and style of the cell.
         if (row != null) {
            for (int i = 0; i < colCnt; i++) {
               cell = row.createCell(i);
               cell.setCellStyle(m_CellStyles[i]);
            }
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
         if (buyerIds.equalsIgnoreCase("all"))
            buyerIds = getAllBuyerIds();
         else
            buyerIds = getAllBuyerIds(buyerIds);

         try {
            //
            // customer information based on accounts or store number.
            // Gross sales are net of credits.
            sql.setLength(0);
            sql.append("select ");
            sql.append("warehouse.name as warehouse, ");
            sql.append("dept_num, ");
            sql.append("buyer.name as bname, ");
            sql.append("vendor.name AS vname, ");
            sql.append("item_entity_attr.item_id, ");
            sql.append("item_entity_attr.description, ");
            sql.append("qoh, ");
            sql.append("coalesce(po_qty.qty_ordered, 0) as open_pos, ");
            sql.append("qoh + nvl(po_qty.qty_ordered, 0) as total_available, ");
            sql.append("coalesce(ord_qty.qty_ordered, 0) as cust_orders, ");
            sql.append("coalesce(wk4_sales.qty_shipped, 0) as monthly_sales, ");
            sql.append("coalesce(fcst_4_week.fcst_qty, 0) as fcst_4, ");
            sql.append("coalesce(fcst_8_week.fcst_qty, 0) as fcst_8, ");
            sql.append("coalesce(r7_promo_sales.qty_ordered, 0) as r7_promo_sales, ");
            sql.append("coalesce(total_sales.units, 0) as total_promo_order_sales, ");
            sql.append("coalesce(promo_fcst.promo_adj, 0) + coalesce(promo_fcst.promo_plan, 0) as prescient_promo_forecast, ");
            sql.append("coalesce(r7_sales.qty_ordered, 0) as r7_sales, ");
            sql.append("portland_active.active as portland_active, ");
            sql.append("pittston_active.active as pittston_active, ");
            
            sql.append("case coalesce(ord_qty.qty_ordered, 0) ");
            sql.append("   when 0 ");
            sql.append("     then 1.0 ");
            sql.append("   else qoh / nvl(ord_qty.qty_ordered, 0) ");
            sql.append("end as oh_coverage, ");
         
            sql.append("case coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_4_week.fcst_qty, 0) ");
            sql.append("   when 0 ");
            sql.append("     then 1.0 ");
            sql.append("   else coalesce(qoh, 0) / (coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_4_week.fcst_qty, 0)) ");
            sql.append("end as oh_fcst4_coverage, ");
            
            sql.append("case coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_4_week.fcst_qty, 0) ");
            sql.append("   when 0 ");
            sql.append("     then 1.0 ");
            sql.append("   else (qoh + coalesce(po_qty.qty_ordered, 0)) / (coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_4_week.fcst_qty, 0)) ");
            sql.append("end as oh_po_fcst4_coverage, ");
            
            sql.append("case coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_8_week.fcst_qty, 0) ");
            sql.append("   when 0 ");
            sql.append("     then 1.0 ");
            sql.append("   else (qoh + coalesce(po_qty.qty_ordered, 0)) / (coalesce(ord_qty.qty_ordered, 0) + coalesce(fcst_8_week.fcst_qty, 0)) ");
            sql.append("end as oh_po_fcst8_coverage ");
                  
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item using(ejd_item_id) ");
            sql.append(String.format("join warehouse on warehouse.name in (%s) ", warehouseList));
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id and ejd_item_warehouse.warehouse_id = warehouse.warehouse_id ");
            sql.append("join emery_dept on emery_dept.dept_id = ejd_item.dept_id ");
            sql.append("join buyer on buyer.buyer_id = emery_dept.buyer_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");

            if ( usePacket ) {
               sql.append("join promo_item on promo_item.item_ea_id = item_entity_attr.item_ea_id and  ");
               sql.append("   promo_id in ( ");
               sql.append("      select promo_id "); 
               sql.append("      from promotion ");
               sql.append(String.format("      where packet_id in (%s) ", ids));
               sql.append("   ) ");
            }
            else {
               sql.append("join promo_item on promo_item.item_ea_id = item_entity_attr.item_ea_id and ");
               sql.append(String.format("promo_id in (%s) ", ids));
            }

            sql.append("left join ( ");
            sql.append("   select ejd_item_id, decode(active, 1, 'true', 0, 'false') as active ");
            sql.append("   from ejd_item_warehouse ");
            sql.append("   where warehouse_id = 1 ");
            sql.append(") portland_active on portland_active.ejd_item_id = ejd_item.ejd_item_id ");
   
            sql.append("left join ( ");
            sql.append("   select ejd_item_id, decode(active, 1, 'true', 0, 'false') as active ");
            sql.append("   from ejd_item_warehouse ");
            sql.append("   where warehouse_id = 2 ");
            sql.append(") pittston_active on pittston_active.ejd_item_id = ejd_item.ejd_item_id ");
   
            sql.append("left outer join ( ");
            sql.append("   select po_dtl.warehouse, po_dtl.item_ea_id, sum(coalesce(qty_ordered, 0) - coalesce(qty_put_away, 0)) as qty_ordered ");
            sql.append("   from po_dtl ");
            sql.append("   join po_hdr on po_hdr.po_hdr_id = po_dtl.po_hdr_id and po_hdr.status = 'OPEN' ");
            sql.append("   where po_dtl.status = 'OPEN' ");
            sql.append("   group by po_dtl.warehouse, po_dtl.item_ea_id ");
            sql.append(") po_qty on po_qty.warehouse = warehouse.fas_facility_id and  po_qty.item_ea_id = item_entity_attr.item_ea_id ");

            sql.append("left outer join ( ");
            sql.append("   select ");
            sql.append("      order_header.warehouse_id, ");
            sql.append("      order_line.item_ea_id, ");
            sql.append("      sum(order_line.qty_ordered) as qty_ordered ");
            sql.append("   from order_line ");
            sql.append("   join order_header on order_line.order_id = order_header.order_id ");
            sql.append("   join order_status on order_status.order_status_id = order_header.order_status_id and  ");
            sql.append("         order_status.description in ('NEW', 'WAITING FOR INVENTORY', 'WAITING CREDIT APPROVAL') ");
            sql.append("   join order_status line_status on line_status.order_status_id = order_line.order_status_id and line_status.description in ('NEW') ");
            sql.append("   left outer join promotion on promotion.promo_id = order_line.promo_id ");
            sql.append("   where order_line.earliest_ship is null or order_line.earliest_ship <= current_date + 28  or promotion.ship_date <= current_date + 28 ");
            sql.append("   group by order_header.warehouse_id, order_line.item_ea_id ");
            sql.append(") ord_qty on ord_qty.warehouse_id = warehouse.warehouse_id and ord_qty.item_ea_id = item_entity_attr.item_ea_id ");
   
            sql.append("left outer join ( ");
            sql.append("   select warehouse_id, item_nbr, sum(qty_shipped) as qty_shipped ");
            sql.append("   from itemsales ");
            sql.append("   where invoice_date > current_date - 28 ");
            sql.append("   group by warehouse_id, item_nbr ");
            sql.append(") wk4_sales on wk4_sales.warehouse_id = warehouse.warehouse_id and  wk4_sales.item_nbr = item_entity_attr.item_id ");
            
            sql.append("left outer join ( ");
            sql.append("   select * ");
            sql.append("   from fcst_4_week ");
            sql.append(") fcst_4_week on fcst_4_week.whs_name = warehouse.name and fcst_4_week.item_id = item_entity_attr.item_id ");
         
            sql.append("left outer join ( ");
            sql.append("   select * ");
            sql.append("   from fcst_8_week ");
            sql.append(") fcst_8_week on fcst_8_week.whs_name = warehouse.name and fcst_8_week.item_id = item_entity_attr.item_id ");
   
            sql.append("left outer join ( ");
            sql.append("   select promo_adj, promo_plan, whs_name, prod_no ");
            sql.append("   from promo_fcst ");
            sql.append("   where act_end_date between ? and ? ");
            sql.append(") promo_fcst on promo_fcst.whs_name = warehouse.name and promo_fcst.prod_no = item_entity_attr.item_id ");
            
            sql.append("left outer join ( ");
            sql.append("   select item_ea_id, warehouse_id, promo_id, sum(qty_ordered) as qty_ordered ");
            sql.append("   from order_line ");
            sql.append("   join order_header using (order_id) ");
            sql.append("   where order_date >= (current_date - 7) and ol_id not in ( ");
            sql.append("      select distinct ol_id ");
            sql.append("      from ol_id_bestpr ");
            sql.append("   ) ");
            sql.append("   group by item_ea_id, promo_id, warehouse_id ");
            sql.append(") r7_promo_sales on r7_promo_sales.item_ea_id = item_entity_attr.item_ea_id and ");
            sql.append("      r7_promo_sales.warehouse_id = warehouse.warehouse_id and ");
            sql.append("      r7_promo_sales.promo_id = promo_item.promo_id ");
                            
            sql.append("left outer join ( ");
            sql.append("   select item_ea_id, warehouse_id, sum(qty_ordered) as qty_ordered ");
            sql.append("   from order_line ");
            sql.append("   join order_header using(order_id) ");
            sql.append("   where order_date >= (current_date - 7) ");
            sql.append("   group by item_ea_id, warehouse_id ");
            sql.append(") r7_sales on r7_sales.item_ea_id = item_entity_attr.item_ea_id and r7_sales.warehouse_id = warehouse.warehouse_id ");
   
            sql.append("left outer join ( ");
            sql.append("   select warehouse.name as facility, promotion.packet_id, order_line.item_ea_id, sum(order_line.qty_ordered) as units ");
            sql.append("   from order_line ");
            sql.append("   join order_header using(order_id) ");
            sql.append("   join promotion using(promo_id) ");
            sql.append("   join warehouse on warehouse.warehouse_id = order_header.warehouse_id ");
            sql.append("   join order_status ls on ls.order_status_id = order_line.order_status_id and ls.description not in ('BACKORDERED') ");
            sql.append("   where ");
            sql.append("      order_line.promo_id is not null and  ");
            
            if ( usePacket )
               sql.append(String.format("promotion.packet_id in (%s) ", ids));
            else
               sql.append(String.format("promotion.promo_id in (%s) ", ids));
   
            sql.append("   group by warehouse.name, promotion.packet_id, order_line.item_ea_id ");
            sql.append(") total_sales on total_sales.item_ea_id = item_entity_attr.item_ea_id and total_sales.facility = warehouse.name ");

            sql.append("where ");
            //
            //buyer will always be included. if it's "all" in the request we'll just pass in all buyer IDs.
            //this prevents the need for messy logic around when to include "and" in the other statements.
            sql.append(String.format("   buyer.buyer_id in (%s) ", buyerIds));
            
            if (itemIds != null && itemIds.length() > 0 )
               sql.append(String.format("   and item_entity_attr.item_ea_id in (%s) and item_entity_attr.item_type_id < 8 ", itemIds));
            
            if ( vendorIds != null && vendorIds.length() > 0 )
               sql.append(String.format("   and item_entity_attr.vendor_id in (%s) ", vendorIds));
               
            sql.append("order by warehouse.name, oh_fcst4_coverage, dept_num, vendor.name, item_entity_attr.item_id ");
            
            m_PoData = m_EdbConn.prepareStatement(sql.toString());
            
            //log.info(sql.toString());
            isPrepared = true;
         } 
         
         catch (SQLException ex) {
            log.error("[PromoInventoryReport]", ex);
         } 
         
         finally {
            sql = null;
         }
      } 
      else
         log.error("[PromoInventoryReport] prepareStatements - null connection");
      
      return isPrepared;
   }

   /**
    * Sets the parameters of this report.
    *
    * @see Report#setParams(ArrayList)
    */
   public void setParams(ArrayList<Param> params) 
   {
      for (Param p : params) {
         switch (p.name) {
            case "startDate":
               startDate = getSqlDate(p.value);
               break;
            
            case "endDate":
               endDate = getSqlDate(p.value);
               break;
            
            case "itemIds":
               itemIds = p.value;
               break;
            
            case "vendorIds":
               vendorIds = p.value;
               break;
            
            case "buyerIds":
               buyerIds = p.value;
               break;
            
            case "usePacket":
               usePacket = Boolean.parseBoolean(p.value);
               break;
            
            case "ids":
               ids = p.value;
               break;
            
            case "warehouseList":
               if (p.value.equalsIgnoreCase("all"))
                  warehouseList = allWarehouses;
               else {
                  warehouseList = String.format("'%s'", p.value);
               }
               break;
         }
      }

      StringBuffer fileName = new StringBuffer();
      SimpleDateFormat df = new SimpleDateFormat("MM-dd-yy");

      fileName.append(df.format(new Date()));
      fileName.append(String.format("_Promo_Inventory_Report_%s.xls", fixIdsForString(ids)));
      m_FileNames.add(fileName.toString());

      df = null;
   }

   private String fixIdsForString(String ids) 
   {
      return ids.replace("'", "");
   }

   private java.sql.Date getSqlDate(String dateParam) 
   {
      SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");

      Date parsed = new Date();

      try {
         parsed = df.parse(dateParam);
      } 
      
      catch (ParseException e) {
         log.error("[PromoInventoryReport] Failed to parse date.", e);
      }

      return new java.sql.Date(parsed.getTime());
   }

   private String getAllBuyerIds() 
   {
      String sql = "select string_agg(distinct buyer_id, ',') as ids from buyer";

      String rtn = "";

      try (PreparedStatement stmt = m_EdbConn.prepareStatement(sql)) {
         try (ResultSet rs = stmt.executeQuery()) {
            if ( rs.next() ) {
               rtn = rs.getString("ids");
            }
         }
      } 
      
      catch (SQLException e) {
         log.error("[PromoInventoryReport]", e);
      }

      return rtn;
   }

   private String getAllBuyerIds(String name) 
   {
      String sql = "select buyer_id from buyer where name = ?";
      String rtn = "";

      try (PreparedStatement stmt = m_EdbConn.prepareStatement(sql)) {
         stmt.setString(1, name);
         
         try (ResultSet rs = stmt.executeQuery()) {
            if ( rs.next() ) {
               rtn = rs.getString("buyer_id");
            }
         }
      } 
      
      catch (SQLException e) {
         log.error("[PromoInventoryReport]", e);
      }

      return rtn;
   }

   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    * Note -
    * The styles came from the original Excel worksheet.
    */
   private void setupWorkbook() 
   {  
      XSSFCellStyle styleTxtC = null;       // Text centered
      XSSFCellStyle styleTxtL = null;       // Text left justified
      XSSFCellStyle styleInt = null;        // 0 decimals and a comma
      XSSFCellStyle csIntRed = null;        // 0 decimals and a comma, red foreground
      XSSFCellStyle styleDouble = null;     // numeric 1 decimal and a comma
      XSSFCellStyle stylePercent = null;    // Integer percent with comma
      XSSFCellStyle csPercentRed = null;      // Integer percent with comma red
      XSSFCellStyle csPercentOrange = null; // Integer percent with comma orange
      XSSFCellStyle csPercentGold = null;      // Integer percent with comma gold
      XSSFCellStyle csPercentYellow = null; // Integer percent with comma yellow
      XSSFDataFormat format = null;
      XSSFFont font = null;

      format = m_Wrkbk.createDataFormat();

      font = m_Wrkbk.createFont();
      font.setFontHeightInPoints((short) 8);
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

      m_CellStyles = new XSSFCellStyle[]{
            styleTxtC,    // col 0 warehouse
            styleTxtC,    // col 1 dept
            styleTxtL,    // col 2 vendor name
            styleTxtL,    // col 3 buyer
            styleTxtC,    // col 4 item id
            styleTxtL,    // col 5 item desc
            styleInt,     // col 6 qoh
            styleInt,     // col 7 open pos
            styleInt,      // col 8 qoh + pos
            styleInt,      // col 9 rolling 7 promo sales
            styleInt,      // col 10 total promo sales
            styleInt,      // col 11 prescient promo forecast
            styleInt,      // col 12 rolling 7 day unit sales
            styleInt,     // col 13 monthly sales
            styleInt,     // col 14 open orders
            styleInt,      // col 15 4 week forecast
            styleInt,      // col 16 8 week forecast
            stylePercent, // col 17 qoh coverage
            stylePercent, // col 18 qoh + 4 week forecast coverage
            stylePercent, // col 19 qoh + 4 week forecast + open po coverage
            stylePercent,  // col 20 qoh + 8 week forecast + open po coverage
            styleTxtL,  // col 21 portland active
            styleTxtL  // col 22 pittston active
      };

      m_CellStylesEx = new XSSFCellStyle[]{
            csIntRed,        // special style #1
            csPercentRed,    // special style #2
            csPercentOrange, // special style #3
            csPercentYellow, // special style #4
            csPercentGold      // special style #5
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


   /*public static void main(String... args) throws SQLException 
   {
        System.out.println(Calendar.getInstance().getTime());
        BasicConfigurator.configure();
        PromoInventoryReport pir = new PromoInventoryReport();
        pir.log = Logger.getLogger(Report.class);
        ArrayList<Param> params = new ArrayList<>();

        //params.add(new Param("Date", "07/01/2016", "startDate")); //start date
        //params.add(new Param("Date", "07/30/2016", "endDate")); //end date

        params.add(new Param("Date", "08/01/2018", "startDate")); //start date
        params.add(new Param("Date", "09/30/2018", "endDate")); //end date

        params.add(new Param("String", "", "itemIds")); //item ids
        params.add(new Param("String", "", "vendorIds")); //vendor ids
        params.add(new Param("String", "all", "buyerIds")); //buyer ids

        //params.add(new Param("boolean", "true", "usePacket")); //usepacket
        //params.add(new Param("String", "'544'", "ids")); //ids

        params.add(new Param("boolean", "false", "usePacket")); //usepacket
        params.add(new Param("String", "'1560', '1561', '6704', '6705', '6706'", "ids")); //ids

        params.add(new Param("String", "all", "warehouseList")); //warehouses

        Connection conn;
        Properties connProps = new Properties();
        connProps.put("user", "ejd");
        connProps.put("password", "boxer");
        conn = DriverManager.getConnection("jdbc:edb://172.30.1.33/emery_jensen", connProps);
        pir.m_EdbConn = conn;
        pir.setParams(params);

        System.out.println(
              String.format("Start date is %s. End date is %s. Item ids are %s. Vendor ids are %s. Buyer ids are %s. Use packet? %b. Ids are %s.",
              pir.startDate, pir.endDate, pir.itemIds, pir.vendorIds, pir.buyerIds, pir.usePacket, pir.ids)
        );

        pir.m_FilePath = "C:/Users/jfisher/workspace/RptServer/reports/";
        boolean res = pir.createReport();
        System.out.println(res);
        System.out.println(Calendar.getInstance().getTime());
    }
*/
}
