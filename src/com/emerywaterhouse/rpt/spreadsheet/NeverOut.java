/**
 * File: NeverOut.java
 * Description: Report that shows the never out items and open POs
 *
 * @author Jeff Fisher
 *
 * Create Date: 11/11/2009
 * Last Update: $Id: NeverOut.java,v 1.6 2013/02/12 18:44:13 prichter Exp $
 *
 * History:
 *    $Log: NeverOut.java,v $
 *    Revision 1.6.6.6 2017/08/03 9:45 sjaguilar
 *    Bob the builder! Can we fix it? Yes we can!
 *    Fixed for Michael, inventory is apparently messed up.
 *    
 *    Revision 1.6  2013/02/12 18:44:13  prichter
 *    Remove restriction to Portland items
 *
 *    Revision 1.5  2012/09/25 19:01:48  prichter
 *    Correct some misspellings.  Changed the sort order.
 *
 *    Revision 1.4  2012/04/02 19:48:20  prichter
 *    Rewrote to use Prescient forecast rather than historical sales
 *
 *    Revision 1.3  2010/01/07 11:16:27  prichter
 *    Use nvl function in query to handle any items that have no sales during one of the reporting periods.
 *
 *    Revision 1.2  2009/11/18 02:40:52  jfisher
 *    Removed unused vars.
 *
 *    Revision 1.1  2009/11/18 02:39:20  jfisher
 *    Production Version
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
import java.util.Date;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
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

public class NeverOut extends Report
{
   private static final short maxCols = 17;
   
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
   private XSSFCellStyle[] m_CellStyles;
   private XSSFCellStyle[] m_CellStylesEx;
   
   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;
      
   /**
    * Default constructor
    */
   public NeverOut()
   {
      super();
      
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet("Never Out Report");      
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
      
      XSSFRow row = null;
      XSSFCell  cell = null;
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
         poData = m_PoData.executeQuery();

         while ( poData.next() && m_Status == RptServer.RUNNING ) {
            item = poData.getString("item_id");
            dept = poData.getString("dept_num");
            setCurAction(String.format(msg, item, dept));
            qoh = poData.getInt("qty_avail");
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
               if( qoh <= 0 )
                  cell.setCellStyle(m_CellStylesEx[csIntRed]);
                                 
               row.getCell(colNum++).setCellValue(poData.getInt("open_pos"));
               row.getCell(colNum++).setCellValue(poData.getInt("total_available"));
               row.getCell(colNum++).setCellValue(monthSales);
               row.getCell(colNum++).setCellValue(poData.getInt("cust_orders"));
               row.getCell(colNum++).setCellValue(poData.getInt("fcst_4"));
               row.getCell(colNum++).setCellValue(poData.getInt("fcst_8"));
               
               // If on hand coverage < 100%, fill with red background
               cell = row.getCell(colNum++);
               cell.setCellValue(poData.getDouble("oh_coverage"));               
               if( poData.getInt("oh_coverage") < 1 )
                  cell.setCellStyle(m_CellStylesEx[csPercentRed]);
               
               // If on hand + 4 week forecast coverage < 100%, fill with orange background
               cell = row.getCell(colNum++);
               cell.setCellValue(poData.getDouble("oh_fcst4_coverage"));               
               if( poData.getInt("oh_fcst4_coverage") < 1 )
                  cell.setCellStyle(m_CellStylesEx[csPercentOrange]);
               
               
               // If on hand + 8 week forecast coverage < 100% (including open po's), 
               // fill with yellow background
               cell = row.getCell(colNum++);
               cell.setCellValue(poData.getDouble("oh_po_fcst4_coverage"));               
               if( poData.getInt("oh_po_fcst4_coverage") < 1 )
                  cell.setCellStyle(m_CellStylesEx[csPercentGold]);
               
               // If on hand + 8 week forecast coverage < 100% (including open po's), 
               // fill with yellow background
               cell = row.getCell(colNum++);
               cell.setCellValue(poData.getDouble("oh_po_fcst8_coverage"));               
               if( poData.getInt("oh_po_fcst8_coverage") < 1 )
                  cell.setCellStyle(m_CellStylesEx[csPercentYellow]);
               
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

         log.fatal("exception:", ex);
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
      XSSFCellStyle csCaption = null;      
      XSSFFont font = null;
      XSSFCell cell = null;
      XSSFRow row = null;
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
            row.getCell(colNum).setCellValue(new XSSFRichTextString(df.format(new Date())+ "\nQty On\nHand"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Open\nPOs"));
            m_Sheet.setColumnWidth(colNum++, 3000);
            row.getCell(colNum).setCellValue(new XSSFRichTextString("Total On\nHand or In\nPipeline"));
            m_Sheet.setColumnWidth(colNum++, 3000);
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
         m_EdbConn = m_RptProc.getEdbConn();
                  
         if ( prepareStatements() )
            created = buildOutputFile();
      }
      
      catch ( Exception ex ) {
         log.fatal("[NeverOut]", ex);
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
   private XSSFRow createRow(int rowNum, int colCnt)
   {
      XSSFRow row = null;
      XSSFCell  cell = null;
      
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
      StringBuffer sql = new StringBuffer(256);      
      boolean isPrepared = false;
      
      if ( m_EdbConn != null ) {
         try {            
           //
           // customer information based on accounts or store number.
           // Gross sales are net of credits.
           sql.setLength(0);
           sql.append("select ");
           sql.append("   warehouse.name as warehouse, "); 
           sql.append("   dept_num, buyer.name as bname, vendor.name as vname, ");  
           sql.append("   item_entity_attr.item_id, item_entity_attr.description, qoh as qty_avail, ");
           sql.append("   nvl(po_qty.qty_ordered, 0) as open_pos, ");
           sql.append("   qoh + nvl(po_qty.qty_ordered, 0) as total_available, "); 
           sql.append("   nvl(ord_qty.qty_ordered, 0) as cust_orders, ");              
           sql.append("   nvl(wk4_sales.qty_shipped, 0) as monthly_sales, ");  
           sql.append("   nvl(fcst_4_week.fcst_qty, 0) as fcst_4, ");
           sql.append("   nvl(fcst_8_week.fcst_qty, 0) as fcst_8, ");
      
           sql.append("   case nvl(ord_qty.qty_ordered, 0) "); 
           sql.append("      when 0 then 1.0  ");
           sql.append("      else qoh / nvl(ord_qty.qty_ordered, 0) "); 
           sql.append("   end as oh_coverage, ");
      
           sql.append("   case (nvl(ord_qty.qty_ordered, 0) + nvl(fcst_4_week.fcst_qty, 0)) "); 
           sql.append("      when 0 then 1.0 "); 
           sql.append("      else qoh  / (nvl(ord_qty.qty_ordered, 0) + nvl(fcst_4_week.fcst_qty, 0)) "); 
           sql.append("   end as oh_fcst4_coverage, "); 
      
           sql.append("   case nvl(ord_qty.qty_ordered, 0) "); 
           sql.append("      when 0 then 1.0  ");
           sql.append("      else (qoh + nvl(po_qty.qty_ordered, 0))  / (nvl(ord_qty.qty_ordered, 0) + nvl(fcst_4_week.fcst_qty, 0)) "); 
           sql.append("   end as oh_po_fcst4_coverage, ");
      
           sql.append("   case nvl(ord_qty.qty_ordered, 0) "); 
           sql.append("      when 0 then 1.0  ");
           sql.append("      else (qoh + nvl(po_qty.qty_ordered, 0))  / (nvl(ord_qty.qty_ordered, 0) + nvl(fcst_8_week.fcst_qty, 0)) "); 
           sql.append("   end as oh_po_fcst8_coverage "); 
                   
           sql.append("from item_entity_attr ");
           sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
           sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.never_out = 1 ");
           sql.append("join warehouse on warehouse.warehouse_id = ejd_item_warehouse.warehouse_id ");
           sql.append("join emery_dept on emery_dept.dept_id = ejd_item.dept_id  ");
           sql.append("join buyer on buyer.buyer_id = emery_dept.buyer_id ");  
           sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id  ");
   
           sql.append("left outer join ( ");  
           sql.append("   select po_dtl.warehouse, item_nbr, sum(qty_ordered) qty_ordered ");  
           sql.append("   from po_dtl ");  
           sql.append("   join po_hdr on po_hdr.po_hdr_id = po_dtl.po_hdr_id and po_hdr.status = 'OPEN' "); 
           sql.append("   where po_dtl.status = 'OPEN' "); 
           sql.append("   group by po_dtl.warehouse, po_dtl.item_nbr ");  
           sql.append(") po_qty on po_qty.warehouse = warehouse.fas_facility_id and po_qty.item_nbr = item_entity_attr.item_id "); 
          
           sql.append("left outer join ( ");  
           sql.append("   select order_header.warehouse_id, order_line.item_ea_id, sum(order_line.qty_ordered) qty_ordered ");  
           sql.append("   from order_line ");  
           sql.append("   join order_header on order_line.order_id = order_header.order_id  ");
           sql.append("   join order_status on order_status.order_status_id = order_header.order_status_id and ");                
           sql.append("                        order_status.description in ('NEW','WAITING FOR INVENTORY','WAITING CREDIT APPROVAL') "); 
           sql.append("   join order_status line_status on line_status.order_status_id = order_line.order_status_id and "); 
           sql.append("                                    line_status.description in ('NEW') ");                        
           sql.append("   left outer join promotion on promotion.promo_id = order_line.promo_id "); 
           sql.append("   where order_line.earliest_ship is null or "); 
           sql.append("      order_line.earliest_ship <= trunc(now) + 28 or "); 
           sql.append("      promotion.ship_date <= trunc(now) + 28 "); 
           sql.append("   group by order_header.warehouse_id, order_line.item_ea_id ");
           sql.append(") ord_qty on ord_qty.warehouse_id = warehouse.warehouse_id and ord_qty.item_ea_id = item_entity_attr.item_ea_id ");
           
           sql.append("left outer join ( ");              
           sql.append("   select warehouse_id, item_nbr, sum(qty_shipped) qty_shipped ");  
           sql.append("   from itemsales ");              
           sql.append("   where invoice_date > trunc(now) - 28  ");  
           sql.append("   group by warehouse_id, item_nbr ");              
           sql.append(") wk4_sales on wk4_sales.warehouse_id = warehouse.warehouse_id and wk4_sales.item_nbr = item_entity_attr.item_id ");            
           sql.append("left outer join fcst_4_week on fcst_4_week.whs_name = warehouse.name and fcst_4_week.item_id = item_entity_attr.item_id ");
           sql.append("left outer join fcst_8_week on fcst_8_week.whs_name = warehouse.name and fcst_8_week.item_id = item_entity_attr.item_id ");
           
           sql.append("where item_entity_attr.item_type_id not in (8,9) ");
           
           sql.append("order by warehouse.name, ");
           sql.append("   case (nvl(ord_qty.qty_ordered, 0) + nvl(fcst_4_week.fcst_qty, 0)) "); 
           sql.append("      when 0 then 1.0 "); 
           sql.append("      else nvl(qoh, 0)  / (nvl(ord_qty.qty_ordered, 0) + nvl(fcst_4_week.fcst_qty, 0)) "); 
           sql.append("   end, "); 
           sql.append("   dept_num, vendor.name, item_entity_attr.item_id "); 
           m_PoData = m_EdbConn.prepareStatement(sql.toString());
           
           isPrepared = true;
        }
        
        catch ( SQLException ex ) {
           log.error("[NeverOut]", ex);
        }
        
        finally {
           sql = null;
        }         
     }
     else
        log.error("[NeverOut].prepareStatements - null edb or fascor connection");
      
     return isPrepared;
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();      
      SimpleDateFormat df = new SimpleDateFormat("MM dd yy");
      
      fileName.append(df.format(new Date()));
      fileName.append(" Never Out Report.xlsx");
      m_FileNames.add(fileName.toString());
      
      df = null;
   }
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    * Note - 
    *    The styles came from the original Excel worksheet.
    */
   private void setupWorkbook()
   {      
      XSSFCellStyle styleTxtC = null;       // Text centered
      XSSFCellStyle styleTxtL = null;       // Text left justified
      XSSFCellStyle styleInt = null;        // 0 decimals and a comma
      XSSFCellStyle csIntRed = null;        // 0 decimals and a comma, red foreground
      XSSFCellStyle styleDouble = null;     // numeric 1 decimal and a comma
      XSSFCellStyle stylePercent = null;    // Integer percent with comma
      XSSFCellStyle csPercentRed = null;	  // Integer percent with comma red
      XSSFCellStyle csPercentOrange = null; // Integer percent with comma orange
      XSSFCellStyle csPercentGold = null;	  // Integer percent with comma gold
      XSSFCellStyle csPercentYellow = null; // Integer percent with comma yellow
      XSSFDataFormat format = null;
      XSSFFont font = null;
            
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
            
      m_CellStyles = new XSSFCellStyle[] {
      	styleTxtC,    // col 0 warehouse
         styleTxtC,    // col 1 dept
         styleTxtL,    // col 2 vendor name
         styleTxtL,    // col 3 buyer
         styleTxtC,    // col 4 item id        
         styleTxtL,    // col 5 item desc
         styleInt,     // col 6 qoh
         styleInt,     // col 7 open pos
         styleInt,	  // col 8 qoh + pos
         styleInt,     // col 9 monthly sales
         styleInt,     // col 10 open orders
         styleInt,	  // col 11 4 week forecast
         styleInt,	  // col 12 8 week forecast
         stylePercent, // col 13 qoh coverage
         stylePercent, // col 14 qoh + 4 week forecast coverage
         stylePercent, // col 16 qoh + 4 week forecast + open po coverage
         stylePercent  // col 16 qoh + 8 week forecast + open po coverage
      };
      
      m_CellStylesEx = new XSSFCellStyle[] {
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
}
