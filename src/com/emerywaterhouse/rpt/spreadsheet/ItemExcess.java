/**
 * File: ItemExcess.java
 * Description: Item excess report, based on InventoryAging.java @see com.emerywaterhouse.rpt.spreadsheet.InventoryAging
 * 
 * @author Eric Verge
 * 
 * Create Date: 09/12/14
 * Last Update: $Id: ItemExcess.java,v 1.3 2014/12/03 22:56:15 everge Exp $
 * 
 * History:
 *    $Log: ItemExcess.java,v $
 *    Revision 1.3  2014/12/03 22:56:15  everge
 *    Updated prepared statements to use only bind variables. Made logging more explicit. Other minor fixes.
 *
 *    Revision 1.2  2014/09/19 18:26:32  everge
 *    Minor fixes. Filtered out credit sale items.
 *
 *    Revision 1.1  2014/09/15 18:08:56  everge
 *    Initial commit
 *
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
import java.util.HashMap;

import org.apache.axis2.rpc.client.RPCServiceClient;
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

public class ItemExcess extends Report 
{
   
   private static final String PTLD_WHS_ID = "01";
   private static final short MAX_COLS = 46;
   
   RPCServiceClient m_Client;
   
   private PreparedStatement m_SaleData; // item sale data
   private PreparedStatement m_AvgCostData; // from emeryd.iciloc
	   
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
   
   private HashMap<String, Double> m_AvgCostPtld;
   private HashMap<String, Double> m_AvgCostPitt;

   /**
    * Constructor
    */
   public ItemExcess() 
   {
      super();
      
      m_Filtered = true;
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet("EXCESS Page 1");
      m_AvgCostPtld = new HashMap<String, Double>();
      m_AvgCostPitt = new HashMap<String, Double>();
      
      setupWorkbook();
   }
   
   public void finalize() throws Throwable
   {
      if ( m_CellStyles != null ) {
         for ( int i = 0; i < m_CellStyles.length; i++ )
            m_CellStyles[i] = null;
      }
      
      m_AvgCostPtld = null;
      m_AvgCostPitt = null;
      m_Sheet = null;
      m_Wrkbk = null;      
      m_CellStyles = null;
      m_BegDate = null;
      m_EndDate = null;
      m_FilterDate = null;
      m_SaleData = null;
      m_AvgCostData = null;
      m_Dept = null;
            
      super.finalize();
   }
   
   /**
    * Creates Excel file from retrieved data
    * @return true if the file is successfully built
    * @throws FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException 
   {
      FileOutputStream outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
      XSSFRow row = null;
      int rowNum = 0;
      int colNum = 0;
      int pageNum = 1;
      boolean result = false;
      String itemId = null;
      Date statusDate = null;
      double ptldCost = 0.0;
      double pittCost = 0.0;
      double compAvgCost = 0.0;
      double ptldAvgCost = 0.0;
      double pittAvgCost = 0.0;
      int compShipped = 0;
      int ptldShipped = 0;
      int pittShipped = 0;
      int compOnHand = 0;
      int ptldOnHand = 0;
      int pittOnHand = 0;
      int compExcess = 0;
      int ptldExcess = 0;
      int pittExcess = 0;
      int unconsolExcs = 0;
      int ptldUtil = 0;
      int pittUtil = 0;
      int ptldBal = 0;
      int pittBal = 0;
      ResultSet data = null;

      try {
         rowNum = createCaptions();
         
         int i = 1;
         if (m_Filtered && m_FilterDate != null && m_FilterDate.length() > 0) {
             m_SaleData.setString(i++, m_FilterDate);
          }
         if (m_Dept != null && m_Dept.length() > 0) {
            m_SaleData.setString(i++, m_Dept);
         }
         m_SaleData.setString(i++, m_BegDate);
         m_SaleData.setString(i++, m_EndDate);
         m_SaleData.setString(i++, m_BegDate);
         m_SaleData.setString(i++, m_EndDate);
         
         if (m_Filtered && m_FilterDate != null && m_FilterDate.length() > 0) {
            m_SaleData.setString(i++, m_FilterDate);
         }
         
         if (m_Dept != null && m_Dept.length() > 0) {
            m_SaleData.setString(i++, m_Dept);
         }
         
         m_SaleData.setString(i++, m_BegDate);
         m_SaleData.setString(i++, m_EndDate);
         
         data = m_SaleData.executeQuery();
         
         while (data.next() && m_Status == RptServer.RUNNING) {
            itemId = data.getString("item_id");            
            ptldCost = data.getDouble("ptld_cost");
            pittCost = data.getDouble("pitt_cost");
            
            ptldOnHand = data.getInt("oh_port");
            pittOnHand = data.getInt("oh_pitt");
            compOnHand = ptldOnHand + pittOnHand;
            
            ptldAvgCost = m_AvgCostPtld.containsKey(itemId) ? m_AvgCostPtld.get(itemId) : ptldCost;
            if (ptldAvgCost <= 0) 
               ptldAvgCost = ptldCost;
            
            pittAvgCost = m_AvgCostPitt.containsKey(itemId) ? m_AvgCostPitt.get(itemId) : pittCost;
            if (pittAvgCost <= 0) 
               pittAvgCost = ptldCost;    
            
            compAvgCost = (compOnHand == 0 ? 0 : ((ptldOnHand * ptldAvgCost) + (pittOnHand * pittAvgCost)) / compOnHand);
            
            ptldShipped = data.getInt("shipped_port");
            pittShipped = data.getInt("shipped_pitt");
            compShipped = ptldShipped + pittShipped;
            
            compExcess = compOnHand - compShipped;
            if (compExcess < 0) 
               compExcess = 0;
            
            ptldExcess = ptldOnHand - ptldShipped;
            if (ptldExcess < 0) 
               ptldExcess = 0;
            
            pittExcess = pittOnHand - pittShipped;
            if (pittExcess < 0) 
               pittExcess = 0;
            
            unconsolExcs = ptldExcess + pittExcess;
            ptldUtil = ptldShipped - ptldOnHand;
            if (ptldUtil < 0) 
               ptldUtil = 0;
            
            pittUtil = pittShipped - pittOnHand;
            if (pittUtil < 0) 
               pittUtil = 0;
            
            ptldBal = (pittExcess > ptldUtil ? ptldUtil : pittExcess);
            pittBal = (ptldExcess > pittUtil ? pittUtil : ptldExcess);
            
            row = createRow(rowNum++, MAX_COLS);
            colNum = 0;
            
            if ( row != null ) {
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(itemId));
               //row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("disposition")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("disp_port")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("disp_pitt")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("active_port")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("active_pitt")));
               row.getCell(colNum++).setCellValue(sdf.format(data.getDate("setup_date")));
               
               statusDate = data.getDate("ptld_status_date");
               if ( statusDate == null )
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString("N"));               
               else
                  row.getCell(colNum++).setCellValue(sdf.format(statusDate));
               
               statusDate = data.getDate("pitt_status_date");
               if ( statusDate == null )
                  row.getCell(colNum++).setCellValue(new XSSFRichTextString("N"));               
               else
                  row.getCell(colNum++).setCellValue(sdf.format(statusDate));
               
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("velocity_port")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("velocity_pitt")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("description")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("name")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("vendor_id")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("dept_num")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("soq_comment")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("guar_sale")));              
               row.getCell(colNum++).setCellValue(data.getInt("ptld_stock_pack"));
               row.getCell(colNum++).setCellValue(data.getInt("pitt_stock_pack"));               
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(data.getString("nbc")));
               row.getCell(colNum++).setCellValue(data.getInt("convenience_pack_1"));
               row.getCell(colNum++).setCellValue(data.getDouble("cube"));
               row.getCell(colNum++).setCellValue(ptldCost);
               row.getCell(colNum++).setCellValue(pittCost);
               row.getCell(colNum++).setCellValue(compAvgCost > 0 ? compAvgCost : ptldCost);
               row.getCell(colNum++).setCellValue(ptldAvgCost);
               row.getCell(colNum++).setCellValue(pittAvgCost);
               row.getCell(colNum++).setCellValue(compShipped);
               row.getCell(colNum++).setCellValue(compOnHand);
               row.getCell(colNum++).setCellValue(compExcess);
               row.getCell(colNum++).setCellValue(compExcess * compAvgCost);
               row.getCell(colNum++).setCellValue(ptldShipped);
               row.getCell(colNum++).setCellValue(ptldOnHand);
               row.getCell(colNum++).setCellValue(ptldExcess);
               row.getCell(colNum++).setCellValue(ptldExcess * ptldAvgCost);
               row.getCell(colNum++).setCellValue(pittShipped);
               row.getCell(colNum++).setCellValue(pittOnHand);
               row.getCell(colNum++).setCellValue(pittExcess);
               row.getCell(colNum++).setCellValue(pittExcess * pittAvgCost);
               row.getCell(colNum++).setCellValue(unconsolExcs);
               row.getCell(colNum++).setCellValue(unconsolExcs * compAvgCost);
               row.getCell(colNum++).setCellValue(ptldUtil);
               row.getCell(colNum++).setCellValue(ptldBal);
               row.getCell(colNum++).setCellValue(ptldBal * ptldAvgCost);
               row.getCell(colNum++).setCellValue(pittUtil);
               row.getCell(colNum++).setCellValue(pittBal);
               row.getCell(colNum++).setCellValue(pittBal * pittAvgCost);
            }
            
            if ( rowNum > 65000 ) {
               pageNum++;
               m_Sheet = m_Wrkbk.createSheet("EXCESS Page " + pageNum);
               rowNum = createCaptions();
            }
         }
         
         m_Wrkbk.write(outFile);
         data.close();

         result = true;
      }
      
      catch (Exception ex) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[ItemExcess]", ex);
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
         sdf = null;
         row = null;
         itemId = null;
         statusDate = null;
         data = null;
      }
      
      return result;
   }
   
   private void closeStatements() 
   {
      closeStmt(m_SaleData);
      closeStmt(m_AvgCostData);
   }
   
   /**
    * Sets up the Excel sheet with titles and column headings
    * @return the number of the next empty row
    */
   private int createCaptions() 
   {
      if ( m_Sheet == null )
         return 0;
      
      XSSFFont font;
      XSSFRow row = null;
      XSSFCell cell = null;
      XSSFCellStyle styleTitle; // bold, left-aligned style
      XSSFCellStyle styleCaption;  // bold, center-aligned style
      
      CellRangeAddress region = null;
      int rowNum = 0;
      int colNum = 0;
      String title = String.format("EXCESS Report: %s - %s", m_BegDate, m_EndDate);
      
      // build font for title and column captions
      font = m_Wrkbk.createFont();
      font.setFontHeightInPoints((short)10);
      font.setFontName("Helvetica");
      font.setBold(true);
      
      styleTitle = m_Wrkbk.createCellStyle();
      styleTitle.setFont(font);
      styleTitle.setAlignment(HorizontalAlignment.LEFT);
      
      styleCaption = m_Wrkbk.createCellStyle();
      styleCaption.setFont(font);
      styleCaption.setAlignment(HorizontalAlignment.CENTER);
      styleCaption.setWrapText(true);
      
      //
      // set the report title
      row = m_Sheet.createRow(rowNum++);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitle);
      cell.setCellValue(new XSSFRichTextString(title));
      
      //
      // Set the filter date if there is one
      title = "Filter: ";
      if ( m_Filtered && m_FilterDate != null )
         title += ("On    Date: " + m_FilterDate);     
      else
         title += "Off";
      
      row = m_Sheet.createRow(rowNum++);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitle);
      cell.setCellValue(new XSSFRichTextString(title));
      
      row = m_Sheet.createRow(rowNum++);
      
      //
      // Merge the title cells.  Gives a better look to the report.
      region = new CellRangeAddress(0, 0, 0, 3);
      m_Sheet.addMergedRegion(region);
      
      region = new CellRangeAddress(1, 1, 0, 3);
      m_Sheet.addMergedRegion(region);
      
      //
      // Build column titles
      row = m_Sheet.createRow(rowNum++);
      
      if ( row != null ) {
         for ( int i = 0; i < MAX_COLS; i++ ) {
            cell = row.createCell(i);
            cell.setCellType(CellType.STRING);
            cell.setCellStyle(styleCaption);
         }
         
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Item #"));
         m_Sheet.setColumnWidth(colNum, 3800);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Status Ptld"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Status Pitt"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Active Ptld"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Active Pitt"));
         
         m_Sheet.setColumnWidth(colNum, 2900);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Add Date"));
         
         m_Sheet.setColumnWidth(colNum, 2900);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Status Date Ptld"));         
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Status Date Ptld"));
         
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Velocity Ptld"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Velocity Pitt"));
         
         m_Sheet.setColumnWidth(colNum, 13000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Description"));
         
         m_Sheet.setColumnWidth(colNum, 6000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Vendor"));
         
         m_Sheet.setColumnWidth(colNum, 2500);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Vendor #"));
         
         m_Sheet.setColumnWidth(colNum, 1400);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Dept"));
         
         m_Sheet.setColumnWidth(colNum, 10000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("SOQ"));
         
         m_Sheet.setColumnWidth(colNum, 2800);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Guar Sale"));
         
         m_Sheet.setColumnWidth(colNum, 2000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Stock Pack Ptld"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Stock Pack Pitt"));
         
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("No Broken Case"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Conv Pack 1"));
         
         m_Sheet.setColumnWidth(colNum, 3000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Cube"));
         
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Cost Ptld"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Cost Pitt"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Company Avg Cost"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Portland Avg Cost"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Pittston Avg Cost"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Company Units Shipped"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Company OH"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Company Excess Units"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Company Excess $'s"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Portland Units Shipped"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Portland OH"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Portland Excess Units"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Portland Excess $'s"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Pittston Units Shipped"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Pittston OH"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Pittston Excess Units"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Pittston Excess $'s"));
         
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Unconsolidated Excess Units"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Unconsolidated Excess $'s"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Portland Potential Utilization"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("To Portland Potential Balance Units"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("To Portland Potential Balnace $'s"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Pittston Potential Utilization"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("To Pittston Potential Balance Units"));
         row.getCell(colNum).setCellValue(new XSSFRichTextString("To Pittston Potential Balnace $'s"));
      }
      
      // freeze column titles and item# column to ease scrolling through data
      m_Sheet.createFreezePane(1, rowNum);
      
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
         if ( m_EdbConn == null && m_RptProc != null )
            m_EdbConn = m_RptProc.getEdbConn();
         
         if ( prepareStatements() ) {
            splitCostData();
            created = buildOutputFile(); 
         }
      }
      
      catch ( Exception ex ) {
         log.fatal("[ItemExcess]", ex);
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
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements() 
   {
      StringBuffer sql = new StringBuffer();
      
      boolean isPrepared = false;
      
      if ( m_EdbConn != null ) {
         try {
            sql.append("select ");
            sql.append("   item_entity_attr.item_id, ");
            //sql.append("   nvl(item_act1.disposition, item_act2.disposition) as disposition, ");
            sql.append("   item_act1.disposition as disp_port, ");
            sql.append("   item_act2.disposition as disp_pitt, ");
            sql.append("   coalesce(item_act1.active, 'N') as active_port, ");
            sql.append("   coalesce(item_act2.active, 'N') active_pitt, ");
            sql.append("   setup_date,  ");
            sql.append("   item_act1.status_date as ptld_status_date, ");
            sql.append("   item_act2.status_date as pitt_status_date, ");
            sql.append("   coalesce(velocity_port, 'N') as  velocity_port, ");
            sql.append("   coalesce(velocity_pitt, 'N') as velocity_pitt, ");
            sql.append("   item_entity_attr.description, item_entity_attr.vendor_id, vendor.name, emery_dept.dept_num, soq_comment, ");
            sql.append("   decode(vendor_guaranteed, 0, 'N', 'Y') as guar_sale, ");
            sql.append("   item_act1.stock_pack as ptld_stock_pack, ");
            sql.append("   item_act2.stock_pack as pitt_stock_pack, ");
            sql.append("   decode(broken_case.broken_case_id, 1, 'N', 'Y') as NBC, ");
            sql.append("   convenience_pack_1, ");
            sql.append("   coalesce(item_act1.cube, item_act2.cube) as cube, ");
            sql.append("   coalesce(ptld_price.buy, 0) as ptld_cost, ");
            sql.append("   coalesce(pitt_price.buy, 0) as pitt_cost, ");
            sql.append("   coalesce(item_ship1.shipped, 0) as shipped_port, ");
            sql.append("   coalesce(item_act1.qoh, 0) as OH_port, ");
            sql.append("   coalesce(item_ship2.shipped, 0) as shipped_pitt, ");
            sql.append("   coalesce(item_act2.qoh, 0) as OH_pitt ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            
            if ( m_Filtered && m_FilterDate != null && m_FilterDate.length() > 0 ) {
               sql.append(" and ejd_item.setup_date <= to_date(?, 'mm/dd/yyyy') ");
            }
            
            sql.append("left outer join ejd_item_price ptld_price on ptld_price.ejd_item_id = item_entity_attr.ejd_item_id and ptld_price.warehouse_id = 1 ");
            sql.append("left outer join ejd_item_price pitt_price on pitt_price.ejd_item_id = item_entity_attr.ejd_item_id and pitt_price.warehouse_id = 2 ");
            sql.append("left outer join ( ");
            sql.append("   select ");
            sql.append("   ejd_item_id, decode(can_plan, 0, 'N', 'Y') as active, disposition, coalesce(item_velocity.velocity, 'N') as velocity_port, status_date, ");
            sql.append("   qoh, cube, stock_pack ");
            sql.append("   from ejd_item_warehouse ");
            sql.append("   join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id ");
            sql.append("   join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id ");
            sql.append("   where ejd_item_warehouse.warehouse_id = 1 ");
            sql.append(") item_act1 on item_act1.ejd_item_id = ejd_item.ejd_item_id ");
            sql.append("left outer join ( ");
            sql.append("   select ");
            sql.append("   ejd_item_id, decode(can_plan, 0, 'N', 'Y') as active, disposition, coalesce(item_velocity.velocity, 'N') as velocity_pitt, status_date, ");
            sql.append("   qoh, cube, stock_pack ");
            sql.append("   from ejd_item_warehouse ");
            sql.append("   join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id ");
            sql.append("   join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id ");
            sql.append("   where ejd_item_warehouse.warehouse_id = 2 ");
            sql.append(") item_act2 on item_act2.ejd_item_id = ejd_item.ejd_item_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join emery_dept on emery_dept.dept_id = ejd_item.dept_id ");
            
            if ( m_Dept != null && m_Dept.length() > 0 ) {
               sql.append(" and emery_dept.dept_num = ? ");
            }
            
            sql.append("join broken_case on broken_case.broken_case_id = ejd_item.broken_case_id ");
            sql.append("left outer join ( ");
            sql.append("   select item_id, sum(qty_shipped) as shipped ");
            sql.append("   from sale_item ");
            sql.append("   join cust_warehouse on cust_warehouse.customer_id = sale_item.customer_id and cust_warehouse.warehouse_id = 1 ");
            sql.append("   where invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') and ");
            sql.append("      sale_type in ('WAREHOUSE','TOOLREPAIR','APG WHS') and tran_type <> 'CREDIT' ");
            sql.append("   group by item_id ");
            sql.append(") item_ship1 on item_ship1.item_id = item_entity_attr.item_id ");
            sql.append("left outer join ( ");
            sql.append("   select item_id, sum(qty_shipped) shipped ");
            sql.append("   from sale_item ");
            sql.append("   join cust_warehouse on cust_warehouse.customer_id = sale_item.customer_id and cust_warehouse.warehouse_id = 2 ");
            sql.append("   where invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') and ");
            sql.append("      sale_type in ('WAREHOUSE','TOOLREPAIR','APG WHS') and ");
            sql.append("      tran_type <> 'CREDIT' ");
            sql.append("   group by item_id ");
            sql.append(") item_ship2 on item_ship2.item_id = item_entity_attr.item_id ");
            sql.append("where  ");
            sql.append("   item_entity_attr.item_type_id < 8 and   ");
            sql.append("   (coalesce(item_ship1.shipped, 0) > 0 or coalesce(item_ship2.shipped, 0) > 0) and (coalesce(item_act1.qoh, 0) > 0 or coalesce(item_act2.qoh, 0) > 0)  ");
            
            //
            // union in items with quantity, but no sales
            sql.append("union ");
            sql.append("select ");
            sql.append("   item_entity_attr.item_id, ");
           // sql.append("   nvl(item_act1.disposition, item_act2.disposition) as disposition, ");
            sql.append("   item_act1.disposition as disp_port, ");
            sql.append("   item_act2.disposition as disp_pitt, ");
            sql.append("   coalesce(item_act1.active, 'N') as active_port, ");
            sql.append("   coalesce(item_act2.active, 'N') active_pitt, ");
            sql.append("   setup_date,  ");
            sql.append("   item_act1.status_date as ptld_status_date, ");
            sql.append("   item_act2.status_date as pitt_status_date, ");
            sql.append("   coalesce(velocity_port, 'N') as  velocity_port, ");
            sql.append("   coalesce(velocity_pitt, 'N') as velocity_pitt, ");
            sql.append("   item_entity_attr.description, item_entity_attr.vendor_id, vendor.name, emery_dept.dept_num, soq_comment, ");
            sql.append("   decode(vendor_guaranteed, 0, 'N', 'Y') as guar_sale, ");
            sql.append("   item_act1.stock_pack as ptld_stock_pack, ");
            sql.append("   item_act2.stock_pack as pitt_stock_pack, ");
            sql.append("   decode(broken_case.broken_case_id, 1, 'N', 'Y') as NBC, ");
            sql.append("   convenience_pack_1, ");
            sql.append("   coalesce(item_act1.cube, item_act2.cube) as cube, ");
            sql.append("   coalesce(ptld_price.buy, 0) as ptld_cost, ");
            sql.append("   coalesce(pitt_price.buy, 0) as pitt_cost, ");
            sql.append("   0 as shipped_port, ");
            sql.append("   coalesce(item_act1.qoh, 0) as OH_port, ");
            sql.append("   0 as shipped_pitt, ");
            sql.append("   coalesce(item_act2.qoh, 0) as OH_pitt ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            
            if ( m_Filtered && m_FilterDate != null && m_FilterDate.length() > 0 ) {
               sql.append(" and ejd_item.setup_date <= to_date(?, 'mm/dd/yyyy') ");
            }
            
            sql.append("left outer join ejd_item_price ptld_price on ptld_price.ejd_item_id = item_entity_attr.ejd_item_id and ptld_price.warehouse_id = 1 ");
            sql.append("left outer join ejd_item_price pitt_price on pitt_price.ejd_item_id = item_entity_attr.ejd_item_id and pitt_price.warehouse_id = 2 ");
            sql.append("left outer join ( ");
            sql.append("   select ");
            sql.append("   ejd_item_id, decode(can_plan, 0, 'N', 'Y') as active, disposition, coalesce(item_velocity.velocity, 'N') as velocity_port, status_date, ");
            sql.append("   qoh, cube, stock_pack ");
            sql.append("   from ejd_item_warehouse ");
            sql.append("   join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id ");
            sql.append("   join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id ");
            sql.append("   where ejd_item_warehouse.warehouse_id = 1 ");
            sql.append(") item_act1 on item_act1.ejd_item_id = ejd_item.ejd_item_id ");
            sql.append("left outer join ( ");
            sql.append("   select ");
            sql.append("   ejd_item_id, decode(can_plan, 0, 'N', 'Y') as active, disposition, coalesce(item_velocity.velocity, 'N') as velocity_pitt, status_date, ");
            sql.append("   qoh, cube, stock_pack ");
            sql.append("   from ejd_item_warehouse ");
            sql.append("   join item_disp on item_disp.disp_id = ejd_item_warehouse.disp_id ");
            sql.append("   join item_velocity on item_velocity.velocity_id = ejd_item_warehouse.velocity_id ");
            sql.append("   where ejd_item_warehouse.warehouse_id = 2 ");
            sql.append(") item_act2 on item_act2.ejd_item_id = ejd_item.ejd_item_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join emery_dept on emery_dept.dept_id = ejd_item.dept_id ");
            
            if ( m_Dept != null && m_Dept.length() > 0 ) {
               sql.append(" and emery_dept.dept_num = ? ");
            }
            
            sql.append("join broken_case on broken_case.broken_case_id = ejd_item.broken_case_id ");
            
            sql.append("where ");
            sql.append("   item_entity_attr.item_type_id < 8 and  ");
            sql.append("   item_entity_attr.item_id in ( ");
            sql.append("      select item_id ");
            sql.append("      from ejd_item_warehouse ");
            sql.append("      join item_entity_attr on item_entity_attr.ejd_item_id = ejd_item_warehouse.ejd_item_id and item_type_id not in (8,9) ");
            sql.append("      where qoh > 0 and warehouse_id in (1,2) ");
            //sql.append("      minus ");
            sql.append("      except ");
            sql.append("      select distinct item_id ");
            sql.append("      from sale_item ");
            sql.append("      join cust_warehouse cw on cw.customer_id = sale_item.customer_id ");
            sql.append("      where ");
            sql.append("         invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') and ");
            sql.append("         sale_type in ('WAREHOUSE','TOOLREPAIR','APG WHS') and ");
            sql.append("         cw.customer_id not in ('199796','037940') and tran_type <> 'CREDIT' ");
            sql.append("   ) and ");
            sql.append("(coalesce(item_act1.qoh, 0) > 0 or coalesce(item_act2.qoh, 0) > 0) ");
            sql.append("order by item_id");                       
            
            m_SaleData = m_EdbConn.prepareStatement(sql.toString());
                        
            //
            // get average cost data from emeryd
            sql.setLength(0);
            sql.append("select itemno as item_id, location as whs_id, decode(qtyonhand, 0, 0, (totalcost / qtyonhand)) as avg_cost ");
            sql.append("from iciloc ");
            
            m_AvgCostData = m_EdbConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         } 
         
         catch (SQLException ex) {
            log.error("[ItemExcess]", ex);
         } 
         
         finally {
            sql = null;
         }
      }
      
      return isPrepared;
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();      
      String tmp = Long.toString(System.currentTimeMillis());
      
      for ( Param param : params ) {         
         switch (param.name.toLowerCase()) {
            case "begdate":
               m_BegDate = param.value;
               break;
            case "enddate":
               m_EndDate = param.value;
               break;
            case "filtered":
               m_Filtered = param.value.equalsIgnoreCase("true");
               break;
            case "filterdate":
               m_FilterDate = param.value;
               break;
            case "dept":
               m_Dept = param.value;
               break;
         }
      }
      
      fileName.append("excess");
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
         styleTxtC,  // col 0 item#
         styleTxtC,  // col 1 status portland
         styleTxtC,  // col 2 status pittston
         styleTxtC,  // col 3 active portland
         styleTxtC,  // col 4 active pittston
         styleTxtC,  // col 5 add date
         styleTxtC,  // col 6 status change date ptld
         styleTxtC,  // col 7 status change date pitt
         styleTxtC,  // col 8 velocity portland
         styleTxtC,  // col 9 velocity pittston
         styleTxtL,  // col 10 description
         styleTxtL,  // col 11 vendor name
         styleTxtC,  // col 12 vendor#
         styleTxtC,  // col 13 dept#
         styleTxtL,  // col 14 SOQ
         styleTxtC,  // col 15 Guar Sale
         styleInt,   // col 16 stock pack ptld
         styleInt,   // col 17 stock pack pitt
         styleTxtC,  // col 18 NBC
         styleInt,   // col 19 convenience pack 1
         styleDouble,// col 20 cube
         styleMoney, // col 21 current cost ptld
         styleMoney, // col 22 current cost pitt
         styleMoney, // col 23 company avg cost
         styleMoney, // col 24 portland avg cost
         styleMoney, // col 25 pittston avg cost
         styleInt,   // col 26 company units shipped
         styleInt,   // col 27 company units on-hand
         styleInt,   // col 28 company excess units
         styleMoney, // col 29 company excess $'s
         styleInt,   // col 30 portland units shipped
         styleInt,   // col 31 portland units on-hand
         styleInt,   // col 32 portland excess units
         styleMoney, // col 33 portland excess $'s
         styleInt,   // col 34 pittston units shipped
         styleInt,   // col 35 pittston units on-hand
         styleInt,   // col 36 pittston excess units
         styleMoney, // col 37 pittston excess $'s
         styleInt,   // col 38 unconsolidated excess units
         styleMoney, // col 39 unconsolidated excess $'s
         styleInt,   // col 40 portland potential utilization
         styleInt,   // col 41 to portland potential balance units
         styleMoney, // col 42 to portland potential balance $'s
         styleInt,   // col 43 pittston potential utilization
         styleInt,   // col 44 to pittston potential balance units
         styleMoney  // col 45 to pittston potential balance $'s
      };
      
      styleTxtC = null;
      styleTxtL = null;
      styleInt = null;
      styleDouble = null;
      styleMoney = null;
      format = null;
   }
   
   /**
    * Splits the AvgCostData result set into two HashMaps based on item location
    */
   private void splitCostData() {
      ResultSet data = null;
      
      try {
         data = m_AvgCostData.executeQuery();
         
         while (data.next()) {
            if (data.getString("whs_id").equals(PTLD_WHS_ID)) {
               m_AvgCostPtld.put(data.getString("item_id"), data.getDouble("avg_cost"));
            }
            else {
               m_AvgCostPitt.put(data.getString("item_id"), data.getDouble("avg_cost"));
            }
         }
      } 
      catch (SQLException ex) {
         log.error("[ItemExcess]", ex);
      }
   }
   
   /*For debugging locally only.
   public static void main(String args[]) throws FileNotFoundException, SQLException 
   {
      org.apache.log4j.BasicConfigurator.configure();
      ItemExcess ie = new ItemExcess();
      Param[] parms = new Param[] {
            new Param("begdate", "03/13/2018", "begdate"), 
            new Param("enddate", "03/13/2018","enddate"),
            new Param("filtered", "true", "filtered"),
            //new Param("filteredate", "03/13/2018", "filterdate"),
            new Param("dept", "06", "dept")
            //new Param("dc", "01", "dc") //some kinda weird constructer puts name at the end
      };
      
      ArrayList<Param> parmslist = new ArrayList<Param>();
      for (Param p : parms) {
         parmslist.add(p);
      }
      
      //System.out.println(parmslist.size());
      ie.setParams(parmslist);
      java.util.Properties connProps = new java.util.Properties();
      connProps.put("user", "ejd");
      connProps.put("password", "boxer");

      ie.m_Status = RptServer.RUNNING;
      ie.m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33:5444/emery_jensen", connProps);
      ie.m_FilePath = "C:/Users/JFisher/temp/";
      ie.createReport();
    
      System.out.println("done");
   }*/
}
