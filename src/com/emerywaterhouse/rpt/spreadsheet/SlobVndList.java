/**
 * File: SlobVndList.java
 * Description: SLOB vendor summary.
 *
 * @author Jeff Fisher
 *
 * Create Date: 09/20/2010
 * Last Update: $Id: SlobVndList.java,v 1.1 2010/10/02 12:13:05 jfisher Exp $
 *
 * History:
 *    $Log: SlobVndList.java,v $
 *    Revision 1.1  2010/10/02 12:13:05  jfisher
 *    Initial production release.
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

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

public class SlobVndList extends Report
{
   private static final short MAX_COLS = 6;
   
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
   private String m_FasId;
   private int m_WhsId;
   private String m_WhsName;
   private String m_AccWhsId;
   
   /**
    * 
    */
   public SlobVndList()
   {
      super();
      
      m_Filtered = false;
      m_WhsName = "ALL";
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet("SLOB Summary");
      
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
      m_AccWhsId = null;
      m_WhsName = null;
      m_FasId = null;
            
      super.finalize();
   }
   
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {      
      String msg = "processing vendor %d";
      XSSFRow row = null;
      int varIdx = 1;
      int rowNum = 0;
      int colNum = 0;     
      boolean result = false;
      double slob12 = 0.0;
      double slob24 = 0.0;
      double invTot = 0.0;
      long vndId = 0;
      String vndName = "";
      String vndDept = "";
      FileOutputStream outFile = null;
      ResultSet saleData = null;
      ResultSet dcData = null;
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         if ( m_WhsId > 0 ) {
            try {
               m_GetDCName.setInt(1, m_WhsId);
               dcData = m_GetDCName.executeQuery();
               
               if ( dcData.next() )
                  m_WhsName = dcData.getString("name");
            }
            
            finally {
               closeRSet(dcData);
            }
         }

         rowNum = createCaptions();
         
         //
         // Set the bind variables.  Because they are not fixed we can simply
         // increment in the correct order in the sql statement.
         // param order -
         //    department
         //    warehouse - fascor id
         //    itemsales - warehouse id
         //    itemsales - begin date
         //    itemsales - end date
         //    item - setup date
         if ( m_Dept != null && m_Dept.length() > 0 )
            m_SaleData.setString(varIdx++, m_Dept);
         
         if ( m_WhsId  > 0 ) {
            m_SaleData.setString(varIdx++, m_FasId);
            m_SaleData.setInt(varIdx++, m_WhsId);
         }
         
         m_SaleData.setString(varIdx++, m_BegDate);
         m_SaleData.setString(varIdx++, m_EndDate);
         
         if ( m_Filtered ) {
            if ( m_FilterDate != null && m_FilterDate.length() > 0 )
               m_SaleData.setString(varIdx++, m_FilterDate);
         }
                  
         saleData = m_SaleData.executeQuery();
          
         while ( saleData.next() && m_Status == RptServer.RUNNING ) {
            vndId = saleData.getLong("vendor_id");
            vndName = saleData.getString("name");
            vndDept = saleData.getString("dept_num");
            invTot = saleData.getDouble("total_inventory");
            slob12 = saleData.getDouble("slob12");
            slob24 = saleData.getDouble("slob24");
            
            setCurAction(String.format(msg, vndId));
                    
            row = createRow(rowNum++, MAX_COLS);
            colNum = 0;
            
            if ( row != null ) {
               row.getCell(colNum++).setCellValue(vndId);
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(vndName));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(vndDept));
               row.getCell(colNum++).setCellValue(invTot);
               row.getCell(colNum++).setCellValue(slob12);
               row.getCell(colNum++).setCellValue(slob24);
            }
         }
         
         m_Wrkbk.write(outFile);
         closeRSet(saleData);

         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {        
         try {
            outFile.close();
         }

         catch( Exception e ) {
            ;
         }

         outFile = null;
         row = null;
         saleData = null;
         dcData = null;
      }

      return result;
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_SaleData);
      closeStmt(m_GetDCName);
   }
   
   /**
    * Sets the captions on the report.
    */
   private int createCaptions()
   {
      XSSFCellStyle styleCaption = null;
      XSSFCellStyle styleTitle = null;
      XSSFFont font = null;
      XSSFCell cell = null;
      XSSFRow row = null;
      CellRangeAddress region = null;      
      int rowNum = 0;
      int colNum = 0;      
      StringBuffer title = new StringBuffer("SLOB Summary Report: ");
            
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
      // Create the filter date information
      title.setLength(0);
      if ( m_Filtered ) {
         if ( m_FilterDate != null ) {         
            title.append("Filter: On, Date: ");
            title.append(m_FilterDate);
         }
      }
      else
         title.append("Filter: Off");
      
      //
      // Set the filter date title information
      row = m_Sheet.createRow(rowNum++);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitle);
      cell.setCellValue(new XSSFRichTextString(title.toString()));
      
      //
      // Create warehouse title
      title.setLength(0);
      title.append("DC: ");
      title.append(m_WhsName);
      
      //
      // Set the warehouse title information
      row = m_Sheet.createRow(rowNum++);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitle);
      cell.setCellValue(new XSSFRichTextString(title.toString()));
                  
      //
      // Merge the title cells.  Gives a better look to the report.
      region = new CellRangeAddress(0, 0, 0, 2);
      m_Sheet.addMergedRegion(region);
      
      region = new CellRangeAddress(1, 1, 0, 2);
      m_Sheet.addMergedRegion(region);
      
      region = new CellRangeAddress(2, 2, 0, 2);
      m_Sheet.addMergedRegion(region);
     
           
      //
      // Create the row for the captions.
      row = m_Sheet.createRow(rowNum++);
           
      for ( int i = 0; i < MAX_COLS; i++ ) {
         cell = row.createCell(i);
         cell.setCellStyle(styleCaption);
      }
      
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Vendor ID"));      
      m_Sheet.setColumnWidth(colNum++, 3000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Vendor Name"));
      m_Sheet.setColumnWidth(colNum++, 12000);
      row.getCell(colNum++).setCellValue(new XSSFRichTextString("Dept #"));
      row.getCell(colNum).setCellValue(new XSSFRichTextString("Total Inventory"));      
      m_Sheet.setColumnWidth(colNum++, 4000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("SLOB-52"));
      m_Sheet.setColumnWidth(colNum++, 4000);
      row.getCell(colNum).setCellValue(new XSSFRichTextString("SLOB-104"));
      m_Sheet.setColumnWidth(colNum, 4000);
      
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
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(1000);
      StringBuffer itemSalesJoin = new StringBuffer(250);
      boolean isPrepared = false;
      String accpacLoc = " and il.location = '%s' \r\n";
      
      if ( m_EdbConn != null ) {
         try {            
            itemSalesJoin.append("   left outer join itemsales on item_entity_attr.item_id = itemsales.item_nbr and \r\n");
            itemSalesJoin.append("      %sitemsales.invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') \r\n");
            
            sql.append("select  \r\n"); 
            sql.append("   vendor.vendor_id, vendor.name, emery_dept.dept_num, \r\n");
            sql.append("   round(nvl(sum(avgcost * qtyonhand)::numeric, 0::numeric), 2) as total_inventory, \r\n");            
            sql.append("   round(nvl(sum(decode(sign((qtyonhand - m12sales) * avgcost), -1, 0, (qtyonhand - m12sales) * avgcost))::numeric, 0::numeric), 2) as slob12, \r\n");
            sql.append("   round(nvl(sum(decode(sign((qtyonhand - m24sales) * avgcost), -1, 0, (qtyonhand - m24sales) * avgcost))::numeric, 0::numeric), 2) as slob24 \r\n");
            sql.append("from vendor \r\n");
            sql.append("join ( \r\n");
            sql.append("   select \r\n");
            sql.append("      item_entity_attr.vendor_id, item_entity_attr.item_id, il.qtyonhand, \r\n");
            sql.append("      nvl(sum(qty_shipped), 0) as m12sales, \r\n");
            sql.append("      nvl(sum(qty_shipped)*2, 0) as m24sales, \r\n");             
            sql.append("      round( \r\n");
            sql.append("         decode(il.qtyonhand, 0, \r\n");
            sql.append("         decode(il.lastcost, 0, ejd_item_price.buy, il.lastcost), il.totalcost / il.qtyonhand)::numeric,3 \r\n");
            sql.append("      ) as avgcost \r\n");
            sql.append("   from \r\n");
            sql.append("      item_entity_attr \r\n");
            sql.append("   join item_type on item_entity_attr.item_type_id = item_type.item_type_id and itemtype = 'STOCK' \r\n");
            sql.append("   join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            
            //
            // Add the department filter sql
            if( m_Dept != null && m_Dept.length() > 0 ) {
               sql.append("join emery_dept on ejd_item.dept_id = emery_dept.dept_id and dept_num = ? \r\n");
            }
                        
            sql.append("   join ejd.sage300_iciloc_mv il on il.itemno = item_entity_attr.item_id and il.qtyonhand > 0 \r\n");
            
            //
            // Build the location filter sql.  Need to keep this in order so we can add the correct bind vars.
            // Accpac being the POS that it is, requires padding of the location to help speed up the query.
            if ( m_WhsId > 0 ) {               
               sql.append(String.format(accpacLoc, m_AccWhsId));
               
               sql.append("join ejd_item_warehouse iw on iw.ejd_item_id = item_entity_attr.ejd_item_id \r\n");
               sql.append("join warehouse w on iw.warehouse_id = w.warehouse_id and w.fas_facility_id = ? \r\n");
               
               sql.append(String.format(itemSalesJoin.toString(), " itemsales.warehouse_id = ? and \r\n"));
               
            }
            else
               sql.append(String.format(itemSalesJoin.toString(), ""));
            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_price.warehouse_id = iw.warehouse_id ");
            
            //
            // Add the item setup date filtering sql
            if ( m_Filtered && (m_FilterDate != null && m_FilterDate.length() > 0) ) {
               sql.append("where \r\n");
               sql.append("   ejd_item.setup_date < to_date(?, 'mm/dd/yyyy') \r\n");
            }
            
            sql.append("   group by \r\n");
            sql.append("      item_entity_attr.vendor_id, item_entity_attr.item_id, il.qtyonhand, \r\n");
            sql.append("      round( \r\n");
            sql.append("         decode(il.qtyonhand, 0, \r\n");
            sql.append("         decode(il.lastcost, 0, ejd_item_price.buy, il.lastcost), il.totalcost / il.qtyonhand)::numeric,3 \r\n");
            sql.append("      ) \r\n");
            sql.append(") wosdat on wosdat.vendor_id = vendor.vendor_id \r\n");
            sql.append("join vendor_dept on vendor_dept.vendor_id = vendor.vendor_id \r\n");
            sql.append("join emery_dept on emery_dept.dept_id = vendor_dept.dept_id \r\n");
            
            sql.append("group by vendor.vendor_id, vendor.name, emery_dept.dept_num \r\n");
            sql.append("order by slob12 desc");
                                    
            m_SaleData = m_EdbConn.prepareStatement(sql.toString());
                                                         
            sql.setLength(0);
            sql.append("select name from warehouse where warehouse_id = ?");
            m_GetDCName = m_EdbConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("exception:", ex);
         }
         
         finally {
            sql = null;
            itemSalesJoin = null;
         }         
      }
      else
         log.error("slob summary.prepareStatements - null enterprisedb connection");
      
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
         
         if ( param.name.equals("whsid") )
            m_WhsId = Integer.parseInt(param.value);
         
         if ( param.name.equals("dc") )
            m_FasId = param.value;
         
         if ( param.name.equals("accdc") )
            m_AccWhsId = param.value;
         
         if ( param.name.equalsIgnoreCase("dept") )
            m_Dept = param.value;
      }
      
      fileName.append("slobvndlist");
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
      XSSFCellStyle styleTxtL = null;      // Text left justified
      XSSFCellStyle styleTxtC = null;      // Text center justified
      XSSFCellStyle styleInt = null;       // Style with 0 decimals
      XSSFCellStyle styleMoney = null;     // Money ($#,##0.00_);[Red]($#,##0.00)
      XSSFDataFormat format = m_Wrkbk.createDataFormat();
      
      styleTxtL = m_Wrkbk.createCellStyle();
      styleTxtL.setAlignment(HorizontalAlignment.LEFT);
      
      styleTxtC = m_Wrkbk.createCellStyle();
      styleTxtC.setAlignment(HorizontalAlignment.CENTER);
      
      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat(format.getFormat("0"));
           
      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short)8);
      
      m_CellStyles = new XSSFCellStyle[] {
         styleInt,     // col 0 vendor id
         styleTxtL,    // col 1 vendor name
         styleTxtC,    // col 2 dept num
         styleMoney,   // col 3 total inventory         
         styleMoney,   // col 4 wos 52
         styleMoney    // col 5 wos 144         
      };
      
      styleTxtL = null;
      styleTxtC = null;
      styleInt = null;
      styleMoney = null;
      format = null;
   }
}
