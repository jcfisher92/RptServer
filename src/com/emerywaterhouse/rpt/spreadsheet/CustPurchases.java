package com.emerywaterhouse.rpt.spreadsheet;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.GregorianCalendar;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class CustPurchases extends Report
{
   //added 36 cols for up to three years of monthly totals
   private short BASE_COLS      = 22;
   private int colCnt = BASE_COLS;

   private String m_CustId;
   private String m_BegDate;
   private String m_EndDate;
   private String m_MerchClass; 
   private String m_FlcId;
   private String m_NrhaId;
   private String m_VndId;
   private String m_RMSList;
   private PreparedStatement m_CustPurchases;
   private PreparedStatement m_GetCustNames;   
   private PreparedStatement m_GetConsolidatedCustID;
   private Integer m_NumberOMonths;
   private GregorianCalendar m_StartCal = new GregorianCalendar();
   private GregorianCalendar m_EndCal = new GregorianCalendar();      
   private ArrayList<XSSFCellStyle> m_CellStyles = new ArrayList<XSSFCellStyle>();
   private ArrayList<String> m_MonthList = new ArrayList<String>();

   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;

   //
   // Log4j logger
   private Logger m_Log;


   /**
    * default constructor
    */
   public CustPurchases()
   {      
      super();
      m_Log = Logger.getLogger(RptServer.class);
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable 
    */
   public void finalize() throws Throwable
   {      
      if ( m_CellStyles != null ) {
         for ( int i = 0; i < (BASE_COLS + m_NumberOMonths); i++ )
            m_CellStyles.remove(i);
      }
      m_Sheet = null;
      m_Wrkbk = null;      
      m_CellStyles = null;

      super.finalize();
   }

   /**
    * Executes the queries and builds the output file
    * 
    * @return true if the file was built, false if not.
    * @throws FileNotFoundException
    */

   private boolean buildOutputFile() throws FileNotFoundException
   {      
      XSSFRow row = null;
      FileOutputStream outFile = null;
      ResultSet CustPurchases = null;
      ResultSet ConsolidatedID = null;
      short rowNum = 1;
      boolean result = false;
      int colNum = 0;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         rowNum = createCaptions();
         m_GetConsolidatedCustID.setString(1, m_CustId);  
         ConsolidatedID = m_GetConsolidatedCustID.executeQuery();  //this is to get parent account for customer SKU's
         
         if ( ConsolidatedID.next() && m_Status == RptServer.RUNNING ) {
            if ( ConsolidatedID.getString("cust_cons_id") != null ) 
               m_CustId = ConsolidatedID.getString("cust_cons_id");
         }
         
         m_CustPurchases.setString(1, m_BegDate);
         m_CustPurchases.setString(2, m_EndDate);
         m_CustPurchases.setString(3, m_CustId); 

         CustPurchases = m_CustPurchases.executeQuery();

         while ( CustPurchases.next() && m_Status == RptServer.RUNNING ) {
            row = createRow(rowNum);
            row.getCell(0).setCellValue(new XSSFRichTextString(CustPurchases.getString("item_nbr")));
            row.getCell(1).setCellValue(new XSSFRichTextString(CustPurchases.getString("customer_sku")));
            row.getCell(2).setCellValue(new XSSFRichTextString(CustPurchases.getString("rms_id")));
            row.getCell(3).setCellValue(new XSSFRichTextString(CustPurchases.getString("vendor")));
            row.getCell(4).setCellValue(new XSSFRichTextString(CustPurchases.getString("nrha_id")));
            row.getCell(5).setCellValue(new XSSFRichTextString(CustPurchases.getString("mdc_id")));
            row.getCell(6).setCellValue(new XSSFRichTextString(CustPurchases.getString("flc_id")));
            row.getCell(7).setCellValue(new XSSFRichTextString(CustPurchases.getString("upc_code")));           
            row.getCell(8).setCellValue(new XSSFRichTextString(CustPurchases.getString("vendor_item_num")));
            row.getCell(9).setCellValue(new XSSFRichTextString(CustPurchases.getString("description")));
            row.getCell(10).setCellValue(new XSSFRichTextString(CustPurchases.getString("sell_price")));
            row.getCell(11).setCellValue(new XSSFRichTextString(CustPurchases.getString("retail_price")));
            row.getCell(12).setCellValue(CustPurchases.getDouble("rtla"));
            row.getCell(13).setCellValue(CustPurchases.getDouble("rtlb"));
            row.getCell(14).setCellValue(CustPurchases.getDouble("rtlc"));
            row.getCell(15).setCellValue(CustPurchases.getDouble("rtld"));
            row.getCell(16).setCellValue(new XSSFRichTextString(CustPurchases.getString("unit")));
            row.getCell(17).setCellValue(CustPurchases.getDouble("stock_pack"));
            row.getCell(18).setCellValue(CustPurchases.getDouble("sen_code"));
            row.getCell(19).setCellValue(CustPurchases.getDouble("grantotale"));
            row.getCell(20).setCellValue(CustPurchases.getDouble("nopromo"));
            row.getCell(21).setCellValue(CustPurchases.getDouble("grantotale") - CustPurchases.getDouble("nopromo"));

            colNum = 22;
            for ( int i = 0; i < (m_MonthList.size()); i++ ){  // for each month, total qty in a column
               row.getCell(colNum).setCellValue(
                     new XSSFRichTextString(CustPurchases.getString("MY"+m_MonthList.get(i)))
                     );
               colNum ++;
            }
            rowNum++;            
         }
         m_Wrkbk.write(outFile);
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         m_Log.error("exception", ex);
      }

      finally {         
         row = null;         
         closeRSet(CustPurchases);
         closeRSet(ConsolidatedID);                 
         try {
            outFile.close();
         }

         catch( Exception e ) {
            m_Log.error("exception:", e);
         }

         outFile = null;
      }

      return result;
   }

   /**
    * Builds the sql based on the type of filter requested by the user.
    * @return A complete sql statement.
    */
   private String buildSql()
   {
      StringBuffer sql = new StringBuffer(1024);
      String SplitRMSList[];    
   
      sql.append("select \r\n");
      sql.append("   invd.item_nbr, \r\n");
      sql.append("   item_ea_cross.customer_sku, \r\n");
      sql.append("   vendor.name as vendor,  \r\n");
      sql.append("   vendor.vendor_id,  \r\n");
      sql.append("   nrha.nrha_id, mdc.mdc_id, flc.flc_id, \r\n");
      sql.append("   ejd_item_whs_upc.upc_code, \r\n");
      sql.append("   vendor_item_ea_cross.vendor_item_num, \r\n");
      sql.append("   item_entity_attr.description, \r\n");
      sql.append("   max(decode(cust_nbr,cust_nbr,totie ,null)) as grantotale, \r\n");
      sql.append("   max(decode(promo_nbr, null, thispromo,0)) as nopromo, \r\n");

      if ( m_RMSList.length() == 0 )
         sql.append("   max(decode(item_nbr, item_nbr, rms, null)) as rms_id, \r\n");
      else
         sql.append("   rms_item.rms_id, \r\n");

      for ( int i = 0; i < m_MonthList.size(); i++ ){
         sql.append("   max(decode(shipped,");
         sql.append(m_MonthList.get(i));
         sql.append(",thismonth,null)) MY");
         sql.append(m_MonthList.get(i));
         sql.append(", \r\n");
      }       

      sql.append("   (select price from ejd_cust_procs.get_sell_price(invd.cust_nbr, invd.item_ea_id)) as sell_price, \r\n");
      sql.append("   ejd_price_procs.get_retail_price(invd.cust_nbr, invd.item_ea_id) as retail_price, \r\n");
      sql.append("   ejd_item_price.sen_code_id as sen_code, \r\n");
      sql.append("   ejd_item_price.buy, \r\n");
      sql.append("   ejd_item_price.retail_a as rtla, \r\n");
      sql.append("   ejd_item_price.retail_b as rtlb, \r\n");
      sql.append("   ejd_item_price.retail_c as rtlc, \r\n");
      sql.append("   ejd_item_price.retail_d as rtld, \r\n");
      sql.append("   ship_unit.unit, ejd_item_warehouse.stock_pack \r\n");
      sql.append("from \r\n");

      sql.append("( \r\n");
      sql.append("   select \r\n");
      sql.append("      distinct item_nbr, inv_dtl.item_ea_id, cust_nbr, inv_dtl.warehouse, promo_nbr, to_char(invoice_date,'MMYYYY') as shipped, \r\n");
      sql.append("      sum(inv_dtl.qty_shipped) over(partition by item_nbr, inv_dtl.item_ea_id, rms_id, to_char(invoice_date,'MMYYYY')) as thismonth, \r\n");
      sql.append("      sum(inv_dtl.qty_shipped) over(partition by item_nbr, inv_dtl.item_ea_id, rms_id, promo_nbr) as thispromo, \r\n");
      
      if ( m_RMSList.length() == 0 )
         sql.append("      max(rms_id) over (partition by item_nbr, inv_dtl.item_ea_id) as rms_id, \r\n");
      else
         sql.append("      rms_item.rms_id, \r\n");
      
      sql.append("      max(rms_id) over (partition by item_nbr, inv_dtl.item_ea_id) as rms, \r\n");
      sql.append("      sum(inv_dtl.qty_shipped) over(partition by item_nbr, inv_dtl.item_ea_id, rms_id ) as totie \r\n");
      sql.append("   from inv_dtl \r\n");
      sql.append("   left outer join rms_item on inv_dtl.item_ea_id = rms_item.item_ea_id \r\n");
      sql.append("   where exists( \r\n");
      sql.append("      select inv_hdr_id  \r\n");
      sql.append("      from inv_hdr \r\n");
      sql.append("      where \r\n");
      sql.append("         invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') and \r\n");
      sql.append("         cust_nbr = ? and \r\n");
      sql.append("         inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id \r\n");
      sql.append("   ) \r\n");
      sql.append(") invd \r\n");
      sql.append("join warehouse on warehouse.name = invd.warehouse \r\n");
      sql.append("join item_entity_attr on item_entity_attr.item_ea_id = invd.item_ea_id \r\n");
      sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id \r\n");
      sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id and ejd_item_warehouse.warehouse_id = warehouse.warehouse_id\r\n");
      sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_price.warehouse_id = warehouse.warehouse_id\r\n");            
      sql.append("join customer on customer.customer_id = invd.cust_nbr \r\n");
      
      //
      // Optional filters appended here.
      sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id \r\n");      
      if ( m_VndId != null && m_VndId.length() > 0 )
         sql.append(String.format("      and vendor.vendor_id = %s \r\n", m_VndId));
      
      sql.append("join flc on ejd_item.flc_id = flc.flc_id \r\n");      
      if ( m_FlcId != null && m_FlcId.length() > 0 )
         sql.append(String.format("      and ejd_item.flc_id = %s \r\n", m_FlcId));
      
      sql.append("join mdc on mdc.mdc_id = flc.mdc_id \r\n");      
      if ( m_MerchClass != null && m_MerchClass.length() > 0 )
         sql.append(String.format("      and mdc.mdc_id = %s \r\n", m_MerchClass));
      
      sql.append("join nrha on nrha.nrha_id = mdc.nrha_id \r\n");
      if ( m_NrhaId != null && m_NrhaId.length() > 0 )
         sql.append(String.format("      and nrha.nrha_id = %s \r\n", m_NrhaId));
      
      sql.append("join ship_unit on item_entity_attr.ship_unit_id = ship_unit.unit_id \r\n");
      sql.append("left outer join item_ea_cross on item_ea_cross.item_ea_id = item_entity_attr.item_ea_id and item_ea_cross.customer_id = invd.cust_nbr \r\n");
      sql.append("left outer join vendor_item_ea_cross on item_entity_attr.item_ea_id = vendor_item_ea_cross.item_ea_id and  \r\n");
      sql.append("      item_entity_attr.vendor_id = vendor_item_ea_cross.vendor_id\r\n");
      sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = ejd_item.ejd_item_id and  \r\n");
      sql.append("      ejd_item_whs_upc.warehouse_id = warehouse.warehouse_id and primary_upc = 1 \r\n");
      
      if ( m_RMSList.length() > 0 )
         sql.append("join rms_item on rms_item.item_ea_id = item_entity_attr.item_ea_id \r\n");
      
      sql.append("where invd.item_ea_id = item_entity_attr.item_ea_id \r\n");
      
      if ( m_RMSList.length() > 0 ) {         
         SplitRMSList = m_RMSList.split(",");

         for ( int i = 0; i < SplitRMSList.length; i++ ) {  // for each RMSId
            if ( i == 0 ) 
               sql.append(" and (rms_item.rms_id = '");
            else
               sql.append(" or rms_item.rms_id = '"); 

            sql.append(SplitRMSList[i]);             
            sql.append("' ");
         }

         sql.append(") \r\n");
      }
      
      sql.append("\r\n");
      sql.append("group by \r\n");
      sql.append("   item_nbr, vendor.vendor_id, vendor.name,\r\n");
      sql.append("   nrha.nrha_id, mdc.mdc_id, flc.flc_id, \r\n");
      sql.append("   sen_code_id, buy, sell, item_entity_attr.description, \r\n");
      sql.append("   customer_sku, vendor_item_ea_cross.vendor_item_num, \r\n");
      sql.append("   upc_code, ship_unit.unit, stock_pack, \r\n");     
      sql.append("   retail_a, retail_b, retail_c, retail_d, ");
      
      if ( m_RMSList.length() > 0 )
         sql.append("rms_item.rms_id, ");
      
      sql.append("invd.cust_nbr, invd.item_ea_id \r\n");
      
      sql.append("order by \r\n");
      sql.append("   nrha.nrha_id, mdc.mdc_id, flc.flc_id, vendor.name, item_nbr");
            
      //log.info(sql.toString());
      return sql.toString();
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_CustPurchases);
      closeStmt(m_GetConsolidatedCustID);
      closeStmt(m_GetCustNames);
   }

   private String GetCustomerName(String custo) 
   {
      ResultSet CustName = null;
      String ThisCust = null;

      try{
         m_GetCustNames.setString(1,custo);  
         CustName = m_GetCustNames.executeQuery();  
         while (CustName.next() && m_Status == RptServer.RUNNING ) {
            if (CustName.getString("name")!= null) 
               ThisCust = CustName.getString("name");
            else
               ThisCust = "";
         }
      }
      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         m_Log.error("exception", ex);
      }

      finally {                  
         closeRSet(CustName);
      }                 

      return (ThisCust);
   }

   /**
    * Creates the report title and the captions.
    */
   private short createCaptions()
   {
      XSSFFont fontTitle;
      XSSFCellStyle styleTitle;   // Bold, centered      
      XSSFCellStyle styleTitleLeft;   // Bold, Left Justified      
      XSSFRow row = null;
      XSSFCell cell = null;
      short rowNum = 0;
      int real_month;
      GregorianCalendar tempcal = new GregorianCalendar();
      StringBuffer caption = new StringBuffer("Customer Purchases: ");

      if ( m_Sheet == null )
         return 0;

      fontTitle = m_Wrkbk.createFont();
      fontTitle.setFontHeightInPoints((short)10);
      fontTitle.setFontName("Arial");
      fontTitle.setBold(true);

      styleTitle = m_Wrkbk.createCellStyle();
      styleTitle.setFont(fontTitle);
      styleTitle.setAlignment(HorizontalAlignment.CENTER);

      styleTitleLeft = m_Wrkbk.createCellStyle();
      styleTitleLeft.setFont(fontTitle);
      styleTitleLeft.setAlignment(HorizontalAlignment.LEFT);


      //
      // set the report title
      row = m_Sheet.createRow(rowNum);
      cell = row.createCell(0); 
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitleLeft);
      caption.append(" ");
      caption.append(m_BegDate);
      caption.append(" - ");
      caption.append(m_EndDate);

      cell.setCellValue(new XSSFRichTextString(caption.toString()));

      caption.setLength(0);
      rowNum ++;
      row = m_Sheet.createRow(rowNum);
      cell = row.createCell(0);
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(styleTitleLeft);
      caption.append(GetCustomerName(m_CustId));
      caption.append(", ");
      caption.append(m_CustId);


      if (m_VndId != null && m_VndId.length() > 0){
         caption.append(" Vendor: ");
         caption.append(m_VndId);
      }

      if (m_FlcId != null && m_FlcId.length() > 0){
         caption.append(" FLC: ");
         caption.append(m_FlcId);
      }

      if (m_MerchClass != null && m_MerchClass.length() > 0){
         caption.append(" MDC: ");
         caption.append(m_MerchClass);
      }

      if (m_NrhaId != null && m_NrhaId.length() > 0){
         caption.append(" NRHA: ");
         caption.append(m_NrhaId);
      }

      if (m_RMSList != null && m_RMSList.length() > 0){
         caption.append(" RMS: ");
         caption.append(m_RMSList);
      }


      cell.setCellValue(new XSSFRichTextString(caption.toString()));

      rowNum ++;
      rowNum ++;
      row = m_Sheet.createRow(rowNum);

      try {
         if ( row != null ) {
            for ( int i = 0; i < (BASE_COLS + m_NumberOMonths); i++ ) {
               cell = row.createCell(i);
               if (i < BASE_COLS)
                  cell.setCellStyle(styleTitleLeft);
               else
                  cell.setCellStyle(styleTitle);
            }

            row.getCell(0).setCellValue(new XSSFRichTextString("Item Number"));
            row.getCell(1).setCellValue(new XSSFRichTextString("Cust. Sku"));
            m_Sheet.setColumnWidth(1, 3000);
            row.getCell(2).setCellValue(new XSSFRichTextString("Rms#"));
            row.getCell(3).setCellValue(new XSSFRichTextString("Vendor Name"));
            m_Sheet.setColumnWidth(3, 10000);
            row.getCell(4).setCellValue(new XSSFRichTextString("Nrha Dept."));
            row.getCell(5).setCellValue(new XSSFRichTextString("Mdse. Class"));
            row.getCell(6).setCellValue(new XSSFRichTextString("Fine Line Class"));
            m_Sheet.setColumnWidth(6, 2000);
            row.getCell(7).setCellValue(new XSSFRichTextString("Upc-Primary"));
            m_Sheet.setColumnWidth(7, 4000);
            row.getCell(8).setCellValue(new XSSFRichTextString("Mfgr. Part No."));
            m_Sheet.setColumnWidth(8, 5000);
            row.getCell(9).setCellValue(new XSSFRichTextString("Item Description"));
            m_Sheet.setColumnWidth(9, 14000);
            row.getCell(10).setCellValue(new XSSFRichTextString("Cust. Cost"));
            m_Sheet.setColumnWidth(10, 3000);
            row.getCell(11).setCellValue(new XSSFRichTextString("Cust. Retail"));
            m_Sheet.setColumnWidth(11, 3000);
            row.getCell(12).setCellValue(new XSSFRichTextString("A Mkt. Retail"));
            m_Sheet.setColumnWidth(12, 3000);
            row.getCell(13).setCellValue(new XSSFRichTextString("B Mkt. Retail"));
            m_Sheet.setColumnWidth(13, 3000);
            row.getCell(14).setCellValue(new XSSFRichTextString("C Mkt. Retail"));
            m_Sheet.setColumnWidth(14, 3000);
            row.getCell(15).setCellValue(new XSSFRichTextString("D Mkt. Retail"));
            m_Sheet.setColumnWidth(15, 3000);
            row.getCell(16).setCellValue(new XSSFRichTextString("Ship Unit"));
            m_Sheet.setColumnWidth(16, 3000);
            row.getCell(17).setCellValue(new XSSFRichTextString("Shelf Pack"));
            m_Sheet.setColumnWidth(17, 3000);
            row.getCell(18).setCellValue(new XSSFRichTextString("Sensitivity Code"));
            m_Sheet.setColumnWidth(18, 3000);
            row.getCell(19).setCellValue(new XSSFRichTextString("Total Qty Purch"));
            m_Sheet.setColumnWidth(19, 3000);
            row.getCell(20).setCellValue(new XSSFRichTextString("Qty Purch - Reg"));
            m_Sheet.setColumnWidth(20, 3000);
            row.getCell(21).setCellValue(new XSSFRichTextString("Qty Purch - Promo"));
            m_Sheet.setColumnWidth(21, 3000);
            tempcal.set(m_StartCal.get(1), m_StartCal.get(2),1);

            for ( int i = (BASE_COLS); i < (BASE_COLS + m_NumberOMonths); i++ ) {
               caption.setLength(0);
               real_month = tempcal.get(GregorianCalendar.MONTH) + 1;
               caption.append(real_month);
               caption.append('/');
               caption.append(tempcal.get(GregorianCalendar.YEAR));
               row.getCell(i).setCellValue(new XSSFRichTextString(caption.toString()));
               tempcal.add(2,1);
            }
         }
      }

      finally {
         row = null;
         cell = null;
         fontTitle = null;
         styleTitle = null;
         caption = null;
      }

      return ++rowNum;
   }

   /**
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    * 
    * @return The fromatted row of the spreadsheet.
    */
   private XSSFRow createRow(short rowNum)
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
            cell.setCellStyle(m_CellStyles.get(i));
         }
      }

      return row;
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {      
      boolean created = false;
      m_Status = RptServer.RUNNING;
      setCalendar();
      setMonthsToReport();
      setupWorkbook();

      try { 
         m_EdbConn = m_RptProc.getEdbConn();         
         prepareStatements();      
         created = buildOutputFile();            
      }

      catch ( Exception ex ) {
         m_Log.fatal("exception:", ex);
      }

      finally {
         closeStatements(); 

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Prepares the sql queries for execution.
    *     
    */
   private void prepareStatements() throws Exception
   {            
      StringBuffer sql = new StringBuffer(256);

      if ( m_EdbConn != null ) {
         // our main big honking query gets put together in buildsql	  
         m_CustPurchases = m_EdbConn.prepareStatement(buildSql());

         // once per run, find the consolidated account ID for customer sku lookup
         sql.append("select cust_cons_id from cust_consolidate ");
         sql.append("where customer_id =  ?");
         sql.append(" and cons_type_id in ");
         sql.append("(select cons_type_id from consolidate_type ");
         sql.append("where description = 'ITEM XREF')");

         m_GetConsolidatedCustID = m_EdbConn.prepareStatement(sql.toString());

         //we need the customer name in the caption, by god.
         sql.setLength(0);
         sql.append("select name from customer where customer_id = ?");
         m_GetCustNames = m_EdbConn.prepareStatement(sql.toString());
      }
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    * 
    * Because it's possible that this report can be called from some other system, the
    * best way to deal with params is to not go by the order, but by the name.
    *
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fname = new StringBuffer();
      String tm = Long.toString(System.currentTimeMillis()).substring(3);
      int pcount = params.size();
      Param param = null;


      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);


         if ( param.name.equals("customer") )
            m_CustId = param.value;

         if ( param.name.equals("merchclass") )
            m_MerchClass = param.value;

         if ( param.name.equals("nrha") )
            m_NrhaId = param.value;

         if ( param.name.equals("flc") )
            m_FlcId = param.value;

         if ( param.name.equals("vendor") && param.value.trim().length() > 0 )
            m_VndId = (param.value);

         if ( param.name.equals("begdate") )
            m_BegDate = param.value;

         if ( param.name.equals("enddate") )
            m_EndDate = param.value;

         if ( param.name.equals("rmslist") )
            m_RMSList = param.value;         
      }

      //
      // Build the file name.
      fname.append(tm);
      fname.append("-");
      fname.append(m_RptProc.getUid());
      fname.append("cp.xlsx");
      m_FileNames.add(fname.toString());
   }

   private void setCalendar()     
   { 
      int year_begin = Integer.parseInt(m_BegDate.substring(6,10));
      int month_begin = Integer.parseInt(m_BegDate.substring(0,2));
      int year_end = Integer.parseInt(m_EndDate.substring(6,10));
      int month_end = Integer.parseInt(m_EndDate.substring(0,2));

      m_StartCal.set(year_begin,month_begin - 1,1);
      m_EndCal.set(year_end, month_end - 1,1);
   }


   private void setMonthsToReport()
   {

      StringBuffer Monthstring = new StringBuffer("");

      GregorianCalendar tempcal = new GregorianCalendar();
      tempcal.set(m_StartCal.get(1), m_StartCal.get(2),1);
      m_NumberOMonths = 1;
      Monthstring.setLength(0);
      Monthstring.append(tempcal.get(GregorianCalendar.MONTH) + 1);
      Monthstring.append(tempcal.get(GregorianCalendar.YEAR));
      m_MonthList.add(Monthstring.toString());
      while (tempcal.before(m_EndCal)) {
         m_NumberOMonths ++;
         tempcal.add(2,1);
         Monthstring.setLength(0);
         Monthstring.append(tempcal.get(GregorianCalendar.MONTH) + 1);
         Monthstring.append(tempcal.get(GregorianCalendar.YEAR));
         m_MonthList.add(Monthstring.toString());

      }
      colCnt = BASE_COLS + m_NumberOMonths;
   }


   /**
    * Sets up the styles for the cells based on the column data.  Does any other inititialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {      
      XSSFCellStyle styleText;      // Text right justified
      XSSFCellStyle styleInt;       // Style with 0 decimals
      XSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00) 
      XSSFCellStyle stylePct;       // Style with 0 decimals + %

      styleText = m_Wrkbk.createCellStyle();
      //styleText.setFont(m_FontData);
      styleText.setAlignment(HorizontalAlignment.LEFT);

      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);

      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short)8);

      stylePct = m_Wrkbk.createCellStyle();
      stylePct.setAlignment(HorizontalAlignment.RIGHT);
      stylePct.setDataFormat((short)9);


      m_CellStyles.add(styleText);    // col 0 item_nbr
      m_CellStyles.add(styleText);    // col 1 customer sku
      m_CellStyles.add(styleText);    // col 2 RMS
      m_CellStyles.add(styleText);    // col 3 vendor name
      m_CellStyles.add(styleText);    // col 4 nrha dept
      m_CellStyles.add(styleText);    // col 5 mdse class
      m_CellStyles.add(styleText);    // col 6 flc
      m_CellStyles.add(styleText);    // col 7 upc
      m_CellStyles.add(styleText);    // col 8 mfr part
      m_CellStyles.add(styleText);    // col 9 item description
      m_CellStyles.add(styleMoney);    // col 10 cust cost
      m_CellStyles.add(styleMoney);    // col 11 cust retail
      m_CellStyles.add(styleMoney);   // col 12 retail a
      m_CellStyles.add(styleMoney);   // col 13 retail b
      m_CellStyles.add(styleMoney);   // col 14 retail c
      m_CellStyles.add(styleMoney);   // col 15 retail d
      m_CellStyles.add(styleText);     // col 16 ship unit
      m_CellStyles.add(styleInt);     // col 17 stock pack
      m_CellStyles.add(styleInt);     // col 18 sensitivity
      m_CellStyles.add(styleInt);     // col 19 qty purchased
      m_CellStyles.add(styleInt);     // col 20 qty purchased: regular
      m_CellStyles.add(styleInt);     // col 21 qty purchased: promo
      for ( int i = 0; i < (m_MonthList.size()); i++ ){
         m_CellStyles.add(styleInt);     // col 22 and up:  qty shipped to each month
      }
      styleText = null;
      styleInt = null;
      styleMoney = null;
      stylePct = null;
   }

   /**
    * Main method for testing the Rep Shipment output.
    * Can supply a LogDate here if desired for testing the queries on a specific date.
    * @param args
    *
   public static void main(String args[]) {
      CustPurchases cp = new CustPurchases();

      Param p1 = new Param();
      p1.name = "customer";
      p1.value = "000001";
      Param p2 = new Param();
      p2.name = "begdate";
      p2.value = "01/01/2017";
      Param p3 = new Param();
      p3.name = "enddate";
      p3.value = "11/13/2017";
      Param p4 = new Param();
      p4.name = "rmslist";
      //p4.value = "";
      p4.value = "870-B";

      ArrayList<Param> params = new ArrayList<Param>();
      params.add(p1);
      params.add(p2);
      params.add(p3);
      params.add(p4);

      cp.m_FilePath = "C:\\Exp\\";

   	java.util.Properties connProps = new java.util.Properties();
   	connProps.put("user", "ejd");
   	connProps.put("password", "boxer");
   	try {
   		cp.m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33:5444/emery_jensen",connProps);

   		cp.setParams(params);
   		cp.createReport();
   	} 
   	catch (Exception e) {
   		e.printStackTrace();
   	}
   }*/
}

