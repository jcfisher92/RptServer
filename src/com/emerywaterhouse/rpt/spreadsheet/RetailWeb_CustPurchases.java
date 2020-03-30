package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.GregorianCalendar;
import java.lang.Integer;


import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class RetailWeb_CustPurchases extends Report{
   
   //added 36 cols for up to three years of monthly totals
   private short BASE_COLS      = 12;
   private int colCnt = BASE_COLS;
   private final String TAB = "\t";
   private final String NEWLINE = "\r\n"; 
   
   private String m_StoreId;
   private String m_CustId;
   private String m_BegDate;
   private String m_EndDate;
   private String m_Email;
   private String m_ReportName;
   private String m_OutputFormat;
   private String m_ConsolidatedCustID;   
   private String m_MerchClass; 
   private String m_FlcId;
   private String m_NrhaId;
   private String m_VndId;
   private String m_RMSList;
   private PreparedStatement m_CustPurchases;
   private PreparedStatement m_HotBuyPurchases;
   //private Integer m_NumberOMonths;
   //private GregorianCalendar m_StartCal = new GregorianCalendar();
   //private GregorianCalendar m_EndCal = new GregorianCalendar();      
   private ArrayList<HSSFCellStyle> m_CellStyles = new ArrayList<HSSFCellStyle>();
   private ArrayList<String> m_MonthList = new ArrayList<String>();
   
   //
   // workbook entries.
   private HSSFWorkbook m_Wrkbk;
   private HSSFSheet m_Sheet;
   private StringBuffer m_Lines;
      
   //
   // Log4j logger
   private Logger m_Log;
       
   
   /**
    * default constructor
    */
   public RetailWeb_CustPurchases()
   {      
      super();
      m_Log = Logger.getLogger(RptServer.class);
      m_Wrkbk = new HSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
   }
   
   /**
    * Cleanup any allocated resources.
    * @throws Throwable 
    */
   public void finalize() throws Throwable
   {      
      if ( m_CellStyles != null ) {
         for ( int i = 0; i < (BASE_COLS); i++ )
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
      HSSFRow row = null;
      FileOutputStream outFile = null;
      ResultSet CustPurchases = null;
      short rowNum = 1;
      boolean result = false;
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         rowNum = createCaptions();
         
         if( m_ReportName != null && m_ReportName.equals("HotBuyReport") ){
          
            	m_HotBuyPurchases.setString(1,m_StoreId);
            	m_HotBuyPurchases.setString(2,m_CustId);
            	m_HotBuyPurchases.setString(3,m_StoreId);
            	m_HotBuyPurchases.setString(4,m_CustId);
            
        	  CustPurchases = m_HotBuyPurchases.executeQuery();
         }else
         {
                m_CustPurchases.setString(1,m_CustId);
                m_CustPurchases.setString(2,m_CustId);
       
        	CustPurchases = m_CustPurchases.executeQuery();
         }

         while ( CustPurchases.next() && m_Status == RptServer.RUNNING ) {
            row = createRow(rowNum);
            row.getCell(0).setCellValue(new HSSFRichTextString(CustPurchases.getString("ew")));
            row.getCell(1).setCellValue(new HSSFRichTextString(CustPurchases.getString("ord_id")));
            row.getCell(2).setCellValue(new HSSFRichTextString(CustPurchases.getString("oh_line_id")));
            row.getCell(3).setCellValue(new HSSFRichTextString(CustPurchases.getString("order_date")));
            row.getCell(4).setCellValue(new HSSFRichTextString(CustPurchases.getString("email")));
            row.getCell(5).setCellValue(new HSSFRichTextString(CustPurchases.getString("oozer")));
            row.getCell(6).setCellValue(new HSSFRichTextString(CustPurchases.getString("itum")));
            row.getCell(7).setCellValue(new HSSFRichTextString(CustPurchases.getString("upc")));
            row.getCell(8).setCellValue(new HSSFRichTextString(CustPurchases.getString("item_description")));
            row.getCell(9).setCellValue(new HSSFRichTextString(CustPurchases.getString("qty")));
            row.getCell(10).setCellValue(new HSSFRichTextString(CustPurchases.getString("price")));
            row.getCell(11).setCellValue(new HSSFRichTextString(CustPurchases.getString("ext_sell")));  
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
    * Executes the queries and builds the output file
    * 
    * @return true if the file was built, false if not.
    * @throws FileNotFoundException
    */
  
   private boolean buildOutputFileText() throws FileNotFoundException
   {      
      FileOutputStream outFile = null;
      ResultSet CustPurchases = null;
      boolean result = false;
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
    	  
         createCaptionsText();
         
         if( m_ReportName != null && m_ReportName.equals("HotBuyReport") ){
            	m_HotBuyPurchases.setString(1,m_StoreId);
            	m_HotBuyPurchases.setString(2,m_CustId);
            	m_HotBuyPurchases.setString(3,m_StoreId);
            	m_HotBuyPurchases.setString(4,m_CustId);
             
        	  CustPurchases = m_HotBuyPurchases.executeQuery();
         }else
         {
        	
                m_CustPurchases.setString(1,m_CustId);
                m_CustPurchases.setString(2,m_CustId);
            
        	CustPurchases = m_CustPurchases.executeQuery();
         }

         while ( CustPurchases.next() && m_Status == RptServer.RUNNING ) {
        	 m_Lines.append(CustPurchases.getString("ew") + TAB); 
        	 m_Lines.append(CustPurchases.getString("ord_id") + TAB);
        	 m_Lines.append(CustPurchases.getString("oh_line_id") + TAB);
        	 m_Lines.append(CustPurchases.getString("order_date") + TAB);
        	 m_Lines.append(CustPurchases.getString("email") + TAB);
        	 m_Lines.append(CustPurchases.getString("oozer") + TAB);
        	 m_Lines.append(CustPurchases.getString("itum") + TAB);
        	 m_Lines.append(CustPurchases.getString("upc") + TAB);
        	 m_Lines.append(CustPurchases.getString("item_description") + TAB);
        	 m_Lines.append(CustPurchases.getString("qty") + TAB);
        	 m_Lines.append(CustPurchases.getString("price") + TAB);
        	 m_Lines.append(CustPurchases.getString("ext_sell") + TAB);
        	 m_Lines.append( NEWLINE );
         }
         
         outFile.write(m_Lines.toString().getBytes());
         m_Lines.delete(0, m_Lines.length());
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         m_Log.error("exception", ex);
      }

      finally {         
         closeRSet(CustPurchases);
         try {
            outFile.close();
         }

         catch( Exception e ) {
            m_Log.error("exception:", e);
         }

         outFile = null;
         m_Lines = null;
      }
      
      return result;
   }

   
   /**
    * Builds the sql based on the type of filter requested by the user.
    * @return A complete sql statement.
    */

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_CustPurchases);
      closeStmt(m_HotBuyPurchases);
   }
   
   /**
    * Creates the report title and the captions.
    */
   private short createCaptions()
   {
      HSSFFont fontTitle;
      HSSFCellStyle styleTitle;   // Bold, centered      
      HSSFCellStyle styleTitleLeft;   // Bold, Left Justified      
      HSSFRow row = null;
      HSSFCell cell = null;
      short rowNum = 0;
      //GregorianCalendar tempcal = new GregorianCalendar();
      //StringBuffer caption = new StringBuffer("Retail Web Order History ");
      StringBuffer caption = new StringBuffer(100);
      //StringBuffer captionEndUser = new StringBuffer("End User Email: ");
       
      if ( m_Sheet == null )
         return 0;
      
      caption.setLength(0);
      
      fontTitle = m_Wrkbk.createFont();
      fontTitle.setFontHeightInPoints((short)10);
      fontTitle.setFontName("Arial");
      fontTitle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
      
      styleTitle = m_Wrkbk.createCellStyle();
      styleTitle.setFont(fontTitle);
      styleTitle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
      
      styleTitleLeft = m_Wrkbk.createCellStyle();
      styleTitleLeft.setFont(fontTitle);
      styleTitleLeft.setAlignment(HSSFCellStyle.ALIGN_LEFT);     
      
      //
      // set the report title
      row = m_Sheet.createRow(rowNum);
      cell = row.createCell(0); 
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(styleTitleLeft);
      
      if( m_ReportName != null && m_ReportName.equals("HotBuyReport") ){
    	 caption.append("Hot Buy Purchase History Report");
      }
      else if( m_ReportName != null && m_ReportName.equals("EndUserReport") ){
    	 caption.append("Purchase History Report By End User");
    	 caption.append(" Email: "+m_Email);
      }
      else {
    	 caption.append("Purchase History Report");
      }
      
      cell.setCellValue(new HSSFRichTextString(caption.toString()));
      
      caption.setLength(0);
      rowNum ++;
      row = m_Sheet.createRow(rowNum);
      cell = row.createCell(0);
      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellStyle(styleTitleLeft);
      
      if (m_CustId != null && m_CustId.length() > 0){
         caption.append("Customer: ");
         caption.append(m_CustId);
      }
      
      
      if (m_BegDate != null && m_EndDate != null)
         caption.append(" ");
         caption.append(m_BegDate);
         caption.append(" - ");
         caption.append(m_EndDate);

      
      cell.setCellValue(new HSSFRichTextString(caption.toString()));

      rowNum ++;
      rowNum ++;
      row = m_Sheet.createRow(rowNum);

      try {
         if ( row != null ) {
            for ( int i = 0; i < (BASE_COLS + 1); i++ ) {
               cell = row.createCell(i);
               if (i < BASE_COLS)
                // cell.setCellStyle(styleTitleLeft);
               //else
                 cell.setCellStyle(styleTitle);
                              
            }
            
            row.getCell(0).setCellValue(new HSSFRichTextString("EW"));
            m_Sheet.setColumnWidth(0, 1000);            
            row.getCell(1).setCellValue(new HSSFRichTextString("Ord_ID"));
            m_Sheet.setColumnWidth(1, 2000);
            row.getCell(2).setCellValue(new HSSFRichTextString("Line_ID"));
            row.getCell(3).setCellValue(new HSSFRichTextString("Order_Date"));
            m_Sheet.setColumnWidth(3, 3000);
            row.getCell(4).setCellValue(new HSSFRichTextString("Email"));
            row.getCell(5).setCellValue(new HSSFRichTextString("User Name"));
            m_Sheet.setColumnWidth(5, 5000);
            row.getCell(6).setCellValue(new HSSFRichTextString("Item_id"));
            //m_Sheet.setColumnWidth(6, 2000);
            row.getCell(7).setCellValue(new HSSFRichTextString("UPC"));
            m_Sheet.setColumnWidth(7, 3500);
            row.getCell(8).setCellValue(new HSSFRichTextString("Description"));
            m_Sheet.setColumnWidth(8,15000);
            row.getCell(9).setCellValue(new HSSFRichTextString("Qty Ord"));
            m_Sheet.setColumnWidth(9, 2000);
            row.getCell(10).setCellValue(new HSSFRichTextString("Price"));
            m_Sheet.setColumnWidth(10, 2000);
            row.getCell(11).setCellValue(new HSSFRichTextString("Ext_sell"));
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
    * Creates the report title and the captions.
    */
   private void createCaptionsText()
   {
	   m_Lines = new StringBuffer();

       m_Lines.append("Packet\tTitle\tVendor Id\tVendor Name\tMessage\tItem Description\tItem Id\tSKU\tUPC\tStock Pack\tNBC\tShip Unit\t");
       m_Lines.append("Cost\tPromo Cost\tRetail\tRetail C\t");
       m_Lines.append("Terms\tDeadline\tUnits Ordered\t");
       
       m_Lines.append("EW\tOrd_ID\tLine_ID\tOrder_Date\tEmail\tUser Name\t");
       m_Lines.append("Item_id\tUPc\tDescription\tQty Ord\tPrice\tExt_sell\t");

       m_Lines.append( NEWLINE );
   }
   
   
   /**
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    * 
    * @return The fromatted row of the spreadsheet.
    */
   private HSSFRow createRow(short rowNum)
   {
      HSSFRow row = null;
      HSSFCell cell = null;
      
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
   @Override

   public boolean createReport()
   {
      boolean created = false; 
      m_Status = RptServer.RUNNING;
      //setCalendar();
      //setMonthsToReport();
      setupWorkbook();
      

      try {         
         m_PgConn = m_RptProc.getPgConn();
         prepareStatements();    
         
         if( m_OutputFormat != null && m_OutputFormat.equals("excel") )
            created = buildOutputFile();
         else
        	created = buildOutputFileText();
       
      }

      catch ( Exception ex ) {
         log.fatal("exception:", ex);
      }
      
      finally {
         closeStatements();
        // DbUtils.closeDbConn(m_PgConn, null, null);
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
      
     if ( m_PgConn != null ) {
        sql.setLength(0);
        //there may be Emery Items and non-Emery itmes in this order
        //up to the "union" this query will get the Emery items, after it the non-Emery items
        sql.append("select 'Y' EW,oh.oh_id ord_id, ol.oh_line_id, oh.created_on order_date,   ");
        sql.append("oh.email_address email, wu.first_name||' '||wu.last_name oozer, ol.emery_item_id itum, ");
        sql.append("item.upc,  ");
        sql.append("ol.item_description, ol.qty, ol.price, (ol.qty * ol.price) ext_sell  ");
        sql.append("from order_history oh ");
        sql.append("join web_user wu on wu.user_id = oh.user_id ");
        sql.append("join order_history_line ol on ol.oh_id = oh.oh_id and ol.emery_item_id is not null ");
        sql.append("join item on item.item_id = ol.emery_item_id  ");
        sql.append("where oh.customer_id = ? ");

        if (m_BegDate != null && m_BegDate.length() > 0){
           sql.append("and oh.created_on >= to_date('");
           sql.append(m_BegDate);
           sql.append("','mm/dd/yyyy') ");
        }   
        if (m_EndDate != null && m_BegDate.length() > 0){
           sql.append("and oh.created_on <= to_date('");
           sql.append(m_EndDate);
           sql.append("','mm/dd/yyyy') ");
        } 
        
        //
        //This is about end user purchases
        if ( m_Email != null && !m_Email.equals("") ){ 
            sql.append("and oh.email_address = '");
            sql.append(m_Email);
            sql.append("'");
        } 
        
        sql.append("union ");  // this will get any non-Emery items
        sql.append("select 'N' EW,oh.oh_id ord_id, ol.oh_line_id, oh.created_on order_date,   ");
        sql.append("oh.email_address email, wu.first_name||' '||wu.last_name oozer, ol.emery_item_id itum, ");
        sql.append("s_item.upc,  ");
        sql.append("ol.item_description, ol.qty, ol.price, (ol.qty * ol.price) ext_sell  ");
        sql.append("from order_history oh ");
        sql.append("join web_user wu on wu.user_id = oh.user_id ");
        sql.append("join order_history_line ol on ol.oh_id = oh.oh_id and ol.store_item_id is not null ");
        sql.append("join s_item on s_item.item_id = ol.store_item_id  ");
        sql.append("where oh.customer_id = ? ");

        if (m_BegDate != null && m_BegDate.length() > 0){
           sql.append("and oh.created_on >= to_date('");
           sql.append(m_BegDate);
           sql.append("','mm/dd/yyyy') ");
        }   
        if (m_EndDate != null && m_BegDate.length() > 0){
           sql.append("and oh.created_on <= to_date('");
           sql.append(m_EndDate);
           sql.append("','mm/dd/yyyy') ");
        }
        
        //
        //This is about end user purchases
        if ( m_Email != null && !m_Email.equals("") ){
            sql.append("and oh.email_address = '");
            sql.append(m_Email);
            sql.append("'");
        } 
        
        sql.append("order by oozer, order_date, ord_id, itum ");
                        
        m_CustPurchases = m_PgConn.prepareStatement(sql.toString());
       
        
        //
        //Hot buy purchases
        sql.setLength(0);
        //there may be Emery Items and non-Emery itmes in this order
        //up to the "union" this query will get the Emery items, after it the non-Emery items
        sql.append("select 'Y' EW,oh.oh_id ord_id, ol.oh_line_id, oh.created_on order_date,   ");
        sql.append("oh.email_address email, wu.first_name||' '||wu.last_name oozer, ol.emery_item_id itum, ");
        sql.append("item.upc,  ");
        sql.append("ol.item_description, ol.qty, ol.price, (ol.qty * ol.price) ext_sell  ");
        sql.append("from order_history oh ");
        sql.append("join web_user wu on wu.user_id = oh.user_id ");
        sql.append("join order_history_line ol on ol.oh_id = oh.oh_id and ol.emery_item_id is not null ");
        sql.append("join item on item.item_id = ol.emery_item_id  ");
        sql.append("join store_item_hotbuy sih on sih.item_id = item.item_id and sih.store_id = int4(?) ");
        sql.append("where oh.customer_id = ? ");
 
        if (m_BegDate != null && m_BegDate.length() > 0){
           sql.append("and oh.created_on >= to_date('");
           sql.append(m_BegDate);
           sql.append("','mm/dd/yyyy') ");
        }   
        if (m_EndDate != null && m_BegDate.length() > 0){
           sql.append("and oh.created_on <= to_date('");
           sql.append(m_EndDate);
           sql.append("','mm/dd/yyyy') ");
        } 
        
        //
        //This is about end user purchases
        if ( m_Email != null && !m_Email.equals("") ){ 
            sql.append("and oh.email_address = '");
            sql.append(m_Email);
            sql.append("'");
        } 
        
        sql.append("union ");  // this will get any non-Emery items
        sql.append("select 'N' EW,oh.oh_id ord_id, ol.oh_line_id, oh.created_on order_date,   ");
        sql.append("oh.email_address email, wu.first_name||' '||wu.last_name oozer, ol.emery_item_id itum, ");
        sql.append("s_item.upc,  ");
        sql.append("ol.item_description, ol.qty, ol.price, (ol.qty * ol.price) ext_sell  ");
        sql.append("from order_history oh ");
        sql.append("join web_user wu on wu.user_id = oh.user_id ");
        sql.append("join order_history_line ol on ol.oh_id = oh.oh_id and ol.store_item_id is not null ");
        sql.append("join s_item on s_item.item_id = ol.store_item_id and s_item.store_id = int4(?) and s_item.hotbuy = true");
        sql.append(" where oh.customer_id = ? ");

        if (m_BegDate != null && m_BegDate.length() > 0){
           sql.append("and oh.created_on >= to_date('");
           sql.append(m_BegDate);
           sql.append("','mm/dd/yyyy') ");
        }   
        if (m_EndDate != null && m_BegDate.length() > 0){
           sql.append("and oh.created_on <= to_date('");
           sql.append(m_EndDate);
           sql.append("','mm/dd/yyyy') ");
        }
        
        sql.append("order by oozer, order_date, ord_id, itum ");
                        
        m_HotBuyPurchases = m_PgConn.prepareStatement(sql.toString());
        
        
        //we need the customer name in the caption, by god.
       // sql.setLength(0);
       // sql.append("select name from customer where customer_id = ?");
       // m_GetCustNames = m_PgConn.prepareStatement(sql.toString());
           
        
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
         
         if ( param.name.equals("cust_id") )
            m_CustId = param.value;
                
         if ( param.name.equals("store_id") )
             m_StoreId = param.value;
         
         if ( param.name.equals("begdate") )
            m_BegDate = param.value;
          
         if ( param.name.equals("enddate") )
            m_EndDate = param.value;
         
         if ( param.name.equals("email") )
             m_Email = param.value;
         
         if ( param.name.equals("reportName") )
             m_ReportName = param.value;
         
         if ( param.name.equals("outputFormat") )
             m_OutputFormat = param.value;
         
      }
      
      if( m_OutputFormat != null && m_OutputFormat.equals("excel") ){
         //
         // Build the file name.
         fname.append(tm);
         fname.append("-");
         fname.append(m_RptProc.getUid());
         fname.append("cp.xls");
         m_FileNames.add(fname.toString());
      }
      else{
         //
         // Build the file name.
         fname.append(tm);
         fname.append("-");
         fname.append(m_RptProc.getUid());
         fname.append("cp.prn");
         m_FileNames.add(fname.toString());
      }
      
      
   }
      
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other inititialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {      
      HSSFCellStyle styleText;      // Text right justified
      HSSFCellStyle styleInt;       // Style with 0 decimals
      HSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00) 
      HSSFCellStyle stylePct;       // Style with 0 decimals + %
      
      styleText = m_Wrkbk.createCellStyle();
      //styleText.setFont(m_FontData);
      styleText.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      
      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      styleInt.setDataFormat((short)3);

      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      styleMoney.setDataFormat((short)8);
      
      stylePct = m_Wrkbk.createCellStyle();
      stylePct.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      stylePct.setDataFormat((short)9);
      
      
      m_CellStyles.add(styleText);    // col 0 EW Item -- Yes or No?
      m_CellStyles.add(styleText);    // col 1 Order History Id
      m_CellStyles.add(styleText);    // col 2 Order History Line ID
      m_CellStyles.add(styleText);    // col 3 Order Date
      m_CellStyles.add(styleText);    // col 4 Email
      m_CellStyles.add(styleText);    // col 5 Web User Name
      m_CellStyles.add(styleText);    // col 6 Item Id
      m_CellStyles.add(styleText);    // col 7 upc
      m_CellStyles.add(styleText);    // col 8 item Description
      m_CellStyles.add(styleInt);    // col 9  qty ord
      m_CellStyles.add(styleMoney);    // col 10 retail_price
      m_CellStyles.add(styleMoney);    // col 11 retail price * qty ord
      styleText = null;
      styleInt = null;
      styleMoney = null;
      stylePct = null;
   }   

}
