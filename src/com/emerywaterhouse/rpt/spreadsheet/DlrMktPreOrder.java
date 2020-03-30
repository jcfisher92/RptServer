/**
 * File: DlrMktPreOrder.java
 * Description: Dealer Market pre show orders
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 12/07/2006
 * Last Update: $Id: DlrMktPreOrder.java,v 1.7 2009/02/18 15:15:57 jfisher Exp $
 * 
 * History: 
 *    $Log: DlrMktPreOrder.java,v $
 *    Revision 1.7  2009/02/18 15:15:57  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.6  2008/10/30 16:41:49  jfisher
 *    Fixed potential null warnings and missing java doc tags
 *
 *    Revision 1.5  2008/10/29 21:11:08  jfisher
 *    Fixed some warnings
 *
 *    Revision 1.4  2007/01/17 16:23:10  jfisher
 *    Added vendor id to the report
 *
 *    Revision 1.3  2006/12/20 16:54:01  jfisher
 *    Fixed the a problem with the customer specific query
 *
 *      
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.web.WebReport;
import com.emerywaterhouse.websvc.Param;


public class DlrMktPreOrder extends Report
{
   private static final short MAX_COLS = 10;
   
   private PreparedStatement m_CustOrdData;
   
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private HSSFCellStyle[] m_CellStyles;
   
   //
   // workbook entries.
   private HSSFWorkbook m_Wrkbk;
   private HSSFSheet m_Sheet;
   
   private int m_ShowId;
   private String m_CustId;  
   private boolean m_SingleCust;
      
   //
   // web_report crappola
   private WebReport m_WebRpt;
   private String m_OutFormat;
   private int m_WebRptId;
   private int m_Cnt = 0;
   
   /**
    * 
    */
   public DlrMktPreOrder()
   {
      super();
      
      m_Wrkbk = new HSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      m_SingleCust = false;
      
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
     
      m_CustId = null;
      m_WebRpt = null;
           
      super.finalize();
   }
   
   /**
    * Builds the email message that will be sent to the customer.  This overrides the default message
    * built by the RptProcessor.
    *
    * @param rptName
    * @return EMail message String
    */
   private String buildEmailText(String rptName)
   {
      StringBuffer msg = new StringBuffer();

      msg.append("The following report are ready for you to pick up:\r\n");
      msg.append("\tDealer Market Pre-order\r\n\r\n");
      msg.append("To view your reports:\r\n");
      msg.append("\thttp://www.emeryonline.com/emerywh/subscriber/my_account/report_list.jsp\r\n\r\n");
      msg.append("If you have any questions or suggestions, please contact help@emeryonline.com\r\n");
      msg.append("or call 800-283-0236 ext. 1.");

      return msg.toString();
   }
      
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      HSSFRow row = null;
      int rowNum = 0;
      int colNum = 0;
      FileOutputStream outFile = null;
      ResultSet custOrdData = null;
      boolean result = false;
      String itemId = null;
      String po = null;

      m_FileNames.set(0, m_FileNames.get(0) + m_WebRpt.getWebReportId() + ".xls" );
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         rowNum = createCaptions();         
        
         if ( m_SingleCust ) {            
            m_CustOrdData.setInt(1, m_ShowId);
            m_CustOrdData.setString(2, m_CustId);
         }
         else {
            m_CustOrdData.setInt(1, m_ShowId);
            m_CustOrdData.setInt(2, m_ShowId);
         }
         
         custOrdData = m_CustOrdData.executeQuery();

         while ( custOrdData.next() && m_Status == RptServer.RUNNING ) {
            m_CustId = custOrdData.getString("customer_id");
            itemId = custOrdData.getString("item_id");
            po = custOrdData.getString("po_num");
            
            setCurAction("processing customer: " + m_CustId + " " + itemId);
            row = createRow(rowNum++, MAX_COLS);
            colNum = 0;
                        
            if ( row != null ) {               
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(m_CustId));
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(custOrdData.getString("custname")));
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(custOrdData.getString("name")));
               row.getCell(colNum++).setCellValue(custOrdData.getInt("vendor_id"));
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(itemId));
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(custOrdData.getString("description")));
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(custOrdData.getString("qty_ordered")));               
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(custOrdData.getString("stkpack_nbc")));
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(custOrdData.getString("packet_id")));
               row.getCell(colNum++).setCellValue(new HSSFRichTextString(po != null ? po : ""));
               
               m_Cnt++;
            }
         }
         
         m_Wrkbk.write(outFile);
         custOrdData.close();

         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[DlrMktPreOrder]", ex);
         
         m_WebRpt.setComments("Unable to build report " + ex.getMessage());
         m_WebRpt.setStatus("ERROR");
      }

      finally {         
         row = null;
                  
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[DlrMktPreOrder]", e);
         }

         outFile = null;
      }

      m_WebRpt.setLineCount(m_Cnt);
      
      return result;
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_CustOrdData);
   }
   
   /**
    * Sets the captions on the report.
    */
   private int createCaptions()
   {
      HSSFRow row = null;      
      int rowNum = 0;
      int colNum = 0;
            
      if ( m_Sheet == null )
         return 0;
      
      //
      // Create the row for the captions.
      row = m_Sheet.createRow(rowNum++);
      
      if ( row != null ) {
         for ( int i = 0; i < MAX_COLS; i++ ) {
            row.createCell(i);            
         }
            
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Cust ID"));
         row.getCell(colNum).setCellValue(new HSSFRichTextString("Cust Name"));
         m_Sheet.setColumnWidth(colNum++, 6000);
         row.getCell(colNum).setCellValue(new HSSFRichTextString("Vnd Name"));
         m_Sheet.setColumnWidth(colNum++, 10000);
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Vnd Id"));
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Item ID"));      
         row.getCell(colNum).setCellValue(new HSSFRichTextString("Descr"));
         m_Sheet.setColumnWidth(colNum++, 10000);
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Qty Ord"));
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("StkPk/NBC"));      
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("Packet"));
         row.getCell(colNum++).setCellValue(new HSSFRichTextString("PO"));
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
      String fileName = null;
      
      try {         
         m_EdbConn = m_RptProc.getEdbConn();
         
         if ( prepareStatements() )
            created = buildOutputFile();
         //
         // Only send this if this was requested from the web.
         if ( created && m_WebRptId > 0 ) {
            fileName = m_FileNames.get(0);

            if ( m_RptProc.getZipped() ) {
               fileName = fileName.substring(0, fileName.indexOf('.')+1) + "zip";
               m_WebRpt.setZipped(true);
            }
            else
               m_WebRpt.setZipped(false);

            m_WebRpt.setFileName(fileName);
            m_WebRpt.setStatus("COMPLETE");

            //
            // Save the web_report entry to the database
            m_WebRpt.update();
            m_EdbConn.commit();

            m_RptProc.setEmailMsg(buildEmailText(fileName));
         }
      }
      
      catch ( Exception ex ) {
         log.fatal("[DlrMktPreOrder]", ex);
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
   private HSSFRow createRow(int rowNum, int colCnt)
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
            cell.setCellStyle(m_CellStyles[i]);
         }
      }

      return row;
   }
   
   /**
    * Retrieve the web_report record from the database.  The web_report_id is
    * stored in parameter 15.  The record should have been created by the
    * PurchHistRptSrv servlet, but if not found, will be created here.
    *
    * @return boolean
    */
   private boolean getWebReport()
   {    
      boolean result = true;
      
      //
      // set the connection in the WebReport bean
      try {
         m_WebRpt = new WebReport();
         m_WebRpt.setConnection(m_RptProc.getEdbConn());
         m_WebRpt.setReportName("Dealer Market Pre-order");
      }

      catch ( Exception ex ) {
         log.error("[DlrMktPreOrder] Unable to set connection in web report ", ex);
         result = false;
      }

      //
      // Load the web_report id from the 15th parameter.  If an id was passed,
      // load that web_report record into the WebReport bean.  Otherwise, create
      // a new web_report record.
      try {
         if ( m_WebRptId >= 0 )
            m_WebRpt.load(m_WebRptId);
         else
            m_WebRpt.insert();
      }

      catch ( Exception ex ) {
         log.error("[DlrMktPreOrder] unable to create a web_report record ", ex);
         result = false;
      }

      return result;
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
            sql.append("select ");
            sql.append("order_header.customer_id, customer.name custname, vendor.name, ");
            sql.append("vendor.vendor_id, item_entity_attr.item_id, item_entity_attr.description, qty_ordered, ");
            sql.append("stock_pack || decode(broken_case.description, 'ALLOW BROKEN CASES', ' ', 'N') as stkpack_nbc, ");
            sql.append("promotion.packet_id, po_num ");
            sql.append("from ");
            sql.append("promotion ");
            sql.append("join order_line on order_line.promo_id = promotion.promo_id ");
            sql.append("join order_header on order_header.order_id = order_line.order_id ");
            sql.append("join customer on customer.customer_id = order_header.customer_id ");
            sql.append("join item_entity_attr on item_entity_attr.item_ea_id = order_line.item_ea_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join broken_case on broken_case.broken_case_id = ejd_item.broken_case_id ");
            
            sql.append("join cust_warehouse on cust_warehouse.customer_id = customer.customer_id and whs_priority = 1 ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = ejd_item.ejd_item_id and ");
            sql.append("      ejd_item_warehouse.warehouse_id = cust_warehouse.warehouse_id ");
            sql.append("where promotion.packet_id in ( ");
            sql.append("   select packet_id ");
            sql.append("   from show_packet ");
            sql.append("   where show_id = ? ");
            sql.append(") and ");
            
            //
            // If we get a customer id for parameter, only look for that customer.  Otherwise, get them all.
            if ( m_SingleCust )
               sql.append("   order_header.customer_id = ? ");
            else {              
               sql.append("   order_header.customer_id in ( ");
               sql.append("       select customer_id ");
               sql.append("       from show_cust ");
               sql.append("       where show_id = ? ");
               sql.append("   ) ");
            }
            
            sql.append("order by customer_id, vendor.name, item_entity_attr.item_id " );

            m_CustOrdData = m_EdbConn.prepareStatement(sql.toString());
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[DlrMktPreOrder]", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("[DlrMktPreOrder].prepareStatements - null Edb connection");
      
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
      String email = null;
      
      m_OutFormat = "EXCEL";
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                  
         if ( param.name.equals("custid") )
            m_CustId = param.value;
         
         if ( param.name.equals("showid") )
            m_ShowId = Integer.parseInt(param.value);
      
         if ( param.name.equalsIgnoreCase("webreportid") ) {
            try {
               m_WebRptId = Integer.parseInt(param.value);
            }
   
            catch ( Exception ex ) {
               m_WebRptId = -1;
            }
         }
         
         if ( param.name.equalsIgnoreCase("email") )
            email = param.value;
      }
      
      fileName.append("dlrmktpreord");
      fileName.append("-");
      
      if ( m_CustId != null && m_CustId.length() == 6 ) {
         fileName.append(m_CustId);
         fileName.append("-");
         m_SingleCust = true;
      }
      
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xls");
      m_FileNames.add(fileName.toString());
      
      if ( getWebReport() ) {
         try {
            m_WebRpt.setFileName(m_FileNames.get(0));
            m_WebRpt.setEMail(email);
            m_WebRpt.setDocFormat(m_OutFormat);
            m_WebRpt.setLineCount(0);
            m_WebRpt.setZipped(m_RptProc.getZipped());
            m_WebRpt.setStatus("RUNNING");
            m_WebRpt.update();
            m_WebRpt.getConnection().commit();
         }
         
         catch ( Exception e ) {
            m_WebRpt.setComments("Unable to set web_report parameters " + e.getMessage());

            try {
               m_WebRpt.update();
               m_EdbConn.commit();
            }
            
            catch ( Exception ex ) {
            }

            log.error("[DlrMktPreOrder]", e);
         }
      }
   }
      
   /**
    * Sets up the styles for the cells based on the column data.  Does any other inititialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {      
      HSSFCellStyle styleText;      // Text left justified
      HSSFCellStyle styleInt;       // Style with 0 decimals
      HSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00) 
            
      styleText = m_Wrkbk.createCellStyle();      
      styleText.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      
      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      styleInt.setDataFormat((short)3);

      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      styleMoney.setDataFormat((short)8);
      
      m_CellStyles = new HSSFCellStyle[] {
         styleText,    // col 0 cust id
         styleText,    // col 1 cust name
         styleText,    // col 2 vnd name
         styleText,    // col 3 vnd id 
         styleText,    // col 4 item id
         styleText,    // col 5 item desc
         styleInt,     // col 6 qty ord
         styleText,    // col 7 stock pack/nbc
         styleText,    // col 8 packet
         styleText,    // col 9 po         
      };
      
      styleText = null;
      styleInt = null;
      styleMoney = null;
   }
}
