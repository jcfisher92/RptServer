/**
 * File: UnitDollarSalesCust.java
 * Description: The unit dollar sales by customer excel spreadsheet.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 05/20/2005
 * Last Update: $Id: UnitDollarSalesCust.java,v 1.21 2009/04/24 15:15:21 pdavidson Exp $
 * 
 * History
 *    $Log: UnitDollarSalesCust.java,v $
 *    Revision 1.21  2009/04/24 15:15:21  pdavidson
 *    Show sales and margin breakdown by warehouse
 *
 *    Revision 1.20  2009/04/08 04:50:24  pdavidson
 *    Fixed all subreport createCaptions methods to not recreate the caption row each time when adding the columns
 *
 *    Revision 1.19  2009/04/08 04:29:25  pdavidson
 *    Fixed all subreport createCaptions methods to not recreate the caption row each time when adding the columns
 *
 *    Revision 1.18  2009/03/27 01:48:24  jfisher
 *    Fixed missing parentheses in the vendor query
 *
 *    Revision 1.17  2009/03/23 18:10:05  jfisher
 *    Added the sale type to the where clause to remove transfers
 *
 *    Revision 1.16  2009/03/04 20:49:30  jfisher
 *    Fixed the createRow bug when adding captions.
 *
 *    Revision 1.15  2009/02/18 16:53:10  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.14  2008/10/30 16:41:15  jfisher
 *    Fixed warnings and added generic type to the constructor class for the subreport.
 *
 *    Revision 1.13  2007/12/12 06:20:24  jfisher
 *    Added new fields for SSarkozy
 *
 *    Revision 1.12  2007/01/17 16:25:48  jfisher
 *    Removed some annoying white space that eclipse sometimes adds.
 *
 *    Revision 1.11  2006/09/05 18:52:23  jfisher
 *    Put in a check for a null sub report to stop the null pointer exception
 *
 *    Revision 1.10  2006/04/28 13:41:30  jfisher
 *    Fixed problem with not having complete sub report name and missing nullary constructor
 *
 *    Revision 1.9  2006/03/14 14:25:50  jfisher
 *    changed the inner class instatiation so it doesn't use the getName method.
 *
 *    Revision 1.8  2006/03/13 13:26:01  jfisher
 *    Modified the inner classes to use the initCells method.
 *
 *    Revision 1.7  2006/02/17 13:54:15  jfisher
 *    *** empty log message ***
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.lang.reflect.Constructor;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class UnitDollarSalesCust extends Report
{
   //
   // Reporting periods
   private static final int PT_ROLLING     = 0;
   //private static final int PT_SPECIFY     = 1;
   
   //
   // Report type filter identifiers
   private static final int RT_VENDOR      = 0;
   private static final int RT_VENDOR_FLC  = 1;
   private static final int RT_FLC         = 2;
   private static final int RT_ITEM        = 3;
   private static final int RT_RMS_SUM     = 4;
   private static final int RT_RMS_DTL     = 5;
   
   //
   // Column widths
   private static final int CW_ADDR1       = 8000;
   private static final int CW_ADDR2       = 4000;
   private static final int CW_CITY        = 4000;
   private static final int CW_CUST_ID     = 3000;
   private static final int CW_CUST_NAME   = 9000;
   private static final int CW_DESC        = 9000;
   private static final int CW_FLC_NAME    = 8000;
   private static final int CW_LINES       = 4000;
   private static final int CW_SELL        = 4000;
   private static final int CW_VND_ID      = 3000;
   private static final int CW_VND_NAME    = 9000;
   private static final int CW_UNITS       = 4000;
      
   //
   // cust data array index identifiers
   private static final int C_NAME  = 0;
   private static final int C_ADDR1 = 1;
   private static final int C_ADDR2 = 2;
   private static final int C_CITY  = 3;
   private static final int C_STATE = 4;
   private static final int C_ZIP   = 5;
   private static final int C_PHONE = 6;
   private static final int C_FAX   = 7;
   
   //
   // Inner class identifiers.  This is needed because using class.getClassName forces the class to 
   // be loaded by the system class loader.  These classes need to be loaded by the EmLoader to be
   // dynamic and work with the outer class.
   private static final String UDSC_VND      = "com.emerywaterhouse.rpt.spreadsheet.UnitDollarSalesCust$UDSCVnd";
   private static final String UDSC_VNDFLC   = "com.emerywaterhouse.rpt.spreadsheet.UnitDollarSalesCust$UDSCVndFlc";
   private static final String UDSC_FLC      = "com.emerywaterhouse.rpt.spreadsheet.UnitDollarSalesCust$UDSCFlc";
   private static final String UDSC_ITEM     = "com.emerywaterhouse.rpt.spreadsheet.UnitDollarSalesCust$UDSCItem";
   private static final String UDSC_RMSSUM   = "com.emerywaterhouse.rpt.spreadsheet.UnitDollarSalesCust$UDSCRmsSum";
   private static final String UDSC_RMSDTL   = "com.emerywaterhouse.rpt.spreadsheet.UnitDollarSalesCust$UDSCRmsDtl";
   
   private String m_BegDate;
   private String m_EndDate;   
   private String m_FlcId;
   private String m_ItemId;
   private String m_RmsIds;
   private int m_RptType;
   private int m_PeriodType;
   private int m_VndId;
   private String[] m_CustDat;
   
   private PreparedStatement m_ItemSales;
   private PreparedStatement m_CustInfo;
   private PreparedStatement m_CustPhone;
   
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle m_CSTitle;     // Bold, left justified
   private XSSFCellStyle m_CSCaption;   // Bold, centered
   private XSSFCellStyle m_CSText;      // Text right justified
   private XSSFCellStyle m_CSInt;       // Style with 0 decimals
   private XSSFCellStyle m_CSMoney;     // Money ($#,##0.00_);[Red]($#,##0.00) 
   private XSSFCellStyle m_CSPct;       // Style with 0 decimals + %
      
   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;
   
   /**
    * default constructor
    */
   public UnitDollarSalesCust()
   {
      super();
      
      m_CustDat = new String[8];
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      XSSFFont font = m_Wrkbk.createFont();
      
      try {
         font.setFontHeightInPoints((short)10);
         font.setFontName("Arial");
         font.setBold(true);
         
         m_CSText = m_Wrkbk.createCellStyle();      
         m_CSText.setAlignment(HorizontalAlignment.LEFT);
         
         m_CSInt = m_Wrkbk.createCellStyle();
         m_CSInt.setAlignment(HorizontalAlignment.RIGHT);
         m_CSInt.setDataFormat((short)3);
   
         m_CSMoney = m_Wrkbk.createCellStyle();
         m_CSMoney.setAlignment(HorizontalAlignment.RIGHT);
         m_CSMoney.setDataFormat((short)8);
         
         m_CSPct = m_Wrkbk.createCellStyle();
         m_CSPct.setAlignment(HorizontalAlignment.RIGHT);
         m_CSPct.setDataFormat((short)9);
         
         m_CSTitle = m_Wrkbk.createCellStyle();
         m_CSTitle.setFont(font);
         m_CSTitle.setAlignment(HorizontalAlignment.LEFT);
         
         m_CSCaption = m_Wrkbk.createCellStyle();
         m_CSCaption.setFont(font);
         m_CSCaption.setAlignment(HorizontalAlignment.CENTER);
      }
      
      finally {
         font = null;
      }
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable 
    */
   public void finalize() throws Throwable
   {      
      for ( int i = 0; i < m_CustDat.length; i++ )
         m_CustDat[i] = null;
      
      m_CustDat = null;
      m_CSText = null;
      m_CSInt = null;
      m_CSMoney = null;
      m_CSPct = null;
      m_CSTitle = null;
      m_Sheet = null;
      m_Wrkbk = null;      
            
      
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
      boolean result = false;      
      FileOutputStream outFile = null;      
      SubRpt subRpt = null;
      int rowNum = 0;
            
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         switch ( m_RptType ) {
            case RT_VENDOR:
               subRpt = createSubReport(UDSC_VND);
            break;
            
            case RT_VENDOR_FLC:
               subRpt = createSubReport(UDSC_VNDFLC);
            break;
            
            case RT_FLC:
               subRpt = createSubReport(UDSC_FLC);
            break;
            
            case RT_ITEM:
               subRpt = createSubReport(UDSC_ITEM);
            break;
            
            case RT_RMS_SUM:
               subRpt = createSubReport(UDSC_RMSSUM);
            break;
               
            case RT_RMS_DTL:
               subRpt = createSubReport(UDSC_RMSDTL);
            break;
         }
         
         if ( subRpt != null ) {
            subRpt.setWrkbk(m_Wrkbk);
            subRpt.setSheet(m_Sheet);
            rowNum = createCaptions(subRpt);
            subRpt.build(outFile, ++rowNum);
            m_Wrkbk.write(outFile);
            result = true;
         }
         else {
            log.error("[UnitDollarSalesCust] null sub report object");            
         }
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         log.error("[UnitDollarSalesCust]", ex);
      }

      finally {         
         subRpt = null;
                  
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[UnitDollarSalesCust]", e);
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
      closeStmt(m_ItemSales);
      closeStmt(m_CustInfo);
      closeStmt(m_CustPhone);
      
      m_ItemSales = null;
      m_CustInfo = null;
      m_CustPhone = null;
   }
   
   /**
    * Creates the report title and the captions.
    */
   private int createCaptions(SubRpt subRpt)
   {
      XSSFRow row = null;
      XSSFCell cell = null;
      int rowNum = 0;
      
      if ( m_Sheet == null )
         return 0;
      
      try {
         //
         // set the main report title
         row = m_Sheet.createRow(rowNum);
         cell = row.createCell(0);
         cell.setCellType(CellType.STRING);
         cell.setCellStyle(m_CSTitle);
         cell.setCellValue(new XSSFRichTextString("Unit & Dollar Customer Sales Report"));
               
         rowNum = 2;      
         subRpt.createCaptions(rowNum);
      }
      
      finally {
         row = null;
         cell = null;
      }
            
      return ++rowNum;
   }
   
   /**
    * Creates the report.
    * 
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {      
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {
         m_EdbConn = m_RptProc.getEdbConn();
         prepareStatements();
         created = buildOutputFile();
      }
      
      catch ( Exception ex ) {
         log.fatal("[UnitDollarSalesCust]", ex);
      }
      
      finally {
         closeStatements();
         
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }
      
      return created;
   }

   /**
    * Dynamically instantiates a sub report inner class.
    * @param className The name of the inner class to instantiate.
    * 
    * @return A SubRpt reference to the class that was instantiated, or null if there was an error.
    */
   private SubRpt createSubReport(String className)
   {
      SubRpt subRpt = null;
      Class<?> c = null;
      Constructor<?> constr;
      Object obj = null;
      
      if ( className != null ) {
         try {
            c = RptServer.getLoader().loadClass(className);
            constr = c.getConstructor(new Class[]{this.getClass()});
            obj = constr.newInstance(new Object[] {this});
            subRpt = (SubRpt)obj;
         }
         
         catch ( Exception ex ) {
            log.error("[UnitDollarSalesCust]", ex);
         }
      }
      else
         log.warn("UnitDollarSalesCust null class name for sub report");
      
      return subRpt;
   }
   
   

   /**
    * Gets the customer data in an array.  The specific fields are:
    *    0 = name
    *    1 = addr1
    *    2 = addr2
    *    3 = city
    *    4 = state
    *    5 = zip
    *    7 = bus phone
    *    8 = bus fax
    *    
    * @param custId The customer id to look up data for.
    * @return An array of customer data based on the customer id passed in.
    * 
    * @throws SQLException 
    */
   private String[] getCustDat(String custId) throws SQLException
   {      
      ResultSet custAddr = null;
      ResultSet custPhone = null;
      
      if ( custId != null && custId.length() == 6 ) {
         m_CustInfo.setString(1, custId);         
         custAddr = m_CustInfo.executeQuery();
         
         try {
            while ( custAddr.next() ) {
               m_CustDat[0] = custAddr.getString("name");
               m_CustDat[1] = custAddr.getString("addr1");
               m_CustDat[2] = custAddr.getString("addr2");
               m_CustDat[3] = custAddr.getString("city");
               m_CustDat[4] = custAddr.getString("state");
               m_CustDat[5] = custAddr.getString("postal_code");
               
               //
               // Check the shipping address first.  If it's been populated then we're done.
               // Otherwise we'll go onto the mailing address.
               if ( (m_CustDat[1] != null && m_CustDat[1].length() > 0) || (m_CustDat[2] != null && m_CustDat[2].length() > 0) )
                  break;
            }
         }
         
         finally {
            closeRSet(custAddr);
            custAddr = null;
         }
         
         m_CustPhone.setString(1, custId);
         custPhone = m_CustPhone.executeQuery();
         
         try {
            while ( custPhone.next() ) {
               if ( custPhone.getString("phone_type").equals("BUSINESS"))
                  m_CustDat[6] = custPhone.getString("phone_number");
               else
                  m_CustDat[7] = custPhone.getString("phone_number");
            }
         }
         
         finally {
            closeRSet(custPhone);
            custPhone = null;
         }
         
         //
         // prevent any null data from showing up in the report.
         for ( int i = 0; i < m_CustDat.length; i++ )
            if ( m_CustDat[i] == null )
               m_CustDat[i] = "";
      }
      
      return m_CustDat;
   }
      
   /**
    * Builds the sql for the time period of the report.
    * 
    * @return The sql snippet
    */
   private String getPeriod()
   {
      StringBuffer sql = new StringBuffer();
      
      //
      // Set the period
      if ( m_PeriodType == PT_ROLLING )
         sql.append("invoice_date between (current_date - 365) and current_date ");
      else {
         sql.append(String.format("invoice_date between to_date('%s', 'mm/dd/yyyy') and ", m_BegDate));
         sql.append(String.format(" to_date('%s', 'mm/dd/yyyy') ", m_EndDate));
      }
      
      return sql.toString();
   }
   
   /**
    * Convenience method that generates the caption text for the period of the report.
    * 
    * @return String containing the caption.
    */
   private String getPeriodCaption()
   {
      StringBuffer caption = new StringBuffer();
      
      if ( m_PeriodType == PT_ROLLING )      
         caption.append("Rolling 12 Months");
      else {
         caption.append(m_BegDate);
         caption.append(" - ");
         caption.append(m_EndDate);
      }
      
      return caption.toString();
   }
   
   /**
    * Prepares the sql queries for execution.
    *     
    */
   private void prepareStatements() throws Exception
   {            
      StringBuffer sql = new StringBuffer(256);
            
      if ( m_EdbConn != null ) {
         sql.append("select customer.customer_id, name, addr1,  addr2, ");
         sql.append("city, state, postal_code ");
         sql.append("from customer ");
         sql.append("left outer join ( ");
         sql.append("   select cust_address.customer_id, addr1, addr2, city, ");
         sql.append("   state, postal_code, addr_link_type.description ");
         sql.append("   from cust_address" );
         sql.append("   join cust_addr_link on cust_addr_link.cust_addr_id = cust_address.cust_addr_id ");
         sql.append("   join addr_link_type on addr_link_type.addr_link_type_id = cust_addr_link.addr_link_type_id and ");
         sql.append("      (addr_link_type.description = 'SHIPPING' or addr_link_type.description = 'MAILING') ");         
         sql.append(") ca on ca.customer_id = customer.customer_id ");
         sql.append("where customer.customer_id = ? ");
         sql.append("order by description desc");
         
         m_CustInfo = m_EdbConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select phone_number, phone_type ");
         sql.append("from cust_contact ");
         sql.append("join cust_contact_type on cust_contact_type.cct_id = cust_contact.cct_id and description = 'ACCOUNTS PAYABLE'");
         sql.append("join emery_contact_phone on emery_contact_phone.ec_id = cust_contact.ec_id ");
         sql.append("join contact_phone on contact_phone.cont_phone_id = emery_contact_phone.cont_phone_id ");
         sql.append("join phone_type on phone_type.phone_type_id = contact_phone.phone_type_id and ");
         sql.append("      phone_type in ('BUSINESS','BUSINESS FAX') ");
         sql.append("where customer_id = ? ");         
         sql.append("order by phone_type");
                  
         m_CustPhone = m_EdbConn.prepareStatement(sql.toString());
      }
   }
   
   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    * 
    * Because it's possible that this report can be called from some other system, the
    * best way to deal with params is to not go by the order, but by the name.
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fname = new StringBuffer();
      String tm = Long.toString(System.currentTimeMillis()).substring(3);
      int pcount = params.size();
      Param param = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
         
         if ( param.name.equals("prdtype") && param.value.trim().length() > 0  )
            m_PeriodType = Integer.parseInt(param.value);
         
         if ( param.name.equals("rpttype") && param.value.trim().length() > 0  )
            m_RptType = Integer.parseInt(param.value);
         
         if ( param.name.equals("flc") )
            m_FlcId = param.value.trim();
         
         if ( param.name.equals("vendor") && param.value.trim().length() > 0 )
            m_VndId = Integer.parseInt(param.value);
         
         if ( param.name.equals("item") )
            m_ItemId = param.value.trim();
         
         if ( param.name.equals("rmsid") )
            m_RmsIds = param.value.trim();
                     
         if ( param.name.equals("begdate") )
            m_BegDate = param.value.trim();
         
         if ( param.name.equals("enddate") )
            m_EndDate = param.value.trim();
      }
      
      //
      // Build the file name.
      fname.append(tm);
      fname.append("-");
      fname.append(m_RptProc.getUid());
      fname.append("udsc.xlsx");
      m_FileNames.add(fname.toString());
   }
   
   
   
   /**
    * Vendor subreport class.  Creates the data based on the vendor filter parameter.
    */
   public class UDSCVnd extends SubRpt
   {      
      public UDSCVnd()
      {
         super();
      }
 
      /**
       * @see SubRpt#build(FileOutputStream out, int rowNum)
       */
      public boolean build(FileOutputStream out, int rowNum)
      {
         boolean processed = false;
         ResultSet itemSales = null;
         XSSFRow row = null;
         String custId = null;
         int col;
         double sold = 0.0;
         double cost = 0.0;
         double margin = 0.0;
         
         try {
            m_ItemSales = m_EdbConn.prepareStatement(buildSql());            
            m_ItemSales.setString(1, Integer.toString(m_VndId));
                        
            itemSales = m_ItemSales.executeQuery();
   
            while ( itemSales.next() && m_Status == RptServer.RUNNING ) {
               row = createDataRow(rowNum++);
               custId = itemSales.getString("cust_nbr");
               setCurAction("processing customer: " + custId);
               getCustDat(custId);
               col = 0;
               
               row.getCell(col).setCellValue(new XSSFRichTextString(itemSales.getString("name")));
               row.getCell(++col).setCellValue(m_VndId);
               row.getCell(++col).setCellValue(new XSSFRichTextString(custId));               
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_NAME]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR1]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR2]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_CITY]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_STATE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ZIP]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_PHONE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_FAX]));
               row.getCell(++col).setCellValue(itemSales.getInt("extqty"));
               row.getCell(++col).setCellValue(itemSales.getDouble("lines"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extsell"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extcost"));
               
               sold = itemSales.getDouble("extsell");
               cost = itemSales.getDouble("extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);

               // Portland numbers
               sold = itemSales.getDouble("portland_extsell");
               cost = itemSales.getDouble("portland_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("portland_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
               
               // Pittston numbers
               sold = itemSales.getDouble("pittston_extsell");
               cost = itemSales.getDouble("pittston_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("pittston_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
            }
            
            processed = true;
         }
         
         catch ( Exception ex ) {
            m_ErrMsg.append(ex.getClass().getName() + "\r\n");
            m_ErrMsg.append(ex.getMessage());
            
            log.fatal("[UnitDollarSalesCust] ", ex);
         }
         
         finally {
            closeRSet(itemSales);
         }
         
         return processed;
      }
      
      /**
       * Builds the sql for the vendor filter.
       * 
       * @see SubRpt#buildSql()
       */
      public String buildSql()
      {
         StringBuffer sql = new StringBuffer();
         
         sql.append("select vendor_id, vendor.name, cust_nbr, ");
         sql.append("sum(qty_shipped) as extqty, sum(ext_sell) as extsell, ");
         sql.append("sum(ext_cost) as extcost, count(*) lines, ");
         
         // This was the most efficient way to get at this data from within the same query.  If we add another
         // warehouse in future, which is unlikely, it will have to be added to here, if that ever happens.
         sql.append("sum(decode(warehouse, 'PORTLAND', qty_shipped, 0)) as portland_extqty, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_sell, 0)) as portland_extsell, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_cost, 0)) as portland_extcost, ");
         sql.append("sum(decode(warehouse, 'PITTSTON', qty_shipped, 0)) as pittston_extqty, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_sell, 0)) as pittston_extsell, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_cost, 0)) as pittston_extcost ");         
         sql.append("from inv_dtl ");
         sql.append("join vendor on vendor.vendor_id = to_number(inv_dtl.vendor_nbr) ");
         sql.append("where sale_type = 'WAREHOUSE' and vendor_nbr = ? and ");
         sql.append(getPeriod());         
         sql.append("group by vendor_id, vendor.name, cust_nbr");
                           
         return sql.toString();
      }
      
      /**
       * Creates the captions for the vendor filter.
       * 
       * @see SubRpt#createCaptions(int rowNum)
       */
      public int createCaptions(int rowNum)
      {
         StringBuffer caption = new StringBuffer();
         XSSFRow row = null;         
         int col = 0;
         
         caption.append("Select Vendor:  Time Frame: ");
         caption.append(getPeriodCaption());
         
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSTitle);
         createCaptionCell(row, col, caption.toString());
         
         rowNum++;
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSCaption);
         createCaptionCell(row, col, "Vendor Name");
         m_Sheet.setColumnWidth(col++, CW_VND_NAME);
         createCaptionCell(row, col, "Vendor ID");
         m_Sheet.setColumnWidth(col++, CW_VND_ID);
         createCaptionCell(row, col, "Customer ID");
         m_Sheet.setColumnWidth(col++, CW_CUST_ID);
         createCaptionCell(row, col, "Customer Name");
         m_Sheet.setColumnWidth(col++, CW_CUST_NAME);
         createCaptionCell(row, col, "Address 1");
         m_Sheet.setColumnWidth(col++, CW_ADDR1);
         createCaptionCell(row, col, "Address 2");
         m_Sheet.setColumnWidth(col++, CW_ADDR2);
         createCaptionCell(row, col, "City");
         m_Sheet.setColumnWidth(col++, CW_CITY);
         createCaptionCell(row, col++, "State");
         createCaptionCell(row, col++, "Zip");
         createCaptionCell(row, col++, "Phone");
         createCaptionCell(row, col++, "Fax");
         createCaptionCell(row, col, "Units");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Lines");
         m_Sheet.setColumnWidth(col++, CW_LINES);
         createCaptionCell(row, col, "Dollars");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Cost");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin%");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Portland");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Pittston");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
                  
         return rowNum;
      }
      
      /**
       * Initializes the cell styles array.
       */
      protected void initCellStyles() 
      {
         m_CellStyles = new XSSFCellStyle[] {
            m_CSText,    // col 0 vnd name
            m_CSText,    // col 1 vnd id
            m_CSText,    // col 2 cust#
            m_CSText,    // col 3 cust name
            m_CSText,    // col 4 addr1
            m_CSText,    // col 5 addr2
            m_CSText,    // col 6 city
            m_CSText,    // col 7 state
            m_CSText,    // col 8 zip
            m_CSText,    // col 9 phone
            m_CSText,    // col 10 fax
            m_CSInt,     // col 11 units
            m_CSInt,     // col 12 lines
            m_CSMoney,   // col 13 dollars
            m_CSMoney,   // col 14 cost
            m_CSPct,     // col 15 margin%
            m_CSMoney,   // col 16 margin$
            
            m_CSInt,     // col 17 units portland
            m_CSMoney,   // col 18 dollars portland
            m_CSMoney,   // col 19 cost portland
            m_CSPct,     // col 20 margin% portland
            m_CSMoney,   // col 21 margin$ portland
            
            m_CSInt,     // col 22 units pittston
            m_CSMoney,   // col 23 dollars pittston
            m_CSMoney,   // col 24 cost pittston
            m_CSPct,     // col 25 margin% pittston
            m_CSMoney    // col 26 margin$ pittston
         };
      }
      
   }
   
   
   
   /**
    * Sub report class to handle the vendor/FLC filter.
    */
   public class UDSCVndFlc extends SubRpt
   {
      /**
       * Overridden constructor that initializes the sub report.
       */
      public UDSCVndFlc() 
      {
         super();
      }

      /**
       * Builds the report data.  Gets called from the main report buildOutputFile method.
       * @see com.emerywaterhouse.rpt.spreadsheet.SubRpt#build(java.io.FileOutputStream, int)
       */
      @Override
      public boolean build(FileOutputStream out, int rowNum)
      {         
         boolean processed = false;
         ResultSet itemSales = null;
         XSSFRow row = null;
         String custId = null;
         int col;
         String flcStr = null;
         double sold = 0.0;
         double cost = 0.0;
         double margin = 0.0;
         
         try {
            m_ItemSales = m_EdbConn.prepareStatement(buildSql());            
            m_ItemSales.setString(1, Integer.toString(m_VndId));            
            m_ItemSales.setString(2, m_FlcId);
            
            itemSales = m_ItemSales.executeQuery();
   
            while ( itemSales.next() && m_Status == RptServer.RUNNING ) {
               row = createDataRow(rowNum++);
               custId = itemSales.getString("cust_nbr");
               flcStr = itemSales.getString("flc") + " " + itemSales.getString("description");
               
               setCurAction("processing customer: " + custId);
               getCustDat(custId);
               col = 0;
               
               row.getCell(col).setCellValue(new XSSFRichTextString(itemSales.getString("name")));
               row.getCell(++col).setCellValue(m_VndId);
               row.getCell(++col).setCellValue(new XSSFRichTextString(flcStr));
               row.getCell(++col).setCellValue(new XSSFRichTextString(custId));               
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_NAME]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR1]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR2]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_CITY]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_STATE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ZIP]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_PHONE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_FAX]));
               row.getCell(++col).setCellValue(itemSales.getInt("extqty"));
               row.getCell(++col).setCellValue(itemSales.getInt("lines"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extsell"));
               
               row.getCell(++col).setCellValue(itemSales.getDouble("extcost"));
               
               sold = itemSales.getDouble("extsell");
               cost = itemSales.getDouble("extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);

               // Portland numbers
               sold = itemSales.getDouble("portland_extsell");
               cost = itemSales.getDouble("portland_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("portland_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
               
               // Pittston numbers
               sold = itemSales.getDouble("pittston_extsell");
               cost = itemSales.getDouble("pittston_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("pittston_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
            }
            
            processed = true;
         }
         
         catch ( Exception ex ) {
            m_ErrMsg.append(ex.getClass().getName() + "\r\n");
            m_ErrMsg.append(ex.getMessage());
            
            log.fatal("[UnitDollarSalesCust] ", ex);
         }
         
         finally {
            closeRSet(itemSales);
         }
         
         return processed;
      }
      
      /**
       * Builds the sql for the vendor flc filter.
       * 
       * @see SubRpt#buildSql()
       */
      public String buildSql()
      {
         StringBuffer sql = new StringBuffer();
         
         sql.append("select vendor_id, vendor.name, flc, description, ");
         sql.append("cust_nbr, sum(qty_shipped) as extqty, sum(ext_sell) as extsell, ");
         sql.append("sum(ext_cost) as extcost, count(*) as lines, ");
         
         // This was the most efficient way to get at this data from within the same query.  If we add another
         // warehouse in future, which is unlikely, it will have to be added to here, if that ever happens.
         sql.append("sum(decode(warehouse, 'PORTLAND', qty_shipped, 0)) as portland_extqty, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_sell, 0)) as portland_extsell, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_cost, 0)) as portland_extcost, ");
         sql.append("sum(decode(warehouse, 'PITTSTON', qty_shipped, 0)) as pittston_extqty, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_sell, 0)) as pittston_extsell, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_cost, 0)) as pittston_extcost ");            
         sql.append("from inv_dtl ");
         sql.append("join vendor on vendor.vendor_id = to_number(inv_dtl.vendor_nbr) ");
         sql.append("join flc on flc.flc_id = inv_dtl.flc ");
         sql.append("where sale_type = 'WAREHOUSE' and vendor_nbr = ? and flc = ? and ");
         sql.append(getPeriod());            
         sql.append("group by vendor_id, vendor.name, flc, description, cust_nbr");
                  
         return sql.toString();
      }
      
      /**
       * Creates the captions for the vendor/flc filter.
       * 
       * @see SubRpt#createCaptions(int rowNum)
       */
      public int createCaptions(int rowNum)
      {
         StringBuffer caption = new StringBuffer();
         XSSFRow row = null;         
         int col = 0;
         
         caption.append("Select Vendor FLC:  Time Frame: ");
         caption.append(getPeriodCaption());
         
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSTitle);
         createCaptionCell(row, col, caption.toString());
         
         rowNum++;
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSCaption);
         createCaptionCell(row, col, "Vendor Name");
         m_Sheet.setColumnWidth(col++, CW_VND_NAME);
         createCaptionCell(row, col, "Vendor ID");
         m_Sheet.setColumnWidth(col++, CW_VND_ID);
         createCaptionCell(row, col, "FLC & Name");
         m_Sheet.setColumnWidth(col++, CW_FLC_NAME);
         createCaptionCell(row, col, "Customer ID");
         m_Sheet.setColumnWidth(col++, CW_CUST_ID);
         createCaptionCell(row, col, "Customer Name");
         m_Sheet.setColumnWidth(col++, CW_CUST_NAME);
         createCaptionCell(row, col, "Address 1");
         m_Sheet.setColumnWidth(col++, CW_ADDR1);
         createCaptionCell(row, col, "Address 2");
         m_Sheet.setColumnWidth(col++, CW_ADDR2);
         createCaptionCell(row, col, "City");
         m_Sheet.setColumnWidth(col++, CW_CITY);
         createCaptionCell(row, col++, "State");
         createCaptionCell(row, col++, "Zip");
         createCaptionCell(row, col++, "Phone");
         createCaptionCell(row, col++, "Fax");
         createCaptionCell(row, col, "Units");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Lines");
         m_Sheet.setColumnWidth(col++, CW_LINES);
         createCaptionCell(row, col, "Dollars");
         m_Sheet.setColumnWidth(col++, CW_SELL);
                  
         createCaptionCell(row, col, "Cost");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin%");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Portland");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Pittston");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         return rowNum;
      }
      
      /**
       * Initializes the cell styles array.
       */
      @Override
      protected void initCellStyles() 
      {
         m_CellStyles = new XSSFCellStyle[] {
            m_CSText,    // col 0 vnd name
            m_CSText,    // col 1 vnd id
            m_CSText,    // col 2 FLC# and name
            m_CSText,    // col 3 cust#
            m_CSText,    // col 4 cust name
            m_CSText,    // col 5 addr1
            m_CSText,    // col 6 addr2
            m_CSText,    // col 7 city
            m_CSText,    // col 8 state
            m_CSText,    // col 9 zip
            m_CSText,    // col 10 phone
            m_CSText,    // col 11 fax
            m_CSInt,     // col 12 units
            m_CSInt,     // col 13 lines
            m_CSMoney,   // col 14 dollars
            
            m_CSMoney,   // col 15 cost
            m_CSPct,     // col 16 margin%
            m_CSMoney,   // col 17 margin$
            
            m_CSInt,     // col 18 units portland
            m_CSMoney,   // col 19 dollars portland
            m_CSMoney,   // col 20 cost portland
            m_CSPct,     // col 21 margin% portland
            m_CSMoney,   // col 22 margin$ portland
            
            m_CSInt,     // col 23 units pittston
            m_CSMoney,   // col 24 dollars pittston
            m_CSMoney,   // col 25 cost pittston
            m_CSPct,     // col 26 margin% pittston
            m_CSMoney    // col 27 margin$ pittston
         };
      }
   }
   
   
   
   /**
    * The FLC filter sub report class
    */
   public class UDSCFlc extends SubRpt
   {
      /**
       * Overridden constructor that initializes the sub report.
       */
      public UDSCFlc() 
      {
         super();
      }

      /**
       * Builds the report file.  Gets called from the buildOutputFile method in the main class.
       * @see com.emerywaterhouse.rpt.spreadsheet.SubRpt#build(java.io.FileOutputStream, int)
       */
      @Override
      public boolean build(FileOutputStream out, int rowNum)
      {         
         boolean processed = false;
         ResultSet itemSales = null;
         XSSFRow row = null;
         String custId = null;
         int col;
         double sold = 0.0;
         double cost = 0.0;
         double margin = 0.0;
         
         try {
            m_ItemSales = m_EdbConn.prepareStatement(buildSql());            
            m_ItemSales.setString(1, m_FlcId);
            
            itemSales = m_ItemSales.executeQuery();
   
            while ( itemSales.next() && m_Status == RptServer.RUNNING ) {
               row = createDataRow(rowNum++);
               custId = itemSales.getString("cust_nbr");
               m_VndId = itemSales.getInt("vendor_nbr");
                              
               setCurAction("processing customer: " + custId);
               getCustDat(custId);
               col = 0;
               
               row.getCell(col).setCellValue(new XSSFRichTextString(m_FlcId));
               row.getCell(++col).setCellValue(new XSSFRichTextString(itemSales.getString("vendor_name")));
               row.getCell(++col).setCellValue(m_VndId);               
               row.getCell(++col).setCellValue(new XSSFRichTextString(custId));               
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_NAME]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR1]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR2]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_CITY]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_STATE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ZIP]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_PHONE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_FAX]));
               row.getCell(++col).setCellValue(itemSales.getInt("extqty"));
               row.getCell(++col).setCellValue(itemSales.getInt("lines"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extsell"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extcost"));
               
               sold = itemSales.getDouble("extsell");
               cost = itemSales.getDouble("extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);

               // Portland numbers
               sold = itemSales.getDouble("portland_extsell");
               cost = itemSales.getDouble("portland_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("portland_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
               
               // Pittston numbers
               sold = itemSales.getDouble("pittston_extsell");
               cost = itemSales.getDouble("pittston_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("pittston_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
            }
            
            processed = true;
         }
         
         catch ( Exception ex ) {
            m_ErrMsg.append(ex.getClass().getName() + "\r\n");
            m_ErrMsg.append(ex.getMessage());
            
            log.fatal("[UnitDollarSalesCust] ", ex);
         }
         
         finally {
            closeRSet(itemSales);
         }
         
         return processed;
      }
      
      /**
       * Builds the sql for the flc filter.
       * 
       * @see SubRpt#buildSql()
       */
      public String buildSql()
      {
         StringBuffer sql = new StringBuffer();
         
         sql.append("select flc, vendor_nbr, vendor_name, ");
         sql.append("cust_nbr, sum(qty_shipped) extqty, sum(ext_sell) extsell, sum(ext_cost) extcost, count(*) lines, ");
         
         // This was the most efficient way to get at this data from within the same query.  If we add another
         // warehouse in future, which is unlikely, it will have to be added to here, if that ever happens.
         sql.append("sum(decode(warehouse, 'PORTLAND', qty_shipped, 0)) as portland_extqty, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_sell, 0)) as portland_extsell, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_cost, 0)) as portland_extcost, ");
         sql.append("sum(decode(warehouse, 'PITTSTON', qty_shipped, 0)) as pittston_extqty, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_sell, 0)) as pittston_extsell, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_cost, 0)) as pittston_extcost ");         
         sql.append("from inv_dtl ");
         sql.append("where sale_type = 'WAREHOUSE' and flc = ? and ");
         sql.append(getPeriod());                  
         sql.append("group by flc, vendor_nbr, vendor_name, cust_nbr");
                           
         return sql.toString();
      }
      
      /**
       * Creates the captions for the FLC filter.
       * 
       * @see SubRpt#createCaptions(int rowNum)
       */
      public int createCaptions(int rowNum)
      {
         StringBuffer caption = new StringBuffer();
         XSSFRow row = null;         
         int col = 0;
         
         caption.append("Select FLC:  Time Frame: ");
         caption.append(getPeriodCaption());
         
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSTitle);
         createCaptionCell(row, col, caption.toString());
         
         rowNum++;
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSCaption);
         
         createCaptionCell(row, col++, "FLC");         
         createCaptionCell(row, col, "Vendor Name");
         m_Sheet.setColumnWidth(col++, CW_VND_NAME);
         createCaptionCell(row, col, "Vendor ID");
         m_Sheet.setColumnWidth(col++, CW_VND_ID);         
         createCaptionCell(row, col, "Customer ID");
         m_Sheet.setColumnWidth(col++, CW_CUST_ID);
         createCaptionCell(row, col, "Customer Name");
         m_Sheet.setColumnWidth(col++, CW_CUST_NAME);
         createCaptionCell(row, col, "Address 1");
         m_Sheet.setColumnWidth(col++, CW_ADDR1);
         createCaptionCell(row, col, "Address 2");
         m_Sheet.setColumnWidth(col++, CW_ADDR2);
         createCaptionCell(row, col, "City");
         m_Sheet.setColumnWidth(col++, CW_CITY);
         createCaptionCell(row, col++, "State");
         createCaptionCell(row, col++, "Zip");
         createCaptionCell(row, col++, "Phone");
         createCaptionCell(row, col++, "Fax");
         createCaptionCell(row, col, "Units");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Lines");
         m_Sheet.setColumnWidth(col++, CW_LINES);
         createCaptionCell(row, col, "Dollars");
         m_Sheet.setColumnWidth(col++, CW_SELL);
                  
         createCaptionCell(row, col, "Cost");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin%");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Portland");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Pittston");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         return rowNum;
      }
      
      /**
       * Initializes the cell styles array.
       */
      protected void initCellStyles() 
      {
         m_CellStyles = new XSSFCellStyle[] {
            m_CSText,    // col 0 FLC#
            m_CSText,    // col 1 vnd name
            m_CSText,    // col 2 vnd id
            m_CSText,    // col 3 cust#
            m_CSText,    // col 4 cust name
            m_CSText,    // col 5 addr1
            m_CSText,    // col 6 addr2
            m_CSText,    // col 7 city
            m_CSText,    // col 8 state
            m_CSText,    // col 9 zip
            m_CSText,    // col 10 phone
            m_CSText,    // col 11 fax
            m_CSInt,     // col 12 units
            m_CSInt,     // col 13 lines
            m_CSMoney,   // col 14 dollars
            
            m_CSMoney,   // col 15 cost
            m_CSPct,     // col 16 margin%
            m_CSMoney,   // col 17 margin$
            
            m_CSInt,     // col 18 units portland
            m_CSMoney,   // col 19 dollars portland
            m_CSMoney,   // col 20 cost portland
            m_CSPct,     // col 21 margin% portland
            m_CSMoney,   // col 22 margin$ portland
            
            m_CSInt,     // col 23 units pittston
            m_CSMoney,   // col 24 dollars pittston
            m_CSMoney,   // col 25 cost pittston
            m_CSPct,     // col 26 margin% pittston
            m_CSMoney    // col 27 margin$ pittston
         };
      }
   }
   
   
   
   /**
    * The item filter sub report class
    *
    */
   public class UDSCItem extends SubRpt
   {
      public UDSCItem()
      {
         super();
      }
      
      /**
       * Overridden constructor that initializes the sub report.
       * @param wrkbk The Workbook object
       * @param sheet The sheet object
       */
      public UDSCItem(XSSFWorkbook wrkbk, XSSFSheet sheet) 
      {
         super();
      }

      /**
       * @see com.emerywaterhouse.rpt.spreadsheet.SubRpt#build(java.io.FileOutputStream, int)
       */
      @Override
      public boolean build(FileOutputStream out, int rowNum)
      {         
         boolean processed = false;
         ResultSet itemSales = null;
         XSSFRow row = null;
         String custId = null;
         int col;
         double sold = 0.0;
         double cost = 0.0;
         double margin = 0.0;
         
         try {
            m_ItemSales = m_EdbConn.prepareStatement(buildSql());
            m_ItemSales.setString(1, m_ItemId);
            
            itemSales = m_ItemSales.executeQuery();
   
            while ( itemSales.next() && m_Status == RptServer.RUNNING ) {
               row = createDataRow(rowNum++);
               custId = itemSales.getString("cust_nbr");
               m_VndId = itemSales.getInt("vendor_nbr");
               m_FlcId = itemSales.getString("flc");
                              
               setCurAction("processing customer: " + custId);
               getCustDat(custId);
               col = 0;
               
               row.getCell(col).setCellValue(new XSSFRichTextString(itemSales.getString("vendor_name")));
               row.getCell(++col).setCellValue(m_VndId);
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_ItemId));               
               row.getCell(++col).setCellValue(new XSSFRichTextString(itemSales.getString("description")));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_FlcId));
               row.getCell(++col).setCellValue(new XSSFRichTextString(custId));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_NAME]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR1]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR2]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_CITY]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_STATE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ZIP]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_PHONE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_FAX]));
               row.getCell(++col).setCellValue(itemSales.getInt("extqty"));
               row.getCell(++col).setCellValue(itemSales.getInt("lines"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extsell"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extcost"));
               
               sold = itemSales.getDouble("extsell");
               cost = itemSales.getDouble("extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);

               // Portland numbers
               sold = itemSales.getDouble("portland_extsell");
               cost = itemSales.getDouble("portland_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("portland_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
               
               // Pittston numbers
               sold = itemSales.getDouble("pittston_extsell");
               cost = itemSales.getDouble("pittston_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("pittston_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
            }
            
            processed = true;
         }
         
         catch ( Exception ex ) {
            m_ErrMsg.append(ex.getClass().getName() + "\r\n");
            m_ErrMsg.append(ex.getMessage());
            
            log.fatal("[UnitDollarSalesCust] ", ex);
         }
         
         finally {
            closeRSet(itemSales);
         }
         
         return processed;
      }
      
      /**
       * Builds the sql for the item filter.
       * 
       * @see SubRpt#buildSql()
       */
      public String buildSql()
      {
         StringBuffer sql = new StringBuffer();
         
         sql.append("select vendor_nbr, vendor_name, item_entity_attr.item_id, item_entity_attr.description, ");
         sql.append("flc, cust_nbr, sum(qty_shipped) as extqty, sum(ext_sell) as extsell, ");
         sql.append("sum(ext_cost) as extcost, count(*) as lines, ");
         
         // This was the most efficient way to get at this data from within the same query.  If we add another
         // warehouse in future, which is unlikely, it will have to be added to here, if that ever happens.
         sql.append("sum(decode(warehouse, 'PORTLAND', qty_shipped, 0)) portland_extqty, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_sell, 0)) portland_extsell, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_cost, 0)) portland_extcost, ");
         sql.append("sum(decode(warehouse, 'PITTSTON', qty_shipped, 0)) pittston_extqty, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_sell, 0)) pittston_extsell, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_cost, 0)) pittston_extcost ");         
         sql.append("from inv_dtl ");
         sql.append("join item_entity_attr on item_entity_attr.item_ea_id = inv_dtl.item_ea_id ");
         sql.append("where sale_type = 'WAREHOUSE' and item_nbr = ? and ");         
         sql.append(getPeriod());         
         sql.append("group by vendor_nbr, vendor_name, item_entity_attr.item_id, item_entity_attr.description, flc, cust_nbr");
                  
         return sql.toString();
      }
      
      /**
       * Creates the captions for the item filter.
       * 
       * @see SubRpt#createCaptions(int rowNum)
       */
      public int createCaptions(int rowNum)
      {
         StringBuffer caption = new StringBuffer();
         XSSFRow row = null;
         int col = 0;
         
         caption.append("Select Item:  Time Frame: ");
         caption.append(getPeriodCaption());
         
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSTitle);
         createCaptionCell(row, col, caption.toString());
         
         rowNum++;
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSCaption);
                           
         createCaptionCell(row, col, "Vendor Name");
         m_Sheet.setColumnWidth(col++, CW_VND_NAME);
         createCaptionCell(row, col, "Vendor ID");
         m_Sheet.setColumnWidth(col++, CW_VND_ID);
         createCaptionCell(row, col++, "Item");
         createCaptionCell(row, col, "Description");
         m_Sheet.setColumnWidth(col++, CW_DESC);
         createCaptionCell(row, col++, "FLC");
         createCaptionCell(row, col, "Customer ID");
         m_Sheet.setColumnWidth(col++, CW_CUST_ID);
         createCaptionCell(row, col, "Customer Name");
         m_Sheet.setColumnWidth(col++, CW_CUST_NAME);
         createCaptionCell(row, col, "Address 1");
         m_Sheet.setColumnWidth(col++, CW_ADDR1);
         createCaptionCell(row, col, "Address 2");
         m_Sheet.setColumnWidth(col++, CW_ADDR2);
         createCaptionCell(row, col, "City");
         m_Sheet.setColumnWidth(col++, CW_CITY);
         createCaptionCell(row, col++, "State");
         createCaptionCell(row, col++, "Zip");
         createCaptionCell(row, col++, "Phone");
         createCaptionCell(row, col++, "Fax");
         createCaptionCell(row, col, "Units");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Lines");
         m_Sheet.setColumnWidth(col++, CW_LINES);
         createCaptionCell(row, col, "Dollars");
         m_Sheet.setColumnWidth(col++, CW_SELL);
                  
         createCaptionCell(row, col, "Cost");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin%");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Portland");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Pittston");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         return rowNum;
      }
      
      /**
       * Initializes the cell styles array.
       */
      protected void initCellStyles() 
      {
         m_CellStyles = new XSSFCellStyle[] {            
            m_CSText,    // col 0 vnd name
            m_CSText,    // col 1 vnd id
            m_CSText,    // col 2 item id
            m_CSText,    // col 3 item desc
            m_CSText,    // col 4 FLC#
            m_CSText,    // col 5 cust#
            m_CSText,    // col 6 cust name
            m_CSText,    // col 7 addr1
            m_CSText,    // col 8 addr2
            m_CSText,    // col 9 city
            m_CSText,    // col 10 state
            m_CSText,    // col 11 zip
            m_CSText,    // col 12 phone
            m_CSText,    // col 13 fax
            m_CSInt,     // col 14 units
            m_CSInt,     // col 14 lines
            m_CSMoney,   // col 15 dollars
            
            m_CSMoney,   // col 16 cost
            m_CSPct,     // col 17 margin%
            m_CSMoney,   // col 18 margin$
            
            m_CSInt,     // col 19 units portland
            m_CSMoney,   // col 20 dollars portland
            m_CSMoney,   // col 21 cost portland
            m_CSPct,     // col 22 margin% portland
            m_CSMoney,   // col 23 margin$ portland
            
            m_CSInt,     // col 24 units pittston
            m_CSMoney,   // col 25 dollars pittston
            m_CSMoney,   // col 26 cost pittston
            m_CSPct,     // col 27 margin% pittston
            m_CSMoney    // col 28 margin$ pittston
         };
      }
   }
   
   
   
   /**
    * The RMS summary sub report
    */
   public class UDSCRmsSum extends SubRpt
   {
      /**
       * Overridden constructor that initializes the sub report.       
       */
      public UDSCRmsSum() 
      {
         super();
      }

      /**
       * @see com.emerywaterhouse.rpt.spreadsheet.SubRpt#build(java.io.FileOutputStream, int)
       */
      @Override
      public boolean build(FileOutputStream out, int rowNum)
      {         
         boolean processed = false;
         ResultSet itemSales = null;
         XSSFRow row = null;
         String custId = null;
         int col;
         double sold = 0.0;
         double cost = 0.0;
         double margin = 0.0;
         
         try {
            m_ItemSales = m_EdbConn.prepareStatement(buildSql());
            itemSales = m_ItemSales.executeQuery();
   
            while ( itemSales.next() && m_Status == RptServer.RUNNING ) {
               row = createDataRow(rowNum++);
               custId = itemSales.getString("cust_nbr");
                              
               setCurAction("processing customer: " + custId);
               getCustDat(custId);
               col = 0;
               
               row.getCell(col).setCellValue(new XSSFRichTextString(itemSales.getString("rms_id")));
               
               row.getCell(++col).setCellValue(new XSSFRichTextString(custId));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_NAME]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR1]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR2]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_CITY]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_STATE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ZIP]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_PHONE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_FAX]));
               row.getCell(++col).setCellValue(itemSales.getInt("extqty"));
               row.getCell(++col).setCellValue(itemSales.getInt("lines"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extsell"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extcost"));
               
               sold = itemSales.getDouble("extsell");
               cost = itemSales.getDouble("extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);

               // Portland numbers
               sold = itemSales.getDouble("portland_extsell");
               cost = itemSales.getDouble("portland_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("portland_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
               
               // Pittston numbers
               sold = itemSales.getDouble("pittston_extsell");
               cost = itemSales.getDouble("pittston_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("pittston_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
            }
            
            processed = true;
         }
         
         catch ( Exception ex ) {
            m_ErrMsg.append(ex.getClass().getName() + "\r\n");
            m_ErrMsg.append(ex.getMessage());
            
            log.fatal("[UnitDollarSalesCust] ", ex);
         }
         
         finally {
            closeRSet(itemSales);
         }
         
         return processed;
      }
      
      /**
       * Builds the sql for the rms summary filter.
       * 
       * @see SubRpt#buildSql()
       */
      public String buildSql()
      {
         StringBuffer sql = new StringBuffer();
                  
         sql.append("select rms_id, cust_nbr,  ");
         sql.append("sum(qty_shipped) extqty, sum(ext_sell) extsell, sum(ext_cost) extcost, count(*) lines, ");
         
         // This was the most efficient way to get at this data from within the same query.  If we add another
         // warehouse in future, which is unlikely, it will have to be added to here, if that ever happens.
         sql.append("sum(decode(warehouse, 'PORTLAND', qty_shipped, 0)) portland_extqty, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_sell, 0)) portland_extsell, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_cost, 0)) portland_extcost, ");
         sql.append("sum(decode(warehouse, 'PITTSTON', qty_shipped, 0)) pittston_extqty, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_sell, 0)) pittston_extsell, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_cost, 0)) pittston_extcost ");
         sql.append("from inv_dtl ");
         sql.append("join rms_item on rms_item.item_ea_id = inv_dtl.item_ea_id ");
         sql.append(String.format("where sale_type = 'WAREHOUSE' and rms_id in (%s) and ", m_RmsIds));         
         sql.append(getPeriod());
         sql.append("group by rms_id, cust_nbr");
                  
         return sql.toString();
      }
      
      /**
       * Creates the captions for the RMS summary filter.
       * 
       * @see SubRpt#createCaptions(int rowNum)
       */
      public int createCaptions(int rowNum)
      {
         StringBuffer caption = new StringBuffer();
         XSSFRow row = null;         
         int col = 0;
         
         caption.append("Select RMS Summary:  Time Frame: ");
         caption.append(getPeriodCaption());
         
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSTitle);
         createCaptionCell(row, col, caption.toString());
         
         rowNum++;
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSCaption);
         
         createCaptionCell(row, col++, "RMS ID");
         createCaptionCell(row, col, "Customer ID");
         m_Sheet.setColumnWidth(col++, CW_CUST_ID);
         createCaptionCell(row, col, "Customer Name");
         m_Sheet.setColumnWidth(col++, CW_CUST_NAME);
         createCaptionCell(row, col, "Address 1");
         m_Sheet.setColumnWidth(col++, CW_ADDR1);
         createCaptionCell(row, col, "Address 2");
         m_Sheet.setColumnWidth(col++, CW_ADDR2);
         createCaptionCell(row, col, "City");
         m_Sheet.setColumnWidth(col++, CW_CITY);
         createCaptionCell(row, col++, "State");
         createCaptionCell(row, col++, "Zip");
         createCaptionCell(row, col++, "Phone");
         createCaptionCell(row, col++, "Fax");
         createCaptionCell(row, col, "Units");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Lines");
         m_Sheet.setColumnWidth(col++, CW_LINES);
         createCaptionCell(row, col, "Dollars");
         m_Sheet.setColumnWidth(col++, CW_SELL);
                  
         createCaptionCell(row, col, "Cost");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin%");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Portland");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Pittston");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         return rowNum;
      }
      
      /**
       * Initializes the cell styles array.
       */
      protected void initCellStyles() 
      {
         m_CellStyles = new XSSFCellStyle[] {            
            m_CSText,    // col 0 rms id            
            m_CSText,    // col 1 cust#
            m_CSText,    // col 2 cust name
            m_CSText,    // col 3 addr1
            m_CSText,    // col 4 addr2
            m_CSText,    // col 5 city
            m_CSText,    // col 6 state
            m_CSText,    // col 7 zip
            m_CSText,    // col 8 phone
            m_CSText,    // col 9 fax
            m_CSInt,     // col 10 units
            m_CSInt,     // col 11 lines
            m_CSMoney,   // col 12 dollars
            
            m_CSMoney,   // col 13 cost
            m_CSPct,     // col 14 margin%
            m_CSMoney,   // col 15 margin$
            
            m_CSInt,     // col 16 units portland
            m_CSMoney,   // col 17 dollars portland
            m_CSMoney,   // col 18 cost portland
            m_CSPct,     // col 19 margin% portland
            m_CSMoney,   // col 20 margin$ portland
            
            m_CSInt,     // col 21 units pittston
            m_CSMoney,   // col 22 dollars pittston
            m_CSMoney,   // col 23 cost pittston
            m_CSPct,     // col 24 margin% pittston
            m_CSMoney    // col 25 margin$ pittston
         };
      }
   }
   
   
   
   /**
    * Creates teh RMS detail sub report based on the rms detail filter.
    */
   public class UDSCRmsDtl extends SubRpt
   {
      /**
       * Overridden constructor that initializes the sub report.       
       */
      public UDSCRmsDtl() 
      {
         super();
      }

      /**
       * @see com.emerywaterhouse.rpt.spreadsheet.SubRpt#build(java.io.FileOutputStream, int)
       */
      @Override
      public boolean build(FileOutputStream out, int rowNum)
      {         
         boolean processed = false;
         ResultSet itemSales = null;
         XSSFRow row = null;
         String custId = null;
         int col;
         double sold = 0.0;
         double cost = 0.0;
         double margin = 0.0;
         
         try {
            m_ItemSales = m_EdbConn.prepareStatement(buildSql());
            itemSales = m_ItemSales.executeQuery();
   
            while ( itemSales.next() && m_Status == RptServer.RUNNING ) {
               row = createDataRow(rowNum++);
               custId = itemSales.getString("cust_nbr");
               m_FlcId = itemSales.getString("flc");
               m_ItemId = itemSales.getString("item_id");
               
               setCurAction("processing customer: " + custId);
               getCustDat(custId);
               col = 0;
               
               row.getCell(col).setCellValue(new org.apache.poi.xssf.usermodel.XSSFRichTextString(itemSales.getString("rms_id")));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_ItemId));
               row.getCell(++col).setCellValue(new XSSFRichTextString(itemSales.getString("description")));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_FlcId));
               row.getCell(++col).setCellValue(new XSSFRichTextString(custId));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_NAME]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR1]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ADDR2]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_CITY]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_STATE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_ZIP]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_PHONE]));
               row.getCell(++col).setCellValue(new XSSFRichTextString(m_CustDat[C_FAX]));
               row.getCell(++col).setCellValue(itemSales.getInt("extqty"));
               row.getCell(++col).setCellValue(itemSales.getInt("lines"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extsell"));
               row.getCell(++col).setCellValue(itemSales.getDouble("extcost"));
               
               sold = itemSales.getDouble("extsell");
               cost = itemSales.getDouble("extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);

               // Portland numbers
               sold = itemSales.getDouble("portland_extsell");
               cost = itemSales.getDouble("portland_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("portland_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
               
               // Pittston numbers
               sold = itemSales.getDouble("pittston_extsell");
               cost = itemSales.getDouble("pittston_extcost");
               margin = sold - cost;
               row.getCell(++col).setCellValue(itemSales.getInt("pittston_extqty"));
               row.getCell(++col).setCellValue(sold);
               row.getCell(++col).setCellValue(cost);
               row.getCell(++col).setCellValue(sold > 0? margin/sold: 0.0);
               row.getCell(++col).setCellValue(margin);
            }
            
            processed = true;
         }
         
         catch ( Exception ex ) {
            m_ErrMsg.append(ex.getClass().getName() + "\r\n");
            m_ErrMsg.append(ex.getMessage());
            
            log.fatal("fatal exception: ", ex);
         }
         
         finally {
            closeRSet(itemSales);
         }
         
         return processed;
      }
      
      /**
       * Builds the sql for the rms detail filter.
       * 
       * @see SubRpt#buildSql()
       */
      public String buildSql()
      {
         StringBuffer sql = new StringBuffer();
                 
         sql.append("select rms_id, item_entity_attr.item_id, item_entity_attr.description, flc, cust_nbr, ");
         sql.append("sum(qty_shipped) as extqty, sum(ext_sell) as extsell, ");
         sql.append("sum(ext_cost) as extcost, count(*) as lines, ");
         
         // This was the most efficient way to get at this data from within the same query.  If we add another
         // warehouse in future, which is unlikely, it will have to be added to here, if that ever happens.
         sql.append("sum(decode(warehouse, 'PORTLAND', qty_shipped, 0)) portland_extqty, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_sell, 0)) portland_extsell, "); 
         sql.append("sum(decode(warehouse, 'PORTLAND', ext_cost, 0)) portland_extcost, ");
         sql.append("sum(decode(warehouse, 'PITTSTON', qty_shipped, 0)) pittston_extqty, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_sell, 0)) pittston_extsell, "); 
         sql.append("sum(decode(warehouse, 'PITTSTON', ext_cost, 0)) pittston_extcost ");         
         sql.append("from inv_dtl ");
         sql.append("join rms_item on rms_item.item_ea_id = inv_dtl.item_ea_id " );
         sql.append("join item_entity_attr on item_entity_attr.item_ea_id = inv_dtl.item_ea_id " );
         sql.append(String.format("where sale_type = 'WAREHOUSE' and rms_id in (%s) and ", m_RmsIds));         
         sql.append(getPeriod());         
         sql.append("group by rms_id, item_entity_attr.item_id, item_entity_attr.description, flc, cust_nbr");
         
         return sql.toString();
      }
      
      /**
       * Creates the captions for the RMS detail filter.
       * 
       * @see SubRpt#createCaptions(int rowNum)
       */
      public int createCaptions(int rowNum)
      {
         StringBuffer caption = new StringBuffer();
         XSSFRow row = null;         
         int col = 0;
         
         caption.append("Select RMS Detail:  Time Frame: ");
         caption.append(getPeriodCaption());
         
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSTitle);
         createCaptionCell(row, col, caption.toString());
         
         rowNum++;
         row = m_Sheet.createRow(rowNum);
         setCaptionStyle(m_CSCaption);
                           
         createCaptionCell(row, col++, "RMS ID");
         createCaptionCell(row, col++, "Item");
         createCaptionCell(row, col, "Description");
         m_Sheet.setColumnWidth(col++, CW_DESC);
         createCaptionCell(row, col++, "FLC");
         createCaptionCell(row, col, "Customer ID");
         m_Sheet.setColumnWidth(col++, CW_CUST_ID);
         createCaptionCell(row, col, "Customer Name");
         m_Sheet.setColumnWidth(col++, CW_CUST_NAME);
         createCaptionCell(row, col, "Address 1");
         m_Sheet.setColumnWidth(col++, CW_ADDR1);
         createCaptionCell(row, col, "Address 2");
         m_Sheet.setColumnWidth(col++, CW_ADDR2);
         createCaptionCell(row, col, "City");
         m_Sheet.setColumnWidth(col++, CW_CITY);
         createCaptionCell(row, col++, "State");
         createCaptionCell(row, col++, "Zip");
         createCaptionCell(row, col++, "Phone");
         createCaptionCell(row, col++, "Fax");
         createCaptionCell(row, col, "Units");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Lines");
         m_Sheet.setColumnWidth(col++, CW_LINES);
         createCaptionCell(row, col, "Dollars");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Cost");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin%");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Portland");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Portland");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         createCaptionCell(row, col, "Units Pittston");
         m_Sheet.setColumnWidth(col++, CW_UNITS);
         createCaptionCell(row, col, "Dollars Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Cost Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin% Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         createCaptionCell(row, col, "Margin$ Pittston");
         m_Sheet.setColumnWidth(col++, CW_SELL);
         
         return rowNum;
      }
      
      /**
       * Initializes the cell styles array.
       */
      protected void initCellStyles() 
      {
         m_CellStyles = new XSSFCellStyle[] {            
            m_CSText,    // col 0 rms id
            m_CSText,    // col 1 item id
            m_CSText,    // col 2 item desc
            m_CSText,    // col 3 FLC#
            m_CSText,    // col 4 cust#
            m_CSText,    // col 5 cust name
            m_CSText,    // col 6 addr1
            m_CSText,    // col 7 addr2
            m_CSText,    // col 8 city
            m_CSText,    // col 9 state
            m_CSText,    // col 10 zip
            m_CSText,    // col 11 phone
            m_CSText,    // col 12 fax
            m_CSInt,     // col 13 units
            m_CSInt,     // col 14 lines
            m_CSMoney,   // col 15 dollars
            
            m_CSMoney,   // col 16 cost
            m_CSPct,     // col 17 margin%
            m_CSMoney,   // col 18 margin$
            
            m_CSInt,     // col 19 units portland
            m_CSMoney,   // col 20 dollars portland
            m_CSMoney,   // col 21 cost portland
            m_CSPct,     // col 22 margin% portland
            m_CSMoney,   // col 23 margin$ portland
            
            m_CSInt,     // col 24 units pittston
            m_CSMoney,   // col 25 dollars pittston
            m_CSMoney,   // col 26 cost pittston
            m_CSPct,     // col 27 margin% pittston
            m_CSMoney    // col 28 margin$ pittston
         };
      }
   }
}
