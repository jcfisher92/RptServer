/**
 * File: CustProfitability.java
 * Description: Report showing customer profitability based on different factors.  Also determines rebates.
 *
 * @author Jeff Fisher
 *
 * Create Date: 11/01/2009
 * Last Update: $Id: CustProfitability.java,v 1.3 2010/06/02 19:42:20 jfisher Exp $
 *
 * History:
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.ByteArrayInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.ObjectInputStream;
import java.lang.reflect.Method;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Iterator;

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

import com.emerywaterhouse.ora.Rebate;
import com.emerywaterhouse.ora.RebateProcs;
import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class CustProfitability extends Report
{
   private static final short MAX_COLS = 35;
   private final long effSalesMin  = 500000;
   
   private final String effRebName     = "EFFICIENCY";
   private final String vndRebDa       = "Distribution America";
   private final String vndRebMonth    = "Monthly";
   private final String vndRebQtr      = "Quarterly";
   private final String vndRebGrowth   = "Growth";
      
   private final int effRebColCnt  = 9;
   private final int qtrRebColCnt  = 7;
    
   private PreparedStatement m_BensonReb;
   private PreparedStatement m_CotFrtReb;
   private PreparedStatement m_CustData;   
   private PreparedStatement m_EffRebData;
   private PreparedStatement m_HarBaseReb;
   private PreparedStatement m_HarRecReb;
   private PreparedStatement m_HasRebate;
   
   private PreparedStatement m_PrevYrSales;
   private PreparedStatement m_QrtRebates;
   private PreparedStatement m_Rebate;
   private PreparedStatement m_VndReb;
   private PreparedStatement m_VndRebSales;
   
   private ResultSet m_QrtRebData;
   
   private ArrayList<String> m_ProcData;
     
   //
   // Params
   private java.sql.Date m_BegDate;
   private java.sql.Date m_EndDate;
   private GregorianCalendar m_BegCal;
   private GregorianCalendar m_EndCal;
     
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;
   
   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;
   private XSSFSheet m_Sheet2;
   private XSSFSheet m_Sheet3;
   private int m_RowNum2;
   private int m_RowNum3;
        
   /**
    * Default constructor.
    */
   public CustProfitability()
   {
      super();
      
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet("Customer Data");
      m_Sheet2 = m_Wrkbk.createSheet("Efficiency Detail");
      m_Sheet3 = m_Wrkbk.createSheet("Quarterly Detail");   
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
      
      if ( m_ProcData != null )
         m_ProcData.clear();
      
      m_Sheet = null;
      m_Sheet2 = null;
      m_Sheet3 = null;      
      m_Wrkbk = null;      
      m_CellStyles = null;
      m_BegDate = null;
      m_EndDate = null;
      m_BegCal = null;
      m_EndCal = null;
      m_ProcData = null;
      
      m_BensonReb = null;
      m_CotFrtReb = null;
      m_CustData = null;      
      m_EffRebData = null;
      m_HasRebate = null;
      m_VndRebSales = null;
      m_PrevYrSales = null;
      m_QrtRebates = null;
      m_QrtRebData = null;      
      m_Rebate = null;
      m_VndReb = null;
      
      super.finalize();
   }
      
   /**
    * Adds a row to the rebate detail worksheet.
    * @param custId The current customer number with a rebate
    * @param rebSales 
    * @param compSales
    * @param prevSales
    * @param lineVal
    * @param creditTot
    * @param crPct
    */
   private void addRebDetail(String custId, double rebate, double rebSales, double compSales, double prevSales, double lineVal, double creditTot, double crPct)
   {
      XSSFCellStyle styleCaption;
      XSSFCellStyle styleTitle;
      XSSFFont font;
      XSSFCell cell = null;
      XSSFRow row = null;
      int colNum = 0;
      
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
      
      if ( m_Sheet2 != null ) {
         row = m_Sheet2.createRow(m_RowNum2++);
         
         //
         // set the type and style of the cell.
         if ( row != null ) {
            for ( int i = 0; i < effRebColCnt; i++ ) {
               cell = row.createCell(i);
               cell.setCellStyle(m_CellStyles[i]);
            }
            
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(custId));
            row.getCell(colNum++).setCellValue(rebate);            
            row.getCell(colNum++).setCellValue(rebSales);
            row.getCell(colNum++).setCellValue(compSales);
            row.getCell(colNum++).setCellValue(prevSales);
            row.getCell(colNum++).setCellValue(rebSales - effSalesMin);
            row.getCell(colNum++).setCellValue(lineVal);
            row.getCell(colNum++).setCellValue(creditTot);
            row.getCell(colNum++).setCellValue(crPct);
         }         
      }
      
      styleCaption = null;
      styleTitle = null;
      font = null;
      cell = null;
      row = null;
   }
   
   /**
    * Adds the detail data for the quarterly rebates.
    * 
    * @param r A reference to a rebate object.
    */
   private void addQtrRebDetail(Rebate r)
   {
      XSSFCellStyle style;
      XSSFFont font;
      XSSFCell cell = null;
      XSSFRow row = null;
      int colNum = 0;
      
      font = m_Wrkbk.createFont();
      font.setFontHeightInPoints((short)10);
      font.setFontName("Arial");
      
      style = m_Wrkbk.createCellStyle();      
      style.setFont(font);
      //style.setAlignment(HorizontalAlignment.CENTER);
      //TODO fix the cell alignment.
            
      try {
         if ( m_Sheet3 != null ) {
            row = m_Sheet3.createRow(m_RowNum3++);
            
            //
            // set the type and style of the cell.
            if ( row != null ) {
               for ( int i = 0; i < qtrRebColCnt; i++ ) {
                  cell = row.createCell(i);
                  cell.setCellStyle(style);
               }
               
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(r.custId));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(r.rebate));
               row.getCell(colNum++).setCellValue(r.q1Reb);
               row.getCell(colNum++).setCellValue(r.q2Reb);
               row.getCell(colNum++).setCellValue(r.q3Reb);
               row.getCell(colNum++).setCellValue(r.q4Reb);
               row.getCell(colNum++).setCellValue(r.rebateTot);
            }         
         }
      }
      
      finally {
         style = null;      
         font = null;
         cell = null;
         row = null;
      }
   }
   
   /**
    * Calculates the Benson base cost rebate.
    * @param custId the customer id
    * @return A reference to a Rebate object.
    */
   public Rebate bensonBaseRebate(String custId)
   {
      ResultSet rs = null;
      Rebate reb = null;
      
      try {
         m_BensonReb.setDate(1, m_BegDate);
         m_BensonReb.setDate(2, m_EndDate);
         
         rs = m_BensonReb.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Benson base cost";         
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
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
      int pageNum = 1;
      String msg = "processing page %d, row %d, customer %s";
      FileOutputStream outFile = null;      
      ResultSet custData = null;
      boolean result = false;      
      String custId = null;
      double coopReb = 0.0;
      double effReb = 0.0;
      double qrtReb = 0.0;
      double otherReb = 0.0;
      double totReb = 0.0;
      double cashDisc = 0.0;
      double grossSales = 0.0;
      double netSales = 0.0;
      double cogs = 0.0;
      double margin = 0.0;
      double daVndReb = 0.0;
      //double daVndRebTot = 0.0;
      double qtrlyVndReb = 0.0;
      //double qtrlyVndRebTot = 0.0;
      double monthlyVndReb = 0.0;
      //double monthlyVndRebTot = 0.0;
      double growthVndReb = 0.0;
      //double growthVndRebTot = 0.0;      
      double totVndReb = 0.0;
      double arTot = 0.0;
      double arPast = 0.0;
      double arPct = 0.0;
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      rowNum = createCaptions();
      
      try {
         //
         // Not sure if we'll need these totals.  The work is done if we need them.
         setCurAction("getting dafe rebate totals");
         //daVndRebTot = getVndReb(vndRebDa);
         setCurAction("getting monthly rebate totals");
         //monthlyVndRebTot = getVndReb(vndRebMonth);
         setCurAction("getting quarterly rebate totals");
         //qtrlyVndRebTot = getVndReb(vndRebQtr);
         setCurAction("getting growth rebate totals");
         //growthVndRebTot = getVndReb(vndRebGrowth);
         setCurAction("getting pittson rebate totals");
         //pittVndRebTot = getVndReb(vndRebPitt);
         
         setCurAction("getting customer rebate totals");
         m_CustData.setDate(1, m_BegDate);
         m_CustData.setDate(2, m_EndDate);
         m_CustData.setDate(3, m_BegDate);
         m_CustData.setDate(4, m_EndDate);         
         custData = m_CustData.executeQuery();

         while ( custData.next() && m_Status == RptServer.RUNNING ) {
            custId = custData.getString("customer_id");
            setCurAction(String.format(msg, pageNum, rowNum, custId));
      
            //
            // Get the data for the internal calculations
            grossSales = custData.getDouble("gsales");
            cogs = custData.getDouble("cogs");
            arTot = custData.getDouble("ar_tot");
            arPast = custData.getDouble("ar_past");
            
            //
            // Get the dealer rebates
            coopReb = getCoopRebate(custId);
            effReb = getEfficiencyRebate(custId);            
            qrtReb = getQuarterlyRebate(custId);
            otherReb = getOtherRebate(custId);
            totReb = coopReb + effReb + qrtReb + otherReb;
            
            //
            //
            netSales = grossSales - totReb - cashDisc;
            margin = grossSales - cogs;
                        
            arPct = netSales > 0 ? (arPast/netSales) : 0;
            
            //
            // Calculate the vendor rebate dollars associated with each customer.
            // This is dollar sales as a % of total rebate program.
            daVndReb = getVndRebSales(custId, vndRebDa);
            monthlyVndReb = getVndRebSales(custId, vndRebMonth);
            qtrlyVndReb = getVndRebSales(custId, vndRebQtr);
            growthVndReb = getVndRebSales(custId, vndRebGrowth);
            //TODO Add pittson in here
            totVndReb = daVndReb + monthlyVndReb + qtrlyVndReb + growthVndReb;
            
            
            row = createRow(rowNum++, MAX_COLS);
            colNum = 0;
            
            if ( row != null ) {
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(custId));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("cname")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("tm")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("city")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("state")));
               row.getCell(colNum++).setCellValue(custData.getInt("trips"));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("type")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("affiliation")));
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(custData.getString("warehouse")));
               row.getCell(colNum++).setCellValue(grossSales);
               row.getCell(colNum++).setCellValue(coopReb);
               row.getCell(colNum++).setCellValue(effReb);
               row.getCell(colNum++).setCellValue(qrtReb);
               row.getCell(colNum++).setCellValue(otherReb);
               row.getCell(colNum++).setCellValue(totReb);
               row.getCell(colNum++).setCellValue(cashDisc);
               row.getCell(colNum++).setCellValue(netSales);
               row.getCell(colNum++).setCellValue(cogs);
               row.getCell(colNum++).setCellValue(margin);
               row.getCell(colNum++).setCellValue(custData.getInt("orders"));
               row.getCell(colNum++).setCellValue(custData.getInt("lines"));
               row.getCell(colNum++).setCellValue(custData.getInt("units"));
               row.getCell(colNum++).setCellValue(custData.getDouble("cr_tot"));
               row.getCell(colNum++).setCellValue(custData.getInt("cr_code"));
               row.getCell(colNum++).setCellValue(custData.getInt("cr_mems"));
               row.getCell(colNum++).setCellValue(custData.getInt("cr_lines"));
               row.getCell(colNum++).setCellValue(custData.getInt("cr_units"));
               row.getCell(colNum++).setCellValue(arTot);
               row.getCell(colNum++).setCellValue(arPast);
               row.getCell(colNum++).setCellValue(arPct);
               row.getCell(colNum++).setCellValue(monthlyVndReb);
               row.getCell(colNum++).setCellValue(qtrlyVndReb);
               row.getCell(colNum++).setCellValue(growthVndReb);
               row.getCell(colNum++).setCellValue(daVndReb);
               row.getCell(colNum++).setCellValue(totVndReb); 
            }
            
            if ( rowNum > 65000 ) {
               m_Sheet.createFreezePane(1, 3);
               pageNum++;
               m_Sheet = m_Wrkbk.createSheet("Customer Data pg " + pageNum);
               rowNum = createCaptions();               
            }
         }
         
         m_Wrkbk.write(outFile);
         custData.close();

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
         custData = null;         
      }

      return result;
   }
  
   /**
    * Calls the stored procedure to get the busy beaver Stanley rebate.
    * 
    * @param custId not used in this method.
    * @return A reference to a Rebate object.
    */
   public Rebate busyBeaverOSIRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
      
      try {
         m_Rebate.setInt(1, RebateProcs.bbOsiReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
         
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Busy Beaver OSI";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Calls the stored procedure to get the busy beaver receiving rebate.
    * 
    * @param custId not used in this method.
    * @return A reference to a Rebate object.
    */
   public Rebate busyBeaverRecRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
            
      try {
         m_Rebate.setInt(1, RebateProcs.bbRecReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
      
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Busy Beaver Receiving";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Calls the stored procedure to get the busy beaver Stanley rebate.
    * 
    * @param custId not used in this method.
    * @return A reference to a Rebate object.
    */
   public Rebate busyBeaverStanleyRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
            
      try {
         m_Rebate.setInt(1, RebateProcs.bbStnReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
     
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Busy Beaver Stanley";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Calls the stored procedure to get the Busy Beaver vendor rebate
    * 
    * @param custId the busy beaver customer number.
    * @return A reference to a Rebate object.
    */
   public Rebate busyBeaverVndRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
            
      try {
         m_Rebate.setInt(1, RebateProcs.bbVndReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
         
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Busy Beaver Vendor";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Close all the open queries
    */
   private void closeStatements()
   {      
      closeStmt(m_BensonReb);
      closeStmt(m_CotFrtReb);
      closeStmt(m_CustData);
      closeStmt(m_VndRebSales);
      closeStmt(m_EffRebData);
      closeStmt(m_HarBaseReb);
      closeStmt(m_HarRecReb);
      closeStmt(m_HasRebate);      
      closeStmt(m_PrevYrSales);
      closeStmt(m_QrtRebates);
      closeRSet(m_QrtRebData);      
      closeStmt(m_Rebate);
      closeStmt(m_VndReb);
   }

   /**
    * Converts an object from the RAW output from the query to a java object
    * used by the rest of the program.
    * 
    * @param o The raw object data from the resultset.
    * @return The actual object ref or null.
    * @throws Exception
    */
   public Object convertObject(Object o) throws Exception
   {
      byte[]b = null;
      Object obj = null;
      ByteArrayInputStream bs = null;
      ObjectInputStream os = null;
      
      if ( o != null ) {
         try {
            b = (byte[])o;      
            bs = new ByteArrayInputStream(b);
            os = new ObjectInputStream(bs);
            obj = os.readObject();
         }

         finally {
            if ( os != null )
               os.close();
            
            if ( bs != null )
               bs.close();
            
            b = null;
            os = null;
            bs = null;
         }
      }
      
      return obj;
   }
   
   /**
    * Calls the stored proc to get the cottle rebate.
    * 
    * @param custId not used in this method.
    * @return A reference to a Rebate object.
    */
   public Rebate cottleFreightRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
            
      try {
         m_CotFrtReb.setDate(1, m_BegDate);
         m_CotFrtReb.setDate(2, m_EndDate);
       
         rs = m_CotFrtReb.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Cottle";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
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

      if ( m_Sheet != null ) {
         //
         // set the report title
         row = m_Sheet.createRow(rowNum++);
         cell = row.createCell(0);
         cell.setCellType(CellType.STRING);
         cell.setCellStyle(styleTitle);
         cell.setCellValue(
            new XSSFRichTextString(
               String.format("Customer Profitability Report: %s to %s ", m_BegDate, m_EndDate)
           )
         );
        
         //
         // Merge the title cells.  Gives a better look to the report.
         region = new CellRangeAddress(0, 0, 0, 5);
         m_Sheet.addMergedRegion(region);
         
         //
         // Create the row for the captions.
         row = m_Sheet.createRow(rowNum++);
         row.setHeight(rowHeight);
         
         for ( int i = 0; i < MAX_COLS; i++ ) {
            cell = row.createCell(i);
            cell.setCellStyle(styleCaption);
         }
         
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Customer\nID"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Customer\nName"));
         m_Sheet.setColumnWidth(colNum++, 5000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("TM"));
         m_Sheet.setColumnWidth(colNum++, 4000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("City"));
         m_Sheet.setColumnWidth(colNum++, 4000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("State"));      
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Delivery\nTrips"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Type"));
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Affiliation"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Warehouse"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Gross Sales"));
         m_Sheet.setColumnWidth(colNum++, 4000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Coop"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Efficiency"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Qtrly"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Other"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Total\nRebates"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Cash\nDiscounts"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Net Sales"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("COGS"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Shipping\nMargin"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Orders"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Lines"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Units"));
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Credits"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Restock"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("CrMems"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("CrLines"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("CrUnits"));
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Current"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Past Due"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Past\nAs\n% of\nSales"));
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Vnd Reb\nMonthly"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Vnd Reb\nQuarterly"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Vnd Reb\nGrowth"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Vnd Reb\nDA"));
         m_Sheet.setColumnWidth(colNum++, 3000);
         row.getCell(colNum).setCellValue(new XSSFRichTextString("Vnd Reb\nTotal"));
         m_Sheet.setColumnWidth(colNum++, 3000);
      }
      
      if ( m_Sheet2 != null ) {
         colNum = 0;
         row = m_Sheet2.createRow(m_RowNum2++);
         cell = row.createCell(0);
         cell.setCellType(CellType.STRING);
         cell.setCellStyle(styleTitle);
         cell.setCellValue(new XSSFRichTextString("Efficiency Rebate Detail"));
        
         //
         // Merge the title cells.  Gives a better look to the report.
         region = new CellRangeAddress(0, 0, 0, 4);
         m_Sheet2.addMergedRegion(region);
         
         //
         // Add the second row for the captions
         row = m_Sheet2.createRow(m_RowNum2++);
         for ( int i = 0; i < effRebColCnt; i++ ) {
            cell = row.createCell(i);
            cell.setCellStyle(styleCaption);
         }
         
         //
         // Add the captions for the efficiency rebate data.         
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Cust ID"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Rebate"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Reb Sales"));         
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Comp Sales"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Prev Sales"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Applied Sales"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Line Val"));         
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Credit Tot"));
         row.getCell(colNum++).setCellValue(new XSSFRichTextString("Credit Pct"));         
      }
      
      m_RowNum3 = createQtrRebCaption();
      
      font = null;
      styleTitle = null;
      styleCaption = null;
      region = null;
      
      return rowNum;
   }
   
   /**
    * Creates the captions for the quarterly rebate detail
    * @return The next row number to use.
    */
   private int createQtrRebCaption()
   {
      XSSFCellStyle styleCaption;
      XSSFCellStyle styleTitle;
      XSSFFont font;
      XSSFCell cell = null;
      XSSFRow row = null;
      CellRangeAddress region = null;      
      int rowNum = 0;
      int colNum = 0;
           
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
      
      try {
         if ( m_Sheet3 != null ) {            
            row = m_Sheet3.createRow(rowNum);
            cell = row.createCell(0);
            cell.setCellType(CellType.STRING);
            cell.setCellStyle(styleTitle);
            cell.setCellValue(new XSSFRichTextString("Quarterly Rebate Detail"));
           
            //
            // Merge the title cells.  Gives a better look to the report.
            region = new CellRangeAddress(rowNum, rowNum, 0, 2);
            m_Sheet3.addMergedRegion(region);
            rowNum++;
            
            //
            // Add the second row for the captions
            row = m_Sheet3.createRow(rowNum++);
            for ( int i = 0; i < qtrRebColCnt; i++ ) {
               cell = row.createCell(i);
               cell.setCellStyle(styleCaption);
            }
            
            //
            // Add the captions for the efficiency rebate data.         
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Cust ID"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Rebate Name"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Q1 Tot"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Q2 Tot"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Q3 Tot"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Q4 Tot"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString("Rebate Total"));         
         }
      }
      
      finally {
         font = null;
         styleTitle = null;
         styleCaption = null;
         region = null;
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
         m_FasConn = m_RptProc.getFasConn();
         
         if ( prepareStatements() )
            created = buildOutputFile();
      }
      
      catch ( Exception ex ) {
         log.fatal("exception:", ex);
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
    * Calculates the coop rebate for a customer.
    * 
    * @param custId The customer to get the rebates dollars for.
    * @return The amount of the rebate.
    */
   private double getCoopRebate(String custId)
   {      
      return 0;
   }  
   
   /**
    * Calculates the efficiency rebate for a customer.
    * 
    * @param custId The customer to get the rebates dollars for.
    * @return The amount of the rebate.
    * @throws SQLException 
    */
   private double getEfficiencyRebate(String custId) throws SQLException
   {
      double rebate = 0.0;
      boolean hasRebate = false;
      ResultSet hasReb = null;
      ResultSet rebData = null;      
      double creditTot = 0;
      double rebSales = 0;
      double compSales = 0;
      double crPct = 0;
      double prevSales = 0;
      double lineVal = 0;
      
      try {
         m_HasRebate.setString(1, effRebName);
         m_HasRebate.setString(2, custId);
         hasReb = m_HasRebate.executeQuery();
         
         if ( hasReb.next() )
            hasRebate = hasReb.getInt(1) > 0;
         
         //
         // If the customer is signed up for the rebate, do the calculations.
         // See the rebate table for the calculation instructions.
         if ( hasRebate ) {
            m_EffRebData.setDate(1, m_BegDate);
            m_EffRebData.setDate(2, m_EndDate);
            m_EffRebData.setDate(3, m_BegDate);
            m_EffRebData.setDate(4, m_EndDate);
            m_EffRebData.setDate(5, m_BegDate);
            m_EffRebData.setDate(6, m_EndDate);
            m_EffRebData.setString(7, custId);
            
            rebData = m_EffRebData.executeQuery();
            
            if ( rebData.next() ) {
               creditTot = rebData.getDouble("c2_credits");
               rebSales = rebData.getDouble("reb_sales");
               lineVal = rebData.getDouble("line_val");
               compSales = rebData.getDouble("comp_sales");
               crPct = (creditTot/rebSales) * 100;
               prevSales = getPrevYrSales(custId);
               
               //
               // Rule 1: reason code 2 credits must be less than 1% net warehouse sales available
               // for rebate.
               if ( crPct < 1 ) {
                  //
                  // Rule 2: current year sales must be greater than the previous years sales.
                  // The compare sales data includes more sale types than just the warehouse sales
                  // used to calculate the actual rebate.
                  if ( compSales > prevSales ) {
                     if ( lineVal >= 20 && lineVal < 25 )
                        rebate = (rebSales - effSalesMin) * .01;
                     else {
                        if ( lineVal >= 25 && lineVal < 30 )
                           rebate = (rebSales - effSalesMin) * .015;
                        else {
                           if ( lineVal >= 30 )
                              rebate = (rebSales - effSalesMin) * .02;
                        }
                     }
                  }
               }
               
               addRebDetail(custId, rebate, rebSales, compSales, prevSales, lineVal, creditTot, crPct);
            }
         }
      }
                  
      finally {
         closeRSet(hasReb);
         closeRSet(rebData);
         
         hasReb = null;
         rebData = null;
      }
      
      return rebate;
   }
   
   /**
    * Calculates the rebates classified as "other" for a customer.
    * 
    * @param custId The customer to get the rebates dollars for.
    * @return The amount of the rebate.
    */
   private double getOtherRebate(String custId)
   {
      // TODO Auto-generated method stub
      return 0;
   }

   /**
    * Calculates the previous years sales for a customer
    * 
    * @param custId The customer to get the sales data for.
    * @return The amount of sales from the previous year.
    */
   private double getPrevYrSales(String custId)
   {
      double sales = 0;
      ResultSet prevYrSales = null;
      Calendar curDate = Calendar.getInstance();
      GregorianCalendar prevBeg = (GregorianCalendar)m_BegCal.clone();
      GregorianCalendar prevEnd = (GregorianCalendar)m_EndCal.clone();
            
      try {
         prevBeg.roll(Calendar.YEAR, false);
         
         //
         // Check to see if we're in the current year.  If we are, figure out what month
         // and set the time frame for the previous year end to the correct month and day to get accurate
         // sales comparisons.  eg Requesting current year and current month is June.  
         // If we're in a previous year, we can just role back one year because
         // we have a complete years worth of data.
         if ( curDate.get(Calendar.YEAR) == m_BegCal.get(Calendar.YEAR) ) {
            prevEnd.roll(Calendar.YEAR, false);
            prevEnd.set(Calendar.MONTH, curDate.get(Calendar.MONTH));
            prevEnd.set(Calendar.DAY_OF_MONTH, curDate.get(Calendar.DAY_OF_MONTH));
         }
         else            
            prevEnd.roll(Calendar.YEAR, false);
                  
         m_PrevYrSales.setDate(1, new java.sql.Date(prevBeg.getTimeInMillis()));
         m_PrevYrSales.setDate(2, new java.sql.Date(prevEnd.getTimeInMillis()));
         m_PrevYrSales.setString(3, custId);
         prevYrSales = m_PrevYrSales.executeQuery();
         
         if ( prevYrSales.next() )
            sales = prevYrSales.getDouble(1);
      }
      
      catch ( SQLException ex ) {         
         log.error("getEfficiencyRebate", ex);
      }
      
      finally {
         closeRSet(prevYrSales);
         prevYrSales = null;
      }
      
      return sales;
   }
   
   /**
    * Calculates the quarterly rebate for a customer.
    * 
    * @param custId The customer to get the rebates dollars for.
    * @return The amount of the rebate.
    * @throws Exception  
    */
   private double getQuarterlyRebate(String custId) throws Exception
   {
      boolean canStop = false;
      String procData = null;
      Iterator<String> iter = null;      
      Class<?> c = null;
      Method meth = null;      
      double rebate = 0;
      Rebate r = null;
      
      try {
         //
         // If this is the first time through, create the list and set the
         // captions for the detail page.         
         if ( m_ProcData == null )
            m_ProcData = new ArrayList<String>();
         else
            m_ProcData.clear();
         
         //
         // Only execute this one time and then iterate through it each time
         // after.  We only need the list of customers with quarterly rebates once.
         if ( m_QrtRebData == null )
            m_QrtRebData = m_QrtRebates.executeQuery();
         else
            m_QrtRebData.beforeFirst();
         
         //
         // The data is sorted by customer id.  Once the customer id is found in the 
         // results, the next time it's not found we can end the search.
         while ( m_QrtRebData.next() ) {
            if ( m_QrtRebData.getString("customer_id").equalsIgnoreCase(custId) ) {
               procData = m_QrtRebData.getString("proc_data");
               
               //
               // If there's a rebate and it has a proc associated with it, add it to the list.
               // We us and create the array list here to avoid recreating it every time.
               if ( procData != null && procData.length() > 0 )
                  m_ProcData.add(procData);
               
               canStop = true;               
            }
            else {
               if ( canStop )
                  break;
            }
         }
         
         //
         // Execute the methods found for each of the rebates.
         // TODO At some point check for the actual number of parameters and maybe the type.
         // For now we'll just force the customer id.
         iter = m_ProcData.iterator();
         
         while ( iter.hasNext() ) {
            c = this.getClass();
            meth = c.getMethod(iter.next(), new Class[]{Class.forName("java.lang.String")});
            
            if ( meth != null ) {
               r = (Rebate)meth.invoke(this, new Object[]{custId});
               
               if ( r != null ) {
                  rebate += r.rebateTot;
                  addQtrRebDetail(r);
               }
               else {
                  log.error(
                     String.format("CustProfitability.getQuarterlyRebate %s null rebate object", meth.getName())
                  );
               }
            }
         }
      }
      
      finally {
         iter = null;
      }
      
      return rebate;
   }
   
   /**
    * Calculates the amount of the rebate belongs to the customer based on the customer
    * sales for the vendors in the actual rebate program.
    * 
    * @param custId The customer id to calc the sales for.
    * @param rebType The vendor rebate program
    * 
    * @return The percent of the of the rebate the customer has earned for Emery.
    */
   private double getVndRebSales(String custId, String rebType)
   {      
      double sales = 0.0;
      ResultSet rs = null;
      String rebParam = "%%%s%%";
      
      try {
         m_VndRebSales.setString(1, custId);
         m_VndRebSales.setDate(2, m_BegDate);
         m_VndRebSales.setDate(3, m_EndDate);
         m_VndRebSales.setString(4, String.format(rebParam, rebType));
         m_VndRebSales.setDate(5, m_BegDate);
         m_VndRebSales.setDate(6, m_EndDate);
         
         rs = m_VndRebSales.executeQuery();
         rs.next();
         sales = rs.getDouble(1);         
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
           
      return sales;
   }
   
   /**
    * Gets the total dollar amount of the growth vendor rebate accrued during the 
    * reporting time period.
    * 
    * @param rebType The type of vendor rebate.  Based on the rebate description.
    * @return The amount of the rebate.
    * 
    * @throws SQLException
    */
   @SuppressWarnings("unused")
   private double getVndReb(String rebType) throws SQLException
   {
      double rebate = 0.0;
      ResultSet rs = null;
      String rebParam = "%%%s%%";
      
      try {
         m_VndReb.setString(1, String.format(rebParam, rebType));
         m_VndReb.setDate(2, m_BegDate);
         m_VndReb.setDate(3, m_EndDate);
         
         rs = m_VndReb.executeQuery();
         rs.next();
         rebate = rs.getDouble(1);         
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return rebate;
   }
  
   /**
    * Calls the stored proc to get the rebate.
    * 
    * @param custId - not used in this method.
    * @return A reference to a Rebate object.
    */
   public Rebate hammondWhsRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
            
      try {
         m_Rebate.setInt(1, RebateProcs.hamWhsReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
        
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Hammond Warehouse Sales";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Calls the stored proc to get the rebate.
    * 
    * @param custId - not used in this method.
    * @return A reference to a Rebate object.
    */
   public Rebate hancockWhsRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
            
      try {
         m_Rebate.setInt(1, RebateProcs.hanWhsQtrReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
         
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Hancock Warehouse Sales";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Calls the stored proc to get the rebate.
    * 
    * @param custId - not used in this method.
    * @return A reference to a Rebate object.
    */
   public Rebate harringtonBaseRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
            
      try {
         m_HarBaseReb.setDate(1, m_BegDate);
         m_HarBaseReb.setDate(2, m_EndDate);
       
         rs = m_HarBaseReb.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Harrington base cost";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Calculates the Harrington receiving rebate
    * 
    * @param custId - not used in this method.
    * @return A reference to a Rebate object.
    */
   public Rebate harringtonRecRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
      
      try {
         m_HarRecReb.setDate(1, m_BegDate);
         m_HarRecReb.setDate(2, m_EndDate);
      
         rs = m_HarRecReb.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Harrington Receiving";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Calculates the LMC warehouse sales rebate
    * 
    * @param custId The customer id.
    * @return A reference to a Rebate object.
    */
   public Rebate lmcWhsRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
      
      try {
         m_Rebate.setInt(1, RebateProcs.lmcWhsReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
        
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "LMC 1% Warehouse Sales";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Calculates the Midcape warehouse sales rebate
    * 
    * @param custId - The midcape account id
    * @return A reference to a Rebate object.
    */
   public Rebate midcapeWhsRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
      
      try {
         m_Rebate.setInt(1, RebateProcs.midWhsReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
         
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "Midcape Warehouse Sales";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Calculates the RP Johnson warehouse sales rebate
    * 
    * @param custId - The account id
    * @return A reference to a Rebate object.
    */
   public Rebate rpjWhsRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
      
      try {
         m_Rebate.setInt(1, RebateProcs.rpjWhsReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
         
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "RP Johnson Warehouse Sales";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(1024);      
      boolean isPrepared = false;
      
      if ( m_EdbConn != null && m_FasConn != null ) {
         try {
            //sql.setLength(0);
            //sql.append("select rebate_procs.getBensonRebate(?, ?) as rebate "); 
            //m_BensonReb = m_EdbConn.prepareStatement(sql.toString()); // TODO doesnt exist in edb
            
            //sql.setLength(0);
            //sql.append("select rebate_procs.getCottleRebate(?, ?) as rebate ");
            //m_CotFrtReb = m_EdbConn.prepareStatement(sql.toString()); // TODO doesnt exist in edb
            
            //sql.setLength(0);
            //sql.append("select rebate_procs.getHarringtonBaseRebate(?, ?) as rebate ");
            //m_HarBaseReb = m_EdbConn.prepareStatement(sql.toString()); // TODO doesnt exist in edb
            
            //sql.setLength(0);
            //sql.append("select rebate_procs.getHarringtonRecRebate(?, ?) as rebate ");
            //m_HarRecReb = m_EdbConn.prepareStatement(sql.toString()); // TODO doesnt exist in edb
                       
            //sql.setLength(0);
            //sql.append("select rebate_procs.getRebate(?, ?, ?, ?) as rebate ");
            //m_Rebate = m_EdbConn.prepareStatement(sql.toString()); // TODO doesnt exist in edb
            
            //
            // customer information based on accounts or store number.
            // Gross sales are net of credits.
            sql.setLength(0);
            sql.append("select ");
            sql.append("   customer.customer_id, customer.name as cname, repname as tm, ");
            sql.append("   cav.city, cav.state, ");
            sql.append("   (");
            sql.append("      select count(trip_sched.fascor_id) ");
            sql.append("      from trip_stop_sched tss ");
            sql.append("      join trip_sched on trip_sched.ts_id = tss.ts_id ");
            sql.append("      where tss.customer_id = customer.customer_id ");
            sql.append("   ) trips, ");
            sql.append("   \"class\" as type, buying_group.name as affiliation, warehouse.name as warehouse, ");
            sql.append("   ar_tot, ar_cur, ar_past, gsales, cogs, orders, lines, units, ");
            sql.append("   cr_tot, cr_mems, cr_lines, cr_units, cr_code ");
            sql.append("from customer ");
            sql.append("join ( ");
            sql.append("   select distinct ejd.cust_procs.findtopparent(customer.customer_id) as customer_id ");
            sql.append("   from customer ");
            sql.append("   join customer_status cs on cs.cust_status_id = customer.cust_status_id and description <> 'INACTIVE' ");
            sql.append(") acct on acct.customer_id = customer.customer_id ");
            sql.append("join cust_market_view cmv on cmv.customer_id = customer.customer_id and \r\n");
            sql.append("   market = 'CUSTOMER TYPE' and class not in ('EMPLOYEE', 'EMERY', 'BACKHAULS', 'INACTIVE') \r\n");
            sql.append("left outer join cust_address_view cav on cav.customer_id = customer.customer_id and addrtype = 'SHIPPING' ");
            sql.append("left outer join cust_buy_group cbg on cbg.customer_id = customer.customer_id and ");
            sql.append("   cbg.beg_date <= trunc(now()) and (cbg.end_date is null or cbg.end_date >= trunc(now())) ");
            sql.append("left outer join buying_group on buying_group.buy_group_id = cbg.buy_group_id "); 
            sql.append("join cust_warehouse cw on cw.customer_id = customer.customer_id ");
            sql.append("join warehouse on warehouse.warehouse_id = cw.warehouse_id ");
            sql.append("left outer join cust_rep_div_view crdv on crdv.customer_id = customer.customer_id and rep_type = 'SALES REP' ");            
            sql.append("left outer join ( ");
            sql.append("   select ");
            sql.append("      decode(parent_id, null, customer_id, parent_id) as customer_id, sum(amtduehc) as ar_tot, ");
            sql.append("      sum(case when datedue >= to_char(now(), 'yyyymmdd') then amtduehc else 0 end) as ar_cur, ");
            sql.append("      sum(case when datedue < to_char(now(), 'yyyymmdd') then amtduehc else 0 end) as ar_past ");
            sql.append("   from customer c");
            sql.append("   join ejd.sage300_arobl_mv on idcust = c.customer_id and swpaid = 0 ");
            sql.append("   where customer_id in ( ");
            sql.append("      select c2.customer_id ");
            sql.append("      from customer c2 ");
            sql.append("      start with c2.customer_id = customer_id ");
            sql.append("      connect by prior customer_id = parent_id ");
            sql.append("   )");
            sql.append("   group by decode(parent_id, null, customer_id, parent_id) ");
            sql.append(") ar on ar.customer_id = customer.customer_id ");
            sql.append("left outer join ( ");
            sql.append("   select ");
            sql.append("      decode(parent_id, null, customer_id, parent_id) as customer_id, ");
            sql.append("      sum(dollars_shipped) as gsales, sum(cost_shipped) as cogs, ");
            sql.append("      count(inv_hdr.inv_hdr_id) as orders, ");
            sql.append("      sum(lines_shipped) as lines, sum(units_shipped) as units ");
            sql.append("   from customer c ");
            sql.append("   join inv_hdr on inv_hdr.cust_nbr = c.customer_id and ");
            sql.append("      invoice_date between ? and ? ");            
            sql.append("   where c.customer_id in ( ");
            sql.append("      select c2.customer_id ");
            sql.append("      from customer c2");
            sql.append("      start with c2.customer_id = customer_id ");
            sql.append("      connect by prior customer_id = parent_id ");
            sql.append("   )");
            sql.append("   group by decode(parent_id, null, customer_id, parent_id) ");
            sql.append(") sales on sales.customer_id = customer.customer_id ");
            sql.append("left outer join ( ");
            sql.append("   select ");
            sql.append("      decode(parent_id, null, customer_id, parent_id) as customer_id, ");
            sql.append("      abs(sum(ext_sell)) as cr_tot, count(distinct inv_hdr_id) as cr_mems, ");
            sql.append("      count(inv_dtl_id) as cr_lines, abs(sum(qty_shipped)) as cr_units, ");
            sql.append("      sum(case to_number(return_reason_cd) when 3 then 1 else 0 end) as cr_code ");
            sql.append("   from customer c ");
            sql.append("   join inv_dtl on inv_dtl.cust_nbr = c.customer_id and ");
            sql.append("      tran_type = 'CREDIT' and ");
            sql.append("      invoice_date between ? and ? ");
            sql.append("   where c.customer_id in ( ");
            sql.append("      select c2.customer_id ");
            sql.append("      from customer c2 ");
            sql.append("      start with customer_id = customer_id ");
            sql.append("      connect by prior customer_id = parent_id ");
            sql.append("   )");
            sql.append("   group by decode(parent_id, null, customer_id, parent_id) ");
            sql.append(") credits on credits.customer_id = customer.customer_id ");
            sql.append("order by customer.customer_id");
            m_CustData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select ");
            sql.append("   sum(reb_sales) as reb_sales, ");
            sql.append("   round(sum(reb_sales)/sum(lines_shipped), 2) as line_val, ");
            sql.append("   sum(c2_credits) as c2_credits, sum(comp_sales) as comp_sales ");
            sql.append("from ( ");
            sql.append("   select ");
            sql.append("      customer_id, sum(dollars_shipped) as reb_sales, ");
            sql.append("      sum(lines_shipped) as lines_shipped, ");
            sql.append("      ( ");
            sql.append("         select abs(sum(inv_dtl.ext_sell)) ");
            sql.append("         from inv_dtl ");
            sql.append("         where ");
            sql.append("            return_reason_cd = '02' and ");
            sql.append("            inv_dtl.invoice_date between ? and ? and ");
            sql.append("            inv_dtl.tran_type = 'CREDIT' and ");
            sql.append("            inv_dtl.sale_type = 'WAREHOUSE' and ");
            sql.append("            inv_dtl.cust_nbr = customer_id ");
            sql.append("      ) c2_credits, ");
            sql.append("      ( ");
            sql.append("         select sum(ih.dollars_shipped) ");
            sql.append("         from inv_hdr ih ");
            sql.append("         where ");
            sql.append("            ih.invoice_date between ? and ? and ");
            sql.append("            ih.tran_type = 'SALE' and ");
            sql.append("            ih.sale_type in ('WAREHOUSE', 'VIRTUAL', 'APG WHS', 'DROP SHIP', 'BLDGMAT DS') and  ");
            sql.append("            ih.cust_nbr = customer_id ");
            sql.append("      ) comp_sales ");
            sql.append("   from customer ");
            sql.append("   join inv_hdr on inv_hdr.cust_nbr = customer.customer_id and ");
            sql.append("      inv_hdr.invoice_date between ? and ? and ");
            sql.append("      inv_hdr.tran_type = 'SALE' and inv_hdr.sale_type in ('WAREHOUSE', 'ACE DIRECT')");
            sql.append("   where customer_id in ( ");
            sql.append("      select customer_id ");
            sql.append("      from customer ");
            sql.append("      start with customer_id = ? ");
            sql.append("      connect by prior customer_id = parent_id ");
            sql.append("   ) ");
            sql.append("   group by customer_id ");
            sql.append(")");
            m_EffRebData = m_EdbConn.prepareStatement(sql.toString());
                                    
            sql.setLength(0);
            sql.append("select sum(reb_count) as reb_count ");
            sql.append("from customer ");
            sql.append("join ( ");
            sql.append("   select customer_id, count(*) reb_count ");
            sql.append("   from cust_rebate ");
            sql.append("   join rebate on rebate.rebate_id = cust_rebate.rebate_id and rebate.name = ? ");
            sql.append("   group by customer_id ");
            sql.append(") reb on reb.customer_id = customer.customer_id ");
            sql.append("start with customer.customer_id = ? ");
            sql.append("connect by prior customer.customer_id = parent_id");
            m_HasRebate = m_EdbConn.prepareStatement(sql.toString());
            
            //
            // Prior years sales include all warehouse variants and rolls up to the parent account.
            sql.setLength(0);
            sql.append("select sum(dollars_shipped) as sales ");
            sql.append("from customer ");
            sql.append("join inv_hdr on inv_hdr.cust_nbr = customer.customer_id and ");
            sql.append("   inv_hdr.invoice_date between ? and ? and ");
            sql.append("   inv_hdr.sale_type in ('WAREHOUSE', 'ACE DIRECT', 'VIRTUAL', 'APG WHS', 'DROP SHIP', 'BLDGMAT DS') ");
            sql.append("where customer_id in ( ");
            sql.append("  select customer_id ");
            sql.append("   from customer ");
            sql.append("   start with customer_id = ? ");
            sql.append("   connect by prior customer_id = parent_id ");
            sql.append(") ");
            m_PrevYrSales = m_EdbConn.prepareStatement(sql.toString());
                        
            sql.setLength(0);
            sql.append("select customer_id, name, proc_data ");
            sql.append("from rebate ");
            sql.append("join cust_rebate on cust_rebate.rebate_id = rebate.rebate_id ");
            sql.append("where rebate_period_id = ( ");
            sql.append("   select rebate_period_id ");
            sql.append("   from rebate_period ");
            sql.append("   where period = 'QUARTERLY' ");
            sql.append(") ");
            sql.append("order by customer_id");
            m_QrtRebates = m_EdbConn.prepareStatement(sql.toString(),
                  ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
           
            
            sql.setLength(0);
            sql.append("select sum(amount) as rebate ");
            sql.append("from vendor_gl_summary ");
            sql.append("where acctid = ( ");
            sql.append("   select to_number(trim(acctid)) ");
            sql.append("   from account_set_gl acs ");
            sql.append("   where acs.description = 'VENDOR REBATE' and acctdesc like ? ");
            sql.append(") ");
            sql.append("and tran_date between ? and ?");
            m_VndReb = m_EdbConn.prepareStatement(sql.toString());
                        
            sql.setLength(0);
               
            sql.append("select sum(ext_sell) as sales ");
            sql.append("from inv_dtl ");
            sql.append("where ");
            sql.append("   cust_nbr = ? and ");
            sql.append("   invoice_date between ? and ? and ");
            sql.append("   vendor_nbr in ( ");
            sql.append("      select distinct vendor_id ");
            sql.append("      from vendor_gl_summary ");
            sql.append("      where acctid = ( ");
            sql.append("         select to_number(trim(acctid)) ");
            sql.append("         from account_set_gl acs ");
            sql.append("         where acs.description = 'VENDOR REBATE' and acctdesc like ? ");
            sql.append("      ) ");
            sql.append("      and tran_date between ? and ? and vendor_id is not null ");
            sql.append("   ) ");
            m_VndRebSales = m_EdbConn.prepareStatement(sql.toString());
            
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
         log.error("custprofitabilty.prepareStatements - null enterprisedb or fascor connection");
      
      return isPrepared;
   }
   
   /**
    * Calculates the LMC warehouse sales rebate
    * 
    * @param custId - not used in this method.
    * @return A reference to a Rebate object.
    */
   public Rebate prscoDsRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
      
      try {
         m_Rebate.setInt(1, RebateProcs.prscoDsReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
       
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "PRSCO Dropship Sales";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
   }
   
   /**
    * Calculates the LMC warehouse sales rebate
    * 
    * @param custId - The customer that the rebate is for
    * @return A reference to a Rebate object.
    */
   public Rebate prscoWhsRebate(String custId)
   {
      Rebate reb = null;
      ResultSet rs = null;
      
      try {
         m_Rebate.setInt(1, RebateProcs.prscoWhsReb);
         m_Rebate.setString(2, custId);
         m_Rebate.setDate(3, m_BegDate);
         m_Rebate.setDate(4, m_EndDate);
         
         rs = m_Rebate.executeQuery();
         rs.next();
         reb = (Rebate)convertObject(rs.getObject(1));
         
         //
         // setup the rest of the data for output to the 
         // rebate detail page
         reb.custId = custId;
         reb.rebate = "PRSCO Warehouse Sales";
      }
      
      catch ( Exception ex ) {
         log.error("exception", ex);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
      
      return reb;
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
      int year;
      int month;
      int day;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                  
         if ( param.name.equals("begdate") ) {
            year = Integer.parseInt(param.value.substring(6,10));
            month = Integer.parseInt(param.value.substring(0,2)) - 1;
            day = Integer.parseInt(param.value.substring(3,5));
            m_BegCal = new GregorianCalendar(year, month, day);
            m_BegDate = new java.sql.Date(m_BegCal.getTimeInMillis());
         }
         
         if ( param.name.equals("enddate") ) {
            year = Integer.parseInt(param.value.substring(6,10));
            month = Integer.parseInt(param.value.substring(0,2)) - 1;
            day = Integer.parseInt(param.value.substring(3,5));
            m_EndCal = new GregorianCalendar(year, month, day);
            m_EndDate = new java.sql.Date(m_EndCal.getTimeInMillis());
         } 
      }
      
      fileName.append("custprofit");
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
         styleTxtC,    // col 0 cust id
         styleTxtL,    // col 1 cust name
         styleTxtL,    // col 2 tm
         styleTxtL,    // col 3 city         
         styleTxtL,    // col 4 state
         styleInt,     // col 5 trips
         styleTxtC,    // col 6 type
         styleTxtC,    // col 7 affiliation
         styleTxtL,    // col 8 warehouse
         styleMoney,   // col 9 gross sales
         styleMoney,   // col 10 coop
         styleMoney,   // col 11 efficiency
         styleMoney,   // col 12 quarterly
         styleMoney,   // col 13 other
         styleMoney,   // col 14 total rebate
         styleMoney,   // col 15 cash discount
         styleMoney,   // col 16 net sales         
         styleMoney,   // col 17 cogs
         styleMoney,   // col 18 shipping margin
         styleInt,     // col 19 order
         styleInt,     // col 20 lines
         styleInt,     // col 21 units
         styleMoney,   // col 22 credits
         styleInt,     // col 23 restock
         styleInt,     // col 24 cr memos
         styleInt,     // col 25 cr lines
         styleInt,     // col 26 cr units
         styleMoney,   // col 27 ar current
         styleMoney,   // col 28 ar past due
         stylePct,     // col 29 pct of sales
         styleMoney,   // col 30 monthly
         styleMoney,   // col 31 quarterly
         styleMoney,   // col 32 growth
         styleMoney,   // col 33 da
         styleMoney    // col 34 total
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
