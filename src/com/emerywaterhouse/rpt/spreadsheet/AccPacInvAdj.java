/**
 * File: AccPacInvAdj.java
 * Description: Accpac report that shows the inventory adjustments
 *
 * @author Peggy Richter
 *
 * Create Date: 10/27/2006
 * Last Update: $Id: AccPacInvAdj.java,v 1.9 2009/03/19 19:45:42 npasnur Exp $
 * 
 * History
 *   $Log: AccPacInvAdj.java,v $
 *   Revision 1.9  2009/03/19 19:45:42  npasnur
 *   Fixed issue with null values for reason
 *
 *   Revision 1.8  2009/02/25 17:16:45  prichter
 *   Added a warehouse_id and new primary key to table inv_adj_reason
 *
 *   Revision 1.7  2009/02/18 14:41:46  jfisher
 *   Fixed depricated methods after poi upgrade
 *
 *   Revision 1.6  2006/12/06 15:35:58  prichter
 *   Fixed a bug that was causing transactions to show up multiple times if all of the optional fields were not created on each line.
 *
 *   Revision 1.5  2006/11/07 21:22:55  prichter
 *   Modified the query to include unposted and $0 adjustments.  Fixed a problem with selection by batch#
 *
 *   Revision 1.4  2006/11/02 20:03:59  jfisher
 *   Production version
 *
 *    
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

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


public class AccPacInvAdj extends Report
{
   private static final short MAX_COLS = 20;
   
   private PreparedStatement m_AdjReason;
   private PreparedStatement m_InvAdjData;
   
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private XSSFCellStyle[] m_CellStyles;
   
   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;  
   private XSSFFont m_FontBold;
   private XSSFFont m_FontNormal;
   
   // Parameter member variables
   private String m_BegDate;
   private String m_EndDate;
   private String m_ItemId;
   private String m_ReasonCode;
   private String m_Vendor;
   private String m_GlAccount;
   private String m_Batch;

   
   /**
    * 
    */
   public AccPacInvAdj()
   {
      super();
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      setupWorkbook();      
   }

   /**
    * Clean up any resources
    *  
    * @see java.lang.Object#finalize()
    */
   public void finalize() throws Throwable
   {            
      closeStmt(m_InvAdjData);
      closeStmt(m_AdjReason);
      
      m_InvAdjData = null;
      m_AdjReason = null;
      
      if ( m_CellStyles != null ) {
         for ( int i = 0; i < m_CellStyles.length; i++ )
            m_CellStyles[i] = null;
      }
      
      m_Sheet = null;
      m_Wrkbk = null;      
      m_CellStyles = null;
      
      super.finalize();
   }
   

   /**
    * Executes the queries and builds the output file
    * 
    * @return true if the report was successfully built
    * @throws FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      XSSFRow row = null;
      int rowNum = 0;
      int colNum = 0;
      FileOutputStream outFile = null;
      ResultSet invAdjData = null;
      boolean result = false;
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         rowNum = createCaptions();
                  
         invAdjData = m_InvAdjData.executeQuery();
         m_CurAction = "Building output file";
         rowNum = createCaptions();

         while ( invAdjData.next() && getStatus() != RptServer.STOPPED ) {
            row = createRow(rowNum++, MAX_COLS);
            colNum = 0;

            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("FACILITY")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("TRANSDATE")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("FISCYEAR")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("FISCPERIOD")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("BATCHID")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("ITEMNO")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(getItemDescr(invAdjData.getString("ITEMNO"))));
            
            // TransTypes of 2, 4, & 6 are inventory decreases
            int trantype = invAdjData.getInt("TRANSTYPE");
            if ( trantype == 2 || trantype == 4 || trantype == 6 ) {
               row.getCell(colNum++).setCellValue(invAdjData.getInt("QUANTITY") * -1);            
               row.getCell(colNum++).setCellValue(invAdjData.getDouble("EXTCOST") * -1D);
            }
            else {
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("QUANTITY")));            
               row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("EXTCOST")));
            }
            
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("WOFFACCT")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("VNDREC")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(getVendorName(invAdjData.getString("VNDREC"))));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("REASON")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(getReasonDescr(invAdjData.getString("FACILITY").trim(), invAdjData.getString("REASON").trim())));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("LOCATION")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("COMMENTS")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("USERID")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("COSTCODE")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("FUNC")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(invAdjData.getString("TIMESTAMP")));

         }
         
         m_Wrkbk.write(outFile);
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {         
         closeRSet(invAdjData);
         invAdjData = null;
         
         row = null;
                  
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error(e);
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
      closeStmt(m_InvAdjData);
   }
   
   /**
    * Creates the report and column headings
    * 
    * @return short - the next available row#
    */
   private int createCaptions()
   {
      XSSFRow row = null;
      XSSFCell cell = null;      
      int colCnt = 20;
      int col = 0;
      int rw = 0;

      //
      row = m_Sheet.createRow(rw++);
      
      cell = row.createCell( 0);
      cell.setCellType(CellType.STRING);
      cell.getCellStyle().setFont(m_FontBold);
      cell.setCellValue(new XSSFRichTextString("Inventory Adjustment Analysis"));

      //
      // Show the current date
      row = m_Sheet.createRow(rw++);
      cell = row.createCell( 0);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(
         new XSSFRichTextString(new SimpleDateFormat("yyyy/MM/dd").format(new java.util.Date()))
      );

      //
      // Build the column headings      
      row = m_Sheet.createRow(rw++);

      if ( row != null ) {
         for ( int i = 0; i < colCnt; i++ ) {
            cell = row.createCell(i);
            cell.setCellType(CellType.STRING);
            cell.getCellStyle().setFont(m_FontBold);
         }

         m_Sheet.setColumnWidth(col, 1500);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Facility"));

         m_Sheet.setColumnWidth(col, 2100);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Trans Date"));
         
         m_Sheet.setColumnWidth(col, 1500);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Fiscal Yr"));
         
         m_Sheet.setColumnWidth(col, 1000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Period"));
         
         m_Sheet.setColumnWidth(col, 1500);
         row.getCell(col++).setCellValue(new XSSFRichTextString("G/L Batch"));

         m_Sheet.setColumnWidth(col, 2000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Item"));

         m_Sheet.setColumnWidth(col, 6000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Item Description"));

         m_Sheet.setColumnWidth(col, 1000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Qty"));

         m_Sheet.setColumnWidth(col, 2000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Cost"));

         m_Sheet.setColumnWidth(col, 3000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("G/L Account"));

         m_Sheet.setColumnWidth(col, 2000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Vendor"));

         m_Sheet.setColumnWidth(col, 6000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Vendor Name"));

         m_Sheet.setColumnWidth(col, 1800);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Reason"));

         m_Sheet.setColumnWidth(col, 4000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Reason Descr"));

         m_Sheet.setColumnWidth(col, 2800);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Location"));

         m_Sheet.setColumnWidth(col, 4000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Comments"));

         m_Sheet.setColumnWidth(col, 2200);
         row.getCell(col++).setCellValue(new XSSFRichTextString("RF User"));

         m_Sheet.setColumnWidth(col, 2500);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Cost Code"));

         m_Sheet.setColumnWidth(col, 2500);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Function"));

         m_Sheet.setColumnWidth(col, 4000);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Timestamp"));
      }
      
      return rw;
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
         m_SageConn = m_RptProc.getSageConn();
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
   private XSSFRow createRow(int rowNum, short colCnt)
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
    * Returns an item's description
    * @param itemId - the item id
    * @return - the item description
    */
   private String getItemDescr(String itemId)
   {
      Statement stmt = null;
      ResultSet rs = null;
      String result = "**Unknown**"; 
      
      try {
         stmt = m_EdbConn.createStatement();
         rs = stmt.executeQuery("select description from item_entity_attr where item_id = " + itemId);
         
         if ( rs.next() )
            result = rs.getString("description");
      }
      
      catch ( Exception e ) {
         e.printStackTrace();
      }
      
      finally {
         closeRSet(rs);
         closeStmt(stmt);
         rs = null;
         stmt = null;
      }
      
      return result;
   }
   
   /**
    * Returns the description of an adjustment reason
    * 
    * @param location String - the AccPac location code
    * @param reasonCd String - the adjustment reason code
    * @return The description of the adjustment reason
    */
   private String getReasonDescr(String location, String reasonCd)
   {
      ResultSet rs = null;
      String result = "";
      
      try {
      	m_AdjReason.setString(1, location);
      	m_AdjReason.setString(2, reasonCd);
         rs = m_AdjReason.executeQuery();
         
         if ( rs.next() ) 
            result = rs.getString("description");
      }
      
      catch ( Exception e ) {
         log.error("exception", e);
      }
      
      finally {
         closeRSet(rs);
         rs = null;
      }
   
      return result;
   }
   
   /**
    * Returns the name of the vendor of record
    * @param vndId - the vendor id
    * @return String - the vendor name
    */
   private String getVendorName(String vndId)
   {
      Statement stmt = null;
      ResultSet rs = null;
      String result = "**Unknown**";
      
      try {
         stmt = m_EdbConn.createStatement();
         rs = stmt.executeQuery("select name from vendor where vendor_id = " + vndId);
         
         if ( rs.next() )
            result = rs.getString("name");
      }
      
      catch ( Exception e ) {
         e.printStackTrace();
      }
      
      finally {
         closeRSet(rs);    
         closeStmt(stmt);
         rs = null;
         stmt = null;
      }
      
      return result;
   }
   
   /**
    * Prepares the sql queries for execution.
    * 
    * @return true if the statements were succssfully prepared
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;
      
      if ( m_SageConn != null && m_EdbConn != null) {
         try {
         	sql.setLength(0);
         	sql.append("select inv_adj_reason.description ");
         	sql.append("from inv_adj_reason ");
         	sql.append("join warehouse on warehouse.warehouse_id = inv_adj_reason.warehouse_id and ");
         	sql.append("                  warehouse.accpac_wh_id = ? ");
         	sql.append("where adj_reason_code = ?  ");
         	m_AdjReason = m_EdbConn.prepareStatement(sql.toString());
         	
         	sql.setLength(0);
            sql.append("select distinct icaded.LOCATION as facility, icadeh.TRANSDATE, icaded.ITEMNO, icaded.QUANTITY, "); 
            sql.append("		 icaded.EXTCOST, WOFFACCT, loc.VALUE as location, icaded.TRANSTYPE, ");
            sql.append("		 coalesce(reason.VALUE, ' ') as reason, glamf.ACCTDESC,  "); 
            sql.append("		 comments.VALUE as comments, userid.VALUE as userid, vndrec.VALUE as vndrec, "); 
            sql.append("		 costcode.VALUE as costcode, func.VALUE as func, timestamp.VALUE as timestamp, ");
            sql.append("		 icadeh.FISCYEAR, icadeh.FISCPERIOD, gljeh.BATCHID, ichist.DOCNUM, ichist.APP, icadeh.ADJENSEQ ");
            sql.append("from EMEDAT.dbo.ICADEH icadeh ");
            sql.append("left join EMEDAT.dbo.ICADED icaded on icaded.ADJENSEQ = icadeh.ADJENSEQ ");
            sql.append("left join EMEDAT.dbo.ICADEDO loc on loc.ADJENSEQ = icaded.ADJENSEQ and loc.\"LINENO\" = icaded.\"LINENO\" ");
            sql.append("left join EMEDAT.dbo.ICADEDO reason on reason.ADJENSEQ = icaded.ADJENSEQ and reason.\"LINENO\" = icaded.\"LINENO\" ");
            sql.append("left join EMEDAT.dbo.ICADEDO comments on comments.ADJENSEQ = icaded.ADJENSEQ and comments.\"LINENO\" = icaded.\"LINENO\" ");
            sql.append("left join EMEDAT.dbo.ICADEDO userid on userid.ADJENSEQ = icaded.ADJENSEQ and userid.\"LINENO\" = icaded.\"LINENO\" ");
            sql.append("left join EMEDAT.dbo.ICADEDO timestamp on timestamp.ADJENSEQ = icaded.ADJENSEQ and timestamp.\"LINENO\" = icaded.\"LINENO\" ");
            sql.append("left join EMEDAT.dbo.ICADEDO vndrec on vndrec.ADJENSEQ = icaded.ADJENSEQ and vndrec.\"LINENO\" = icaded.\"LINENO\" ");
            sql.append("left join EMEDAT.dbo.ICADEDO func on func.ADJENSEQ = icaded.ADJENSEQ and func.\"LINENO\" = icaded.\"LINENO\" "); 
            sql.append("left join EMEDAT.dbo.ICADEDO costcode on costcode.ADJENSEQ = icaded.ADJENSEQ and costcode.\"LINENO\" = icaded.\"LINENO\" ");
            sql.append("left join EMEDAT.dbo.GLAMF glamf on glamf.ACCTFMTTD = icaded.WOFFACCT ");
            sql.append("left join EMEDAT.dbo.ICHIST ichist on ichist.DOCNUM = icadeh.DOCNUM "); 
            sql.append("left join EMEDAT.dbo.GLJEH gljeh on gljeh.DRILAPP = ichist.APP and gljeh.DRILLDWNLK = ichist.DRILLDWNLK ");
            sql.append("where loc.OPTFIELD = 'LOCATION' and ");
            sql.append("	   reason.OPTFIELD = 'REASONCODE' and ");
            sql.append("	   comments.OPTFIELD = 'COMMENTS' and ");
            sql.append("	   userid.OPTFIELD = 'RFUSER' and ");
            sql.append("	   timestamp.OPTFIELD = 'TIMESTAMP' and "); 
            sql.append("	   vndrec.OPTFIELD = 'VNDREC' and  ");
            sql.append("	   func.OPTFIELD = 'FUNCTION' and  ");
            sql.append("	   costcode.OPTFIELD = 'COSTCODE' ");
           
            if ( m_BegDate.length() == 10 )
               sql.append("and icadeh.TRANSDATE >= " + 
                           m_BegDate.substring(6, 10) + 
                           m_BegDate.substring(0,2) + 
                           m_BegDate.substring(3,5) + " ");
            
            if ( m_EndDate.length() == 10 )
               sql.append("and icadeh.TRANSDATE <= " + 
                           m_EndDate.substring(6, 10) + 
                           m_EndDate.substring(0,2) + 
                           m_EndDate.substring(3,5) + " ");
            
            if ( m_ItemId.length() == 7 )
               sql.append("and icaded.ITEMNO = '" + m_ItemId + "' ");
            
            if ( m_ReasonCode.length() > 0 )
               sql.append("and reason.VALUE = '" + m_ReasonCode + "' ");
            
            if ( m_Vendor.length() > 0 )
               sql.append("and vndrec.VALUE = '" + m_Vendor + "' ");
            
            if ( m_GlAccount.length() > 0 )
               sql.append("and WOFFACCT = '" + m_GlAccount + "' ");
            
            if ( m_Batch.length() > 0 )
               sql.append("and BATCHID = '" + m_Batch + "' ");
            
            sql.append("order by icadeh.ADJENSEQ");
                        
            m_InvAdjData = m_SageConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("exception:", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else if (m_SageConn == null) {
         log.error("AccPacInvAdj.prepareStatements - null sqlserver connection");
      } else {
         log.error("AccPacInvAdj.prepareStatements - null edb connection");
      }
      
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
                  
      fileName.append("acp_invadj");      
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
      
      m_BegDate = growDateString(params.get(0).value.trim());
      m_EndDate = growDateString(params.get(1).value.trim());
      m_ItemId = params.get(2).value.trim();
      m_ReasonCode = params.get(3).value.trim();
      m_Vendor = params.get(4).value.trim();
      m_GlAccount = params.get(5).value.trim();
      m_Batch = params.get(6).value.trim();
   }
   
   /**
    * Strings in ACCPAC are stored as YYYYMMDD. We expect to receive input that is 10 characters long to convert.
    * If the input is not 10 characters in length, then see if the day or month needs to be expanded by adding a 0 to the front of it.
    * (ie, 5th day of the 4th month becoming '05/04' instead of just '5/4')
    * If the length is already 10, we do not need to do anything and can simply return it.
    * @param dateString
    * @return
    */
   public String growDateString(String dateString) {
      String newDate = dateString;
      
      if (newDate.length() < 10) {         
         String firstChunk = dateString.substring(0, dateString.indexOf("/"));
         if (firstChunk.length() == 1) {
            firstChunk = "0" + firstChunk;
         }
         String secondChunk = dateString.substring(firstChunk.length());
         secondChunk = secondChunk.substring(0, secondChunk.indexOf("/"));
         if (secondChunk.length() == 1) {
            secondChunk = "0" + secondChunk;
         }
         newDate = firstChunk + "/" + secondChunk + dateString.substring(dateString.lastIndexOf("/"));
         log.info("Date given smaller than 10 characters - attempted repair from: " + dateString + " to: " + newDate);
      }
      
      return newDate;
   }
      
   /**
    * Sets up the styles for the cells based on the column data.  Does any other inititialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {      
      XSSFCellStyle styleText;      // Text left justified
      XSSFCellStyle styleInt;       // Style with 0 decimals
      XSSFCellStyle styleDec;       // 2 decimal positions
      
      //
      // Create a font that is normal size & bold
      m_FontBold = m_Wrkbk.createFont();
      m_FontBold.setFontHeightInPoints((short)8);
      m_FontBold.setFontName("Arial");
      m_FontBold.setBold(true);
      
      //
      // Create a font that is normal size & bold
      m_FontNormal = m_Wrkbk.createFont();
      m_FontNormal.setFontHeightInPoints((short)8);
      m_FontNormal.setFontName("Arial");
            
      styleText = m_Wrkbk.createCellStyle();      
      styleText.setAlignment(HorizontalAlignment.LEFT);
      styleText.setFont(m_FontNormal);
      
      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);
      styleInt.setFont(m_FontNormal);

      styleDec = m_Wrkbk.createCellStyle();
      styleDec.setAlignment(HorizontalAlignment.RIGHT);
      styleDec.setDataFormat((short)4);
      styleDec.setFont(m_FontNormal);
      
      m_CellStyles = new XSSFCellStyle[] {
         styleText,     // col 0 Facility
         styleText,     // col 1 Trans Date yyyymmdd
         styleText,     // col 2 Fiscal Year
         styleText,     // col 3 Fiscal Period
         styleText,     // col 4 G/L Batch
         styleText,     // col 5 Item#
         styleText,     // col 6 Item Description
         styleInt,      // col 7 Qty
         styleDec,      // col 8 Cost
         styleText,     // col 9 G/L Account
         styleText,     // col 10 vendor of record id
         styleText,     // col 11 vendor name
         styleText,     // col 12 reason code
         styleText,     // col 13 reason description
         styleText,     // col 14 location
         styleText,     // col 15 comments
         styleText,     // col 16 rf user
         styleText,     // col 17 cost code
         styleText,     // col 18 function
         styleText,     // col 19 timestamp
      };
      
      styleText = null;
      styleInt = null;
      styleDec = null;
   }
   
   
   
   /* testing main
   public static void main(String[] args) {
      AccPacInvAdj apia = new AccPacInvAdj();
      apia.m_FilePath = "C:\\exp\\";

      Param p1 = new Param();
      p1.value = "5/28/2018";
      Param p2 = new Param();
      p2.value = "6/7/2018";
      Param p3 = new Param();
      p3.value = "";
      Param p4 = new Param();
      p4.value = "";
      Param p5 = new Param();
      p5.value = "";
      Param p6 = new Param();
      p6.value = "";
      Param p7 = new Param();
      p7.value = "";
      
      ArrayList<Param> params = new ArrayList<Param>();
      params.add(p1);
      params.add(p2);
      params.add(p3);
      params.add(p4);
      params.add(p5);
      params.add(p6);
      params.add(p7);
      apia.setParams(params);
      
      try {
         
         // set 2x connections
         java.util.Properties connProps = new java.util.Properties();
         connProps.put("user", "ejd");
         connProps.put("password", "boxer");
         apia.m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33:5444/emery_jensen",connProps);
         
         java.util.Properties connProps2 = new java.util.Properties();
         connProps2.put("user", "dev");
         connProps2.put("password", "Pin3AppL3Pizza");
         apia.m_SageConn = java.sql.DriverManager.getConnection("jdbc:jtds:sqlserver://ACCPAC-SQL:1433",connProps2);
         
         // run
         apia.createReport();
      
      } catch (Exception e) {
         System.out.println("oh no");
         e.printStackTrace();
      }
   }
   */
   
}
