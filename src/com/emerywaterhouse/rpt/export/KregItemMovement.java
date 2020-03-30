/**
 * File: KregItemMovement.java
 * Description: Item movement report for Kreg Tool items.  Used to calculate commisions.
 *
 * @author Jeff Fisher
 *
 * Create Date: 12/21/2010
 * Last Update: $Id: KregItemMovement.java,v 1.1 2010/12/26 13:22:37 jfisher Exp $
 *
 * History:
 *    $Log: KregItemMovement.java,v $
 *    Revision 1.1  2010/12/26 13:22:37  jfisher
 *    Initial add
 *
 */
package com.emerywaterhouse.rpt.export;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class KregItemMovement extends Report
{
   //
   // IDs for the styles used in the worksheet.
   private static final int stText     = 0;
   private static final int stInt      = 1;
   
   //
   // data
   private PreparedStatement m_ItemData;
   
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private HSSFCellStyle[] m_CellStyles;
   
   //
   // workbook entries.
   private HSSFWorkbook m_Wrkbk;
   private HSSFSheet m_Sheet;
   
   //
   // params
   private Date m_StartDate;
   private Date m_EndDate;
   
   public KregItemMovement()
   {
      super();
      m_Wrkbk = new HSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
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
      
      m_StartDate = null;
      m_EndDate = null;
      m_Sheet = null;
      m_Wrkbk = null;      
      m_CellStyles = null;
      
      super.finalize();
   }
   
   /**
    * adds an integer type cell to current row at the specified column in current sheet
    *
    * @param row The row that the cell will be added to.
    * @param col 0-based column number of spreadsheet cell
    * @param style ID of the Excel style to be used to display cell
    * @param value integer value to be stored in cell
    */
   private void addCell(HSSFRow row, int col, int style, int value)
   {
      HSSFCell cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellValue(value);
      cell.setCellStyle(m_CellStyles[style]);
      cell = null;
   }
   
   /**
    * adds a text type cell to current row at the specified column in current sheet
    *
    * @param row The row that the cell will be added to.
    * @param col 0-based column number of spreadsheet cell
    * @param style ID of the Excel style to be used to display cell
    * @param value text value to be stored in cell
    */
   private void addCell(HSSFRow row, int col, int style, String value)
   {
      HSSFCell cell = row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new HSSFRichTextString(value));
      cell.setCellStyle(m_CellStyles[style]);
      cell = null;
   }
   
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      HSSFRow row = null;      
      FileOutputStream outFile = null;
      ResultSet itemData = null;
      int rowNum = 0;
      boolean result = false;
      SimpleDateFormat fmt = new SimpleDateFormat("MM/dd/yyyy");
            
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         rowNum = createCaptions();
         
         m_ItemData.setDate(1, m_StartDate);
         m_ItemData.setDate(2, m_EndDate);         
         itemData = m_ItemData.executeQuery();

         while ( itemData.next() && m_Status == RptServer.RUNNING ) {
            row = createRow(rowNum++);
         
            if ( row != null ) {               
               addCell(row, 0, stText, itemData.getString("cust_nbr"));               
               addCell(row, 1, stText, itemData.getString("item_nbr"));
               addCell(row, 2, stText, itemData.getString("item_descr"));
               addCell(row, 3, stInt, itemData.getInt("qty"));               
               addCell(row, 4, stText, fmt.format(itemData.getDate("invoice_date")));               
            }
         }
  
         m_Wrkbk.write(outFile);
         closeRSet(itemData);

         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      
         log.fatal("[KregItemMovement]", ex);
      }
   
      finally {
         try {
            outFile.close();
         }
      
         catch( Exception e ) {
            log.error("[KregItemMovement]", e);
         }
         
         itemData = null;
         row = null;
         outFile = null;
         fmt = null;        
      }

      return result;
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_ItemData);
   }
   
   /**
    * Sets the captions on the report.
    */
   private short createCaptions()
   {
      HSSFRow row = null;      
      short rowNum = 0;
                  
      if ( m_Sheet == null )
         return 0;
      
      //
      // Create the row for the captions.
      row = m_Sheet.createRow(rowNum);
      
      //
      // Don't need a cell object because no spacing or font changes.
      if ( row != null ) {
         row.createCell(0).setCellValue(new HSSFRichTextString("cust#"));            
         row.createCell(1).setCellValue(new HSSFRichTextString("item#"));
         row.createCell(2).setCellValue(new HSSFRichTextString("item desc"));
         row.createCell(3).setCellValue(new HSSFRichTextString("qty"));
         row.createCell(4).setCellValue(new HSSFRichTextString("inv date"));
      }
    
      return ++rowNum;
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
         log.fatal("[KregItemMovement]", ex);
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
    * 
    * @return The new row of the spreadsheet.
    */
   private HSSFRow createRow(int rowNum)
   {
      HSSFRow row = null;
      
      if ( m_Sheet != null )
         row = m_Sheet.createRow(rowNum);

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
            sql.append("select");
            sql.append("   cust_nbr, item_nbr, item_descr, sum(qty_shipped) qty, invoice_date ");
            sql.append("from inv_dtl ");
            sql.append("where ");
            sql.append("   vendor_nbr = '710209' and item_nbr is not null and ");
            sql.append("   invoice_date between ? and ? ");
            sql.append("group by cust_nbr, item_nbr, item_descr, invoice_date ");
            sql.append("order by cust_nbr, item_nbr, invoice_date");
            
            m_ItemData = m_EdbConn.prepareStatement(sql.toString());
            isPrepared = true;   
         }
         
         catch ( SQLException ex ) {
            log.error("[KregItemMovement]", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("[KregItemMovement] null db connection");
      
      return isPrepared;
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      int pcount = params.size();
      Param param = null;
      SimpleDateFormat fmt = new SimpleDateFormat("dd-MMM-yyyy");
      String tmp = "";
            
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                            
         if ( param.name.equals("startdate") ) {
            tmp = param.value;
            try {
               m_StartDate = new Date(fmt.parse(param.value).getTime());
            } 
            
            catch (ParseException e) {            
               e.printStackTrace();
            }
         }   
         
         if ( param.name.equals("enddate") ) {
            try {
               m_EndDate = new Date(fmt.parse(param.value).getTime());
            } 
            catch (ParseException e) {            
               e.printStackTrace();
            }
         }
      }
          
      m_FileNames.add(String.format("%s-emery-kregtool.xls", tmp.toLowerCase()));
   }
      
   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {      
      HSSFCellStyle styleText = null;        // Text left justified
      HSSFCellStyle styleInt = null;         // Style with 0 decimals
      
      styleText = m_Wrkbk.createCellStyle();      
      styleText.setAlignment(HorizontalAlignment.LEFT);
      
      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);
          
      m_CellStyles = new HSSFCellStyle[] {
         styleText,
         styleInt
      };
      
      styleText = null;
      styleInt = null;
   }
}
