/**
 * File: SubRpt.java
 * Description: Base class for building pieces of a report.  Should be used when a report
 *    has many different configurations.
 *
 * @author Jeffrey Fisher
 *
 * Create Date: 02/02/2006
 * Last Update: $Id: SubRpt.java,v 1.4 2009/02/18 16:53:10 jfisher Exp $
 * 
 * History 
 *    $Log: SubRpt.java,v $
 *    Revision 1.4  2009/02/18 16:53:10  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileOutputStream;


import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public abstract class SubRpt
{
   //
   // References to the main reports spreadsheet objects.
   private XSSFCellStyle m_CSCaption;
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;
   
   //
   // The cell styles for each of the base columns in the spreadsheet.
   protected XSSFCellStyle[] m_CellStyles;
   
   /**
    * default constructor;
    */
   public SubRpt()
   {
      super();
      initCellStyles();
   }
      
   /**
    * 
    * @param wrkbk A reference to the main reports HSSFWorkbook object.
    * @param sheet A reference to the main report HSSFSheet object.
    */
   public SubRpt(XSSFWorkbook wrkbk, XSSFSheet sheet) 
   {
      this();
      
      m_Wrkbk = wrkbk;
      m_Sheet = sheet;      
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
      m_CSCaption = null;
      
      super.finalize();
   }
   
   /**
    * Place holder method.  Descendant class should implement this method to build the output
    * file.
    * 
    * @param out The output stream the report is going to go on.
    * @param rowNum The row number to start the caption on.
    * 
    * @return true if the build was successful, false if not.
    */
   public abstract boolean build(FileOutputStream out, int rowNum);
   
   /**
    * Place holder for the method that builds the sql statement for the sub report.
    * 
    * @return A string containing the sql statement.
    */
   public String buildSql()
   {
      return "";
   }
   
   /**
    * Place holder for the sub report caption creation routine.  Descendant classes should
    * override this method.
    * 
    * @param rowNum The row number to start the caption on.
    * @return The row number that the caption finished on.
    */
   public int createCaptions(int rowNum)
   {
      return rowNum;
   }
   
   /**
    * Creates a caption cell for the given row.  Create the style if it's not already set and 
    * adds the caption to the cell.
    * 
    * @param row A reference to the row object that the cell will be in.
    * @param col The column for the cell.
    * @param caption The caption for the cell.
    * 
    * @return The newly created cell or null if the row was null.
    */
   protected XSSFCell createCaptionCell(XSSFRow row, int col, String caption)
   {
      XSSFCell cell = null;
      
      if ( row != null ) {
         if ( m_CSCaption == null )
            createDefaultCaptionStyle();
         
         cell = row.createCell(col);
         cell.setCellType(CellType.STRING);
         cell.setCellStyle(m_CSCaption);
         cell.setCellValue(new XSSFRichTextString(caption != null ? caption : ""));
      }
      
      return cell;
   }
   
   /**
    * Creates the default font for the caption of the sub report.  This can be
    * overridden so that a different style and font are used.    
    */
   protected void createDefaultCaptionStyle()
   {
      XSSFFont font = m_Wrkbk.createFont();
      
      try {         
         font.setFontHeightInPoints((short)10);
         font.setFontName("Arial");
         font.setBold(true);
         
         m_CSCaption = m_Wrkbk.createCellStyle();
         m_CSCaption.setFont(font);
         m_CSCaption.setAlignment(HorizontalAlignment.CENTER);
      }
         
      finally {
         font = null;
      }
   }
   
   /**
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    * 
    * @return The formatted row of the spreadsheet.
    */
   protected XSSFRow createDataRow(int rowNum)
   {
      XSSFRow row = null;
      XSSFCell cell = null;
      int colCnt = m_CellStyles.length;
      
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
    * @return Returns the sheet.
    */
   public XSSFSheet getSheet()
   {
      return m_Sheet;
   }

   /**
    * @return Returns the wrkbk.
    */
   public XSSFWorkbook getWrkbk()
   {
      return m_Wrkbk;
   }

   /**
    * Method called by descendant classes to initialize the cell style array.    
    */
   protected abstract void initCellStyles();
   
   /**
    * Sets the cell style for the title and captions.
    * @param style
    */
   public void setCaptionStyle(XSSFCellStyle style)
   {
      if ( style != null )
         m_CSCaption = style;
   }
      
   /**
    * @param sheet The sheet to set.
    */
   public void setSheet(XSSFSheet sheet)
   {
      m_Sheet = sheet;
   }   

   /**
    * @param wrkbk The wrkbk to set.
    */
   public void setWrkbk(XSSFWorkbook wrkbk)
   {
      m_Wrkbk = wrkbk;
   }
   
   
}
