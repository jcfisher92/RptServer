/**
 * 
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

/**
 *
 */
public class CustomerAssortment extends Report 
{
   private static int charWidth = 295;
   private static short fontHeight = 11;
   private static int A = 0;
   private static int B = 1;
   private static int C = 2;
   private static int D = 3;
   private static int E = 4;
   private static int F = 5;
   private static int G = 6;
   private static int H = 7;
   private static int I = 8;
   private static int J = 9;
   private static int K = 10;
   private static int L = 11;
   private static int M = 12;
   private static int N = 13;
   private static int O = 14;
   private static int P = 15;
   private static int Q = 16;
   private static int R = 17;
   private static int S = 18;
   private static int T = 19;
   private static int U = 20;
   private static int V = 21;
   private static int W = 22;
   private static int X = 23;
   private static int Y = 24;
   private static int Z = 25;
   private static int AA = 26;
   private static int AB = 27;
   private static int AC = 28;
   private static int AD = 29;
   private static int AE = 30;
   private static int AF = 31;
   private static int AG = 32;
   private static int AH = 33;
   private static int AI = 34;
   private static int AJ = 35;
   private static int AK = 36;
   private static int AL = 37;
   private static int AM = 38;
   
   private XSSFWorkbook m_Workbook;
   private XSSFSheet m_Sheet1;
   private XSSFSheet m_Sheet2;
   private XSSFSheet m_Sheet3;
   private XSSFSheet m_Sheet4;
   private XSSFSheet m_Sheet5;
   private XSSFSheet m_Sheet6;
   
   private XSSFCellStyle m_StyleHdrL;
   private XSSFCellStyle m_StyleHdrC;
   private XSSFCellStyle m_StyleTxtC;      // Text centered
   private XSSFCellStyle m_StyleTxtL;      // Text left justified
   private XSSFCellStyle m_StyleTxtR;      // Text right justified
   private XSSFCellStyle m_StyleTxtRI;     // Text right justified & italic
   private XSSFCellStyle m_StyleInt;       // Style with 0 decimals
   private XSSFCellStyle m_StyleDouble;    // numeric #,##0.00
   private XSSFCellStyle m_StyleDate;      // mm/dd/yyyy
   private XSSFCellStyle m_StyleFillGrey;
   private XSSFCellStyle m_StyleFillBlack;
   private XSSFCellStyle m_StyleLtGreenPct;
   private XSSFCellStyle m_StyleLtGreenCur;
   private XSSFCellStyle m_StyleLtRedPct;
   private XSSFCellStyle m_StyleLtRedCur;
   private XSSFCellStyle m_StylePct;
   private XSSFCellStyle m_StylePctB;
   private XSSFCellStyle m_StyleCur;
   private XSSFCellStyle m_StyleCurB;
   
   private String m_CustId;
   private int m_SkuTot;
   private double m_DollarTot;
   
   //
   // Counters for totals of units used in the summary.
   private int m_DiscCount;
   private int m_PrevDiscCount;
   private int m_LikeCount;   
   private int m_MatchCount;
   private int m_OMLikeCount;
   private int m_OMMatchCount;
      
   //
   // COGS vars
   private double m_CurMatchCogs;
   private double m_NewMatchCogs;
   private double m_CurLikeCogs;
   private double m_NewLikeCogs;
   private double m_DiscCogs;
   private double m_PrevDiscCogs;
      
   private double m_LikeRetTot;
   private double m_MatchRetTot;
   private double m_CompRetTot;
   private double m_PremRetTot;
   private double m_MktRetTot;
      
   PreparedStatement m_CustData;
   PreparedStatement m_DiscoItems;
   PreparedStatement m_LikeItems;
   PreparedStatement m_MatchedItems;
   PreparedStatement m_OMItems;
   PreparedStatement m_SumTax;
     
   public CustomerAssortment()
   {
      super();
      
      m_DiscCount = 0;
      m_MatchCount = 0;
      m_LikeCount = 0;      
      m_OMLikeCount = 0;
      m_OMMatchCount = 0;
      
      m_SkuTot = 0;
      m_DollarTot = 0.0;
   }
   
   /**
    * adds a numeric type cell to current row at col p_Col in current sheet
    *
    * @param col 0-based column number of spreadsheet cell
    * @param value numeric value to be stored in cell
    * @param style Excel style to be used to display cell
    */
   private Cell addCell(Row row, int col, double value, CellStyle style)
   {
      Cell cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);      
      cell.setCellStyle(style);
      cell.setCellValue(value);
      
      return cell;
   }
   
   /**
    * adds a numeric type cell to current row at col p_Col in current sheet
    *
    * @param col 0-based column number of spreadsheet cell
    * @param value integer value to be stored in cell
    * @param style Excel style to be used to display cell
    */
   private Cell addCell(Row row, int col, int value, CellStyle style)
   {
      Cell cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellStyle(style);
      cell.setCellValue(value);
      
      return cell;
   }
   
   /**
    * adds a text type cell to current row at col p_Col in current sheet
    *
    * @param col 0-based column number of spreadsheet cell
    * @param value text value to be stored in cell
    * @param style Excel style to be used to display cell
    */
   private Cell addCell(Row row, int col, String value, CellStyle style)
   {
      Cell cell = row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(value != null ? value : "" ));
      cell.setCellStyle(style);
      
      return cell;
   }
   
   /**
    * 
    * @param row
    * @param col
    * @param value
    * @param style
    * @return
    */
   private Cell addCellFormula(Row row, int col, String value, CellStyle style)
   {
      Cell cell = row.createCell(col);
      cell.setCellType(CellType.FORMULA);
      cell.setCellFormula(value);
      cell.setCellStyle(style);
      
      return cell;
   }
   
   /**
    * Adds the headers to the discontinued items tab
    * @return The last row number used.
    */
   private int addDiscoItemsHeader()
   {
      int rowNum = 0;
      int col = 0;
      Row row = m_Sheet4.createRow(rowNum);
      
      m_Sheet4.setColumnWidth(col, (25 * charWidth));
      addCell(row, col, "Taxonomy 1", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (25 * charWidth));
      addCell(row, col, "Taxonomy 2", m_StyleTxtC);
            
      m_Sheet4.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Item#", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Customer SKU", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Emery UPC", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Mfg#", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (50 * charWidth));
      addCell(row, col, "Item Description", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Vendor", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "R12 Units", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Current Cost", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Cost Total", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Current Retail", m_StyleTxtC);
            
      m_Sheet4.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Retail Total", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Current Margin", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Stock Pack", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Current OM", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Current UOM", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (10 * charWidth));
      addCell(row, col, "NBC", m_StyleTxtC);
      
      m_Sheet4.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Match", m_StyleTxtC);
      
      return rowNum;
   }
   
   /**
    * Adds the headers to the like items tab
    * @return The last row number used.
    */
   private int addLikeItemsHeader()
   {
      int rowNum = 0;
      int col = 0;
      Row row = m_Sheet3.createRow(rowNum);
      
      m_Sheet3.setColumnWidth(col, (25 * charWidth));
      addCell(row, col, "Taxonomy 1", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (25 * charWidth));
      addCell(row, col, "Taxonomy 2", m_StyleTxtC);
            
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Item#", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Customer SKU", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Emery UPC", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Mfg#", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (50 * charWidth));
      addCell(row, col, "Item Description", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Vendor", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD Item#", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD UPC", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD Mfg", m_StyleTxtC);
            
      m_Sheet3.setColumnWidth(++col, (50 * charWidth));
      addCell(row, col, "EJD Description", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD Vendor", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "R12 Units", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Cost", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New Cost", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Cost Total", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New Cost Total", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Cost Variance", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Retail", m_StyleTxtC);
            
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Retail Total", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Margin", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Competitve Retail", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Competitve Retail Total", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Competitive Margin %", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Market Retail", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Market Retail Total", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Market Margin %", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Premium Retail", m_StyleTxtC);
            
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Premium Retail Total", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Premium Margin %", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New OM", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New UOM", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (10 * charWidth));
      addCell(row, col, "NBC", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Match", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Stock Pack", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current OM", m_StyleTxtC);
      
      m_Sheet3.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current UOM", m_StyleTxtC);
      
      return rowNum;
   }
   
   /**
    * Adds the matched items headers
    * @return The last row number used.
    */
   private int addMatchedItemsHeader()
   {
      int rowNum = 0;
      int col = 0;   
      Row row = m_Sheet2.createRow(rowNum);
      
      m_Sheet2.setColumnWidth(col, (25 * charWidth));
      addCell(row, col, "Taxonomy 1", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (25 * charWidth));
      addCell(row, col, "Taxonomy 2", m_StyleTxtC);
            
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Item#", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Customer SKU", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Emery UPC", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Mfg#", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (50 * charWidth));
      addCell(row, col, "Item Description", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Vendor", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD Item#", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD UPC", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD Mfg", m_StyleTxtC);
            
      m_Sheet2.setColumnWidth(++col, (50 * charWidth));
      addCell(row, col, "EJD Description", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD Vendor", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "R12 Units", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Cost", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New Cost", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Cost Total", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New Cost Total", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Cost Variance", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Retail", m_StyleTxtC);
            
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Retail Total", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Margin", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Competitve Retail", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Competitve Retail Total", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Competitive Margin %", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Market Retail", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Market Retail Total", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Market Margin %", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Premium Retail", m_StyleTxtC);
            
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Premium Retail Total", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Premium Margin %", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New OM", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New UOM", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (10 * charWidth));
      addCell(row, col, "NBC", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Match", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Stock Pack", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current OM", m_StyleTxtC);
      
      m_Sheet2.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current UOM", m_StyleTxtC);
      
      return rowNum;
   }
   
   /**
    * Adds the headers to the OM items tab
    * @return The last row number used.
    */
   private int addOMItemsHeader()
   {
      int rowNum = 0;
      int col = 0;
      Row row = m_Sheet5.createRow(rowNum);
      
      m_Sheet5.setColumnWidth(col, (25 * charWidth));
      addCell(row, col, "Taxonomy 1", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (25 * charWidth));
      addCell(row, col, "Taxonomy 2", m_StyleTxtC);
            
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Item#", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Customer SKU", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Emery UPC", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Mfg#", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (50 * charWidth));
      addCell(row, col, "Item Description", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Vendor", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD Item#", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD UPC", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD Mfg", m_StyleTxtC);
            
      m_Sheet5.setColumnWidth(++col, (50 * charWidth));
      addCell(row, col, "EJD Description", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "EJD Vendor", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Stock Pack", m_StyleTxtC);
  
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current OM", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New OM", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "OM Difference", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current UOM", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New UOM", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "NBC", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "R12 Units", m_StyleTxtC);
            
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Cost", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New Cost", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Cost Total", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "New Cost Total", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Cost Variance", m_StyleTxtC);
     
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Retail", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Retail Total", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Current Margin", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Competitve Retail", m_StyleTxtC);
            
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Competitive Retail Total", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Competitive Margin %", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Market Retail", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Market Retail Total", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (10 * charWidth));
      addCell(row, col, "Market Margin %", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Premium Retail", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Premium Retail Total", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "Premium Margin %", m_StyleTxtC);
      
      m_Sheet5.setColumnWidth(++col, (20 * charWidth));
      addCell(row, col, "match", m_StyleTxtC);
      
      return rowNum;
   }
   
   /**
    * Adds the headers to the Summary Tax Match tab
    * @return The last row number used.
    */
   private int addSummaryTaxHeader()
   {
      int rowNum = 0;
      int col = 0;
      Row row = m_Sheet6.createRow(rowNum);
      
      m_Sheet6.setColumnWidth(col, (30 * charWidth));
      addCell(row, col, "Taxonomy 1", m_StyleTxtC);
      
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "R12 Units", m_StyleTxtC);
            
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Current Cost Total", m_StyleTxtC);
      
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "New Cost Total", m_StyleTxtC);
            
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Current Retail Total", m_StyleTxtC);
      
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Current Margin %", m_StyleTxtC);
            
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Competitive Retail Total", m_StyleTxtC);
      
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Competitive Margin %", m_StyleTxtC);
            
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Market Retail Total", m_StyleTxtC);
      
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, " Market Margin %", m_StyleTxtC);
      
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Premium Retail Total", m_StyleTxtC);
            
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Premium Margin %", m_StyleTxtC);
      
      m_Sheet6.setColumnWidth(++col, (15 * charWidth));
      addCell(row, col, "Match", m_StyleTxtC);
      
      return rowNum;
   }
   
   /**
    * Builds the discontinued items tab/sheet
    * @throws SQLException 
    */
   private void buildDiscoItems() throws SQLException
   {
      int rowNum = addDiscoItemsHeader();
      int r12Units = 0;
      Row row = null;
      ResultSet rs = null;
      double costTotal = 0;
      String match = null;;
      
      m_DiscoItems.setString(1, m_CustId);
      m_DiscoItems.setString(2, m_CustId);
      m_DiscoItems.setString(3, m_CustId);
      m_DiscoItems.setString(4, m_CustId);
      
      try {
         m_CurAction = "retrieving discontinued items sheet data";
         rs = m_DiscoItems.executeQuery();
         
         m_CurAction = "building discontinuted items sheet";
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            costTotal = rs.getDouble("cur_cost_total");
            match = rs.getString("match");
            r12Units = rs.getInt("r12_units");
            
            row = m_Sheet4.createRow(++rowNum);
                        
            addCell(row, A, rs.getString("taxonomy_1"), m_StyleTxtL);                       
            addCell(row, B, rs.getString("taxonomy_2"), m_StyleTxtL);
            addCell(row, C, rs.getString("item_id"), m_StyleTxtL);
            addCell(row, D, rs.getString("customer_sku"), m_StyleTxtL);
            addCell(row, E, rs.getString("emery_upc"), m_StyleTxtL);
            addCell(row, F, rs.getString("mfrnum"), m_StyleTxtL);
            addCell(row, G, rs.getString("item_description"), m_StyleTxtL);
            addCell(row, H, rs.getString("cur_vendor"), m_StyleTxtL);
            addCell(row, I, r12Units, m_StyleInt);
            addCell(row, J, rs.getDouble("cur_cost"), m_StyleDouble);            
            addCell(row, K, costTotal, m_StyleDouble);            
            addCell(row, L, rs.getDouble("cur_retail"), m_StyleDouble);
            addCell(row, M, rs.getDouble("cur_ret_tot"), m_StyleDouble);
            addCell(row, N, rs.getDouble("cur_margin"), m_StyleDouble);
            addCell(row, O, rs.getInt("stock_pack"), m_StyleInt);
            addCell(row, P, rs.getInt("cur_om"), m_StyleInt);
            addCell(row, Q, rs.getString("cur_uom"), m_StyleTxtL);
            addCell(row, R, rs.getString("nbc"), m_StyleTxtL);
            addCell(row, S, match, m_StyleTxtL);
            
            if ( r12Units > 0 ) {
               if ( match != null && match.length() > 0 ) {
                  if ( match.equalsIgnoreCase("discontinued") ) {
                     m_DiscCount++;
                     m_DiscCogs += costTotal;
                  }
                  else {
                     if ( match.equalsIgnoreCase("previously discontinued") ) {
                        m_PrevDiscCount++;
                        m_PrevDiscCogs += costTotal;
                     }
                  }
               }
            }
                        
            match = null;            
            costTotal = 0.0;
            r12Units = 0;
         }
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
         row = null;
      }
   }
   
   /**
    * Builds the like items tab/sheet
    * @throws SQLException 
    */
   private void buildLikeItems() throws SQLException
   {
      int rowNum = addLikeItemsHeader();
      int r12Units = 0;
      
      double curCostTot = 0.0;
      double curRetTot = 0.0;
      double newCostTot = 0.0;
      double premRetTot = 0.0;
      double mktRetTot = 0.0;
      double compRetTot = 0.0;
      Row row = null;
      ResultSet rs = null;
      
      m_LikeCount = 0;
      m_LikeItems.setString(1, m_CustId);
      m_LikeItems.setString(2, m_CustId);
      m_LikeItems.setString(3, m_CustId);
      m_LikeItems.setString(4, m_CustId);
      m_LikeItems.setString(5, m_CustId);
      
      try {
         m_CurAction = "retrieving like items sheet data";
         rs = m_LikeItems.executeQuery();
         
         m_CurAction = "building like items sheet";
         while ( rs.next() && m_Status == RptServer.RUNNING ) {            
            curCostTot = rs.getDouble("cur_cost_total");
            newCostTot = rs.getDouble("new_cost_total");
            curRetTot = rs.getDouble("cur_ret_tot");
            premRetTot = rs.getDouble("premium_ret_tot");
            mktRetTot = rs.getDouble("market_ret_tot");            
            compRetTot = rs.getDouble("comp_ret_tot");
            r12Units = rs.getInt("r12_units");
            
            if ( newCostTot <= 0 )
               newCostTot = curCostTot;
            
            if ( mktRetTot <= 0 )
               mktRetTot = curRetTot;
            
            if ( compRetTot <= 0 )
               compRetTot = curRetTot;
            
            if ( premRetTot <= 0 )
               premRetTot = curRetTot;
            
            row = m_Sheet3.createRow(++rowNum);
            
            addCell(row, A, rs.getString("taxonomy_1"), m_StyleTxtL);                       
            addCell(row, B, rs.getString("taxonomy_2"), m_StyleTxtL);
            addCell(row, C, rs.getString("item_id"), m_StyleTxtL);
            addCell(row, D, rs.getString("customer_sku"), m_StyleTxtL);
            addCell(row, E, rs.getString("emery_upc"), m_StyleTxtL);
            addCell(row, F, rs.getString("mfrnum"), m_StyleTxtL);
            addCell(row, G, rs.getString("item_description"), m_StyleTxtL);
            addCell(row, H, rs.getString("cur_vendor"), m_StyleTxtL);
            addCell(row, I, rs.getString("ejd_item"), m_StyleTxtL);
            addCell(row, J, rs.getString("ejd_upc"), m_StyleTxtL);
            addCell(row, K, rs.getString("ejd_mfrnum"), m_StyleTxtL);
            addCell(row, L, rs.getString("ejd_description"), m_StyleTxtL);
            addCell(row, M, rs.getString("ejd_vendor_name"), m_StyleTxtL);
            addCell(row, N, r12Units, m_StyleInt);
            addCell(row, O, rs.getDouble("cur_cost"), m_StyleDouble);
            addCell(row, P, rs.getDouble("new_cost"), m_StyleDouble);
            addCell(row, Q, curCostTot, m_StyleDouble);
            addCell(row, R, newCostTot, m_StyleDouble);
            addCell(row, S, rs.getDouble("cost_variance"), m_StyleDouble);
            addCell(row, T, rs.getDouble("cur_retail"), m_StyleDouble);
            addCell(row, U, curRetTot, m_StyleDouble);
            addCell(row, V, rs.getDouble("cur_margin"), m_StyleDouble);
            addCell(row, W, rs.getDouble("comp_ret"), m_StyleDouble);
            addCell(row, X, compRetTot, m_StyleDouble);
            addCell(row, Y, rs.getDouble("comp_margin_pct"), m_StyleDouble);
            addCell(row, Z, rs.getDouble("market_ret"), m_StyleDouble);
            addCell(row, AA, mktRetTot, m_StyleDouble);
            addCell(row, AB, rs.getDouble("market_margin_pct"), m_StyleDouble);
            addCell(row, AC, rs.getDouble("premium_ret"), m_StyleDouble);
            addCell(row, AD, premRetTot, m_StyleDouble);
            addCell(row, AE, rs.getDouble("premium_margin_pct"), m_StyleDouble);
            addCell(row, AF, rs.getInt("new_om"), m_StyleInt);
            addCell(row, AG, rs.getString("new_uom"), m_StyleTxtL);
            addCell(row, AH, rs.getString("nbc"), m_StyleTxtL);
            addCell(row, AI, rs.getString("match"), m_StyleTxtL);
            addCell(row, AJ, rs.getInt("stock_pack"), m_StyleInt);
            addCell(row, AK, rs.getInt("cur_om"), m_StyleInt);
            addCell(row, AL, rs.getString("cur_uom"), m_StyleTxtL);
            
            if ( r12Units > 0 ) {
               m_LikeCount++;
               m_CurLikeCogs += curCostTot;
               m_NewLikeCogs += newCostTot;
               
               m_LikeRetTot += curRetTot;
               m_CompRetTot += compRetTot;
               m_PremRetTot += premRetTot;
               m_MktRetTot += mktRetTot;
            }
            
            curCostTot = 0.0;
            newCostTot = 0.0;
            curRetTot = 0.0;
            compRetTot = 0.0;
            mktRetTot = 0.0;
            premRetTot = 0.0;
            r12Units = 0;
         }
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
         row = null;
      }      
   }
   
   /**
    * Builds the matched items tab/sheet
    * @throws SQLException 
    */
   private void buildMatchedItems() throws SQLException
   {
      int rowNum = addMatchedItemsHeader();      
      int r12Units = 0;
      double curCostTot = 0.0;
      double curRetTot = 0.0;
      double newCostTot = 0.0;
      double premRetTot = 0.0;
      double mktRetTot = 0.0;
      double compRetTot = 0.0;
      
      Row row = null;
      ResultSet rs = null;
      
      m_MatchedItems.setString(1, m_CustId);
      m_MatchedItems.setString(2, m_CustId);
      m_MatchedItems.setString(3, m_CustId);
      m_MatchedItems.setString(4, m_CustId);
      m_MatchedItems.setString(5, m_CustId);
      
      try {
         m_CurAction = "retrieving matched items data";
         rs = m_MatchedItems.executeQuery();
      
         m_CurAction = "building matched items sheet";
         while ( rs.next() && m_Status == RptServer.RUNNING ) {            
            row = m_Sheet2.createRow(++rowNum);            
            curCostTot = rs.getDouble("cur_cost_total");            
            newCostTot = rs.getDouble("new_cost_total");
            curRetTot = rs.getDouble("cur_ret_tot");
            premRetTot = rs.getDouble("premium_ret_tot");
            mktRetTot = rs.getDouble("market_ret_tot");            
            compRetTot = rs.getDouble("comp_ret_tot");
            r12Units = rs.getInt("r12_units");
            
            if ( newCostTot <= 0 )
               newCostTot = curCostTot;
            
            if ( mktRetTot <= 0 )
               mktRetTot = curRetTot;
            
            if ( compRetTot <= 0 )
               compRetTot = curRetTot;
            
            if ( premRetTot <= 0 )
               premRetTot = curRetTot;
            
            addCell(row, A, rs.getString("taxonomy_1"), m_StyleTxtL);                       
            addCell(row, B, rs.getString("taxonomy_2"), m_StyleTxtL);
            addCell(row, C, rs.getString("item_id"), m_StyleTxtL);
            addCell(row, D, rs.getString("customer_sku"), m_StyleTxtL);
            addCell(row, E, rs.getString("emery_upc"), m_StyleTxtL);
            addCell(row, F, rs.getString("mfrnum"), m_StyleTxtL);
            addCell(row, G, rs.getString("item_description"), m_StyleTxtL);
            addCell(row, H, rs.getString("cur_vendor"), m_StyleTxtL);
            addCell(row, I, rs.getString("ejd_item"), m_StyleTxtL);
            addCell(row, J, rs.getString("ejd_upc"), m_StyleTxtL);
            addCell(row, K, rs.getString("ejd_mfrnum"), m_StyleTxtL);
            addCell(row, L, rs.getString("ejd_description"), m_StyleTxtL);
            addCell(row, M, rs.getString("ejd_vendor_name"), m_StyleTxtL);
            addCell(row, N, r12Units, m_StyleInt);
            addCell(row, O, rs.getDouble("cur_cost"), m_StyleDouble);
            addCell(row, P, rs.getDouble("new_cost"), m_StyleDouble);
            addCell(row, Q, curCostTot, m_StyleDouble);
            addCell(row, R, newCostTot, m_StyleDouble);
            addCell(row, S, rs.getDouble("cost_variance"), m_StyleDouble);
            addCell(row, T, rs.getDouble("cur_retail"), m_StyleDouble);
            addCell(row, U, curRetTot, m_StyleDouble);
            addCell(row, V, rs.getDouble("cur_margin"), m_StyleDouble);
            addCell(row, W, rs.getDouble("comp_ret"), m_StyleDouble);
            addCell(row, X, compRetTot, m_StyleDouble);
            addCell(row, Y, rs.getDouble("comp_margin_pct"), m_StyleDouble);
            addCell(row, Z, rs.getDouble("market_ret"), m_StyleDouble);
            addCell(row, AA, mktRetTot, m_StyleDouble);
            addCell(row, AB, rs.getDouble("market_margin_pct"), m_StyleDouble);
            addCell(row, AC, rs.getDouble("premium_ret"), m_StyleDouble);
            addCell(row, AD, premRetTot, m_StyleDouble);
            addCell(row, AE, rs.getDouble("premium_margin_pct"), m_StyleDouble);
            addCell(row, AF, rs.getInt("new_om"), m_StyleInt);
            addCell(row, AG, rs.getString("new_uom"), m_StyleTxtL);
            addCell(row, AH, rs.getString("nbc"), m_StyleTxtL);
            addCell(row, AI, rs.getString("match"), m_StyleTxtL);
            addCell(row, AJ, rs.getInt("stock_pack"), m_StyleInt);
            addCell(row, AK, rs.getInt("cur_om"), m_StyleInt);
            addCell(row, AL, rs.getString("cur_uom"), m_StyleTxtL);
            
            if ( r12Units > 0 ) {
               m_MatchCount++;
               m_CurMatchCogs += curCostTot;
               m_NewMatchCogs += newCostTot;
               m_MatchRetTot += curRetTot;
               
               m_CompRetTot += compRetTot;
               m_PremRetTot += premRetTot;
               m_MktRetTot += mktRetTot;
            }
            
            curCostTot = 0.0;
            newCostTot = 0.0;
            curRetTot = 0.0;
            compRetTot = 0.0;
            mktRetTot = 0.0;
            premRetTot = 0.0;
            r12Units = 0;
         }
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
         row = null;
      }
   }
   
   /**
    * Builds the om items tab/sheet
    * @throws SQLException 
    */
   private void buildOMItems() throws SQLException
   {
      int rowNum = addOMItemsHeader();
      Row row = null;
      ResultSet rs = null;
      double curCostTot = 0;
      double newCostTot = 0;
      String match = null;
      
      m_OMItems.setString(1, m_CustId);
      m_OMItems.setString(2, m_CustId);
      m_OMItems.setString(3, m_CustId);
      m_OMItems.setString(4, m_CustId);
      m_OMItems.setString(5, m_CustId);
      
      try {
         m_CurAction = "retrieving om items data";
         rs = m_OMItems.executeQuery();
         
         m_CurAction = "building matched items sheet";
         
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            curCostTot = rs.getDouble("cur_cost_tot");
            newCostTot = rs.getDouble("new_cost_tot");
            match = rs.getString("match");
            
            if ( newCostTot <= 0 )
               newCostTot = curCostTot;
            
            row = m_Sheet5.createRow(++rowNum);
            
            addCell(row, A, rs.getString("taxonomy_1"), m_StyleTxtL);                       
            addCell(row, B, rs.getString("taxonomy_2"), m_StyleTxtL);
            addCell(row, C, rs.getString("item_id"), m_StyleTxtL);
            addCell(row, D, rs.getString("customer_sku"), m_StyleTxtL);
            addCell(row, E, rs.getString("emery_upc"), m_StyleTxtL);
            addCell(row, F, rs.getString("mfrnum"), m_StyleTxtL);
            addCell(row, G, rs.getString("item_description"), m_StyleTxtL);
            addCell(row, H, rs.getString("cur_vendor"), m_StyleTxtL);
            addCell(row, I, rs.getString("ejd_item"), m_StyleTxtL);
            addCell(row, J, rs.getString("ejd_upc"), m_StyleTxtL);
            addCell(row, K, rs.getString("ejd_mfrnum"), m_StyleTxtL);
            addCell(row, L, rs.getString("ejd_description"), m_StyleTxtL);
            addCell(row, M, rs.getString("ejd_vendor_name"), m_StyleTxtL);
            addCell(row, N, rs.getInt("stock_pack"), m_StyleInt);
            addCell(row, O, rs.getInt("cur_om"), m_StyleInt);
            addCell(row, P, rs.getInt("new_om"), m_StyleInt);
            addCell(row, Q, rs.getInt("om_diff"), m_StyleInt);
            addCell(row, R, rs.getString("cur_uom"), m_StyleTxtC);
            addCell(row, S, rs.getString("new_uom"), m_StyleTxtC);
            addCell(row, T, rs.getString("nbc"), m_StyleTxtC);
            addCell(row, U, rs.getInt("r12_units"), m_StyleInt);            
            addCell(row, V, rs.getDouble("cur_cost"), m_StyleDouble);
            addCell(row, W, rs.getDouble("new_cost"), m_StyleDouble);
            addCell(row, X, curCostTot, m_StyleDouble);
            addCell(row, Y, newCostTot, m_StyleDouble);
            addCell(row, Z, rs.getDouble("cost_variance"), m_StyleDouble);
            addCell(row, AA, rs.getDouble("cur_retail"), m_StyleDouble);
            addCell(row, AB, rs.getDouble("cur_ret_tot"), m_StyleDouble);
            addCell(row, AC, rs.getDouble("cur_margin"), m_StyleDouble);
            addCell(row, AD, rs.getDouble("comp_ret"), m_StyleDouble);
            addCell(row, AE, rs.getDouble("comp_ret_tot"), m_StyleDouble);
            addCell(row, AF, rs.getDouble("comp_margin_pct"), m_StyleDouble);
            addCell(row, AG, rs.getDouble("market_ret"), m_StyleDouble);
            addCell(row, AH, rs.getDouble("market_ret_tot"), m_StyleDouble);
            addCell(row, AI, rs.getDouble("market_margin_pct"), m_StyleDouble);
            addCell(row, AJ, rs.getDouble("premium_ret"), m_StyleDouble);
            addCell(row, AK, rs.getDouble("premium_ret_tot"), m_StyleDouble);
            addCell(row, AL, rs.getDouble("premium_margin_pct"), m_StyleDouble);
            addCell(row, AM, match, m_StyleTxtL);
            
            
            if ( match != null && match.length() > 0 ) {
               if ( match.equalsIgnoreCase("matched item") )
                  m_OMMatchCount++;
               else {
                  if ( match.equalsIgnoreCase("like item") )
                     m_OMLikeCount++;
               }
            }
            
            curCostTot = 0.0;
            newCostTot = 0.0;
            match = null;
         }
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
         row = null;
      }      
   }
   
   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      boolean result = false;
      
      try {
         setupWorkbook();
         
         if (m_Status == RptServer.RUNNING )
            buildMatchedItems();
         
         if (m_Status == RptServer.RUNNING )
            buildLikeItems();
         
         if (m_Status == RptServer.RUNNING )
            buildDiscoItems();
         
         if (m_Status == RptServer.RUNNING )
            buildOMItems();
         
         if (m_Status == RptServer.RUNNING )
            buildSummaryTax();
         
         if (m_Status == RptServer.RUNNING )
            buildSummary();
         
         m_Workbook.write(outFile);
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[CustomerAssortment]", ex);
      }

      finally {
         ;
      }
      
      return result;
   }
   
   /**
    * Builds the summary tab/sheet
    * @throws SQLException 
    */
   private void buildSummary() throws SQLException
   {
      Row row = m_Sheet1.createRow(1);      
      ResultSet rs = null;
      String custName = "";
      String custType = "";
      
      //
      // Conditional formatting rules for this sheet.
      XSSFColor ltGreen = new XSSFColor(new java.awt.Color(198, 239, 206));
      XSSFColor ltRed = new XSSFColor(new java.awt.Color(242, 220, 219));            
      XSSFSheetConditionalFormatting sheetCF = m_Sheet1.getSheetConditionalFormatting();
      
      ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "0");
      PatternFormatting fillR1 = rule1.createPatternFormatting();
      fillR1.setFillBackgroundColor(ltGreen);
      fillR1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
      FontFormatting fontR1 = rule1.createFontFormatting();
      fontR1.setFontColorIndex(IndexedColors.GREEN.getIndex());
            
      ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, "0");
      PatternFormatting fillR2 = rule2.createPatternFormatting();
      fillR2.setFillBackgroundColor(ltRed);
      fillR2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
      FontFormatting fontR2 = rule2.createFontFormatting();
      fontR2.setFontColorIndex(IndexedColors.RED.getIndex());
      
      CellRangeAddress[] regions = {
         CellRangeAddress.valueOf("G10:G12"),
         CellRangeAddress.valueOf("H10:H12"),
         CellRangeAddress.valueOf("B24:B26"),
         CellRangeAddress.valueOf("C31:C33"),
         CellRangeAddress.valueOf("G31:G33"),
         CellRangeAddress.valueOf("H31:H33"),
      };

      sheetCF.addConditionalFormatting(regions, rule1, rule2);
      
      m_Sheet1.setColumnWidth(A, (28 * charWidth));
      m_Sheet1.setColumnWidth(B, (28 * charWidth));
      m_Sheet1.setColumnWidth(C, (22 * charWidth));
      m_Sheet1.setColumnWidth(D, (22 * charWidth));
      m_Sheet1.setColumnWidth(E, (22 * charWidth));
      m_Sheet1.setColumnWidth(F, (22 * charWidth));
      m_Sheet1.setColumnWidth(G, (12 * charWidth));
      m_Sheet1.setColumnWidth(H, (12 * charWidth));
      
      m_CurAction = "building summary sheet";
      
      try {
         m_CustData.setString(1, m_CustId);
         rs = m_CustData.executeQuery();
         
         if ( rs.next() ) {
            custName = rs.getString("name");
            custType = rs.getString("cust_type");
         }
         
         DbUtils.closeDbConn(null, null, rs);
         
         m_SkuTot = m_DiscCount + m_PrevDiscCount + m_LikeCount + m_MatchCount;// + m_OMCount;         
         m_DollarTot = m_DiscCogs + m_PrevDiscCogs + m_CurLikeCogs + m_CurMatchCogs;// + m_CurOMCogs;
                  
         //
         // Cust info group
         addCell(row, A, "Retailer", m_StyleHdrL);
         addCell(row, B, String.format("%s #%s", custName, m_CustId), m_StyleTxtC);
         
         row = m_Sheet1.createRow(2);
         addCell(row, A, "Retailer Type", m_StyleHdrL);
         addCell(row, B, custType, m_StyleTxtL);
         
         row = m_Sheet1.createRow(3);
         addCell(row, A, "Distribution Center", m_StyleHdrL);
         addCell(row, B, "Wilton", m_StyleTxtL);
         
         row = m_Sheet1.createRow(4);
         addCell(row, A, "Delivery Schedule", m_StyleHdrL);
         addCell(row, B, "-", m_StyleTxtL);
         
         row = m_Sheet1.createRow(5);
         addCell(row, A, "Total SKU Count with Emery", m_StyleHdrL);
         addCell(row, B, m_SkuTot, m_StyleInt);
         
         row = m_Sheet1.createRow(6);
         addCell(row, A, "Total Sales $$", m_StyleHdrL);
         addCell(row, B, m_DollarTot, m_StyleCurB);
                
         //
         //Assortment summary lines header
         row = m_Sheet1.createRow(8);
         addCell(row, A, "Assortment Integrity", m_StyleHdrC);            
         addCell(row, B, "# of SKUs", m_StyleHdrC);
         addCell(row, C, "% of SKUs", m_StyleHdrC);
         addCell(row, D, "% of Dollar Purchases", m_StyleHdrC);
         addCell(row, E, "Current Customer COGS", m_StyleHdrC);
         addCell(row, F, "New Purchased $$", m_StyleHdrC);
         addCell(row, G, "$$ Var", m_StyleHdrC);
         addCell(row, H, "$$ Savings", m_StyleHdrC);
         
         row = m_Sheet1.createRow(9);
         //
         addCell(row, A, "Exact Match Items", m_StyleTxtRI);
         addCell(row, B, m_MatchCount, m_StyleInt);
         addCellFormula(row, C, "B10/B6", m_StylePct);
         addCellFormula(row, D, "E10/B7", m_StylePct);
         addCell(row, E, m_CurMatchCogs, m_StyleCur);
         addCell(row, F, m_NewMatchCogs, m_StyleCur);
         addCellFormula(row, G, "E10-F10", m_StyleCur);
         addCellFormula(row, H, "(E10-F10)/E10", m_StylePct);
         
                  
         row = m_Sheet1.createRow(10);
         addCell(row, A, "\"Like\" Match Items", m_StyleTxtRI);
         addCell(row, B, m_LikeCount, m_StyleInt);
         addCellFormula(row, C, "B11/B6", m_StylePct);
         addCellFormula(row, D, "E11/B7", m_StylePct);
         addCell(row, E, m_CurLikeCogs, m_StyleCur);
         addCell(row, F, m_NewLikeCogs, m_StyleCur);
         addCellFormula(row, G, "E11-F11", m_StyleCur);
         addCellFormula(row, H, "(E11-F11)/E11", m_StylePct);
                           
         row = m_Sheet1.createRow(11);
         addCellFormula(row, E, "SUM(E10:E11)", m_StyleCur);         
         addCellFormula(row, F, "SUM(F10:F11)", m_StyleCurB);         
         addCellFormula(row, G, "SUM(G10:G11)", m_StyleCurB);
         addCellFormula(row, H, "(E12-F12)/E12", m_StylePct);
                     
         row = m_Sheet1.createRow(12);
         addCell(row, A, "", m_StyleFillGrey);
         addCell(row, B, "", m_StyleFillGrey);
         addCell(row, C, "", m_StyleFillGrey);
         addCell(row, D, "", m_StyleFillGrey);
         addCell(row, E, "", m_StyleFillGrey);
         addCell(row, F, "", m_StyleFillGrey);
         addCell(row, G, "", m_StyleFillGrey);
         addCell(row, H, "", m_StyleFillGrey);
         
         //
         // Discontinued data rows
         row = m_Sheet1.createRow(13);
         addCell(row, A, "Discontinued - No Match*", m_StyleTxtRI);
         addCell(row, B, m_DiscCount, m_StyleInt);
         addCellFormula(row, C, "B14/B6", m_StylePct);
         addCellFormula(row, D, "E14/B7", m_StylePct);
         addCell(row, E, m_DiscCogs, m_StyleCur);
         
         row = m_Sheet1.createRow(14);
         addCell(row, A, "Discontinued - Previously", m_StyleTxtRI);
         addCell(row, B, m_PrevDiscCount, m_StyleInt);
         addCellFormula(row, C, "B15/B6", m_StylePct);
         addCellFormula(row, D, "E15/B7", m_StylePct);
         addCell(row, E, m_PrevDiscCogs, m_StyleCur);
         
         row = m_Sheet1.createRow(15);
         addCellFormula(row, B, "SUM(B10:B11,B14:B15)", m_StyleInt);
         addCellFormula(row, C, "SUM(C10:C15)", m_StylePctB);
         addCellFormula(row, D, "SUM(D10:D15)", m_StylePctB);
         addCellFormula(row, E, "SUM(E14:E15)", m_StyleCurB);
                 
         row = m_Sheet1.createRow(16);
         addCellFormula(row, B, "B6-B16", m_StyleInt);
         addCellFormula(row, E, "SUM(E12+E16)", m_StyleCurB);
                     
         row = m_Sheet1.createRow(17);
         addCell(row, A, "Unit of Measure Considerations", m_StyleHdrC);
         addCell(row, B, "# of SKUs", m_StyleHdrC);
         addCell(row, C, "% Total Compared", m_StyleHdrC);
         
         row = m_Sheet1.createRow(18);
         addCell(row, A, "Total # of SKU's Compared", m_StyleTxtRI);
         addCellFormula(row, B, "SUM(B10:B11)", m_StyleInt);
         addCell(row, D, "missing Wilton OM on ", m_StyleTxtL);
         addCellFormula(row, E, "SUM(B10:B11)-B19", m_StyleInt);
         
         row = m_Sheet1.createRow(19);
         addCell(row, A, "Exact Match SKUs", m_StyleTxtRI);
         addCell(row, B, m_OMMatchCount, m_StyleInt);
         addCellFormula(row, C, "B20/B19", m_StylePct);
         
         row = m_Sheet1.createRow(20);
         addCell(row, A, "\"Like\" Match SKUs", m_StyleTxtRI);
         addCell(row, B, m_OMLikeCount, m_StyleInt);
         addCellFormula(row, C, "B21/B19", m_StylePct);
         
         row = m_Sheet1.createRow(22);
         addCell(row, A, "Cost Of Goods", m_StyleHdrC);
         addCell(row, B, "% Lower Than Current Pricing", m_StyleHdrC);
         
         row = m_Sheet1.createRow(23);
         addCell(row, A, "Exact Match Items", m_StyleTxtRI);
         addCellFormula(row, B, "H10", m_StylePct);
         
         row = m_Sheet1.createRow(24);
         addCell(row, A, "\"Like\" Match Items", m_StyleTxtRI);
         addCellFormula(row, B, "H11", m_StylePct);
         
         row = m_Sheet1.createRow(25);
         addCell(row, A, "Overall COGS Savings", m_StyleTxtRI);
         addCellFormula(row, B, "H12", m_StylePct);
         
         row = m_Sheet1.createRow(27);
         addCell(row, A, "Retail Pricing", m_StyleHdrC);
         addCell(row, B, "Sales @ Retail", m_StyleHdrC);
         addCell(row, C, "Sales @ Retail % Change", m_StyleHdrC);
         addCell(row, D, "Customer COGS", m_StyleHdrC);
         addCell(row, E, "GM$s", m_StyleHdrC);
         addCell(row, F, "GM%", m_StyleHdrC);
         addCell(row, G, "% Var", m_StyleHdrC);
         addCell(row, H, "GM$ Var", m_StyleHdrC);
         
         row = m_Sheet1.createRow(28);
         addCell(row, A, "Current Estimated GM", m_StyleTxtRI);
         addCell(row, B, m_LikeRetTot + m_MatchRetTot, m_StyleInt);
         addCellFormula(row, D, "E12", m_StyleCur);
         addCellFormula(row, E, "B29-D29", m_StyleCur);
         addCellFormula(row, F, "E29/B29", m_StylePct);
                     
         row = m_Sheet1.createRow(29);
         addCell(row, A, "", m_StyleFillBlack);
         addCell(row, B, "", m_StyleFillBlack);
         addCell(row, C, "", m_StyleFillBlack);
         addCell(row, D, "", m_StyleFillBlack);
         addCell(row, E, "", m_StyleFillBlack);
         addCell(row, F, "", m_StyleFillBlack);
         addCell(row, G, "", m_StyleFillBlack);
         addCell(row, H, "", m_StyleFillBlack);
         
         row = m_Sheet1.createRow(30);
         addCell(row, A, "New Estimated GM @\r\nCompetitive Retail", m_StyleTxtRI);
         addCell(row, B, m_CompRetTot, m_StyleInt);
         addCellFormula(row, C, "(B31-B29)/B29", m_StylePct);
         addCellFormula(row, D, "F12", m_StyleCur);
         addCellFormula(row, E, "B31-D31", m_StyleCur);
         addCellFormula(row, F, "E31/B31", m_StylePct);
         addCellFormula(row, G, "F31-F29", m_StylePct);
         addCellFormula(row, H, "E31-E29", m_StyleCur);
         
         row = m_Sheet1.createRow(31);
         addCell(row, A, "New Estimated GM @\r\nMarket Retail", m_StyleTxtRI);
         addCell(row, B, m_MktRetTot, m_StyleInt);
         addCellFormula(row, C, "(B32-B29)/B29", m_StylePct);
         addCellFormula(row, D, "F12", m_StyleCur);
         addCellFormula(row, E, "B32-D32", m_StyleCur);
         addCellFormula(row, F, "E32/B32", m_StylePct);
         addCellFormula(row, G, "F32-F29", m_StylePct);
         addCellFormula(row, H, "E32-E29", m_StyleCur);
         
         row = m_Sheet1.createRow(32);
         addCell(row, A, "New Estimated GM @\r\nPremium Retail", m_StyleTxtRI);
         addCell(row, B, m_PremRetTot, m_StyleInt);
         addCellFormula(row, C, "(B33-B29)/B29", m_StylePct);
         addCellFormula(row, D, "F12", m_StyleCur);
         addCellFormula(row, E, "B33-D33", m_StyleCur);
         addCellFormula(row, F, "E33/B33", m_StylePct);
         addCellFormula(row, G, "F33-F29", m_StylePct);
         addCellFormula(row, H, "E33-E29", m_StyleCur);
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
         row = null;
      }
   }
   
   /**
    * Builds the taxonomy tab/sheet
    * @throws SQLException 
    */
   private void buildSummaryTax() throws SQLException
   {
      int rowNum = addSummaryTaxHeader();
      Row row = null;
      ResultSet rs = null;
      
      m_SumTax.setString(1, m_CustId);
      m_SumTax.setString(2, m_CustId);
      m_SumTax.setString(3, m_CustId);
      m_SumTax.setString(4, m_CustId);
      
      try {
         m_CurAction = "retrieving summary tax sheet data";
         rs = m_SumTax.executeQuery();
         
         m_CurAction = "building summary tax sheet";
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            row = m_Sheet6.createRow(++rowNum);
            
            addCell(row, A, rs.getString("taxonomy_1"), m_StyleTxtL);
            addCell(row, B, rs.getInt("r12_units"), m_StyleInt);
            addCell(row, C, rs.getDouble("cur_cost_tot"), m_StyleDouble);
            addCell(row, D, rs.getDouble("new_cost_tot"), m_StyleDouble);
            addCell(row, E, rs.getDouble("cur_ret_tot"), m_StyleDouble);
            addCell(row, F, rs.getDouble("cur_margin_pct"), m_StyleDouble);
            addCell(row, G, rs.getDouble("comp_ret_total"), m_StyleDouble);
            addCell(row, H, rs.getDouble("comp_margin_pct"), m_StyleDouble);
            addCell(row, I, rs.getDouble("market_ret_tot"), m_StyleDouble);
            addCell(row, J, rs.getDouble("market_margin_pct"), m_StyleDouble);
            addCell(row, K, rs.getDouble("premium_ret_tot"), m_StyleDouble);
            addCell(row, L, rs.getDouble("premium_margin_pct"), m_StyleDouble);            
            addCell(row, M, rs.getString("match"), m_StyleTxtL);
         }
      }
      
      finally {
         DbUtils.closeDbConn(null, null, rs);
         rs = null;
         row = null;
      }
   }
   
   /**
    * Closes prepared statements and cleans up member variables
    */
   protected void cleanup()
   {
      m_Sheet1 = null;
      m_Sheet2 = null;
      m_Sheet3 = null;
      m_Sheet4 = null;
      m_Sheet5 = null;
      m_Sheet6 = null;      
      m_Workbook = null;
      
      m_StyleHdrL = null;
      m_StyleHdrC = null;
      m_StyleTxtC = null;
      m_StyleTxtL = null;
      m_StyleTxtR = null;
      m_StyleTxtRI = null;
      m_StyleInt = null;
      m_StyleDouble = null;
      m_StyleFillGrey = null;
      m_StyleFillBlack = null;
      m_StyleLtGreenPct = null;
      m_StyleLtGreenCur = null;
      m_StyleLtRedPct = null;
      m_StyleLtRedCur = null;
      m_StylePct = null;
      m_StyleCur = null;
      m_StylePctB = null;
      m_StyleCurB = null;
      
      DbUtils.closeDbConn(null, m_MatchedItems, null);
      DbUtils.closeDbConn(null, m_DiscoItems, null);
      DbUtils.closeDbConn(null, m_LikeItems, null);
      DbUtils.closeDbConn(null, m_OMItems, null);
      DbUtils.closeDbConn(null, m_SumTax, null);
      DbUtils.closeDbConn(null, m_CustData, null);
      
      m_MatchedItems = null;
      m_DiscoItems = null;
      m_LikeItems = null;
      m_OMItems = null;
      m_SumTax = null;
      m_CustData = null;
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
         prepareStatements();
         created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[CustomerAssortment]", ex);
      }

      finally {
        cleanup();

        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * defines Excel fonts and styles
    */
   private void defineStyles()
   {      
      XSSFDataFormat format = m_Workbook.createDataFormat();
      XSSFColor darkBlue = new XSSFColor(new java.awt.Color(0, 32, 96));
      XSSFColor medGrey = new XSSFColor(new java.awt.Color(64, 64, 64));
             
      try {
         //
         // Normal Font
         Font fontNorm = m_Workbook.createFont();
         fontNorm.setFontName("Calibri");
         fontNorm.setFontHeightInPoints(fontHeight);
         
         //
         // defines bold font
         Font fontBold = m_Workbook.createFont();
         fontBold.setFontName("Calibri");
         fontBold.setFontHeightInPoints(fontHeight);
         fontBold.setBold(true);
         
         //
         // defines bold header font
         Font fontBoldS = m_Workbook.createFont();
         fontBoldS.setFontName("Calibri");
         fontBoldS.setFontHeightInPoints(fontHeight);
         fontBoldS.setBold(true);
         fontBoldS.setColor(IndexedColors.WHITE.getIndex());
         
         //
         // defines italic font
         Font fontItalic = m_Workbook.createFont();
         fontItalic.setFontName("Calibri");
         fontItalic.setItalic(true);
         fontItalic.setFontHeightInPoints(fontHeight);
         
         //
         // defines green font
         Font fontGreen = m_Workbook.createFont();
         fontGreen.setFontName("Calibri");         
         fontGreen.setFontHeightInPoints(fontHeight);
         fontGreen.setColor(IndexedColors.GREEN.getIndex());
         
         //
         // defines green font
         Font fontRed = m_Workbook.createFont();
         fontRed.setFontName("Calibri");         
         fontRed.setFontHeightInPoints(fontHeight);
         fontRed.setColor(IndexedColors.RED.getIndex());
         
         //
         // defines style column header, left-justified
         m_StyleHdrL = m_Workbook.createCellStyle();
         m_StyleHdrL.setFont(fontBoldS);
         m_StyleHdrL.setAlignment(HorizontalAlignment.LEFT);
         m_StyleHdrL.setVerticalAlignment(VerticalAlignment.TOP);
         m_StyleHdrL.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         m_StyleHdrL.setFillForegroundColor(darkBlue);
         m_StyleHdrL.setBorderTop(BorderStyle.THIN);
         m_StyleHdrL.setBorderBottom(BorderStyle.THIN);
         m_StyleHdrL.setBorderLeft(BorderStyle.THIN);
         m_StyleHdrL.setBorderRight(BorderStyle.THIN);
                           
         m_StyleHdrC = m_Workbook.createCellStyle();
         m_StyleHdrC.setFont(fontBoldS);
         m_StyleHdrC.setAlignment(HorizontalAlignment.LEFT);
         m_StyleHdrC.setVerticalAlignment(VerticalAlignment.TOP);
         m_StyleHdrC.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         m_StyleHdrC.setFillForegroundColor(darkBlue);
         m_StyleHdrC.setBorderTop(BorderStyle.THIN);
         m_StyleHdrC.setBorderBottom(BorderStyle.THIN);
         m_StyleHdrC.setBorderLeft(BorderStyle.THIN);
         m_StyleHdrC.setBorderRight(BorderStyle.THIN);
         
         m_StyleTxtL = m_Workbook.createCellStyle();
         m_StyleTxtL.setAlignment(HorizontalAlignment.LEFT);
         m_StyleTxtL.setFont(fontNorm);
         m_StyleTxtL.setBorderTop(BorderStyle.THIN);
         m_StyleTxtL.setBorderBottom(BorderStyle.THIN);
         m_StyleTxtL.setBorderLeft(BorderStyle.THIN);
         m_StyleTxtL.setBorderRight(BorderStyle.THIN);
         
         m_StyleTxtC = m_Workbook.createCellStyle();
         m_StyleTxtC.setAlignment(HorizontalAlignment.CENTER);
         m_StyleTxtC.setFont(fontNorm);
         m_StyleTxtC.setWrapText(true);
         m_StyleTxtC.setBorderTop(BorderStyle.THIN);
         m_StyleTxtC.setBorderBottom(BorderStyle.THIN);
         m_StyleTxtC.setBorderLeft(BorderStyle.THIN);
         m_StyleTxtC.setBorderRight(BorderStyle.THIN);
         
         m_StyleTxtR = m_Workbook.createCellStyle();
         m_StyleTxtR.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleTxtR.setFont(fontNorm);
         m_StyleTxtR.setBorderTop(BorderStyle.THIN);
         m_StyleTxtR.setBorderBottom(BorderStyle.THIN);
         m_StyleTxtR.setBorderLeft(BorderStyle.THIN);
         m_StyleTxtR.setBorderRight(BorderStyle.THIN);
         
         
         m_StyleTxtRI = m_Workbook.createCellStyle();
         m_StyleTxtRI.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleTxtRI.setFont(fontItalic);
         m_StyleTxtRI.setWrapText(true);
         m_StyleTxtRI.setBorderTop(BorderStyle.THIN);
         m_StyleTxtRI.setBorderBottom(BorderStyle.THIN);
         m_StyleTxtRI.setBorderLeft(BorderStyle.THIN);
         m_StyleTxtRI.setBorderRight(BorderStyle.THIN);
         
         m_StyleInt = m_Workbook.createCellStyle();
         m_StyleInt.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleInt.setDataFormat(format.getFormat("0"));
         m_StyleInt.setFont(fontNorm);
         m_StyleInt.setBorderTop(BorderStyle.THIN);
         m_StyleInt.setBorderBottom(BorderStyle.THIN);
         m_StyleInt.setBorderLeft(BorderStyle.THIN);
         m_StyleInt.setBorderRight(BorderStyle.THIN);
         
         m_StyleDouble = m_Workbook.createCellStyle();
         m_StyleDouble.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleDouble.setDataFormat(format.getFormat("#,##0.00"));
         m_StyleDouble.setFont(fontNorm);
         m_StyleDouble.setBorderTop(BorderStyle.THIN);
         m_StyleDouble.setBorderBottom(BorderStyle.THIN);
         m_StyleDouble.setBorderLeft(BorderStyle.THIN);
         m_StyleDouble.setBorderRight(BorderStyle.THIN);
         
         m_StyleDate = m_Workbook.createCellStyle();
         m_StyleDate.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleDate.setDataFormat(format.getFormat("mm/dd/yyyy"));
         m_StyleDate.setFont(fontNorm);
         m_StyleDate.setBorderTop(BorderStyle.THIN);
         m_StyleDate.setBorderBottom(BorderStyle.THIN);
         m_StyleDate.setBorderLeft(BorderStyle.THIN);
         m_StyleDate.setBorderRight(BorderStyle.THIN);
         
         m_StyleFillGrey = m_Workbook.createCellStyle();
         m_StyleFillGrey.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         m_StyleFillGrey.setFillForegroundColor(medGrey);
         
         m_StyleFillBlack = m_Workbook.createCellStyle();
         m_StyleFillBlack.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         m_StyleFillBlack.setFillForegroundColor(IndexedColors.BLACK.getIndex());
         
         m_StyleLtGreenPct = m_Workbook.createCellStyle();
         m_StyleLtGreenPct.setFont(fontGreen);
         m_StyleLtGreenPct.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleLtGreenPct.setDataFormat(format.getFormat("0.00%"));
         m_StyleLtGreenPct.setBorderTop(BorderStyle.THIN);
         m_StyleLtGreenPct.setBorderBottom(BorderStyle.THIN);
         m_StyleLtGreenPct.setBorderLeft(BorderStyle.THIN);
         m_StyleLtGreenPct.setBorderRight(BorderStyle.THIN);
         
         m_StyleLtGreenCur = m_Workbook.createCellStyle();
         m_StyleLtGreenCur.setFont(fontGreen);
         m_StyleLtGreenCur.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleLtGreenCur.setDataFormat(format.getFormat("$#,##0_);[Red]($#,##0)"));
         m_StyleLtGreenCur.setBorderTop(BorderStyle.THIN);
         m_StyleLtGreenCur.setBorderBottom(BorderStyle.THIN);
         m_StyleLtGreenCur.setBorderLeft(BorderStyle.THIN);
         m_StyleLtGreenCur.setBorderRight(BorderStyle.THIN);
                  
         m_StyleLtRedPct = m_Workbook.createCellStyle();
         m_StyleLtRedPct.setFont(fontGreen);
         m_StyleLtRedPct.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleLtRedPct.setDataFormat(format.getFormat("0.00%"));
         m_StyleLtRedPct.setBorderTop(BorderStyle.THIN);
         m_StyleLtRedPct.setBorderBottom(BorderStyle.THIN);
         m_StyleLtRedPct.setBorderLeft(BorderStyle.THIN);
         m_StyleLtRedPct.setBorderRight(BorderStyle.THIN);
         
         m_StyleLtRedCur = m_Workbook.createCellStyle();
         m_StyleLtRedCur.setFont(fontGreen);
         m_StyleLtRedCur.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleLtRedCur.setDataFormat(format.getFormat("$#,##0_);[Red]($#,##0)"));
         m_StyleLtRedCur.setBorderTop(BorderStyle.THIN);
         m_StyleLtRedCur.setBorderBottom(BorderStyle.THIN);
         m_StyleLtRedCur.setBorderLeft(BorderStyle.THIN);
         m_StyleLtRedCur.setBorderRight(BorderStyle.THIN);
         
         m_StylePct = m_Workbook.createCellStyle();
         m_StylePct.setFont(fontNorm);
         m_StylePct.setAlignment(HorizontalAlignment.RIGHT);
         m_StylePct.setDataFormat(format.getFormat("0.00%"));
         m_StylePct.setBorderTop(BorderStyle.THIN);
         m_StylePct.setBorderBottom(BorderStyle.THIN);
         m_StylePct.setBorderLeft(BorderStyle.THIN);
         m_StylePct.setBorderRight(BorderStyle.THIN);
         
         m_StyleCur = m_Workbook.createCellStyle();
         m_StyleCur.setFont(fontNorm);
         m_StyleCur.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleCur.setDataFormat(format.getFormat("$#,##0_);[Red]($#,##0)"));
         m_StyleCur.setBorderTop(BorderStyle.THIN);
         m_StyleCur.setBorderBottom(BorderStyle.THIN);
         m_StyleCur.setBorderLeft(BorderStyle.THIN);
         m_StyleCur.setBorderRight(BorderStyle.THIN);
         
         m_StylePctB = m_Workbook.createCellStyle();
         m_StylePctB.setFont(fontBold);
         m_StylePctB.setAlignment(HorizontalAlignment.RIGHT);
         m_StylePctB.setDataFormat(format.getFormat("0.00%"));
         m_StylePctB.setBorderTop(BorderStyle.THIN);
         m_StylePctB.setBorderBottom(BorderStyle.THIN);
         m_StylePctB.setBorderLeft(BorderStyle.THIN);
         m_StylePctB.setBorderRight(BorderStyle.THIN);
         
         m_StyleCurB = m_Workbook.createCellStyle();
         m_StyleCurB.setFont(fontBold);
         m_StyleCurB.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleCurB.setDataFormat(format.getFormat("$#,##0_);[Red]($#,##0)"));
         m_StyleCurB.setBorderTop(BorderStyle.THIN);
         m_StyleCurB.setBorderBottom(BorderStyle.THIN);
         m_StyleCurB.setBorderLeft(BorderStyle.THIN);
         m_StyleCurB.setBorderRight(BorderStyle.THIN);
      }
      
      finally {      
         format = null;
         darkBlue = null;
         medGrey = null;         
      }
   }
   
   /**
    * 
    * @return
    * @throws SQLException
    */
   private boolean prepareStatements() throws SQLException
   {
      boolean isPrepared = false;
      StringBuffer sql = new StringBuffer();
      
      if ( m_EdbConn != null ) {
         try {
            sql.append("select customer_id, name, class as cust_type ");
            sql.append("from customer ");
            sql.append("join cust_market_view using(customer_id) ");
            sql.append("where customer_id = ? and market = 'CUSTOMER TYPE' ");
            m_CustData = m_EdbConn.prepareStatement(sql.toString());
                                    
            sql.setLength(0);
            sql.append("select ");
            sql.append("    c.taxonomy_1, ");
            sql.append("    c.taxonomy_2, ");
            sql.append("    c.item_id, ");
            sql.append("    c.customer_sku, ");
            sql.append("    c.emery_upc, ");
            sql.append("    c.mfrnum, ");
            sql.append("    c.item_description, ");
            sql.append("    c.cur_vendor, ");
            sql.append("    c.ejd_item, ");
            sql.append("    c.ejd_upc, ");
            sql.append("    c.ejd_mfrnum, ");
            sql.append("    c.ejd_description, ");
            sql.append("    c.ejd_vendor_name, ");
            sql.append("    c.r12_units, ");
            sql.append("    c.cur_cost, ");
            sql.append("    c.new_cost, ");
            sql.append("    c.r12_units * c.cur_cost as cur_cost_total, ");
            sql.append("    c.r12_units * c.new_cost as new_cost_total, ");
            sql.append("    (c.r12_units * c.new_cost ) - (c.r12_units * c.cur_cost) as cost_variance, ");
            sql.append("    c.cur_retail, ");
            sql.append("    c.cur_ret_tot, ");
            sql.append("    CASE WHEN c.cur_retail IS NULL OR c.cur_cost IS NULL OR c.retail_pack IS NULL ");
            sql.append("    THEN 0 ");
            sql.append("    WHEN c.cur_retail = 0 OR c.retail_pack = 0 ");
            sql.append("    THEN 0 ");
            sql.append("    ELSE round(((c.cur_retail - (c.cur_cost / c.retail_pack)) / c.cur_retail) * 100, 4) ");
            sql.append("    END AS cur_margin, ");
            sql.append("    c.comp_ret, ");
            sql.append("    c.comp_ret_tot, ");
            sql.append("    CASE WHEN c.comp_ret IS NULL OR c.new_cost IS NULL OR c.ejd_ret_pack IS NULL ");
            sql.append("      THEN 0 ");
            sql.append("    WHEN c.comp_ret = 0 OR c.ejd_ret_pack = 0 ");
            sql.append("      THEN 0 ");
            sql.append("    ELSE round(((c.comp_ret - (c.new_cost / c.ejd_ret_pack)) / c.comp_ret) * 100, 4) ");
            sql.append("    END                 AS comp_margin_pct, ");
            sql.append("    c.market_ret, ");
            sql.append("    c.market_ret_tot, ");
            sql.append("    CASE WHEN c.market_ret IS NULL OR c.new_cost IS NULL OR c.ejd_ret_pack IS NULL ");
            sql.append("      THEN 0 ");
            sql.append("    WHEN c.market_ret = 0 OR c.ejd_ret_pack = 0 ");
            sql.append("      THEN 0 ");
            sql.append("    ELSE round(((c.market_ret - (c.new_cost / c.ejd_ret_pack)) / c.market_ret) * 100, 4) ");
            sql.append("    END AS market_margin_pct, ");
            sql.append("    c.premium_ret, ");
            sql.append("    c.premium_ret_tot, ");
            sql.append("    CASE WHEN c.premium_ret IS NULL OR c.new_cost IS NULL OR c.ejd_ret_pack IS NULL ");
            sql.append("      THEN 0 ");
            sql.append("    WHEN c.premium_ret = 0 OR c.ejd_ret_pack = 0 ");
            sql.append("      THEN 0 ");
            sql.append("    ELSE round(((c.premium_ret - (c.new_cost / c.ejd_ret_pack)) / c.premium_ret) * 100, 4) ");
            sql.append("    END                 AS premium_margin_pct, ");
            sql.append("    c.new_om, ");
            sql.append("    c.new_uom, ");
            sql.append("    c.nbc, ");
            sql.append("    c.match, ");
            sql.append("    c.stock_pack, ");
            sql.append("    c.cur_om, ");
            sql.append("    c.cur_uom ");
            sql.append("from ( ");
            sql.append("    SELECT ");
            sql.append("    b.dept AS taxonomy_1, ");
            sql.append("    b.category AS taxonomy_2, ");
            sql.append("    b.item_id, ");
            sql.append("    b.customer_sku, ");
            sql.append("    b.emery_upc, ");
            sql.append("    b.mfrnum, ");
            sql.append("    b.Emery_Description AS item_description, ");
            sql.append("    b.vendor_name AS cur_vendor, ");
            sql.append("    b.ace_sku AS ejd_item, ");
            sql.append("    b.ejd_upc, ");
            sql.append("    b.ejd_mfrnum, ");
            sql.append("    b.ejd_description, ");
            sql.append("    b.ejd_vendor_name, ");
            sql.append("    case when b.sell_price > (b.ejd_sell_price * 2) ");
            sql.append("         then b.r12_units * b.retail_pack ");
            sql.append("         else b.r12_units ");
            sql.append("    end as r12_units, ");
            sql.append(" ");
            sql.append("    case when b.sell_price > (b.ejd_sell_price * 2)  then b.sell_price / b.retail_pack else b.sell_price end as cur_cost, ");
            sql.append(" ");
            sql.append("    case when b.ejd_sell_price > (b.sell_price * 2)  then b.ejd_sell_price / b.ejd_ret_pack  else b.ejd_sell_price end as new_cost, ");
            sql.append("    b.ret_price AS cur_retail, ");
            sql.append("    b.r12_units * b.ret_price AS cur_ret_tot, ");
            sql.append("    CASE WHEN b.ret_price IS NULL OR b.sell_price IS NULL OR b.retail_pack IS NULL ");
            sql.append("    THEN 0 ");
            sql.append("    WHEN b.ret_price = 0 OR b.retail_pack = 0 ");
            sql.append("    THEN 0 ");
            sql.append("    ELSE round(((b.ret_price - (b.sell_price / b.retail_pack)) / b.ret_price) * 100, 4) ");
            sql.append("    END AS cur_margin, ");
            sql.append("    b.ejd_reta AS comp_ret, ");
            sql.append("    b.ejd_reta  * b.r12_units AS comp_ret_tot, ");
            sql.append("    CASE WHEN b.ejd_reta IS NULL OR b.ejd_sell_price IS NULL OR b.ejd_ret_pack IS NULL ");
            sql.append("    THEN 0 ");
            sql.append("    WHEN b.ejd_reta = 0 OR b.ejd_ret_pack = 0 ");
            sql.append("    THEN 0 ");
            sql.append("    ELSE round(((b.ejd_reta - (b.ejd_sell_price / b.ejd_ret_pack)) / b.ejd_reta) * 100, 4) ");
            sql.append("    END AS comp_margin_pct, ");
            sql.append("    b.ejd_retb AS market_ret, ");
            sql.append("    b.ejd_retb  *  b.r12_units AS market_ret_tot, ");
            sql.append("    CASE WHEN b.ejd_retb IS NULL OR b.ejd_sell_price IS NULL OR b.ejd_ret_pack IS NULL ");
            sql.append("    THEN 0 ");
            sql.append("    WHEN b.ejd_retb = 0 OR b.ejd_ret_pack = 0 ");
            sql.append("    THEN 0 ");
            sql.append("    ELSE round(((b.ejd_retb - (b.ejd_sell_price / b.ejd_ret_pack)) / b.ejd_retb) * 100, 4) ");
            sql.append("    END AS market_margin_pct, ");
            sql.append("    b.ejd_retc AS premium_ret, ");
            sql.append("    b.ejd_retc  *  b.r12_units AS premium_ret_tot, ");
            sql.append("    CASE WHEN b.ejd_retc IS NULL OR b.ejd_sell_price IS NULL OR b.ejd_ret_pack IS NULL ");
            sql.append("    THEN 0 ");
            sql.append("    WHEN b.ejd_retc = 0 OR b.ejd_ret_pack = 0 ");
            sql.append("    THEN 0 ");
            sql.append("    ELSE round(((b.ejd_retc - (b.ejd_sell_price / b.ejd_ret_pack)) / b.ejd_retc) * 100, 4) ");
            sql.append("    END AS premium_margin_pct, ");
            sql.append("    b.EJD_OM AS new_om, ");
            sql.append("    b.EJD_UOM AS new_uom, ");
            sql.append("    b.nbc, ");
            sql.append("    b.matchtype AS MATCH, ");
            sql.append("    b.stock_pack, ");
            sql.append("    b.om AS cur_om, ");
            sql.append("    b.Emery_UOM AS cur_uom, ");
            sql.append("    b.retail_pack, ");
            sql.append("    b.ejd_ret_pack ");
            sql.append("    FROM ");
            sql.append("    ( ");
            sql.append("    SELECT ");
            sql.append("    CASE WHEN cust.parent_id IS NULL ");
            sql.append("    THEN a.cust_nbr ");
            sql.append("    ELSE cust.parent_id ");
            sql.append("    END AS cust_nbr, ");
            sql.append("    CASE WHEN prnt.name IS NULL ");
            sql.append("    THEN cust.name ");
            sql.append("    ELSE prnt.name ");
            sql.append("    END AS cust_name, ");
            sql.append("    a.item_ea_id, ");
            sql.append("    iea.item_id, ");
            sql.append("    iex.customer_sku, ");
            sql.append("    aix1.ace_sku, ");
            sql.append("    iea.description AS emery_description, ");
            sql.append("    iea2.description AS ejd_description, ");
            sql.append("    ven.name AS vendor_name, ");
            sql.append("    vix.vendor_item_num AS mfrnum, ");
            sql.append("    vixa.vendor_item_num AS ejd_mfrnum, ");
            sql.append("    tax1.taxonomy AS dept, ");
            sql.append("    tax2.taxonomy AS CATEGORY, ");
            sql.append("    tax3.taxonomy AS subcategory, ");
            sql.append("    ( SELECT upc_code ");
            sql.append("    FROM ejd_item_whs_upc ");
            sql.append("    WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11 ");
            sql.append("    ORDER BY primary_upc DESC ");
            sql.append("    LIMIT 1) AS ejd_upc, ");
            sql.append("    ( SELECT upc_code ");
            sql.append("    FROM ejd_item_whs_upc ");
            sql.append("    WHERE ejd_item_id = iea.ejd_item_id AND warehouse_id IN (1, 2) ");
            sql.append("    ORDER BY primary_upc DESC ");
            sql.append("    LIMIT 1) AS emery_upc, ");
            sql.append("    eiw.stock_pack, ");
            sql.append("    CASE WHEN ejd.broken_case_id <> 1 ");
            sql.append("    THEN eiw.stock_pack ");
            sql.append("    ELSE 1 ");
            sql.append("    END AS om, ");
            sql.append("    iea.retail_pack, ");
            sql.append("    (ejd_cust_procs.get_sell_price(?, a.item_ea_id)).price AS sell_price, ");
            sql.append(" ");
            sql.append("    CASE WHEN iea2.item_ea_id IS NULL ");
            sql.append("    THEN 0 ");
            sql.append("    ELSE (ejd_cust_procs.get_sell_price(( SELECT CASE WHEN proxy_cust_id IS NULL THEN '008229' ELSE proxy_cust_id END AS proxy_cust_id ");
            sql.append("    FROM ejd.cust_migration WHERE customer_id = ?), iea2.item_ea_id)).price ");
            sql.append("    END AS ejd_sell_price, ");
            sql.append(" ");
            sql.append("    ejd_price_procs.get_retail_price(?, a.item_ea_id) AS ret_price, ");
            sql.append("    a.r12_units, ");
            sql.append("    CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("    THEN 0 ");
            sql.append("    ELSE ( SELECT retail_a ");
            sql.append("    FROM ejd_item_price ");
            sql.append("    WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("    END AS ejd_reta, ");
            sql.append("    CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("    THEN 0 ");
            sql.append("    ELSE ( SELECT retail_b ");
            sql.append("    FROM ejd_item_price ");
            sql.append("    WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("    END AS ejd_retb, ");
            sql.append("    CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("    THEN 0 ");
            sql.append("    ELSE ( SELECT retail_c ");
            sql.append("    FROM ejd_item_price ");
            sql.append("    WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("    END AS ejd_retc, ");
            sql.append("    iea2.retail_pack AS ejd_ret_pack, ");
            sql.append("    ( SELECT stock_pack ");
            sql.append("    FROM ejd_item_warehouse ");
            sql.append("    WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) AS ejd_om, ");
            sql.append("    CASE WHEN aix1.match = 1 THEN 'Matched Item' ");
            sql.append("    WHEN aix1.match = 2 THEN 'Like Item' ");
            sql.append("    WHEN eiw.disp_id <> 1 THEN 'PREVIOUSLY DISCONTINUED' ");
            sql.append("    ELSE 'DISCONTINUED' END AS MatchType, ");
            sql.append("    ven2.name AS EJD_Vendor_Name, ");
            sql.append("    decode(ejd.broken_case_id, 1, '', 'NBC') AS nbc, ");
            sql.append("    su1.name AS Emery_UOM, ");
            sql.append("    su2.name AS EJD_UOM ");
            sql.append("    FROM ( ");
            sql.append("    SELECT ");
            sql.append("    cust_nbr, ");
            sql.append("    item_ea_id, ");
            sql.append("    sum(r12_units) AS r12_units ");
            sql.append("    FROM ( ");
            sql.append("    SELECT DISTINCT ");
            sql.append("    CASE WHEN cust.parent_id IS NULL ");
            sql.append("    THEN dtl.cust_nbr ");
            sql.append("    ELSE cust.parent_id ");
            sql.append("    END AS cust_nbr, ");
            sql.append("    dtl.item_ea_id, ");
            sql.append("    sum(dtl.qty_shipped) AS r12_units ");
            sql.append("    FROM sa.inv_dtl dtl ");
            sql.append("    INNER JOIN customer cust ON cust.customer_id = dtl.cust_nbr ");
            sql.append("    WHERE dtl.invoice_date > CURRENT_DATE - 365 AND dtl.cust_nbr IN ( ");
            sql.append("    SELECT DISTINCT customer_id ");
            sql.append("    FROM customer ");
            sql.append("    WHERE (customer_id = ? OR parent_id = ?)) ");
            sql.append("    GROUP BY cust.parent_id, dtl.cust_nbr, dtl.item_ea_id ");
            sql.append("    ) b ");
            sql.append("    GROUP BY cust_nbr, item_ea_id) a ");
            sql.append("    LEFT OUTER JOIN ejd.item_entity_attr iea ON iea.item_ea_id = a.item_ea_id ");
            sql.append("    LEFT OUTER JOIN ejd.ship_unit su1 ON su1.unit_id = iea.ship_unit_id ");
            sql.append("    INNER JOIN ejd.customer cust ON cust.customer_id = a.cust_nbr ");
            sql.append("    LEFT OUTER JOIN ejd.customer prnt ON prnt.customer_id = cust.parent_id ");
            sql.append("    INNER JOIN ejd.cust_warehouse cw ON cust.customer_id = cw.customer_id ");
            sql.append("    INNER JOIN ejd.ejd_item ejd ON ejd.ejd_item_id = iea.ejd_item_id ");
            sql.append("    LEFT OUTER JOIN ejd.ejd_item_warehouse eiw ");
            sql.append("    ON eiw.ejd_item_id = iea.ejd_item_id AND eiw.warehouse_id = cw.warehouse_id ");
            sql.append("    LEFT OUTER JOIN ejd.item_ea_taxonomy tax3 ON tax3.taxonomy_id = iea.taxonomy_id ");
            sql.append("    LEFT OUTER JOIN ejd.item_ea_taxonomy tax2 ON tax2.taxonomy_id = tax3.parent_id ");
            sql.append("    LEFT OUTER JOIN ejd.item_ea_taxonomy tax1 ON tax1.taxonomy_id = tax2.parent_id ");
            sql.append("    INNER JOIN ejd.vendor ven ON iea.vendor_id = ven.vendor_id ");
            sql.append("    LEFT OUTER JOIN ejd.vendor_item_ea_cross vix ON vix.vendor_id = iea.vendor_id AND vix.item_ea_id = iea.item_ea_id ");
            sql.append("    LEFT OUTER JOIN ( ");
            sql.append("    SELECT ");
            sql.append("    aix1.item_id AS item_id, ");
            sql.append("    aix1.ace_sku AS ace_sku, ");
            sql.append("    aix1.match, ");
            sql.append("    CASE WHEN aix.item_id IS NULL ");
            sql.append("    THEN aix1.item_id ");
            sql.append("    ELSE aix.item_id END AS aix_item_id ");
            sql.append("    FROM scratch_pad.ace_item_xref_1 aix1 ");
            sql.append("    LEFT OUTER JOIN ejd.ace_item_xref aix ON aix.ace_sku = aix1.ace_sku ");
            sql.append("    ) AS aix1 ON aix1.item_id = iea.item_id ");
            sql.append("    LEFT OUTER JOIN ejd.item_ea_cross iex ON iex.customer_id = cust.customer_id AND iex.item_ea_id = iea.item_ea_id ");
            sql.append("    LEFT OUTER JOIN ejd.item_entity_attr iea2 ON iea2.item_id = aix1.aix_item_id AND iea2.item_type_id = 8 ");
            sql.append("    LEFT OUTER JOIN ejd.ship_unit su2 ON su2.unit_id = iea2.ship_unit_id ");
            sql.append("    LEFT OUTER JOIN ejd.vendor_item_ea_cross vixa ");
            sql.append("    ON vixa.vendor_id = iea2.vendor_id AND vixa.item_ea_id = iea2.item_ea_id ");
            sql.append("    LEFT OUTER JOIN ejd.vendor ven2 ON ven2.vendor_id = iea2.vendor_id ");
            sql.append("    ORDER BY tax1.taxonomy, ven.name, tax2.taxonomy, iea.item_id DESC ");
            sql.append("    ) b ");
            sql.append("    WHERE b.matchtype = 'Matched Item' and b.dept <> 'Store Supplies' ");
            sql.append("    ) c ");
            m_MatchedItems = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("SELECT ");
            sql.append("    b.dept              AS taxonomy_1, ");
            sql.append("    b.category          AS taxonomy_2, ");
            sql.append("    b.item_id, ");
            sql.append("    b.customer_sku, ");
            sql.append("    b.emery_upc, ");
            sql.append("    b.mfrnum, ");
            sql.append("    b.Emery_Description AS item_description, ");
            sql.append("    b.vendorname        AS cur_vendor, ");
            sql.append("    b.ace_sku           AS ejd_item, ");
            sql.append("    b.ejd_upc, ");
            sql.append("    b.ejd_mfrnum, ");
            sql.append("    b.ejd_description, ");
            sql.append("    b.ejd_vendor_name, ");
            sql.append("    b.r12_units, ");
            sql.append("    b.sellprice         AS cur_cost, ");
            sql.append("    b.ejdsellprice      AS new_cost, ");
            sql.append("    b.r12_units * b.sellprice as cur_cost_total, ");
            sql.append("    b.r12_units * b.ejdsellprice as new_cost_total, ");
            sql.append("    (b.r12_units * b.ejdsellprice ) - (b.r12_units * b.sellprice) as cost_variance, ");
            sql.append("    b.retailprice       AS cur_retail, ");
            sql.append("    b.r12_units * b.retailprice as cur_ret_tot, ");
            sql.append("    CASE WHEN b.retailprice IS NULL OR b.sellprice IS NULL OR b.retail_pack IS NULL ");
            sql.append("      THEN 0 ");
            sql.append("    WHEN b.retailprice = 0 OR b.retail_pack = 0 ");
            sql.append("      THEN 0 ");
            sql.append("    ELSE round(((b.retailprice - (b.sellprice / b.retail_pack)) / b.retailprice) * 100, 4) ");
            sql.append("    END                 AS cur_margin, ");
            sql.append("    b.EJDRetailA        AS comp_ret, ");
            sql.append("    b.EJDRetailA  * b.r12_units     AS comp_ret_tot, ");
            sql.append("    CASE WHEN b.EJDRetailA IS NULL OR b.ejdsellprice IS NULL OR b.EJDRetailPack IS NULL ");
            sql.append("      THEN 0 ");
            sql.append("    WHEN b.EJDRetailA = 0 OR b.EJDRetailPack = 0 ");
            sql.append("      THEN 0 ");
            sql.append("    ELSE round(((b.EJDRetailA - (b.ejdsellprice / b.EJDRetailPack)) / b.EJDRetailA) * 100, 4) ");
            sql.append("    END                 AS comp_margin_pct, ");
            sql.append("    b.EJDRetailB        AS market_ret, ");
            sql.append("    b.EJDRetailB  *  b.r12_units     AS market_ret_tot, ");
            sql.append("    CASE WHEN b.EJDRetailB IS NULL OR b.ejdsellprice IS NULL OR b.EJDRetailPack IS NULL ");
            sql.append("      THEN 0 ");
            sql.append("    WHEN b.EJDRetailB = 0 OR b.EJDRetailPack = 0 ");
            sql.append("      THEN 0 ");
            sql.append("    ELSE round(((b.EJDRetailB - (b.ejdsellprice / b.EJDRetailPack)) / b.EJDRetailB) * 100, 4) ");
            sql.append("    END                 AS market_margin_pct, ");
            sql.append("    b.EJDRetailC        AS premium_ret, ");
            sql.append("    b.EJDRetailC  *  b.r12_units      AS premium_ret_tot, ");
            sql.append("    CASE WHEN b.EJDRetailC IS NULL OR b.ejdsellprice IS NULL OR b.EJDRetailPack IS NULL ");
            sql.append("      THEN 0 ");
            sql.append("    WHEN b.EJDRetailC = 0 OR b.EJDRetailPack = 0 ");
            sql.append("      THEN 0 ");
            sql.append("    ELSE round(((b.EJDRetailC - (b.ejdsellprice / b.EJDRetailPack)) / b.EJDRetailC) * 100, 4) ");
            sql.append("    END                 AS premium_margin_pct, ");
            sql.append("    b.EJD_OM            AS new_om, ");
            sql.append("    b.EJD_UOM           as new_uom, ");
            sql.append("    b.nbc, ");
            sql.append("    b.MatchType         as match, ");
            sql.append("    b.stock_pack, ");
            sql.append("    b.om                AS cur_om, ");
            sql.append("    b.Emery_UOM         as cur_uom ");
            sql.append("  FROM ");
            sql.append("    ( ");
            sql.append("      SELECT ");
            sql.append("        CASE WHEN cust.parent_id IS NULL ");
            sql.append("          THEN a.cust_nbr ");
            sql.append("        ELSE cust.parent_id ");
            sql.append("        END                                                             AS cust_nbr, ");
            sql.append("        CASE WHEN prnt.name IS NULL ");
            sql.append("          THEN cust.name ");
            sql.append("        ELSE prnt.name ");
            sql.append("        END                                                             AS custName, ");
            sql.append("        a.item_ea_id, ");
            sql.append("        iea.item_id, ");
            sql.append("        iex.customer_sku, ");
            sql.append("        aix1.ace_sku, ");
            sql.append("        iea.description                                                 AS emery_Description, ");
            sql.append("        iea2.description                                                AS ejd_description, ");
            sql.append("        ven.name                                                        AS vendorName, ");
            sql.append("        vix.vendor_item_num                                             AS MfrNum, ");
            sql.append("        vixa.vendor_item_num                                            AS ejd_MfrNum, ");
            sql.append("        tax1.taxonomy                                                   AS dept, ");
            sql.append("        tax2.taxonomy                                                   AS category, ");
            sql.append("        tax3.taxonomy                                                   AS SubCategory, ");
            sql.append("        (SELECT upc_code ");
            sql.append("         FROM ejd_item_whs_upc ");
            sql.append("         WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11 ");
            sql.append("         ORDER BY primary_upc DESC ");
            sql.append("         LIMIT 1)                                                       AS EJD_UPC, ");
            sql.append("        (SELECT upc_code ");
            sql.append("         FROM ejd_item_whs_upc ");
            sql.append("         WHERE ejd_item_id = iea.ejd_item_id AND warehouse_id IN (1, 2) ");
            sql.append("         ORDER BY primary_upc DESC ");
            sql.append("         LIMIT 1)                                                       AS emery_UPC, ");
            sql.append("        eiw.stock_pack, ");
            sql.append("        CASE WHEN ejd.broken_case_id <> 1 ");
            sql.append("          THEN eiw.stock_pack ");
            sql.append("        ELSE 1 ");
            sql.append("        END                                                             AS OM, ");
            sql.append("        iea.retail_pack, ");
            sql.append("        (ejd_cust_procs.get_sell_price(?, a.item_ea_id)).price AS sellPrice, ");
            sql.append("        CASE WHEN iea2.item_ea_id IS NULL ");
            sql.append("          THEN 0 ");
            sql.append("          ELSE (ejd_cust_procs.get_sell_price(( SELECT CASE WHEN proxy_cust_id IS NULL THEN '008229' ELSE proxy_cust_id END AS proxy_cust_id ");
            sql.append("          FROM ejd.cust_migration WHERE customer_id = ?), iea2.item_ea_id)).price ");
            sql.append("        END AS EjdSellPrice, ");
            sql.append("        ejd_price_procs.get_retail_price(?, a.item_ea_id)      AS RetailPrice, ");
            sql.append("        a.r12_units, ");
            sql.append("        CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("          THEN 0 ");
            sql.append("        ELSE (SELECT retail_a ");
            sql.append("              FROM ejd_item_price ");
            sql.append("              WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("        END                                                             AS EJDRetailA, ");
            sql.append("        CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("          THEN 0 ");
            sql.append("        ELSE (SELECT retail_b ");
            sql.append("              FROM ejd_item_price ");
            sql.append("              WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("        END                                                             AS EJDRetailB, ");
            sql.append("        CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("          THEN 0 ");
            sql.append("        ELSE (SELECT retail_c ");
            sql.append("              FROM ejd_item_price ");
            sql.append("              WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("        END                                                             AS EJDRetailC, ");
            sql.append("        iea2.retail_pack                                                AS EJDRetailPack, ");
            sql.append("        (SELECT stock_pack ");
            sql.append("         FROM ejd_item_warehouse ");
            sql.append("         WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11)    AS EJD_OM, ");
            sql.append("        CASE WHEN aix1.match = 1 THEN 'Matched Item' ");
            sql.append("             WHEN aix1.match = 2 THEN  'Like Item' ");
            sql.append("             else 'DISCONTINUED' END     AS MatchType, ");
            sql.append("        ven2.name as EJD_Vendor_Name, ");
            sql.append("        decode(ejd.broken_case_id, 1, '', 'NBC') as nbc, ");
            sql.append("        su1.name as Emery_UOM, ");
            sql.append("        su2.name as EJD_UOM ");
            sql.append("      FROM ( ");
            sql.append("             SELECT ");
            sql.append("               cust_nbr, ");
            sql.append("               item_ea_id, ");
            sql.append("               sum(r12_units) AS r12_units ");
            sql.append("             FROM ( ");
            sql.append("                    SELECT DISTINCT ");
            sql.append("                      CASE WHEN cust.parent_id IS NULL ");
            sql.append("                        THEN dtl.cust_nbr ");
            sql.append("                      ELSE cust.parent_id ");
            sql.append("                      END                  AS cust_nbr, ");
            sql.append("                      dtl.item_ea_id, ");
            sql.append("                      sum(dtl.qty_shipped) AS r12_units ");
            sql.append("                    FROM sa.inv_dtl dtl ");
            sql.append("                      INNER JOIN customer cust ON cust.customer_id = dtl.cust_nbr ");
            sql.append("                    WHERE dtl.invoice_date > current_date - 365 AND dtl.cust_nbr IN ( ");
            sql.append("                      SELECT DISTINCT customer_id ");
            sql.append("                      FROM customer ");
            sql.append("                      WHERE (customer_id = ? OR parent_id = ?)) ");
            sql.append("                    GROUP BY cust.parent_id, dtl.cust_nbr, dtl.item_ea_id ");
            sql.append("                  ) b ");
            sql.append("             GROUP BY cust_nbr, item_ea_id) a ");
            sql.append("        left outer JOIN ejd.item_entity_attr iea ON iea.item_ea_id = a.item_ea_id ");
            sql.append("        left outer join ejd.ship_unit su1 on su1.unit_id = iea.ship_unit_id ");
            sql.append("        INNER JOIN ejd.customer cust ON cust.customer_id = a.cust_nbr ");
            sql.append("        LEFT OUTER JOIN ejd.customer prnt ON prnt.customer_id = cust.parent_id ");
            sql.append("        INNER JOIN ejd.cust_warehouse cw ON cust.customer_id = cw.customer_id ");
            sql.append("        INNER JOIN ejd.ejd_item ejd ON ejd.ejd_item_id = iea.ejd_item_id ");
            sql.append("        left outer JOIN ejd.ejd_item_warehouse eiw ");
            sql.append("          ON eiw.ejd_item_id = iea.ejd_item_id AND eiw.warehouse_id = cw.warehouse_id ");
            sql.append("        left outer JOIN ejd.item_ea_taxonomy tax3 ON tax3.taxonomy_id = iea.taxonomy_id ");
            sql.append("        left outer JOIN ejd.item_ea_taxonomy tax2 ON tax2.taxonomy_id = tax3.parent_id ");
            sql.append("        left outer JOIN ejd.item_ea_taxonomy tax1 ON tax1.taxonomy_id = tax2.parent_id ");
            sql.append("        inner JOIN ejd.vendor ven ON iea.vendor_id = ven.vendor_id ");
            sql.append("        left outer JOIN ejd.vendor_item_ea_cross vix ON vix.vendor_id = iea.vendor_id AND vix.item_ea_id = iea.item_ea_id ");
            sql.append("        LEFT OUTER JOIN ( ");
            sql.append("                          SELECT ");
            sql.append("                            aix1.item_id         AS item_id, ");
            sql.append("                            aix1.ace_sku         AS ace_sku, ");
            sql.append("                            aix1.match, ");
            sql.append("                            CASE WHEN aix.item_id IS NULL ");
            sql.append("                              THEN aix1.item_id ");
            sql.append("                            ELSE aix.item_id END AS aix_item_id ");
            sql.append("                          FROM scratch_pad.ace_item_xref_1 aix1 ");
            sql.append("                            LEFT OUTER JOIN ejd.ace_item_xref aix ON aix.ace_sku = aix1.ace_sku ");
            sql.append("                        ) AS aix1 ON aix1.item_id = iea.item_id ");
            sql.append("        LEFT OUTER JOIN ejd.item_ea_cross iex ON iex.customer_id = cust.customer_id AND iex.item_ea_id = iea.item_ea_id ");
            sql.append("        LEFT OUTER JOIN ejd.item_entity_attr iea2 ON iea2.item_id = aix1.aix_item_id AND iea2.item_type_id = 8 ");
            sql.append("        left outer join ejd.ship_unit su2 on su2.unit_id = iea2.ship_unit_id ");
            sql.append("        LEFT OUTER JOIN ejd.vendor_item_ea_cross vixa ");
            sql.append("          ON vixa.vendor_id = iea2.vendor_id AND vixa.item_ea_id = iea2.item_ea_id ");
            sql.append("        left outer join ejd.vendor ven2 on ven2.vendor_id = iea2.vendor_id ");
            sql.append("      ORDER BY tax1.taxonomy, ven.name, tax2.taxonomy, iea.item_id DESC ");
            sql.append("    ) b ");
            sql.append("where b.MatchType = 'Like Item' and b.dept <> 'Store Supplies' ");
            m_LikeItems = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("SELECT ");
            sql.append("    b.dept AS taxonomy_1, ");
            sql.append("    b.category AS taxonomy_2, ");
            sql.append("    b.item_id, ");
            sql.append("    b.customer_sku, ");
            sql.append("    b.emery_upc, ");
            sql.append("    b.mfrnum, ");
            sql.append("    b.Emery_Description AS item_description, ");
            sql.append("    b.vendorname AS cur_vendor, ");
            sql.append("    b.r12_units, ");
            sql.append("    b.sellprice AS cur_cost, ");
            sql.append("    b.r12_units * b.sellprice as cur_cost_total, ");
            sql.append("    b.retailprice  AS cur_retail, ");
            sql.append("    b.r12_units * b.retailprice as cur_ret_tot, ");
            sql.append("    CASE WHEN b.retailprice IS NULL OR b.sellprice IS NULL OR b.retail_pack IS NULL ");
            sql.append("      THEN 0 ");
            sql.append("    WHEN b.retailprice = 0 OR b.retail_pack = 0 ");
            sql.append("      THEN 0 ");
            sql.append("    ELSE round(((b.retailprice - (b.sellprice / b.retail_pack)) / b.retailprice) * 100, 4) ");
            sql.append("    END AS cur_margin, ");
            sql.append("    b.stock_pack, ");
            sql.append("    b.om  AS cur_om, ");
            sql.append("    b.Emery_UOM as cur_uom, ");
            sql.append("    b.nbc, ");
            sql.append("    b.MatchType as match ");
            sql.append("  FROM ");
            sql.append("    ( ");
            sql.append("      SELECT ");
            sql.append("        CASE WHEN cust.parent_id IS NULL ");
            sql.append("          THEN a.cust_nbr ");
            sql.append("        ELSE cust.parent_id ");
            sql.append("        END                                                             AS Cust_nbr, ");
            sql.append("        CASE WHEN prnt.name IS NULL ");
            sql.append("          THEN cust.name ");
            sql.append("        ELSE prnt.name ");
            sql.append("        END                                                             AS CustName, ");
            sql.append("        a.item_ea_id, ");
            sql.append("        iea.item_id, ");
            sql.append("        iex.customer_sku, ");
            sql.append("        aix1.ace_sku, ");
            sql.append("        iea.description                                                 AS Emery_Description, ");
            sql.append("        iea2.description                                                AS EJD_Description, ");
            sql.append("        ven.name                                                        AS VendorName, ");
            sql.append("        vix.vendor_item_num                                             AS MfrNum, ");
            sql.append("        vixa.vendor_item_num                                            AS EJD_MfrNum, ");
            sql.append("        tax1.taxonomy                                                   AS Dept, ");
            sql.append("        tax2.taxonomy                                                   AS Category, ");
            sql.append("        tax3.taxonomy                                                   AS SubCategory, ");
            sql.append("        (SELECT upc_code ");
            sql.append("         FROM ejd_item_whs_upc ");
            sql.append("         WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11 ");
            sql.append("         ORDER BY primary_upc DESC ");
            sql.append("         LIMIT 1)                                                       AS EJD_UPC, ");
            sql.append("        (SELECT upc_code ");
            sql.append("         FROM ejd_item_whs_upc ");
            sql.append("         WHERE ejd_item_id = iea.ejd_item_id AND warehouse_id IN (1, 2) ");
            sql.append("         ORDER BY primary_upc DESC ");
            sql.append("         LIMIT 1)                                                       AS emery_UPC, ");
            sql.append("        eiw.stock_pack, ");
            sql.append("        CASE WHEN ejd.broken_case_id <> 1 ");
            sql.append("          THEN eiw.stock_pack ");
            sql.append("        ELSE 1 ");
            sql.append("        END                                                             AS OM, ");
            sql.append("        iea.retail_pack, ");
            sql.append("        (ejd_cust_procs.get_sell_price(?, a.item_ea_id)).price AS sellPrice, ");
            sql.append("        CASE WHEN iea2.item_ea_id IS NULL ");
            sql.append("          THEN 0 ");
            sql.append("        ELSE (ejd_cust_procs.get_sell_price('008229', iea2.item_ea_id)).price ");
            sql.append("        END                                                             AS EJDsellPrice, ");
            sql.append("        ejd_price_procs.get_retail_price(?, a.item_ea_id)      AS RetailPrice, ");
            sql.append("        a.r12_units, ");
            sql.append("        CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("          THEN 0 ");
            sql.append("        ELSE (SELECT retail_a ");
            sql.append("              FROM ejd_item_price ");
            sql.append("              WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("        END                                                             AS EJDRetailA, ");
            sql.append("        CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("          THEN 0 ");
            sql.append("        ELSE (SELECT retail_b ");
            sql.append("              FROM ejd_item_price ");
            sql.append("              WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("        END                                                             AS EJDRetailB, ");
            sql.append("        CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("          THEN 0 ");
            sql.append("        ELSE (SELECT retail_c ");
            sql.append("              FROM ejd_item_price ");
            sql.append("              WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("        END                                                             AS EJDRetailC, ");
            sql.append("        iea2.retail_pack                                                AS EJDRetailPack, ");
            sql.append("        (SELECT stock_pack ");
            sql.append("         FROM ejd_item_warehouse ");
            sql.append("         WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11)    AS EJD_OM, ");
            sql.append("        CASE WHEN aix1.match = 1 THEN 'Matched Item' ");
            sql.append("             WHEN aix1.match = 2 THEN  'Like Item' ");
            sql.append("             WHEN eiw.disp_id <> 1 then 'PREVIOUSLY DISCONTINUED' ");                     
            sql.append("             else 'DISCONTINUED' END     AS MatchType, ");
            sql.append("        ven2.name as EJD_Vendor_Name, ");
            sql.append("        decode(ejd.broken_case_id, 1, '', 'NBC') as nbc, ");
            sql.append("        su1.name as Emery_UOM, ");
            sql.append("        su2.name as EJD_UOM ");
            sql.append("      FROM ( ");
            sql.append("             SELECT ");
            sql.append("               cust_nbr, ");
            sql.append("               item_ea_id, ");
            sql.append("               sum(r12_units) AS r12_units ");
            sql.append("             FROM ( ");
            sql.append("                    SELECT DISTINCT ");
            sql.append("                      CASE WHEN cust.parent_id IS NULL ");
            sql.append("                        THEN dtl.cust_nbr ");
            sql.append("                      ELSE cust.parent_id ");
            sql.append("                      END                  AS cust_nbr, ");
            sql.append("                      dtl.item_ea_id, ");
            sql.append("                      sum(dtl.qty_shipped) AS r12_units ");
            sql.append("                    FROM sa.inv_dtl dtl ");
            sql.append("                      INNER JOIN customer cust ON cust.customer_id = dtl.cust_nbr ");
            sql.append("                    WHERE dtl.invoice_date > current_date - 365 AND dtl.cust_nbr IN ( ");
            sql.append("                      SELECT DISTINCT customer_id ");
            sql.append("                      FROM customer ");
            sql.append("                      WHERE (customer_id = ? OR parent_id = ?)) ");
            sql.append("                    GROUP BY cust.parent_id, dtl.cust_nbr, dtl.item_ea_id ");
            sql.append("                  ) b ");
            sql.append("             GROUP BY cust_nbr, item_ea_id) a ");
            sql.append("        left outer JOIN ejd.item_entity_attr iea ON iea.item_ea_id = a.item_ea_id ");
            sql.append("        left outer join ejd.ship_unit su1 on su1.unit_id = iea.ship_unit_id ");
            sql.append("        INNER JOIN ejd.customer cust ON cust.customer_id = a.cust_nbr ");
            sql.append("        LEFT OUTER JOIN ejd.customer prnt ON prnt.customer_id = cust.parent_id ");
            sql.append("        INNER JOIN ejd.cust_warehouse cw ON cust.customer_id = cw.customer_id ");
            sql.append("        INNER JOIN ejd.ejd_item ejd ON ejd.ejd_item_id = iea.ejd_item_id ");
            sql.append("        left outer JOIN ejd.ejd_item_warehouse eiw ON eiw.ejd_item_id = iea.ejd_item_id AND eiw.warehouse_id = cw.warehouse_id ");
            sql.append("        left outer JOIN ejd.item_ea_taxonomy tax3 ON tax3.taxonomy_id = iea.taxonomy_id ");
            sql.append("        left outer JOIN ejd.item_ea_taxonomy tax2 ON tax2.taxonomy_id = tax3.parent_id ");
            sql.append("        left outer JOIN ejd.item_ea_taxonomy tax1 ON tax1.taxonomy_id = tax2.parent_id ");
            sql.append("        inner JOIN ejd.vendor ven ON iea.vendor_id = ven.vendor_id ");
            sql.append("        left outer JOIN ejd.vendor_item_ea_cross vix ON vix.vendor_id = iea.vendor_id AND vix.item_ea_id = iea.item_ea_id ");
            sql.append("        LEFT OUTER JOIN ( ");
            sql.append("                          SELECT ");
            sql.append("                            aix1.item_id         AS item_id, ");
            sql.append("                            aix1.ace_sku         AS ace_sku, ");
            sql.append("                            aix1.match, ");
            sql.append("                            CASE WHEN aix.item_id IS NULL ");
            sql.append("                              THEN aix1.item_id ");
            sql.append("                            ELSE aix.item_id END AS aix_item_id ");
            sql.append("                          FROM scratch_pad.ace_item_xref_1 aix1 ");
            sql.append("                            LEFT OUTER JOIN ejd.ace_item_xref aix ON aix.ace_sku = aix1.ace_sku ");
            sql.append("                        ) AS aix1 ON aix1.item_id = iea.item_id ");
            sql.append("        LEFT OUTER JOIN ejd.item_ea_cross iex ON iex.customer_id = cust.customer_id AND iex.item_ea_id = iea.item_ea_id ");
            sql.append("        LEFT OUTER JOIN ejd.item_entity_attr iea2 ON iea2.item_id = aix1.aix_item_id AND iea2.item_type_id = 8 ");
            sql.append("        left outer join ejd.ship_unit su2 on su2.unit_id = iea2.ship_unit_id ");
            sql.append("        LEFT OUTER JOIN ejd.vendor_item_ea_cross vixa ");
            sql.append("          ON vixa.vendor_id = iea2.vendor_id AND vixa.item_ea_id = iea2.item_ea_id ");
            sql.append("        left outer join ejd.vendor ven2 on ven2.vendor_id = iea2.vendor_id ");
            sql.append("      ORDER BY tax1.taxonomy, ven.name, tax2.taxonomy, iea.item_id DESC ");
            sql.append("    ) b ");
            sql.append("where b.MatchType in ( 'DISCONTINUED', 'PREVIOUSLY DISCONTINUED') and b.dept <> 'Store Supplies' ");
            m_DiscoItems = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select ");
            sql.append("    b.dept              as taxonomy_1, ");
            sql.append("    b.category          as taxonomy_2, ");
            sql.append("    b.item_id, ");
            sql.append("    b.customer_sku, ");
            sql.append("    b.emery_upc, ");
            sql.append("    b.mfrnum, ");
            sql.append("    b.emery_description as item_description, ");
            sql.append("    b.vendorname        as cur_vendor, ");
            sql.append("    b.ace_sku           as ejd_item, ");
            sql.append("    b.ejd_upc, ");
            sql.append("    b.ejd_mfrnum, ");
            sql.append("    b.ejd_description, ");
            sql.append("    b.ejd_vendor_name, ");
            sql.append("    b.stock_pack, ");
            sql.append("    b.om                as cur_om, ");
            sql.append("    b.ejd_om            as new_om, ");
            sql.append("    b.om - b.ejd_om     as om_diff, ");
            sql.append("    b.emery_uom         as cur_uom, ");
            sql.append("    b.ejd_uom           as new_uom, ");
            sql.append("    b.nbc, ");
            sql.append("    b.r12_units, ");
            sql.append("    b.sellprice         as cur_cost, ");
            sql.append("    b.ejdsellprice      as new_cost, ");
            sql.append("    b.r12_units * b.sellprice as cur_cost_tot, ");
            sql.append("    b.r12_units * b.ejdsellprice as new_cost_tot, ");
            sql.append("    (b.r12_units * b.ejdsellprice ) - (b.r12_units * b.sellprice) as cost_variance, ");
            sql.append("    b.retailprice       as cur_retail, ");
            sql.append("    b.r12_units * b.retailprice as cur_ret_tot, ");
            sql.append("    case when b.retailprice is null or b.sellprice is null or b.retail_pack is null ");
            sql.append("      then 0 ");
            sql.append("    when b.retailprice = 0 or b.retail_pack = 0 ");
            sql.append("      then 0 ");
            sql.append("    else round(((b.retailprice - (b.sellprice / b.retail_pack)) / b.retailprice) * 100, 4) ");
            sql.append("    end                 as cur_margin, ");
            sql.append("    b.ejdretaila        as comp_ret, ");
            sql.append("    b.ejdretaila  * b.r12_units  as comp_ret_tot, ");
            sql.append("    case when b.ejdretaila is null or b.ejdsellprice is null or b.ejdretailpack is null ");
            sql.append("      then 0 ");
            sql.append("    when b.ejdretaila = 0 or b.ejdretailpack = 0 ");
            sql.append("      then 0 ");
            sql.append("    else round(((b.ejdretaila - (b.ejdsellprice / b.ejdretailpack)) / b.ejdretaila) * 100, 4) ");
            sql.append("    end                 as comp_margin_pct, ");
            sql.append("    b.ejdretailb        as market_ret, ");
            sql.append("    b.ejdretailb  *  b.r12_units     as market_ret_tot, ");
            sql.append("    case when b.ejdretailb is null or b.ejdsellprice is null or b.ejdretailpack is null ");
            sql.append("      then 0 ");
            sql.append("    when b.ejdretailb = 0 or b.ejdretailpack = 0 ");
            sql.append("      then 0 ");
            sql.append("    else round(((b.ejdretailb - (b.ejdsellprice / b.ejdretailpack)) / b.ejdretailb) * 100, 4) ");
            sql.append("    end                 as market_margin_pct, ");
            sql.append("    b.ejdretailc        as premium_ret, ");
            sql.append("    b.ejdretailc  *  b.r12_units      as premium_ret_tot, ");
            sql.append("    case when b.ejdretailc is null or b.ejdsellprice is null or b.ejdretailpack is null ");
            sql.append("      then 0 ");
            sql.append("    when b.ejdretailc = 0 or b.ejdretailpack = 0 ");
            sql.append("      then 0 ");
            sql.append("    else round(((b.ejdretailc - (b.ejdsellprice / b.ejdretailpack)) / b.ejdretailc) * 100, 4) ");
            sql.append("    end                 as premium_margin_pct, ");
            sql.append("    b.MatchType         as match   ");
            sql.append("  FROM ");
            sql.append("    ( ");
            sql.append("      select ");
            sql.append("        case when cust.parent_id is null ");
            sql.append("          then a.cust_nbr ");
            sql.append("        else cust.parent_id ");
            sql.append("        end                                                             as cust_nbr, ");
            sql.append("        case when prnt.name is null ");
            sql.append("          then cust.name ");
            sql.append("        else prnt.name ");
            sql.append("        end                                                             as custname, ");
            sql.append("        a.item_ea_id, ");
            sql.append("        iea.item_id, ");
            sql.append("        iex.customer_sku, ");
            sql.append("        aix1.ace_sku, ");
            sql.append("        iea.description                                                 as emery_description, ");
            sql.append("        iea2.description                                                as ejd_description, ");
            sql.append("        ven.name                                                        as vendorname, ");
            sql.append("        vix.vendor_item_num                                             as mfrnum, ");
            sql.append("        vixa.vendor_item_num                                            as ejd_mfrnum, ");
            sql.append("        tax1.taxonomy                                                   as dept, ");
            sql.append("        tax2.taxonomy                                                   as category, ");
            sql.append("        tax3.taxonomy                                                   as subcategory, ");
            sql.append("        (select upc_code ");
            sql.append("         from ejd_item_whs_upc ");
            sql.append("         where ejd_item_id = iea2.ejd_item_id and warehouse_id = 11 ");
            sql.append("         order by primary_upc desc ");
            sql.append("         limit 1)                                                       as ejd_upc, ");
            sql.append("        (select upc_code ");
            sql.append("         from ejd_item_whs_upc ");
            sql.append("         where ejd_item_id = iea.ejd_item_id and warehouse_id in (1, 2) ");
            sql.append("         order by primary_upc desc ");
            sql.append("         limit 1)                                                       as emery_upc, ");
            sql.append("        eiw.stock_pack, ");
            sql.append("        case when ejd.broken_case_id <> 1 ");
            sql.append("          then eiw.stock_pack ");
            sql.append("        else 1 ");
            sql.append("        end                                                             as om, ");
            sql.append("        iea.retail_pack, ");
            sql.append("        (ejd_cust_procs.get_sell_price(?, a.item_ea_id)).price as sellprice, ");
            sql.append("        CASE WHEN iea2.item_ea_id IS NULL ");
            sql.append("          THEN 0 ");
            sql.append("          ELSE (ejd_cust_procs.get_sell_price(( SELECT CASE WHEN proxy_cust_id IS NULL THEN '008229' ELSE proxy_cust_id END AS proxy_cust_id ");
            sql.append("          FROM ejd.cust_migration WHERE customer_id = ?), iea2.item_ea_id)).price ");
            sql.append("        END AS ejdsellprice, ");
            sql.append("        ejd_price_procs.get_retail_price(?, a.item_ea_id)      as retailprice, ");
            sql.append("        a.r12_units, ");
            sql.append("        case when iea2.ejd_item_id is null ");
            sql.append("          then 0 ");
            sql.append("        else (select retail_a ");
            sql.append("              from ejd_item_price ");
            sql.append("              where ejd_item_id = iea2.ejd_item_id and warehouse_id = 11) ");
            sql.append("        end                                                             as ejdretaila, ");
            sql.append("        case when iea2.ejd_item_id is null ");
            sql.append("          then 0 ");
            sql.append("        else (select retail_b ");
            sql.append("              from ejd_item_price ");
            sql.append("              where ejd_item_id = iea2.ejd_item_id and warehouse_id = 11) ");
            sql.append("        end                                                             as ejdretailb, ");
            sql.append("        case when iea2.ejd_item_id is null ");
            sql.append("          then 0 ");
            sql.append("        else (select retail_c ");
            sql.append("              from ejd_item_price ");
            sql.append("              where ejd_item_id = iea2.ejd_item_id and warehouse_id = 11) ");
            sql.append("        end                                                             as ejdretailc, ");
            sql.append("        iea2.retail_pack                                                as ejdretailpack, ");
            sql.append("        (select stock_pack ");
            sql.append("         from ejd_item_warehouse ");
            sql.append("         where ejd_item_id = iea2.ejd_item_id and warehouse_id = 11)    as ejd_om, ");
            sql.append("        case when aix1.match = 1 then 'Matched Item' ");
            sql.append("             when aix1.match = 2 then  'Like Item' ");
            sql.append("             else 'DISCONTINUED' end     as matchtype, ");
            sql.append("        ven2.name as ejd_vendor_name, ");
            sql.append("        decode(ejd.broken_case_id, 1, '', 'NBC') as nbc, ");
            sql.append("        su1.name as emery_uom, ");
            sql.append("        su2.name as ejd_uom ");
            sql.append("      from ( ");
            sql.append("             select ");
            sql.append("               cust_nbr, ");
            sql.append("               item_ea_id, ");
            sql.append("               sum(r12_units) as r12_units ");
            sql.append("             from ( ");
            sql.append("                    select distinct ");
            sql.append("                      case when cust.parent_id is null ");
            sql.append("                        then dtl.cust_nbr ");
            sql.append("                      else cust.parent_id ");
            sql.append("                      end                  as cust_nbr, ");
            sql.append("                      dtl.item_ea_id, ");
            sql.append("                      sum(dtl.qty_shipped) as r12_units ");
            sql.append("                    from sa.inv_dtl dtl ");
            sql.append("                      inner join customer cust on cust.customer_id = dtl.cust_nbr ");
            sql.append("                    where dtl.invoice_date > current_date - 365 and dtl.cust_nbr in ( ");
            sql.append("                      select distinct customer_id ");
            sql.append("                      from customer ");
            sql.append("                      where (customer_id = ? or parent_id = ?)) ");
            sql.append("                    group by cust.parent_id, dtl.cust_nbr, dtl.item_ea_id ");
            sql.append("                  ) b ");
            sql.append("             group by cust_nbr, item_ea_id) a ");
            sql.append("        left outer join ejd.item_entity_attr iea on iea.item_ea_id = a.item_ea_id ");
            sql.append("        left outer join ejd.ship_unit su1 on su1.unit_id = iea.ship_unit_id ");
            sql.append("        inner join ejd.customer cust on cust.customer_id = a.cust_nbr ");
            sql.append("        left outer join ejd.customer prnt on prnt.customer_id = cust.parent_id ");
            sql.append("        inner join ejd.cust_warehouse cw on cust.customer_id = cw.customer_id ");
            sql.append("        inner join ejd.ejd_item ejd on ejd.ejd_item_id = iea.ejd_item_id ");
            sql.append("        left outer join ejd.ejd_item_warehouse eiw ");
            sql.append("          on eiw.ejd_item_id = iea.ejd_item_id and eiw.warehouse_id = cw.warehouse_id ");
            sql.append("        left outer join ejd.item_ea_taxonomy tax3 on tax3.taxonomy_id = iea.taxonomy_id ");
            sql.append("        left outer join ejd.item_ea_taxonomy tax2 on tax2.taxonomy_id = tax3.parent_id ");
            sql.append("        left outer join ejd.item_ea_taxonomy tax1 on tax1.taxonomy_id = tax2.parent_id ");
            sql.append("        inner join ejd.vendor ven on iea.vendor_id = ven.vendor_id ");
            sql.append("        left outer join ejd.vendor_item_ea_cross vix on vix.vendor_id = iea.vendor_id and vix.item_ea_id = iea.item_ea_id ");
            sql.append("        left outer join ( ");
            sql.append("                          select ");
            sql.append("                            aix1.item_id         as item_id, ");
            sql.append("                            aix1.ace_sku         as ace_sku, ");
            sql.append("                            aix1.match, ");
            sql.append("                            case when aix.item_id is null ");
            sql.append("                              then aix1.item_id ");
            sql.append("                            else aix.item_id end as aix_item_id ");
            sql.append("                          from scratch_pad.ace_item_xref_1 aix1 ");
            sql.append("                            left outer join ejd.ace_item_xref aix on aix.ace_sku = aix1.ace_sku ");
            sql.append("                        ) as aix1 on aix1.item_id = iea.item_id ");
            sql.append("        left outer join ejd.item_ea_cross iex on iex.customer_id = cust.customer_id and iex.item_ea_id = iea.item_ea_id ");
            sql.append("        left outer join ejd.item_entity_attr iea2 on iea2.item_id = aix1.aix_item_id and iea2.item_type_id = 8 ");
            sql.append("        left outer join ejd.ship_unit su2 on su2.unit_id = iea2.ship_unit_id ");
            sql.append("        left outer join ejd.vendor_item_ea_cross vixa ");
            sql.append("          on vixa.vendor_id = iea2.vendor_id and vixa.item_ea_id = iea2.item_ea_id ");
            sql.append("        left outer join ejd.vendor ven2 on ven2.vendor_id = iea2.vendor_id ");
            sql.append("      order by tax1.taxonomy, ven.name, tax2.taxonomy, iea.item_id desc ");
            sql.append("    ) b ");
            sql.append("where b.om <> b.ejd_om and b.dept <> 'Store Supplies' ");
            m_OMItems = m_EdbConn.prepareStatement(sql.toString());
            
            sql.setLength(0);
            sql.append("select  ");
            sql.append("   c.\"Taxonomy 1\" as taxonomy_1, ");
            sql.append("   sum(c.\"R12 Units\") as r12_units, ");
            sql.append("   sum(c.\"Current Cost Total\") as cur_cost_tot, ");
            sql.append("   sum(c.\"New Cost Total\") as new_cost_tot, ");
            sql.append("   sum(c.\"Current Retail Total\") as cur_ret_tot, ");
            sql.append("   case when sum(c.\"Current Retail Total\") = 0 then 0 else ");
            sql.append("   round(((sum(c.\"Current Retail Total\") - sum(c.\"Current Cost Total\")) /  sum(c.\"Current Retail Total\")) * 100, 4) ");
            sql.append("   end as cur_margin_pct, ");
            sql.append("   sum(c.\"Competitve Retail Total\") as comp_ret_total, ");
            sql.append("   case when sum(c.\"Competitve Retail Total\") = 0 then 0 else ");
            sql.append("   round(((sum(c.\"Competitve Retail Total\") - sum(c.\"New Cost Total\")) /  sum(c.\"Competitve Retail Total\")) * 100, 4) ");
            sql.append("   end as comp_margin_pct, ");
            sql.append("   sum(c.\"Market Retail Total\") as market_ret_tot, ");
            sql.append("   case when sum(c.\"Market Retail Total\") = 0 then 0 else ");
            sql.append("   round(((sum(c.\"Market Retail Total\") - sum(c.\"New Cost Total\")) /  sum(c.\"Market Retail Total\")) * 100, 4) ");
            sql.append("   end  as market_margin_pct, ");
            sql.append("   sum(c.\"Premium Retail Total\") as premium_ret_tot, ");
            sql.append("   case when sum(c.\"Premium Retail Total\") = 0 then 0 else ");
            sql.append("   round(((sum(c.\"Premium Retail Total\") - sum(c.\"New Cost Total\")) /  sum(c.\"Premium Retail Total\")) * 100, 4) ");
            sql.append("   end as premium_margin_pct, ");
            sql.append("   c.match ");
            sql.append("from ( ");
            sql.append("  SELECT ");
            sql.append("  b.dept AS \"Taxonomy 1\", ");
            sql.append("  b.category AS \"Taxonomy 2\", ");
            sql.append("  b.item_id AS \"Item #\", ");
            sql.append("  b.customer_sku AS \"Customer Sku\", ");
            sql.append("  b.emery_UPC AS EmeryUPC, ");
            sql.append("  b.mfrnum AS \"Mfg #\", ");
            sql.append("  b.Emery_Description AS \"Item Description\", ");
            sql.append("  b.vendorname AS \"Current Vendor\", ");
            sql.append("  b.ace_sku AS \"EJD Item#\", ");
            sql.append("  b.EJD_UPC AS EJD_UPC, ");
            sql.append("  b.EJD_MfrNum AS EJD_MFG, ");
            sql.append("  b.EJD_Description AS EJD_Description, ");
            sql.append("  b.EJD_Vendor_Name AS \"EJD Vendor\", ");
            sql.append("  b.r12_units AS \"R12 Units\", ");
            sql.append("  b.sellprice AS \"Current Cost\", ");
            sql.append("  b.ejdsellprice AS \"New Cost\", ");
            sql.append("  b.r12_units * b.sellprice AS \"Current Cost Total\", ");
            sql.append("  b.r12_units * b.ejdsellprice AS \"New Cost Total\", ");
            sql.append("  (b.r12_units * b.ejdsellprice ) - (b.r12_units * b.sellprice) AS \"Cost Variance\", ");
            sql.append("  b.retailprice AS \"Current Retail\", ");
            sql.append("  b.r12_units * b.retailprice AS \"Current Retail Total\", ");
            sql.append("  CASE WHEN b.retailprice IS NULL OR b.sellprice IS NULL OR b.retail_pack IS NULL ");
            sql.append("  THEN 0 ");
            sql.append("  WHEN b.retailprice = 0 OR b.retail_pack = 0 ");
            sql.append("  THEN 0 ");
            sql.append("  ELSE round(((b.retailprice - (b.sellprice / b.retail_pack)) / b.retailprice) * 100, 4) ");
            sql.append("  END AS \"Current Margin\", ");
            sql.append("  b.EJDRetailA AS \"Competitve Retail\", ");
            sql.append("  b.EJDRetailA  * b.r12_units AS \"Competitve Retail Total\", ");
            sql.append("  CASE WHEN b.EJDRetailA IS NULL OR b.ejdsellprice IS NULL OR b.EJDRetailPack IS NULL ");
            sql.append("  THEN 0 ");
            sql.append("  WHEN b.EJDRetailA = 0 OR b.EJDRetailPack = 0 ");
            sql.append("  THEN 0 ");
            sql.append("  ELSE round(((b.EJDRetailA - (b.ejdsellprice / b.EJDRetailPack)) / b.EJDRetailA) * 100, 4) ");
            sql.append("  END AS \"Competitive Margin %\", ");
            sql.append("  b.EJDRetailB AS \"Market Retail\", ");
            sql.append("  b.EJDRetailB  *  b.r12_units AS \"Market Retail Total\", ");
            sql.append("  CASE WHEN b.EJDRetailB IS NULL OR b.ejdsellprice IS NULL OR b.EJDRetailPack IS NULL ");
            sql.append("  THEN 0 ");
            sql.append("  WHEN b.EJDRetailB = 0 OR b.EJDRetailPack = 0 ");
            sql.append("  THEN 0 ");
            sql.append("  ELSE round(((b.EJDRetailB - (b.ejdsellprice / b.EJDRetailPack)) / b.EJDRetailB) * 100, 4) ");
            sql.append("  END AS \"Market Margin %\", ");
            sql.append("  b.EJDRetailC AS \"Premium Retail\", ");
            sql.append("  b.EJDRetailC  *  b.r12_units AS \"Premium Retail Total\", ");
            sql.append("  CASE WHEN b.EJDRetailC IS NULL OR b.ejdsellprice IS NULL OR b.EJDRetailPack IS NULL ");
            sql.append("  THEN 0 ");
            sql.append("  WHEN b.EJDRetailC = 0 OR b.EJDRetailPack = 0 ");
            sql.append("  THEN 0 ");
            sql.append("  ELSE round(((b.EJDRetailC - (b.ejdsellprice / b.EJDRetailPack)) / b.EJDRetailC) * 100, 4) ");
            sql.append("  END AS \"Premium Margin %\", ");
            sql.append("  b.EJD_OM AS \"New OM\", ");
            sql.append("  b.EJD_UOM AS \"New UOM\", ");
            sql.append("  b.nbc AS NBC, ");
            sql.append("  b.MatchType AS MATCH, ");
            sql.append("  b.stock_pack AS \"Stock Pack\", ");
            sql.append("  b.om AS \"Current OM\", ");
            sql.append("  b.Emery_UOM AS \"Current UOM\" ");
            sql.append("  FROM ");
            sql.append("  ( ");
            sql.append("     SELECT ");
            sql.append("     CASE WHEN cust.parent_id IS NULL ");
            sql.append("     THEN a.cust_nbr ");
            sql.append("     ELSE cust.parent_id ");
            sql.append("     END AS Cust_nbr, ");
            sql.append("     CASE WHEN prnt.name IS NULL ");
            sql.append("     THEN cust.name ");
            sql.append("     ELSE prnt.name ");
            sql.append("     END AS CustName, ");
            sql.append("     a.item_ea_id, ");
            sql.append("     iea.item_id, ");
            sql.append("     iex.customer_sku, ");
            sql.append("     aix1.ace_sku, ");
            sql.append("     iea.description AS Emery_Description, ");
            sql.append("     iea2.description AS EJD_Description, ");
            sql.append("     ven.name AS VendorName, ");
            sql.append("     vix.vendor_item_num AS MfrNum, ");
            sql.append("     vixa.vendor_item_num AS EJD_MfrNum, ");
            sql.append("     tax1.taxonomy AS Dept, ");
            sql.append("     tax2.taxonomy AS CATEGORY, ");
            sql.append("     tax3.taxonomy AS SubCategory, ");
            sql.append("     (SELECT upc_code FROM ejd_item_whs_upc WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11 ORDER BY primary_upc DESC LIMIT 1) AS EJD_UPC, ");
            sql.append("     (SELECT upc_code FROM ejd_item_whs_upc WHERE ejd_item_id = iea.ejd_item_id AND warehouse_id IN (1, 2) ORDER BY primary_upc DESC LIMIT 1) AS emery_UPC, ");
            sql.append("     eiw.stock_pack, ");
            sql.append("     CASE WHEN ejd.broken_case_id <> 1 ");
            sql.append("     THEN eiw.stock_pack ");
            sql.append("     ELSE 1 ");
            sql.append("     END AS OM, ");
            sql.append("     iea.retail_pack, ");
            sql.append("     (ejd_cust_procs.get_sell_price(?, a.item_ea_id)).price AS sellPrice, ");
            sql.append("     CASE WHEN iea2.item_ea_id IS NULL ");
            sql.append("     THEN 0 ");
            sql.append("     ELSE (ejd_cust_procs.get_sell_price('008229', iea2.item_ea_id)).price ");
            sql.append("     END AS EJDsellPrice, ");
            sql.append("     ejd_price_procs.get_retail_price(?, a.item_ea_id) AS RetailPrice, ");
            sql.append("     a.r12_units, ");
            sql.append("     CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("     THEN 0 ");
            sql.append("     ELSE ( SELECT retail_a ");
            sql.append("     FROM ejd_item_price ");
            sql.append("     WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("     END AS EJDRetailA, ");
            sql.append("     CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("     THEN 0 ");
            sql.append("     ELSE ( SELECT retail_b ");
            sql.append("     FROM ejd_item_price ");
            sql.append("     WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("     END AS EJDRetailB, ");
            sql.append("     CASE WHEN iea2.ejd_item_id IS NULL ");
            sql.append("     THEN 0 ");
            sql.append("     ELSE ( SELECT retail_c ");
            sql.append("     FROM ejd_item_price ");
            sql.append("     WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) ");
            sql.append("     END AS EJDRetailC, ");
            sql.append("     iea2.retail_pack AS EJDRetailPack, ");
            sql.append("     ( SELECT stock_pack ");
            sql.append("     FROM ejd_item_warehouse ");
            sql.append("     WHERE ejd_item_id = iea2.ejd_item_id AND warehouse_id = 11) AS EJD_OM, ");
            sql.append("     CASE WHEN aix1.match = 1 THEN 'Matched Item' ");
            sql.append("     WHEN aix1.match = 2 THEN 'Like Item' ");
            sql.append("     ELSE 'DISCONTINUED' END AS MatchType, ");
            sql.append("     ven2.name AS EJD_Vendor_Name, ");
            sql.append("     decode(ejd.broken_case_id, 1, '', 'NBC') AS nbc, ");
            sql.append("     su1.name AS Emery_UOM, ");
            sql.append("     su2.name AS EJD_UOM ");
            sql.append("     FROM ( ");
            sql.append("      SELECT ");
            sql.append("      cust_nbr, ");
            sql.append("      item_ea_id, ");
            sql.append("      sum(r12_units) AS r12_units ");
            sql.append("      FROM ( ");
            sql.append("        SELECT DISTINCT ");
            sql.append("        CASE WHEN cust.parent_id IS NULL ");
            sql.append("        THEN dtl.cust_nbr ");
            sql.append("        ELSE cust.parent_id ");
            sql.append("        END AS cust_nbr, ");
            sql.append("        dtl.item_ea_id, ");
            sql.append("        sum(dtl.qty_shipped) AS r12_units ");
            sql.append("        FROM sa.inv_dtl dtl ");
            sql.append("        INNER JOIN customer cust ON cust.customer_id = dtl.cust_nbr ");
            sql.append("        WHERE dtl.invoice_date > current_date - 365 AND dtl.cust_nbr IN ( ");
            sql.append("        SELECT DISTINCT customer_id ");
            sql.append("        FROM customer ");
            sql.append("        WHERE (customer_id = ? OR parent_id = ?)) ");
            sql.append("        GROUP BY cust.parent_id, dtl.cust_nbr, dtl.item_ea_id ");
            sql.append("      ) b ");
            sql.append("      GROUP BY cust_nbr, item_ea_id) a ");
            sql.append("     LEFT OUTER JOIN ejd.item_entity_attr iea ON iea.item_ea_id = a.item_ea_id ");
            sql.append("     LEFT OUTER JOIN ejd.ship_unit su1 ON su1.unit_id = iea.ship_unit_id ");
            sql.append("     INNER JOIN ejd.customer cust ON cust.customer_id = a.cust_nbr ");
            sql.append("     LEFT OUTER JOIN ejd.customer prnt ON prnt.customer_id = cust.parent_id ");
            sql.append("     INNER JOIN ejd.cust_warehouse cw ON cust.customer_id = cw.customer_id ");
            sql.append("     INNER JOIN ejd.ejd_item ejd ON ejd.ejd_item_id = iea.ejd_item_id ");
            sql.append("     LEFT OUTER JOIN ejd.ejd_item_warehouse eiw ");
            sql.append("     ON eiw.ejd_item_id = iea.ejd_item_id AND eiw.warehouse_id = cw.warehouse_id ");
            sql.append("     LEFT OUTER JOIN ejd.item_ea_taxonomy tax3 ON tax3.taxonomy_id = iea.taxonomy_id ");
            sql.append("     LEFT OUTER JOIN ejd.item_ea_taxonomy tax2 ON tax2.taxonomy_id = tax3.parent_id ");
            sql.append("     LEFT OUTER JOIN ejd.item_ea_taxonomy tax1 ON tax1.taxonomy_id = tax2.parent_id ");
            sql.append("     INNER JOIN ejd.vendor ven ON iea.vendor_id = ven.vendor_id ");
            sql.append("     LEFT OUTER JOIN ejd.vendor_item_ea_cross vix ON vix.vendor_id = iea.vendor_id AND vix.item_ea_id = iea.item_ea_id ");
            sql.append("     LEFT OUTER JOIN ( ");
            sql.append("     SELECT ");
            sql.append("     aix1.item_id AS item_id, ");
            sql.append("     aix1.ace_sku AS ace_sku, ");
            sql.append("     aix1.match, ");
            sql.append("     CASE WHEN aix.item_id IS NULL ");
            sql.append("     THEN aix1.item_id ");
            sql.append("     ELSE aix.item_id END AS aix_item_id ");
            sql.append("     FROM scratch_pad.ace_item_xref_1 aix1 ");
            sql.append("     LEFT OUTER JOIN ejd.ace_item_xref aix ON aix.ace_sku = aix1.ace_sku ");
            sql.append("     ) AS aix1 ON aix1.item_id = iea.item_id ");
            sql.append("     LEFT OUTER JOIN ejd.item_ea_cross iex ON iex.customer_id = cust.customer_id AND iex.item_ea_id = iea.item_ea_id ");
            sql.append("     LEFT OUTER JOIN ejd.item_entity_attr iea2 ON iea2.item_id = aix1.aix_item_id AND iea2.item_type_id = 8 ");
            sql.append("     LEFT OUTER JOIN ejd.ship_unit su2 ON su2.unit_id = iea2.ship_unit_id ");
            sql.append("     LEFT OUTER JOIN ejd.vendor_item_ea_cross vixa ");
            sql.append("     ON vixa.vendor_id = iea2.vendor_id AND vixa.item_ea_id = iea2.item_ea_id ");
            sql.append("     LEFT OUTER JOIN ejd.vendor ven2 ON ven2.vendor_id = iea2.vendor_id ");
            sql.append("     ORDER BY tax1.taxonomy, ven.name, tax2.taxonomy, iea.item_id DESC ");
            sql.append("  ) b ");
            sql.append("  where b.MatchType in ( 'Matched Item', 'Like Item') and b.dept <> 'Store Supplies' ");
            sql.append(" ");
            sql.append("  ) c ");
            sql.append(" group by c.MATCH, c.\"Taxonomy 1\" ");
            sql.append(" order by  c.MATCH desc, c.\"Taxonomy 1\" asc ");
            
            m_SumTax = m_EdbConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("[CustomerAssortment]", ex);
         }
         
         finally {
            sql = null;
         }
      }
      
      return isPrepared;
   }
   
   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      int pcount = params.size();
      Param param = null;
      String custName = "";
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
                  
         if ( param.name.equals("custid") )
            m_CustId = param.value;
         
         if ( param.name.equals("custname") )
            custName = param.value;
      }
      
      if ( custName == null || custName.length() == 0 )
         custName = m_CustId;
      
      m_FileNames.add(String.format("%s Summary.xlsx", custName));
   }
   
   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {
      m_Workbook = new XSSFWorkbook();
      m_Sheet1 = m_Workbook.createSheet("Summary");
            
      m_Sheet2 = m_Workbook.createSheet("MatchedItems");
      m_Sheet3 = m_Workbook.createSheet("LikeItems");
      m_Sheet4 = m_Workbook.createSheet("DiscoItems");
      m_Sheet5 = m_Workbook.createSheet("OMItems");
      m_Sheet6 = m_Workbook.createSheet("SummaryTaxMatch");
      
      defineStyles();
   }
}
