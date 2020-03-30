/**
 * $Id: OffSiteInventory.java,v 1.6 2012/05/09 07:13:25 pberggren Exp $
 * 
 * @author pberggren
 * 
 * $Log: OffSiteInventory.java,v $
 * Revision 1.6  2012/05/09 07:13:25  pberggren
 * Debugged SQL
 *
 * Revision 1.5  2012/05/09 05:46:16  pberggren
 * Removed commented code for cleanliness
 *
 * Revision 1.4  2012/05/08 02:57:45  prichter
 * Added cvs tags
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
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
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
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

/**
 * 
 *
 */
public class OffSiteInventory extends Report 
{
		
	    private static final short maxCols = 20;
	    
		//DB Data
		private PreparedStatement m_southPort;
		private PreparedStatement m_foreRiver;
		private PreparedStatement m_pittston;
	   
	   //report objects
	   //private StringBuffer m_Lines;
	   private XSSFWorkbook m_WrkBk;
	   private XSSFSheet m_South;
	   private XSSFSheet m_Fore;
	   private XSSFSheet m_Pitt;
	   //private short m_rownum = 1;
	   
	   //
	   // The cell styles for each of the columns in the spreadsheet.
	   private XSSFCellStyle[] m_CellStyles;
	   private XSSFCellStyle[] m_CellStylesEx;
	   
	   //
	   // Column widths
	   private static final int CW_ItemNum        = 2000;
	   private static final int CW_ItemDesc       = 6500;
	   private static final int CW_VendorNum      = 2000;
	   private static final int CW_Vendor         = 5000;
	   private static final int CW_DeptNum        = 2000;
	   private static final int CW_NeverOut       = 2000;
	   private static final int CW_TotalOH        = 2000;
	   private static final int CW_WareOH         = 2000;
	   private static final int CW_TrailerOH      = 2000;
	   private static final int CW_foreR          = 2000;
	   private static final int CW_PittS          = 2000;
	   private static final int CW_southP         = 2000;
	   private static final int CW_28Day          = 2000;
	   private static final int CW_LY28Day        = 2000;
	   private static final int CW_CurrWOS        = 2000;
	   private static final int CW_LastWOS        = 2000;
	   private static final int CW_FutureOrders   = 2000;
	   private static final int CW_CurrentOrders  = 2000;
	   private static final int CW_AvUnitCost     = 2000;
	   private static final int CW_TotValue       = 2000;
	   
	   /**
	    * Default constructor
	    */
	   public OffSiteInventory()
	   {
	      super();
	      
	      m_WrkBk = new XSSFWorkbook();
	      m_South = m_WrkBk.createSheet("South Portland"); 
	      m_Fore = m_WrkBk.createSheet("Fore River"); 
	      m_Pitt = m_WrkBk.createSheet("Pittston"); 
	      m_MaxRunTime = RptServer.HOUR * 12;      
	            
	      setupWorkbook();
	   }
	   
	   public void finalize() throws Throwable
	   {      
	      if ( m_CellStyles != null ) {
	         for ( int i = 0; i < m_CellStyles.length; i++ )
	            m_CellStyles[i] = null;
	      }
	      
	      if ( m_CellStylesEx != null ) {
	         for ( int i = 0; i < m_CellStylesEx.length; i++ )
	            m_CellStylesEx[i] = null;
	      }
	      
	      m_South = null;     
	      m_Fore = null;   
	      m_Pitt = null;
	      m_WrkBk = null;
	      m_CellStyles = null;
	      m_CellStylesEx = null;
	    	      
	      super.finalize();
	   }
	/**
	* Builds the output file
	* @return true if successful, false if not
	*/
	   
   private boolean buildSpreadsheet() throws FileNotFoundException
   {      	 
	  XSSFRow Row = null;
	  FileOutputStream OutFile = null;
	  ResultSet southP = null;
	  int RowNum = 0;	  
	  boolean Result = false;
	  String itemId = null;
	 	  
	  OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
	  RowNum = createRowCaptions();

	   try {
	         southP = m_southPort.executeQuery();

	         setCurAction("generating spreadsheet "); /* + m_CustId); */

	      while ( southP.next() && m_Status == RptServer.RUNNING ) {
	           try {
	               Row = createRow(m_South, RowNum);
	               itemId = southP.getString("item_id");
	               Row.getCell(0).setCellValue(new XSSFRichTextString(itemId));
	               Row.getCell(1).setCellValue(new XSSFRichTextString(southP.getString("description")));
	               Row.getCell(2).setCellValue(southP.getInt("vendor_id"));
	               Row.getCell(3).setCellValue(new XSSFRichTextString(southP.getString("name")));
	               Row.getCell(4).setCellValue(new XSSFRichTextString(southP.getString("dept_num")));
	               Row.getCell(5).setCellValue(new XSSFRichTextString(southP.getString("Never_Out")));
	               Row.getCell(6).setCellValue(southP.getInt("Total_OH"));
	               Row.getCell(7).setCellValue(southP.getInt("Whs_OH"));
	               Row.getCell(8).setCellValue(southP.getInt("Trailer_OH"));
	               Row.getCell(9).setCellValue(southP.getInt("Fore_R_OH"));
	               Row.getCell(10).setCellValue(southP.getInt("Pitt_OH"));
	               Row.getCell(11).setCellValue(southP.getInt("South_P_OH"));
	               Row.getCell(12).setCellValue(southP.getInt("28_Day"));
	               Row.getCell(13).setCellValue(southP.getInt("LY_28_Day"));
	               Row.getCell(14).setCellValue(southP.getFloat("Cur_WOS"));
	               Row.getCell(15).setCellValue(southP.getFloat("LY_WOS"));
	               Row.getCell(16).setCellValue(southP.getInt("Future_Orders"));
	               Row.getCell(17).setCellValue(southP.getInt("Current_Orders"));
	               Row.getCell(18).setCellValue(southP.getFloat("Avg_Unit_Cost"));
	               Row.getCell(19).setCellValue(southP.getFloat("Total_Value"));
	               
	               RowNum++;
	              
	           }
	            //Check for broken pipe errors.  Stop the process and send email.
	           catch ( Exception e ) {
	               log.error("exception", e);               
	               m_Status = RptServer.STOPPED;
	               //m_DBError = true;

	               m_ErrMsg.append("Your Off Site Inventory Report had the following errors: \r\n");
	               m_ErrMsg.append(e.getClass().getName() + "\r\n");
	               m_ErrMsg.append(e.getMessage());
	           }
	       }
	         
	         m_WrkBk.write(OutFile);
	         Result = true;
	    }

	      catch ( Exception e ) {
	    	  log.error("exception", e);
	      }
	      
	 	 finally {
	 	    DbUtils.closeDbConn(null, null, southP);
	 	    southP = null;
	 	    Row = null;
	 		RowNum = 0;
	 		//colNum = 0;
	 		itemId = null;
	 	    try {
	 	       OutFile.close();
	 	    }

	 	    catch( Exception e ) {
	 	       log.error(e);
	 	    }

	 	    //OutFile = null;
	 	 }
	
		      
	//Populate Fore Sheet
	XSSFRow Row2 = null;
	FileOutputStream OutFile2 = null;
	ResultSet foreR = null;
	int RowNum2 = 0;
	//int colNum2 = 0;
	//boolean Result2 = false;
	String itemId2 = null;
    //m_FileNames.set(0, m_FileNames.get(0) + m_WebRpt.getWebReportId() + ".xls" );
	OutFile2 = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
	      	      
	RowNum2 = createRowCaptions();

	   try {
	       foreR = m_foreRiver.executeQuery();

	       setCurAction("generating spreadsheet ");

	       while ( foreR.next() && m_Status == RptServer.RUNNING ) {
	           try {
	               Row2 = createRow2(m_Fore, RowNum2);
	               itemId2 = foreR.getString("item_id");
	               Row2.getCell(0).setCellValue(new XSSFRichTextString(itemId2));
	               Row2.getCell(1).setCellValue(new XSSFRichTextString(foreR.getString("description")));
	               Row2.getCell(2).setCellValue(foreR.getInt("vendor_id"));
	               Row2.getCell(3).setCellValue(new XSSFRichTextString(foreR.getString("name")));
	               Row2.getCell(4).setCellValue(new XSSFRichTextString(foreR.getString("dept_num")));
	               Row2.getCell(5).setCellValue(new XSSFRichTextString(foreR.getString("Never_Out")));
	               Row2.getCell(6).setCellValue(foreR.getInt("Total_OH"));
	               Row2.getCell(7).setCellValue(foreR.getInt("Whs_OH"));
	               Row2.getCell(8).setCellValue(foreR.getInt("Trailer_OH"));
	               Row2.getCell(9).setCellValue(foreR.getInt("Fore_R_OH"));
	               Row2.getCell(10).setCellValue(foreR.getInt("Pitt_OH"));
	               Row2.getCell(11).setCellValue(foreR.getInt("South_P_OH"));
	               Row2.getCell(12).setCellValue(foreR.getInt("28_Day"));
	               Row2.getCell(13).setCellValue(foreR.getInt("LY_28_Day"));
	               Row2.getCell(14).setCellValue(foreR.getFloat("Cur_WOS"));
	               Row2.getCell(15).setCellValue(foreR.getFloat("LY_WOS"));
	               Row2.getCell(16).setCellValue(foreR.getInt("Future_Orders"));
	               Row2.getCell(17).setCellValue(foreR.getInt("Current_Orders"));
	               Row2.getCell(18).setCellValue(foreR.getFloat("Avg_Unit_Cost"));
	               Row2.getCell(19).setCellValue(foreR.getFloat("Total_Value"));
	               
	               RowNum2++;
	               
	            }
	            //Check for broken pipe errors.  Stop the process and send email.
	            catch ( Exception e ) {
	               log.error("exception", e);               
	               m_Status = RptServer.STOPPED;
	               //m_DBError = true;

	               m_ErrMsg.append("Your Off Site Inventory report had the following errors: \r\n");
	               m_ErrMsg.append(e.getClass().getName() + "\r\n");
	               m_ErrMsg.append(e.getMessage());
	            }
	       }
	         
	         m_WrkBk.write(OutFile2);
	         //Result2 = true;
	   }

	   catch ( Exception e ) {
	     log.error("exception", e);
	   }
	      
	   finally {
	 	    DbUtils.closeDbConn(null, null, foreR);
	 	    foreR = null;

	 	  try {
	 	      OutFile2.close();
	 	  }

	 	  catch( Exception e ) {
	 	       log.error(e);
	      }

	 	    OutFile2 = null;
	   }

	  //Populate Fore Sheet
	  XSSFRow Row3 = null;
	  FileOutputStream OutFile3 = null;
      ResultSet pitts = null;
	  int RowNum3 = 0;
	  //int colNum3 = 0;
	  //boolean Result3 = false;
	  String itemId3 = null;

	  //m_FileNames.set(0, m_FileNames.get(0) + m_WebRpt.getWebReportId() + ".xls" );
	  OutFile3 = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
		      	      
		
	  RowNum3 = createRowCaptions();

		 try {
		        
		 pitts = m_pittston.executeQuery();

		 setCurAction("generating spreadsheet "); /* + m_CustId); */

		    while ( pitts.next() && m_Status == RptServer.RUNNING ) {
			    try {
			      Row3 = createRow3(m_Pitt, RowNum3);
			      itemId3 = pitts.getString("item_id");
			      Row3.getCell(0).setCellValue(new XSSFRichTextString(itemId3));
			      Row3.getCell(1).setCellValue(new XSSFRichTextString(pitts.getString("description")));
			      Row3.getCell(2).setCellValue(pitts.getInt("vendor_id"));
			      Row3.getCell(3).setCellValue(new XSSFRichTextString(pitts.getString("name")));
			      Row3.getCell(4).setCellValue(new XSSFRichTextString(pitts.getString("dept_num")));
			      Row3.getCell(5).setCellValue(new XSSFRichTextString(pitts.getString("Never_Out")));
			      Row3.getCell(6).setCellValue(pitts.getInt("Total_OH"));
			      Row3.getCell(7).setCellValue(pitts.getInt("Whs_OH"));
			      Row3.getCell(8).setCellValue(pitts.getInt("Trailer_OH"));
			      Row3.getCell(9).setCellValue(pitts.getInt("Fore_R_OH"));
			      Row3.getCell(10).setCellValue(pitts.getInt("Pitt_OH"));
			      Row3.getCell(11).setCellValue(pitts.getInt("South_P_OH"));
			      Row3.getCell(12).setCellValue(pitts.getInt("28_Day"));
			      Row3.getCell(13).setCellValue(pitts.getInt("LY_28_Day"));
			      Row3.getCell(14).setCellValue(pitts.getFloat("Cur_WOS"));
			      Row3.getCell(15).setCellValue(pitts.getFloat("LY_WOS"));
			      Row3.getCell(16).setCellValue(pitts.getInt("Future_Orders"));
			      Row3.getCell(17).setCellValue(pitts.getInt("Current_Orders"));
			      Row3.getCell(18).setCellValue(pitts.getFloat("Avg_Unit_Cost"));
			      Row3.getCell(19).setCellValue(pitts.getFloat("Total_Value"));
			               
			      RowNum3++;
			          
                }
			            //Check for broken pipe errors.  Stop the process and send email.
		          catch ( Exception e ) {
		          log.error("exception", e);               
		          m_Status = RptServer.STOPPED;
		          //m_DBError = true;

		          m_ErrMsg.append("Your Off Site Inventory report had the following errors: \r\n");
		          m_ErrMsg.append(e.getClass().getName() + "\r\n");
		          m_ErrMsg.append(e.getMessage());
		          }
		    }
			         
		          m_WrkBk.write(OutFile3);
		 // Result3 = true;
		 
	  }

	  catch ( Exception e ) {
	  	  log.error("exception", e);
	  }
		   
	  finally {
	    DbUtils.closeDbConn(null, null, pitts);
	    pitts = null;

	    try {
	       OutFile3.close();
	    }

	    catch( Exception e ) {
	       log.error(e);
	    }

	    OutFile3 = null;
	 }
	 
		 return Result;
		 //return Result2;
		 //return Result3;
	 	 }
   //}

   
	public int createRowCaptions()
	{
		XSSFCell cell = null;
		XSSFCell cell2 = null;
		XSSFCell cell3 = null;
		XSSFSheet m_South = null;
		XSSFSheet m_Fore = null;
		XSSFSheet m_Pitt = null;
	      XSSFRow row = null;
	      XSSFRow row2 = null;
	      XSSFRow row3 = null;
	      XSSFCellStyle styleCaptionsRow = null;
	      XSSFFont fontCaptionsRow = null;
	      int col = 0;
	      int col2 = 0;
	      int col3 = 0;
	      int rowNum = 0;
	      int rowNum2 = 0;
	      int rowNum3 = 0;
	      short rowHeight = 1000;
	      short rowHeight2 = 1000;
	      short rowHeight3 = 1000;
	      
	      fontCaptionsRow = m_WrkBk.createFont();
	      fontCaptionsRow.setFontHeightInPoints((short)8);
	      fontCaptionsRow.setFontName("Arial");
	      fontCaptionsRow.setBold(true);
	         
	      styleCaptionsRow = m_WrkBk.createCellStyle();
	      styleCaptionsRow.setFont(fontCaptionsRow);
	      styleCaptionsRow.setAlignment(HorizontalAlignment.CENTER);
	      //
	      //Shading
	      styleCaptionsRow.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
	      styleCaptionsRow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	      
	      //
	      //Border
	      styleCaptionsRow.setBorderTop(BorderStyle.THIN);// This is working
	      styleCaptionsRow.setBorderBottom(BorderStyle.THIN);
	      styleCaptionsRow.setBorderLeft(BorderStyle.THIN);
	      styleCaptionsRow.setBorderRight(BorderStyle.THIN);
	      
	      try {
	         while ( m_South != null ) {
	             //
	             // Create the row for the captions.
	             row = m_South.createRow(rowNum++);
	             row.setHeight(rowHeight);
	             
	             for ( int i = 0; i < maxCols; i++ ) {
	                cell = row.createCell(i);
	                cell.setCellStyle(styleCaptionsRow);
	             }
	             
	                //
	                //Rows for South Sheet
	             	rowNum++;
	             	row = m_South.createRow(rowNum);
	             	//row.setHeightInPoints((3*m_South.getDefaultRowHeightInPoints()));
	             	row.setRowStyle(styleCaptionsRow);
	             	col = 0;
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Item #"));
	             	m_South.setColumnWidth(col++, CW_ItemNum);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Item Desc"));
	             	m_South.setColumnWidth(col++, CW_ItemDesc);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Vendor #"));
	             	m_South.setColumnWidth(col++, CW_VendorNum);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Vendor Name"));
	             	m_South.setColumnWidth(col++, CW_Vendor);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Dept #"));
	             	m_South.setColumnWidth(col++, CW_DeptNum);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Never Out"));
	             	m_South.setColumnWidth(col++, CW_NeverOut);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Total OH"));
	             	m_South.setColumnWidth(col++, CW_TotalOH);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Whs OH"));
	             	m_South.setColumnWidth(col++, CW_WareOH);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Trailer OH"));
	             	m_South.setColumnWidth(col++, CW_TrailerOH);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Fore River OH"));
	             	m_South.setColumnWidth(col++, CW_foreR);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Pittston OH"));
	             	m_South.setColumnWidth(col++, CW_PittS);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("SO Port OH"));
	             	m_South.setColumnWidth(col++, CW_southP);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("28 Day Sales"));
	             	m_South.setColumnWidth(col++, CW_28Day);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("LY 28 Day Sales"));
	             	m_South.setColumnWidth(col++, CW_LY28Day);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Cur WOS"));
	             	m_South.setColumnWidth(col++, CW_CurrWOS);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("LY WOS"));
	             	m_South.setColumnWidth(col++, CW_LastWOS);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Future Orders"));
	             	m_South.setColumnWidth(col++, CW_FutureOrders);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Current Orders"));
	             	m_South.setColumnWidth(col++, CW_CurrentOrders);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Average Unit Cost"));
	             	m_South.setColumnWidth(col++, CW_AvUnitCost);
	             	row.getCell(col).setCellValue(new XSSFRichTextString("Total Cost"));
	             	m_South.setColumnWidth(col++, CW_TotValue);
	          }
	      //return rowNum;
	          { 
	          }
         }
	        finally{
        	  //i = 0;
        	 // cell = null;
	      
	   } 
	      
	      
	   try {
	      while ( m_Fore != null ) {
	         //
	         // Create the row for the captions.
		     row2 = m_Fore.createRow(rowNum2++);
		     row2.setHeight(rowHeight2);
		             
		     for ( int i2 = 0; i2 < maxCols; i2++ ) {
		     cell2 = row2.createCell(i2);
		     cell2.setCellStyle(styleCaptionsRow);
		     }
		     
		     	//
		     	//Rows for Fore Sheet
		     	rowNum2++;
		     	row2 = m_Fore.createRow(rowNum2);
	      		//row.setHeightInPoints((3*m_South.getDefaultRowHeightInPoints()));
		     	row2.setRowStyle(styleCaptionsRow);
		     	col2 = 0;
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Item #"));
		     	m_Fore.setColumnWidth(col2++, CW_ItemNum);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Item Desc"));
		     	m_Fore.setColumnWidth(col2++, CW_ItemDesc);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Vendor #"));
		     	m_Fore.setColumnWidth(col2++, CW_VendorNum);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Vendor Name"));
		     	m_Fore.setColumnWidth(col2++, CW_Vendor);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Dept #"));
		     	m_Fore.setColumnWidth(col2++, CW_DeptNum);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Never Out"));
		     	m_Fore.setColumnWidth(col2++, CW_NeverOut);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Total OH"));
		     	m_Fore.setColumnWidth(col2++, CW_TotalOH);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Whs OH"));
		     	m_Fore.setColumnWidth(col2++, CW_WareOH);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Trailer OH"));
		     	m_Fore.setColumnWidth(col2++, CW_TrailerOH);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Fore River OH"));
		     	m_Fore.setColumnWidth(col2++, CW_foreR);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Pittston OH"));
		     	m_Fore.setColumnWidth(col2++, CW_PittS);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("SO Port OH"));
		     	m_Fore.setColumnWidth(col2++, CW_southP);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("28 Day Sales"));
		     	m_Fore.setColumnWidth(col2++, CW_28Day);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("LY 28 Day Sales"));
		     	m_Fore.setColumnWidth(col2++, CW_LY28Day);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Cur WOS"));
		     	m_Fore.setColumnWidth(col2++, CW_CurrWOS);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("LY WOS"));
		     	m_Fore.setColumnWidth(col2++, CW_LastWOS);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Future Orders"));
		     	m_Fore.setColumnWidth(col2++, CW_FutureOrders);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Current Orders"));
		     	m_Fore.setColumnWidth(col2++, CW_CurrentOrders);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Average Unit Cost"));
		     	m_Fore.setColumnWidth(col2++, CW_AvUnitCost);
		     	row2.getCell(col2).setCellValue(new XSSFRichTextString("Total Cost"));
		     	m_Fore.setColumnWidth(col2++, CW_TotValue);
	    }
	   // return rowNum2;
	   }
	   finally{
        //i = 0;
        // cell = null;
        }
	    
	   
	 
	   
	  try {
	       while ( m_Pitt != null ) {
	           //
	           // Create the row for the captions.
	           row3 = m_Pitt.createRow(rowNum3++);
	           row3.setHeight(rowHeight3);
	             
	             for ( int i3 = 0; i3 < maxCols; i3++ ) {
	                cell3 = row3.createCell(i3);
	                cell3.setCellStyle(styleCaptionsRow);
	             }
	             
	             //
	             //Rows for Pitt Sheet
	            rowNum3++;
	      		row3 = m_Pitt.createRow(rowNum3);
	      		//row.setHeightInPoints((3*m_South.getDefaultRowHeightInPoints()));
	      		row3.setRowStyle(styleCaptionsRow);
	      		col3 = 0;
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Item #"));
	      		m_Pitt.setColumnWidth(col3++, CW_ItemNum);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Item Desc"));
	      		m_Pitt.setColumnWidth(col3++, CW_ItemDesc);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Vendor #"));
	      		m_Pitt.setColumnWidth(col3++, CW_VendorNum);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Vendor Name"));
	      		m_Pitt.setColumnWidth(col3++, CW_Vendor);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Dept #"));
	      		m_Pitt.setColumnWidth(col3++, CW_DeptNum);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Never Out"));
	      		m_Pitt.setColumnWidth(col3++, CW_NeverOut);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Total OH"));
	      		m_Pitt.setColumnWidth(col3++, CW_TotalOH);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Whs OH"));
	      		m_Pitt.setColumnWidth(col3++, CW_WareOH);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Trailer OH"));
	      		m_Pitt.setColumnWidth(col3++, CW_TrailerOH);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Fore River OH"));
	      		m_Pitt.setColumnWidth(col3++, CW_foreR);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Pittston OH"));
	      		m_Pitt.setColumnWidth(col3++, CW_PittS);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("SO Port OH"));
	      		m_Pitt.setColumnWidth(col3++, CW_southP);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("28 Day Sales"));
	      		m_Pitt.setColumnWidth(col3++, CW_28Day);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("LY 28 Day Sales"));
	      		m_Pitt.setColumnWidth(col3++, CW_LY28Day);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Cur WOS"));
	      		m_Pitt.setColumnWidth(col3++, CW_CurrWOS);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("LY WOS"));
	      		m_Pitt.setColumnWidth(col3++, CW_LastWOS);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Future Orders"));
	      		m_Pitt.setColumnWidth(col3++, CW_FutureOrders);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Current Orders"));
	      		m_Pitt.setColumnWidth(col3++, CW_CurrentOrders);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Average Unit Cost"));
	      		m_Pitt.setColumnWidth(col3++, CW_AvUnitCost);
	      		row3.getCell(col3).setCellValue(new XSSFRichTextString("Total Cost"));
	      		m_Pitt.setColumnWidth(col3++, CW_TotValue);
	       }
	      // return rowNum3;
	  } 
	   finally {
	      fontCaptionsRow = null;         
	      styleCaptionsRow = null;
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
		    m_EdbConn = m_RptProc.getEdbConn();
		         
		      if ( prepareStatements() )
		    	  if ( prepareStatements2() )
		    		  if ( prepareStatements3() )
		         created = buildSpreadsheet();
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

		   
	   	
	 private XSSFRow createRow(XSSFSheet m_South, int rowNum)
	 {
		  ;
		  int colCnt = 0;
		  XSSFRow row = null;
	      XSSFCell cell = null;
	      
	      if ( m_South == null )
	         return row;

	      row = m_South.createRow(rowNum);

	      //
	      // set the type and style of the cell For South Sheet.
	      if ( row != null ) {
	         for ( int i = 0; i < colCnt; i++ ) {            
	            cell = row.createCell(i);
	            cell.setCellStyle(m_CellStyles[i]);
	         }
	      }
	      return row;
	 }
	 
	 private XSSFRow createRow2(XSSFSheet m_Fore, int rowNum2)
	 {
	
		  int colCnt2 = 0;
		  XSSFRow row2 = null;
	      XSSFCell cell2 = null;
	      
	      if ( m_Fore == null )
		         return row2;

		  row2 = m_Fore.createRow(rowNum2);

		  //
		  // set the type and style of the cell For Fore Sheet.
		  if ( row2 != null ) {
		      for ( int i2 = 0; i2 < colCnt2; i2++ ) {            
		         cell2 = row2.createCell(i2);
		         cell2.setCellStyle(m_CellStyles[i2]);
		     }
		 }
		 return row2;
	 }
	 
	 private XSSFRow createRow3(XSSFSheet m_Pitt, int rowNum3)
	 {
		  int colCnt3 = 0;
		  XSSFRow row3 = null;
	      XSSFCell cell3 = null;
	      
	      if ( m_Pitt == null )
			     return row3;

		  row3 = m_Pitt.createRow(rowNum3);

		  //
		  // set the type and style of the cell For Pitt Sheet.
		  if ( row3 != null ) {
		     for ( int i3 = 0; i3 < colCnt3; i3++ ) {            
		        cell3 = row3.createCell(i3);
		        cell3.setCellStyle(m_CellStyles[i3]);
		     }
		  }
	      return row3;
	 }
	 
	 private boolean prepareStatements()
	   {
	      StringBuffer sql = new StringBuffer();
	      boolean isPrepared = false;

	      if ( m_EdbConn != null ) {
	         try {
	            setCurAction("Preparing Statements");
	            
	            sql.append("select item_entity_attr.item_id, item_entity_attr.description, "); 
	            sql.append("	vendor.vendor_id, vendor.name, emery_dept.dept_num, ");
	            sql.append("	decode(ejd_item_warehouse.never_out, 1, 'Y', 'N') as Never_Out,");
	            sql.append("	oh.oh as Total_OH, whs_oh.oh as Whs_OH, ");
	            sql.append("	trl_oh.oh as Trailer_OH, fr_oh.oh as Fore_R_OH, pit_oh.oh as Pitt_OH, sp_oh.oh as South_P_OH, ");
	            sql.append("	sales_12.sales as \"28_Day\", sales_12_prior.sales as LY_28_Day, ");
	            sql.append("	round(decode(sales_12.sales, null, null, 0, null, whs_oh.oh / sales_12.sales * 4),1) as Cur_WOS, ");
	            sql.append("	round(decode(sales_12_prior.sales, null, null, 0, null, whs_oh.oh / sales_12_prior.sales * 4),1) as LY_WOS, ");
	            sql.append("	future_orders.orders as Future_Orders, ");
	            sql.append("	current_orders.orders as Current_Orders, ");
	            sql.append("	round((iciloc.totalcost / decode(iciloc.qtyonhand, 0, decode(iciloc.lastcost, 0, ejd_item_price.buy, iciloc.lastcost), iciloc.qtyonhand))::numeric,3) as Avg_Unit_Cost, ");
	            sql.append("	oh.oh * round((iciloc.totalcost / decode(iciloc.qtyonhand, 0, decode(iciloc.lastcost, 0, ejd_item_price.buy, iciloc.lastcost), iciloc.qtyonhand))::numeric,3) as Total_Value ");
	            sql.append("from item_entity_attr ");
	            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
	            sql.append("join ejd.sage300_iciloc_mv iciloc on iciloc.itemno = item_entity_attr.item_id and ");
	            sql.append("	iciloc.location = '01' ");
	            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ");
	            sql.append("	ejd_item_warehouse.warehouse_id = 1 ");
	            sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id and ejd_item_price.warehouse_id = ejd_item_warehouse.warehouse_id ");
	            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
	            sql.append("join emery_dept on emery_dept.dept_id = ejd_item.dept_id ");
	            sql.append("join ( ");
	            sql.append("select \"sku\" item_id, sum(\"actual_qty\") as oh ");
	            sql.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql.append("group by \"sku\"  ");
	            sql.append("	) oh on oh.item_id = item_entity_attr.item_id ");
	            sql.append("join ( ");
	            sql.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	           	sql.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql.append("and \"loc_id\" not like 'TRL%' and ");
	            sql.append("	\"loc_id\" not like 'PIT%' and ");
	            sql.append("	\"loc_id\" not like 'FR%' and ");
	            sql.append("	\"loc_id\" not like 'SP%' ");
	            sql.append("group by \"sku\" ");
	            sql.append("	) whs_oh on whs_oh.item_id = item_entity_attr.item_id ");
	            sql.append("left outer join ( ");
	            sql.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql.append("and \"loc_id\" like 'TRL%' ");
	            sql.append("group by \"sku\" ");
	            sql.append("	) trl_oh on trl_oh.item_id = item_entity_attr.item_id ");
	            sql.append("left outer join ( ");
	            sql.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql.append("and \"loc_id\" like 'FR%' ");
	            sql.append("group by \"sku\" ");
	            sql.append("	) fr_oh on fr_oh.item_id = item_entity_attr.item_id ");
	            sql.append("left outer join ( ");
	            sql.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql.append("and \"loc_id\" like 'PIT%' ");
	            sql.append("group by \"sku\" ");
	            sql.append("	) pit_oh on pit_oh.item_id = item_entity_attr.item_id ");
	            sql.append("left outer join ( ");
	            sql.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql.append("and \"loc_id\" like 'SP%' ");
	            sql.append("group by \"sku\" ");
	            sql.append("	) sp_oh on sp_oh.item_id = item_entity_attr.item_id ");
	            sql.append("left outer join ( ");
	            sql.append("select item_nbr, sum(qty_shipped) as sales ");
	            sql.append("from itemsales ");
	            sql.append("where invoice_date >= trunc(now()) - 29 and ");
	            sql.append("	invoice_date <= trunc(now()) - 1 ");
	            sql.append("group by item_nbr ");
	            sql.append(") sales_12 on sales_12.item_nbr = item_entity_attr.item_id ");
	            sql.append("left outer join ( ");
	            sql.append("select item_nbr, sum(qty_shipped) as sales ");
	            sql.append("from itemsales ");
	            sql.append("where invoice_date >= add_months(trunc(now()) - 29, -12) and ");
	            sql.append("	invoice_date <= add_months(trunc(now()) - 1, -12) ");
	            sql.append("group by item_nbr ");
	            sql.append("	) sales_12_prior on sales_12_prior.item_nbr = item_entity_attr.item_id ");
	            sql.append("left outer join ( ");
	            sql.append("select item_id, sum(qty_ordered) as orders ");
	            sql.append("from order_line ");
	            sql.append("join order_header on order_header.order_id = order_line.order_id ");
	            sql.append("where order_line.order_status_id = (select order_status_id from order_status where description = 'NEW') and ");
	            sql.append("	order_header.order_status_id in (select order_status_id from order_status where description in ('NEW','WAITING FOR INVENTORY','WAITING CREDIT APPROVAL')) and ");
	            sql.append("	order_line.earliest_ship is not null and ");
	            sql.append("	order_line.earliest_ship > now() + 7 ");
	            sql.append("group by item_id ");
	            sql.append("	) future_orders on future_orders.item_id = item_entity_attr.item_id ");
	            sql.append("left outer join ( ");
	            sql.append("select item_id, sum(qty_ordered) as orders ");
	            sql.append("from order_line ");
	            sql.append("join order_header on order_header.order_id = order_line.order_id ");
	            sql.append("where order_line.order_status_id = (select order_status_id from order_status where description = 'NEW') and ");
	            sql.append("	order_header.order_status_id in (select order_status_id from order_status where description in ('NEW','WAITING FOR INVENTORY','WAITING CREDIT APPROVAL')) and ");
	            sql.append("	(order_line.earliest_ship is null or order_line.earliest_ship < now() + 7) ");
	            sql.append("group by item_id ");
	            sql.append("	) current_orders on current_orders.item_id = item_entity_attr.item_id ");
	            sql.append("where item_entity_attr.item_id in ( ");
	            sql.append("select distinct \"sku\" ");
	            sql.append("from loc_allocation ");
	            sql.append("where \"sku\" in ( ");
	            sql.append("select distinct \"sku\" ");
	            sql.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql.append("and loc_allocation.\"loc_id\" like 'TRL%' or ");
	            sql.append("	loc_allocation.\"loc_id\" like 'FR%' or ");
	            sql.append("	loc_allocation.\"loc_id\" like 'PIT%' or ");
	            sql.append("	loc_allocation.\"loc_id\" like 'SP%' ");
	            sql.append(") ");
	            sql.append("and loc_allocation.warehouse = 'PORTLAND' ");
	            sql.append(") ");
	            sql.append("and whs_oh.oh is not null ");
	            sql.append("and item_entity_attr.item_type_id in (select item_type_id from item_type where itemtype not in ('ACE', 'EXPANDED ASST')) ");
	            sql.append("order by round(decode(sales_12.sales, null, null, 0, null, whs_oh.oh / sales_12.sales * 4),1) ");
	            m_southPort = m_EdbConn.prepareStatement(sql.toString());
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
	      private boolean prepareStatements2()
		   {	      
	      StringBuffer sql2 = new StringBuffer();
	      boolean isPrepared2 = false;

	      if ( m_EdbConn != null ) {
	         try {
	            setCurAction("Preparing Statements");
	            
	            sql2.append("select item_entity_attr.item_id, item_entity_attr.description, "); 
	            sql2.append("	vendor.vendor_id, vendor.name, emery_dept.dept_num, ");
	            sql2.append("	decode(ejd_item_warehouse.never_out, 1, 'Y', 'N') as Never_Out,");
	            sql2.append("	oh.oh as Total_OH, whs_oh.oh as Whs_OH, ");
	            sql2.append("	trl_oh.oh as Trailer_OH, fr_oh.oh as Fore_R_OH, pit_oh.oh as Pitt_OH, sp_oh.oh as South_P_OH, ");
	            sql2.append("	sales_12.sales as \"28_Day\", sales_12_prior.sales as LY_28_Day, ");
	            sql2.append("	round(decode(sales_12.sales, null, null, 0, null, whs_oh.oh / sales_12.sales * 4),1) as Cur_WOS, ");
	            sql2.append("	round(decode(sales_12_prior.sales, null, null, 0, null, whs_oh.oh / sales_12_prior.sales * 4),1) as LY_WOS, ");
	            sql2.append("	future_orders.orders as Future_Orders, ");
	            sql2.append("	current_orders.orders as Current_Orders, ");
	            sql2.append("	round((iciloc.totalcost / decode(iciloc.qtyonhand, 0, decode(iciloc.lastcost, 0, ejd_item_price.buy, iciloc.lastcost), iciloc.qtyonhand))::numeric,3) as Avg_Unit_Cost, ");
	            sql2.append("	oh.oh * round((iciloc.totalcost / decode(iciloc.qtyonhand, 0, decode(iciloc.lastcost, 0, ejd_item_price.buy, iciloc.lastcost), iciloc.qtyonhand))::numeric,3) as Total_Value ");
	            sql2.append("from item_entity_attr ");
	            sql2.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
	            sql2.append("join ejd.sage300_iciloc_mv iciloc on iciloc.itemno = item_entity_attr.item_id and ");
	            sql2.append("	iciloc.location = '01' ");
	            sql2.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ");
	            sql2.append("	ejd_item_warehouse.warehouse_id = 1 ");
	            sql2.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id and ejd_item_price.warehouse_id = ejd_item_warehouse.warehouse_id ");
	            sql2.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
	            sql2.append("join emery_dept on emery_dept.dept_id = ejd_item.dept_id ");
	            sql2.append("join ( ");
	            sql2.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql2.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql2.append("group by \"sku\"  ");
	            sql2.append("	) oh on oh.item_id = item_entity_attr.item_id ");
	            sql2.append("join ( ");
	            sql2.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	           	sql2.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql2.append("and \"loc_id\" not like 'TRL%' and ");
	            sql2.append("	\"loc_id\" not like 'PIT%' and ");
	            sql2.append("	\"loc_id\" not like 'FR%' and ");
	            sql2.append("	\"loc_id\" not like 'SP%' ");
	            sql2.append("group by \"sku\" ");
	            sql2.append("	) whs_oh on whs_oh.item_id = item_entity_attr.item_id ");
	            sql2.append("left outer join ( ");
	            sql2.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql2.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql2.append("and \"loc_id\" like 'TRL%' ");
	            sql2.append("group by \"sku\" ");
	            sql2.append("	) trl_oh on trl_oh.item_id = item_entity_attr.item_id ");
	            sql2.append("left outer join ( ");
	            sql2.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql2.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql2.append("and \"loc_id\" like 'FR%' ");
	            sql2.append("group by \"sku\" ");
	            sql2.append("	) fr_oh on fr_oh.item_id = item_entity_attr.item_id ");
	            sql2.append("left outer join ( ");
	            sql2.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql2.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql2.append("and \"loc_id\" like 'PIT%' ");
	            sql2.append("group by \"sku\" ");
	            sql2.append("	) pit_oh on pit_oh.item_id = item_entity_attr.item_id ");
	            sql2.append("left outer join ( ");
	            sql2.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql2.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql2.append("and \"loc_id\" like 'SP%' ");
	            sql2.append("group by \"sku\" ");
	            sql2.append("	) sp_oh on sp_oh.item_id = item_entity_attr.item_id ");
	            sql2.append("left outer join ( ");
	            sql2.append("select item_nbr, sum(qty_shipped) as sales ");
	            sql2.append("from itemsales ");
	            sql2.append("where invoice_date >= trunc(now()) - 29 and ");
	            sql2.append("	invoice_date <= trunc(now()) - 1 ");
	            sql2.append("group by item_nbr ");
	            sql2.append(") sales_12 on sales_12.item_nbr = item_entity_attr.item_id ");
	            sql2.append("left outer join ( ");
	            sql2.append("select item_nbr, sum(qty_shipped) as sales ");
	            sql2.append("from itemsales ");
	            sql2.append("where invoice_date >= add_months(trunc(now()) - 29, -12) and ");
	            sql2.append("	invoice_date <= add_months(trunc(now()) - 1, -12) ");
	            sql2.append("group by item_nbr ");
	            sql2.append("	) sales_12_prior on sales_12_prior.item_nbr = item_entity_attr.item_id ");
	            sql2.append("left outer join ( ");
	            sql2.append("select item_id, sum(qty_ordered) as orders ");
	            sql2.append("from order_line ");
	            sql2.append("join order_header on order_header.order_id = order_line.order_id ");
	            sql2.append("where order_line.order_status_id = (select order_status_id from order_status where description = 'NEW') and ");
	            sql2.append("	order_header.order_status_id in (select order_status_id from order_status where description in ('NEW','WAITING FOR INVENTORY','WAITING CREDIT APPROVAL')) and ");
	            sql2.append("	order_line.earliest_ship is not null and ");
	            sql2.append("	order_line.earliest_ship > now() + 7 ");
	            sql2.append("group by item_id ");
	            sql2.append("	) future_orders on future_orders.item_id = item_entity_attr.item_id ");
	            sql2.append("left outer join ( ");
	            sql2.append("select item_id, sum(qty_ordered) as orders ");
	            sql2.append("from order_line ");
	            sql2.append("join order_header on order_header.order_id = order_line.order_id ");
	            sql2.append("where order_line.order_status_id = (select order_status_id from order_status where description = 'NEW') and ");
	            sql2.append("	order_header.order_status_id in (select order_status_id from order_status where description in ('NEW','WAITING FOR INVENTORY','WAITING CREDIT APPROVAL')) and ");
	            sql2.append("	(order_line.earliest_ship is null or order_line.earliest_ship < now() + 7) ");
	            sql2.append("group by item_id ");
	            sql2.append("	) current_orders on current_orders.item_id = item_entity_attr.item_id ");
	            sql2.append("where item_entity_attr.item_id in ( ");
	            sql2.append("select distinct \"sku\" ");
	            sql2.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql2.append("and \"sku\" in ( ");
	            sql2.append("select distinct \"sku\" ");
	            sql2.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql2.append("and loc_allocation.\"loc_id\" like 'TRL%' or ");
	            sql2.append("	loc_allocation.\"loc_id\" like 'FR%' or ");
	            sql2.append("	loc_allocation.\"loc_id\" like 'PIT%' or ");
	            sql2.append("	loc_allocation.\"loc_id\" like 'SP%' ");
	            sql2.append(") ");
	            sql2.append(") ");
	            sql2.append("and fr_oh.oh is not null ");
	            sql2.append("and item_entity_attr.item_type_id in (select item_type_id from item_type where itemtype not in ('ACE', 'EXPANDED ASST')) ");
	            sql2.append("order by round(decode(sales_12.sales, null, null, 0, null, whs_oh.oh / sales_12.sales * 4),1) ");
	            m_foreRiver = m_EdbConn.prepareStatement(sql2.toString());
	            isPrepared2 = true;
	         }
	         
	         catch ( SQLException ex ) {
	            log.error("exception:", ex);
	         }
	         
	         finally {
	            sql2 = null;
	         }         
	      }
	      else
	         log.error("custprofitabilty.prepareStatements - null enterprisedb or fascor connection");
	      
	      return isPrepared2;
	   
	 }
	      
	 private boolean prepareStatements3()
     {	
	      StringBuffer sql3 = new StringBuffer();
	      boolean isPrepared3 = false;

	      if ( m_EdbConn != null ) {
	         try {
	            setCurAction("Preparing Statements");
	            
	            sql3.append("select item_entity_attr.item_id, item_entity_attr.description, "); 
	            sql3.append("	vendor.vendor_id, vendor.name, emery_dept.dept_num, ");
	            sql3.append("	decode(ejd_item_warehouse.never_out, 1, 'Y', 'N') as Never_Out,");
	            sql3.append("	oh.oh as Total_OH, whs_oh.oh as Whs_OH, ");
	            sql3.append("	trl_oh.oh as Trailer_OH, fr_oh.oh as Fore_R_OH, pit_oh.oh as Pitt_OH, sp_oh.oh as South_P_OH, ");
	            sql3.append("	sales_12.sales as \"28_Day\", sales_12_prior.sales as LY_28_Day, ");
	            sql3.append("	round(decode(sales_12.sales, null, null, 0, null, whs_oh.oh / sales_12.sales * 4),1) as Cur_WOS, ");
	            sql3.append("	round(decode(sales_12_prior.sales, null, null, 0, null, whs_oh.oh / sales_12_prior.sales * 4),1) as LY_WOS, ");
	            sql3.append("	future_orders.orders as Future_Orders, ");
	            sql3.append("	current_orders.orders as Current_Orders, ");
	            sql3.append("	round((iciloc.totalcost / decode(iciloc.qtyonhand, 0, decode(iciloc.lastcost, 0, ejd_item_price.buy, iciloc.lastcost), iciloc.qtyonhand))::numeric,3) as Avg_Unit_Cost, ");
	            sql3.append("	oh.oh * round((iciloc.totalcost / decode(iciloc.qtyonhand, 0, decode(iciloc.lastcost, 0, ejd_item_price.buy, iciloc.lastcost), iciloc.qtyonhand))::numeric,3) as Total_Value ");
	            sql3.append("from item_entity_attr ");
	            sql3.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
	            sql3.append("join ejd.sage300_iciloc_mv iciloc on iciloc.itemno = item_entity_attr.item_id and ");
	            sql3.append("	iciloc.location = '01' ");
	            sql3.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ");
	            sql3.append("	ejd_item_warehouse.warehouse_id = 1 ");
	            sql3.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id and ejd_item_price.warehouse_id = ejd_item_warehouse.warehouse_id ");
	            sql3.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
	            sql3.append("join emery_dept on emery_dept.dept_id = ejd_item.dept_id ");
	            sql3.append("join ( ");
	            sql3.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql3.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql3.append("group by \"sku\"  ");
	            sql3.append("	) oh on oh.item_id = item_entity_attr.item_id ");
	            sql3.append("join ( ");
	            sql3.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	           	sql3.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql3.append("and \"loc_id\" not like 'TRL%' and ");
	            sql3.append("	\"loc_id\" not like 'PIT%' and ");
	            sql3.append("	\"loc_id\" not like 'FR%' and ");
	            sql3.append("	\"loc_id\" not like 'SP%' ");
	            sql3.append("group by \"sku\" ");
	            sql3.append("	) whs_oh on whs_oh.item_id = item_entity_attr.item_id ");
	            sql3.append("left outer join ( ");
	            sql3.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql3.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql3.append("and \"loc_id\" like 'TRL%' ");
	            sql3.append("group by \"sku\" ");
	            sql3.append("	) trl_oh on trl_oh.item_id = item_entity_attr.item_id ");
	            sql3.append("left outer join ( ");
	            sql3.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql3.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql3.append("and \"loc_id\" like 'FR%' ");
	            sql3.append("group by \"sku\" ");
	            sql3.append("	) fr_oh on fr_oh.item_id = item_entity_attr.item_id ");
	            sql3.append("left outer join ( ");
	            sql3.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql3.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql3.append("and \"loc_id\" like 'PIT%' ");
	            sql3.append("group by \"sku\" ");
	            sql3.append("	) pit_oh on pit_oh.item_id = item_entity_attr.item_id ");
	            sql3.append("left outer join ( ");
	            sql3.append("select \"sku\" as item_id, sum(\"actual_qty\") as oh ");
	            sql3.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql3.append("and \"loc_id\" like 'SP%' ");
	            sql3.append("group by \"sku\" ");
	            sql3.append("	) sp_oh on sp_oh.item_id = item_entity_attr.item_id ");
	            sql3.append("left outer join ( ");
	            sql3.append("select item_nbr, sum(qty_shipped) as sales ");
	            sql3.append("from itemsales ");
	            sql3.append("where invoice_date >= trunc(now()) - 29 and ");
	            sql3.append("	invoice_date <= trunc(now()) - 1 ");
	            sql3.append("group by item_nbr ");
	            sql3.append(") sales_12 on sales_12.item_nbr = item_entity_attr.item_id ");
	            sql3.append("left outer join ( ");
	            sql3.append("select item_nbr, sum(qty_shipped) as sales ");
	            sql3.append("from itemsales ");
	            sql3.append("where invoice_date >= add_months(trunc(now()) - 29, -12) and ");
	            sql3.append("	invoice_date <= add_months(trunc(now()) - 1, -12) ");
	            sql3.append("group by item_nbr ");
	            sql3.append("	) sales_12_prior on sales_12_prior.item_nbr = item_entity_attr.item_id ");
	            sql3.append("left outer join ( ");
	            sql3.append("select item_id, sum(qty_ordered) as orders ");
	            sql3.append("from order_line ");
	            sql3.append("join order_header on order_header.order_id = order_line.order_id ");
	            sql3.append("where order_line.order_status_id = (select order_status_id from order_status where description = 'NEW') and ");
	            sql3.append("	order_header.order_status_id in (select order_status_id from order_status where description in ('NEW','WAITING FOR INVENTORY','WAITING CREDIT APPROVAL')) and ");
	            sql3.append("	order_line.earliest_ship is not null and ");
	            sql3.append("	order_line.earliest_ship > now() + 7 ");
	            sql3.append("group by item_id ");
	            sql3.append("	) future_orders on future_orders.item_id = item_entity_attr.item_id ");
	            sql3.append("left outer join ( ");
	            sql3.append("select item_id, sum(qty_ordered) as orders ");
	            sql3.append("from order_line ");
	            sql3.append("join order_header on order_header.order_id = order_line.order_id ");
	            sql3.append("where order_line.order_status_id = (select order_status_id from order_status where description = 'NEW') and ");
	            sql3.append("	order_header.order_status_id in (select order_status_id from order_status where description in ('NEW','WAITING FOR INVENTORY','WAITING CREDIT APPROVAL')) and ");
	            sql3.append("	(order_line.earliest_ship is null or order_line.earliest_ship < now() + 7) ");
	            sql3.append("group by item_id ");
	            sql3.append("	) current_orders on current_orders.item_id = item_entity_attr.item_id ");
	            sql3.append("where item_entity_attr.item_id in ( ");
	            sql3.append("select distinct \"sku\" ");
	            sql3.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql3.append("and \"sku\" in ( ");
	            sql3.append("select distinct \"sku\" ");
	            sql3.append("from loc_allocation where warehouse = 'PORTLAND' ");
	            sql3.append("and loc_allocation.\"loc_id\" like 'TRL%' or ");
	            sql3.append("	loc_allocation.\"loc_id\" like 'FR%' or ");
	            sql3.append("	loc_allocation.\"loc_id\" like 'PIT%' or ");
	            sql3.append("	loc_allocation.\"loc_id\" like 'SP%' ");
	            sql3.append(") ");
	            sql3.append(") ");
	            sql3.append("and pit_oh.oh is not null ");
	            sql3.append("and item_entity_attr.item_type_id in (select item_type_id from item_type where itemtype not in ('ACE', 'EXPANDED ASST')) ");
	            sql3.append("order by round(decode(sales_12.sales, null, null, 0, null, whs_oh.oh / sales_12.sales * 4),1) ");
	            m_pittston = m_EdbConn.prepareStatement(sql3.toString());
	            isPrepared3 = true;
	         }
	         
	         catch ( SQLException ex ) {
	            log.error("exception:", ex);
	         }
	         
	         finally {
	            sql3 = null;
	         }         
	      }
	      else
	         log.error("custprofitabilty.prepareStatements - null enterprisedb or fascor connection");
	      
	      return isPrepared3;
     }
   
	 public void setParams(ArrayList<Param> params)
     {
	      StringBuffer fileName = new StringBuffer();      
	      String tmp = Long.toString(System.currentTimeMillis());
	      	      
	      fileName.append("off_site_inventory");      
	      fileName.append("-");
	      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
	      fileName.append(".xlsx");
	      m_FileNames.add(fileName.toString());
	 }
	 
	 private void setupWorkbook()
	 {	      
	      XSSFCellStyle styleTxtC = null;       // Text centered
	      XSSFCellStyle styleTxtL = null;       // Text left justified
	      XSSFCellStyle styleInt = null;        // 0 decimals and a comma
	      XSSFCellStyle csIntRed = null;        // 0 decimals and a comma, red foreground
	      XSSFCellStyle styleDouble = null;     // numeric 1 decimal and a comma
	      XSSFCellStyle styleFloat = null;		//numeric 1 decimal 
	      XSSFCellStyle stylePercent = null;    // Integer percent with comma
	      XSSFCellStyle csPercentRed = null;	  // Integer percent with comma red
	      XSSFCellStyle csPercentOrange = null; // Integer percent with comma orange
	      XSSFCellStyle csPercentGold = null;	  // Integer percent with comma gold
	      XSSFCellStyle csPercentYellow = null; // Integer percent with comma yellow
	      XSSFDataFormat format = null;
	      XSSFFont font = null;
	            
	      format = m_WrkBk.createDataFormat();
	      
	      font = m_WrkBk.createFont();
	      font.setFontHeightInPoints((short)8);
	      font.setFontName("Arial");
	            
	      styleTxtL = m_WrkBk.createCellStyle();
	      styleTxtL.setAlignment(HorizontalAlignment.LEFT);
	      styleTxtL.setFont(font);
	      
	      styleTxtC = m_WrkBk.createCellStyle();
	      styleTxtC.setAlignment(HorizontalAlignment.CENTER);
	      styleTxtC.setFont(font);
	      
	      styleInt = m_WrkBk.createCellStyle();
	      styleInt.setAlignment(HorizontalAlignment.RIGHT);
	      styleInt.setFont(font);
	      styleInt.setDataFormat(format.getFormat("_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)"));
	      
	      styleDouble = m_WrkBk.createCellStyle();
	      styleDouble.setAlignment(HorizontalAlignment.RIGHT);
	      styleDouble.setFont(font);
	      styleDouble.setDataFormat(format.getFormat("_(* #,##0.0_);_(* (#,##0.0);_(* \"-\"??_);_(@_)"));
	      
	      styleFloat = m_WrkBk.createCellStyle();
	      styleFloat.setAlignment(HorizontalAlignment.RIGHT);
	      styleFloat.setFont(font);
	      styleFloat.setDataFormat(format.getFormat("_(* #####0.000_);_(* (#####0.000);_(* \"-\"??_);_(@_)"));
	      
	      stylePercent = m_WrkBk.createCellStyle();
	      stylePercent.setAlignment(HorizontalAlignment.RIGHT);
	      stylePercent.setFont(font);
	      stylePercent.setDataFormat(format.getFormat("#,##0%"));

	      //
	      // These are used in the conditional formatting.
	      csIntRed = m_WrkBk.createCellStyle();
	      csIntRed.setAlignment(HorizontalAlignment.RIGHT);
	      csIntRed.setFont(font);
	      csIntRed.setDataFormat(format.getFormat("_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)"));
	      csIntRed.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	      csIntRed.setFillForegroundColor(IndexedColors.RED.getIndex());
	      
	      csPercentRed = m_WrkBk.createCellStyle();
	      csPercentRed.setAlignment(HorizontalAlignment.RIGHT);
	      csPercentRed.setFont(font);
	      csPercentRed.setDataFormat(format.getFormat("#,##0%"));
	      csPercentRed.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	      csPercentRed.setFillForegroundColor(IndexedColors.RED.getIndex());
	            
	      csPercentOrange = m_WrkBk.createCellStyle();
	      csPercentOrange.setAlignment(HorizontalAlignment.RIGHT);
	      csPercentOrange.setFont(font);
	      csPercentOrange.setDataFormat(format.getFormat("#,##0%"));
	      csPercentOrange.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	      csPercentOrange.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
	      
	      csPercentYellow = m_WrkBk.createCellStyle();
	      csPercentYellow.setAlignment(HorizontalAlignment.RIGHT);
	      csPercentYellow.setFont(font);
	      csPercentYellow.setDataFormat(format.getFormat("#,##0%"));
	      csPercentYellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	      csPercentYellow.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
	      
	      csPercentGold = m_WrkBk.createCellStyle();
	      csPercentGold.setAlignment(HorizontalAlignment.RIGHT);
	      csPercentGold.setFont(font);
	      csPercentGold.setDataFormat(format.getFormat("#,##0%"));
	      csPercentGold.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	      csPercentGold.setFillForegroundColor(IndexedColors.GOLD.getIndex());
	            
	      m_CellStyles = new XSSFCellStyle[] {
	    	 styleTxtL,    // col 0 warehouse
	    	 styleTxtL,    // col 1 dept
	    	 styleInt,    // col 2 vendor name
	         styleTxtL,    // col 3 buyer
	         styleTxtL,    // col 4 item id        
	         styleTxtL,    // col 5 item desc
	         styleInt,     // col 6 qoh
	         styleInt,     // col 7 open pos
	         styleInt,	  // col 8 qoh + pos
	         styleInt,     // col 9 monthly sales
	         styleInt,     // col 10 open orders
	         styleInt,	  // col 11 4 week forecast
	         styleInt,	  // col 12 8 week forecast
	         styleInt, // col 13 qoh coverage
	         stylePercent, // col 14 qoh + 4 week forecast coverage
	         stylePercent, // col 16 qoh + 4 week forecast + open po coverage
	         stylePercent  // col 16 qoh + 8 week forecast + open po coverage
	      };
	      
	      m_CellStylesEx = new XSSFCellStyle[] {
	         csIntRed,        // special style #1
	         csPercentRed,    // special style #2
	         csPercentOrange, // special style #3
	         csPercentYellow, // special style #4
	         csPercentGold	  // special style #5
	      };
	      
	      styleFloat = null;
	      styleTxtC = null;
	      styleTxtL = null;
	      styleInt = null;
	      styleDouble = null;
	      csIntRed = null;
	      csPercentRed = null;
	      csPercentOrange = null;
	      csPercentYellow = null;	      
	      format = null;
	      font = null;   
	 }
}
