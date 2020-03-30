/**
 * Title:			BestBrandsPRAdvChgNotice.java
 * Description:     Best Brands Priced Right Advance Change Notice
 * Company:			Emery-Waterhouse
 * @author			Naresh Pasnur
 * 
 * Create Date: 06/25/2012
 * 
 * History:
 *    $Log: BestBrandsPRAdvChgNotice.java,v $
 *    Revision 1.3  2013/06/03 14:07:22  prichter
 *    Some cosmetic code changes
 *
 *    Revision 1.2  2013/05/07 13:29:54  prichter
 *    Code alignment cleanup
 *
 *    Revision 1.1  2012/07/30 20:49:40  npasnur
 *    initial commit
 *
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.GregorianCalendar;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.hssf.util.HSSFColor;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class BestBrandsPRAdvChgNotice extends Report {
	
	//
	// Parameters
	private String m_CustId;
	private Date m_RunDate;
	private Date m_PriceDate;
	private Date m_LastRunDate;
	private int m_Days = 730;  // number of days of history to report( 24 months )
	private ArrayList<String> m_StoreId; 
	
	//
	// Statements
	private PreparedStatement m_ReportData;
	private PreparedStatement m_ItemDelData;     //Get all detail information of items that have been taken of the list w.r.to last run report
	private PreparedStatement m_ItemDelList;     //Get all items that have been taken of the list w.r.to last run report
	private PreparedStatement m_PurchHist;       //units purchased of an item by the customer
	private PreparedStatement m_FamilyTree;      //finds all customers connected to a parent id
	
   //	
   // POI stuff
   HSSFWorkbook m_Wrkbk = new HSSFWorkbook();
   HSSFSheet m_Sheet = m_Wrkbk.createSheet();
   private HSSFFont m_Font;
   private HSSFFont m_FontSubtitle;
   private HSSFFont m_FontTitle;
   private HSSFFont m_FontBold;
   private HSSFFont m_FontData;
   private HSSFCellStyle m_StyleTitle;    // size 12 Bold centered
   private HSSFCellStyle m_StyleSubtitle; // size 10 left-alligned
   private HSSFCellStyle m_StyleBoldText; //size 8 bold left-alligned
   private HSSFCellStyle m_StyleBoldNbr;  // size 8 bold right-alligned
   private HSSFCellStyle m_StyleBoldCtr;  // size 8 bold centered
   private HSSFCellStyle m_StyleText;  	  // Text right justified
   private HSSFCellStyle m_StyleDec;      // Style with 2 decimals
   private HSSFCellStyle m_StyleTextCtr;  // Text style centered
   private HSSFCellStyle m_StyleInt;      // Style with 0 decimals
   private HSSFCellStyle m_StylePct;      // Style with 0 decimals + %
   private HSSFCellStyle m_ShadeStyleDec; // Style with 0 decimals + % with shading
   private HSSFCellStyle m_StyleLabel;    // Text labels, right justify, 8pt
   
   private int m_Row = 0;
   private int m_Col = 0;
	
   FileOutputStream m_OutFile;  
	
	/**
	 * Constructor
	 */
	public BestBrandsPRAdvChgNotice()
	{
		super();
		
	  m_StoreId = new ArrayList<String>();
		
      //
      // Create the default font for this workbook
      m_Font = m_Wrkbk.createFont();
      m_Font.setFontHeightInPoints((short) 8);
      m_Font.setFontName("Arial");

      //
      // Create a font that is normal size & bold
      m_FontBold = m_Wrkbk.createFont();
      m_FontBold.setFontHeightInPoints((short)8);
      m_FontBold.setFontName("Arial");
      m_FontBold.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

      //
      // Create a font that is normal size
      m_FontData = m_Wrkbk.createFont();
      m_FontData.setFontHeightInPoints((short)8);
      m_FontData.setFontName("Arial");
      
      //
      // Create a font for sub titles
      m_FontSubtitle = m_Wrkbk.createFont();
      m_FontSubtitle.setFontHeightInPoints((short)10);
      m_FontSubtitle.setFontName("Arial");
      m_FontSubtitle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);      

      //
      // Create a font for titles
      m_FontTitle = m_Wrkbk.createFont();
      m_FontTitle.setFontHeightInPoints((short)12);
      m_FontTitle.setFontName("Arial");
      m_FontTitle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

      //
      // Setup the cell styles used in this report
      m_StyleBoldCtr = m_Wrkbk.createCellStyle();
      m_StyleBoldCtr.setFont(m_FontBold);
      m_StyleBoldCtr.setAlignment(HSSFCellStyle.ALIGN_CENTER);
      m_StyleBoldCtr.setWrapText(true);

      m_StyleBoldText = m_Wrkbk.createCellStyle();
      m_StyleBoldText.setFont(m_FontBold);
      m_StyleBoldText.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      m_StyleBoldText.setWrapText(true);

      m_StyleDec = m_Wrkbk.createCellStyle();
      m_StyleDec.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      m_StyleDec.setFont(m_FontData);
      m_StyleDec.setDataFormat((short)8);
      
      //
      //Border
      m_StyleDec.setBorderTop(HSSFCellStyle.BORDER_THIN);
      m_StyleDec.setBorderBottom(HSSFCellStyle.BORDER_THIN);
      m_StyleDec.setBorderLeft(HSSFCellStyle.BORDER_THIN);
      m_StyleDec.setBorderRight(HSSFCellStyle.BORDER_THIN);

      m_StyleInt = m_Wrkbk.createCellStyle();
      m_StyleInt.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      m_StyleInt.setFont(m_FontData);
      m_StyleInt.setDataFormat((short)3);

      m_StyleLabel = m_Wrkbk.createCellStyle();
      m_StyleLabel.setFont(m_Font);
      m_StyleLabel.setAlignment(HSSFCellStyle.ALIGN_RIGHT);

      m_StylePct = m_Wrkbk.createCellStyle();
      m_StylePct.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      m_StylePct.setFont(m_FontData);
      m_StylePct.setDataFormat(HSSFDataFormat.getBuiltinFormat("0%"));
     
      //
      //Border
      m_StylePct.setBorderTop(HSSFCellStyle.BORDER_THIN);
      m_StylePct.setBorderBottom(HSSFCellStyle.BORDER_THIN);
      m_StylePct.setBorderLeft(HSSFCellStyle.BORDER_THIN);
      m_StylePct.setBorderRight(HSSFCellStyle.BORDER_THIN);
            
      m_ShadeStyleDec = m_Wrkbk.createCellStyle();
      m_ShadeStyleDec.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      m_ShadeStyleDec.setFont(m_FontData);
      m_ShadeStyleDec.setDataFormat((short)8);
        
      //
      //Shading
      m_ShadeStyleDec.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
      m_ShadeStyleDec.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
      
      //
      //Border
      m_ShadeStyleDec.setBorderTop(HSSFCellStyle.BORDER_THIN);
      m_ShadeStyleDec.setBorderBottom(HSSFCellStyle.BORDER_THIN);
      m_ShadeStyleDec.setBorderLeft(HSSFCellStyle.BORDER_THIN);
      m_ShadeStyleDec.setBorderRight(HSSFCellStyle.BORDER_THIN);
      
      m_StyleText = m_Wrkbk.createCellStyle();
      m_StyleText.setFont(m_FontData);
      m_StyleText.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      m_StyleText.setWrapText(true);
      
      //
      //Border
      m_StyleText.setBorderTop(HSSFCellStyle.BORDER_THIN);
      m_StyleText.setBorderBottom(HSSFCellStyle.BORDER_THIN);
      m_StyleText.setBorderLeft(HSSFCellStyle.BORDER_THIN);
      m_StyleText.setBorderRight(HSSFCellStyle.BORDER_THIN);

      m_StyleTextCtr = m_Wrkbk.createCellStyle();
      m_StyleTextCtr.setFont(m_FontData);
      m_StyleTextCtr.setAlignment(HSSFCellStyle.ALIGN_CENTER);
      m_StyleTextCtr.setWrapText(true);
      
      m_StyleSubtitle = m_Wrkbk.createCellStyle();
      m_StyleSubtitle.setFont(m_FontSubtitle);
      m_StyleSubtitle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      
      m_StyleTitle = m_Wrkbk.createCellStyle();
      m_StyleTitle.setFont(m_FontTitle);
      m_StyleTitle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
      m_StyleBoldNbr = m_Wrkbk.createCellStyle();
      m_StyleBoldNbr.setFont(m_FontBold);
      m_StyleBoldNbr.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
      m_StyleBoldNbr.setWrapText(true);
	}
	
	/**
	 * Clean up resources
	 */
	public void close() 
	{
    	DbUtils.closeDbConn(null, m_ReportData, null);
    	DbUtils.closeDbConn(null, m_ItemDelList, null);
    	DbUtils.closeDbConn(null, m_ItemDelData, null);
    	DbUtils.closeDbConn(null, m_PurchHist, null);
    	DbUtils.closeDbConn(null, m_FamilyTree, null);
    	DbUtils.closeDbConn(m_OraConn, null, null);
    	m_ReportData = null;
    	m_ItemDelList= null;
    	m_ItemDelData = null;
    	m_PurchHist = null;
    	m_FamilyTree = null;
		m_OraConn = null;		
		m_StoreId.clear();
	}
	
   /**
    * Convenience method that adds a new String type cell with no borders and the specified alignment.
    *
    * @param rowNum int - the row index.
    * @param colNum short - the column index.
    * @param val String - the cell value.
    *
    * @return HSSFCell - the newly added String type cell, or a reference to the existing one.
    */
   private HSSFCell addCell(int rowNum, int colNum, String val, HSSFCellStyle style)
   {
      HSSFCell cell = addCell(rowNum, colNum);

      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
      cell.setCellValue(new HSSFRichTextString(val));
      cell.setCellStyle(style);

      return cell;
   }

   /**
    * Convenience method that adds a new numeric type cell with no borders and the specified alignment.
    *
    * @param rowNum - the row index.
    * @param colNum short - the column index.
    * @param val double - the cell value.
    * @param style HSSFCellStyle - the cell style and format
    *
    * @return HSSFCell - the newly added numeric type cell, or a reference to the existing one.
    */
   private HSSFCell addCell(int rowNum, int colNum, double val, HSSFCellStyle style)
   {
      HSSFCell cell = addCell(rowNum, colNum);

      cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
      cell.setCellStyle(style);
      cell.setCellValue(val);

      return cell;
   }
  
   /**
    * Adds a new cell with the specified borders and horizontal alignment.
    *
    * @param rowNum - the row index.
    * @param colNum short - the column index.
    *
    * @return HSSFCell - the newly added cell, or a reference to the existing one.
    */
   private HSSFCell addCell(int rowNum, int colNum)
   {
      HSSFRow row = addRow(rowNum);
      HSSFCell cell = row.getCell(colNum);

      if ( cell == null )
         cell = row.createCell(colNum);

      row = null;

      return cell;
   }

   /**
    * Adds a new row or returns the existing one.
    *
    * @param rowNum int - the row index.
    * @return HSSFRow - the row object added, or a reference to the existing one.
    */
   private HSSFRow addRow(int rowNum)
   {
      HSSFRow row = m_Sheet.getRow(rowNum);

      if ( row == null )
         row = m_Sheet.createRow(rowNum);

      return row;
   }

	/**
	 * Creates the output file 
	 * @return boolean - true if successful
	 * @throws FileNotFoundException 
	 * @throws SQLException 
	 */
	public boolean buildOutputFile() throws FileNotFoundException, SQLException
	{
	  boolean built = false;		
      StringBuffer fileName = new StringBuffer(); 
      ResultSet rs = null;
      ResultSet rsStores = null;
      SimpleDateFormat fmt = new SimpleDateFormat("MM/dd/yyyy");
      String title = "BBPR Advanced Change Notice for " + fmt.format(m_RunDate);
      
      String custName = getCustName(m_CustId);
      String nbc = null;
      String itemId = null;
      String prevItmProg = "";
      String newItmProg = null;
      String usaItem = null;
      String topSeller = null;
      double curSell = 0.0;
      double curRetail = 0.0;
      double futSell = 0.0;
      double futRetail = 0.0;
      double curMargin = 0.0;
      double futMargin = 0.0;
      int retlPck = 1;
      String tmp = null;
      
      //
      // Build the report file name
      tmp = Long.toString(System.currentTimeMillis());
      fileName.append("BBPR Advanced Change Notice(");
      fileName.append(m_CustId);
      fileName.append(") ");
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(" (");
      fileName.append(new SimpleDateFormat("MM-dd-yyyy").format(m_RunDate));
      fileName.append(") ");
      fileName.append(".xls");
      m_FileNames.add(fileName.toString());
      m_OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      fileName = null;

      m_Sheet.getPrintSetup().setLandscape(true);
      m_Wrkbk.setRepeatingRowsAndColumns(0,-1,-1,0,4);
      m_Sheet.setMargin(HSSFSheet.LeftMargin, .25);
      m_Sheet.setMargin(HSSFSheet.RightMargin, .25);
      m_Sheet.setMargin(HSSFSheet.TopMargin,0.50);
      m_Sheet.setMargin(HSSFSheet.BottomMargin, .50);
      
      m_Sheet.getPrintSetup().setLandscape(true);
      m_Sheet.getHeader().setCenter(HeaderFooter.font("Arial", "Bold") + HeaderFooter.fontSize((short) 12) + title);
      m_Sheet.getHeader().setRight(HeaderFooter.font("Arial", "Bold") + HeaderFooter.fontSize((short) 12) + HeaderFooter.page() + " of " + HeaderFooter.numPages());
      
      m_Row = 0;
      
      m_RptProc.setEmailMsg(
      	"Attached is the Best Brands Priced Right Change Notice for " +
      	custName + "("+ m_CustId + ")" + " for the reporting period starting " +
      	new SimpleDateFormat("MM-dd-yyyy").format(m_RunDate)
      );
      
      addCell(m_Row, 1, "Account: " + m_CustId, m_StyleBoldText);
   	m_Row++;
   	addCell(m_Row, 1, custName, m_StyleBoldText);
      m_Row++;
   	m_Row++;
   	 
   	m_Col = 0;
   	m_Sheet.setColumnWidth(m_Col, 1500);
    
   	createRowCaptions(m_Row);  	
   	
   	m_Row = m_Row + 1;
   	  
   	m_Row++;
      
      try {
        //
   	    //find all stores related to this account and load them into the arraylist
        try {
           m_FamilyTree.setString(1, m_CustId);
           rsStores = m_FamilyTree.executeQuery();
           while ( rsStores.next() ) {
              m_StoreId.add( rsStores.getString("customer_id") );
           }
        }
        catch ( Exception e ) {
            log.error("BestBrandsPRAdvChgNotice: buildOutputFile() : Exception while trying to find all stores for Customer: "+m_CustId, e);
        }
        finally {
      	   DbUtils.closeDbConn(null, null, rsStores);
           rsStores = null;
        }
    	  
        m_ReportData.setString(1, m_CustId);
        m_ReportData.setDate(2, m_RunDate);
      	
        rs = m_ReportData.executeQuery();
      	
        while ( rs.next() ) {
      	  m_Col = 0;
      		      		 
  	        newItmProg = rs.getString("item_program");
  	        
  	        if(newItmProg != null && !newItmProg.equals(prevItmProg)){
  	           createItmProgRow(m_Row, m_Col,newItmProg);
  	           m_Row++; 
  	        }
  	        
  	        prevItmProg = newItmProg;
      		
      		itemId = rs.getString("item_id");
      		setCurAction("Processing customer " + m_CustId + " item " + itemId);
      		addCell(m_Row, m_Col++, rs.getString("department"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("vendor_name"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("mfr_nbr"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("upc_code"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("item_descr"), m_StyleText);
      		addCell(m_Row, m_Col++, rs.getString("unit"), m_StyleText);
      		
      		nbc = rs.getString("nbc"); 
            nbc =  nbc == null ? "" : nbc;
            addCell(m_Row, m_Col++, rs.getString("stock_pack"), m_StyleText);
            addCell(m_Row, m_Col++, nbc, m_StyleText);
       
            topSeller = rs.getString("tsi");
      		topSeller = (topSeller == null? "": "Y");
      		addCell(m_Row, m_Col++, topSeller, m_StyleText);
      	    
      		usaItem = rs.getString("usa_item");
            usaItem = (usaItem == null? "": "Y");
            addCell(m_Row, m_Col++, usaItem, m_StyleText);
      		
            addCell(m_Row, m_Col++, itemId, m_StyleText);
            
            curSell = rs.getDouble("current_sell");
            curRetail = rs.getDouble("current_retail");
            futSell = rs.getDouble("future_sell");
            futRetail = rs.getDouble("future_retail");
            
            //
            //if the item has retail pack > 1, then recalculate the cost and the margins based on retail pack.
            retlPck = rs.getInt("retail_pack");
            
            if(retlPck > 1){
               curSell = curSell/retlPck;
               futSell = futSell/retlPck;
            }
            
            curMargin = (curRetail - curSell) /curRetail;
            futMargin = (futRetail - futSell) /futRetail;
                      
            if(curSell != futSell){
               addCell(m_Row, m_Col++, curSell, m_ShadeStyleDec);
            }
            else{
               addCell(m_Row, m_Col++, curSell, m_StyleDec);
            }
                        
            if(curRetail != futRetail){
                addCell(m_Row, m_Col++, curRetail, m_ShadeStyleDec);
            }
            else{
                addCell(m_Row, m_Col++, curRetail, m_StyleDec);
            }
      		
            addCell(m_Row, m_Col++, Math.round(curMargin*100)+"%", m_StyleText);
                        
            if(curSell != futSell){
                addCell(m_Row, m_Col++, futSell, m_ShadeStyleDec);
            }
            else{
                addCell(m_Row, m_Col++, futSell, m_StyleDec);
            }
                         
            if(curRetail != futRetail){
               addCell(m_Row, m_Col++, futRetail, m_ShadeStyleDec);
            }
            else{
               addCell(m_Row, m_Col++, futRetail, m_StyleDec);
            }
            
            addCell(m_Row, m_Col++, Math.round(futMargin*100)+"%", m_StyleText);
            
            addCell(m_Row, m_Col++, Math.round((futMargin - curMargin)*100)+"%", m_StyleText);
              		
      		addCell(m_Row, m_Col++, unitsSold( m_CustId, itemId ), m_StyleText);
      		
      		if(rs.getInt("new_add") > 0){
          	   addCell(m_Row, m_Col++, "Y", m_StyleText);
            }
            else{
               addCell(m_Row, m_Col++, "", m_StyleText);
            }
      		
      		curSell = 0.0;
            curRetail = 0.0;
            futSell = 0.0;
            futRetail = 0.0;
            
      		m_Row++;
       	}
      	      	
      	m_Row++;
       
      	//
      	//Build items taken off the list from last run report.
      	m_Row = buildItemDelList(m_Row);
      	
      	m_Wrkbk.write(m_OutFile);
      	m_OutFile.close();
			
      	built = true;
      }
      
      catch ( Exception e ) {
      	log.error("exception", e);
      }
      
      finally {
      	m_Wrkbk = null;
      	m_OutFile = null;
      }

		return built;
	}
	
	//
	//Builds items list that have been taken off w.r.to last run report
	private int buildItemDelList(int rowNum)
	{
		ResultSet rsItemList = null;
		ResultSet rsItemData = null;
	    String prevItmProg = "";
	    String newItmProg = null;
	    String itemId = null;
	    String nbc = null;
	    String usaItem = null;
	    String topSeller = null;
	    int itemCnt = 0;
	    double curSell = 0.0;
	    double curRetail = 0.0;
	    double futSell = 0.0;
	    double futRetail = 0.0;
	    double curMargin = 0.0;
	    double futMargin = 0.0;
	    int retlPck = 1;
	    
		try {
			log.info("BBRP report deleted items list build: " + m_CustId + " " + m_LastRunDate.toString() + " " + m_RunDate.toString());
			m_ItemDelList.setString(1, m_CustId); //Use Emery test customer when the report is run for the first time.
			m_ItemDelList.setDate(2, m_LastRunDate);
			m_ItemDelList.setString(3, m_CustId);
			m_ItemDelList.setDate(4, m_RunDate);
	      	
			rsItemList = m_ItemDelList.executeQuery();
	      	
			while ( rsItemList.next() ) {
				m_Col = 0;
	      			  	        
				newItmProg = rsItemList.getString("item_program");
				itemId = rsItemList.getString("item_id");
				setCurAction("Processing customer " + m_CustId + " item " + itemId);
	      		
				try{
					//
					//Get detail information of the item
					m_ItemDelData.setString(1, m_CustId); ///Use Emery test customer when the report is run for the first time.
					m_ItemDelData.setString(2, itemId);
					m_ItemDelData.setString(3, newItmProg);
					m_ItemDelData.setDate(4, m_LastRunDate);
		      		
					rsItemData = m_ItemDelData.executeQuery();
	      		
					if( rsItemData.next() ) {
						if(itemCnt <= 0 ) {
							rowNum = rowNum + 1;
							//
	      			   //Items that have been taken off the list from last run report
	      			   rowNum = createItemDelListHdr(rowNum);
	      			   itemCnt = itemCnt + 1;
	      			}
	      			  
	      			//
						//Add the header only when at least one item is found.
						if(newItmProg != null && !newItmProg.equals(prevItmProg)){
							createItmProgRow(rowNum, m_Col,newItmProg);
							rowNum++; 
	 	  	          }
	 	  	          prevItmProg = newItmProg; 
	      			 	      			   
	 	  	          addCell(rowNum, m_Col++, rsItemData.getString("department"), m_StyleText);
	 	  	          addCell(rowNum, m_Col++, rsItemData.getString("vendor_name"), m_StyleText);
	 	  	          addCell(rowNum, m_Col++, rsItemData.getString("mfr_nbr"), m_StyleText);
	 	  	          addCell(rowNum, m_Col++, rsItemData.getString("upc_code"), m_StyleText);
	 	  	          addCell(rowNum, m_Col++, rsItemData.getString("item_descr"), m_StyleText);
	 	  	          addCell(rowNum, m_Col++, rsItemData.getString("unit"), m_StyleText);
	 	  	          addCell(rowNum, m_Col++, rsItemData.getString("stock_pack"), m_StyleText);
	      		
	 	  	          nbc = rsItemData.getString("nbc"); 
	 	  	          nbc =  nbc == null ? "" : nbc;
	 	  	          addCell(rowNum, m_Col++, nbc, m_StyleText);
	       
	 	  	          topSeller = rsItemData.getString("tsi");
	 	  	          topSeller = (topSeller == null? "": "Y");
	 	  	          addCell(rowNum, m_Col++, topSeller, m_StyleText);
	            	    
	 	  	          usaItem = rsItemData.getString("usa_item");
	 	  	          usaItem = (usaItem == null? "": "Y");
	 	  	          addCell(rowNum, m_Col++, usaItem, m_StyleText);
	                  
	 	  	          addCell(rowNum, m_Col++, itemId, m_StyleText);
	 	  	          
	 	  	          curSell = rsItemData.getDouble("current_sell");
	 	  	          curRetail = rsItemData.getDouble("current_retail");
	 	  	          futSell = rsItemData.getDouble("future_sell");
	 	  	          futRetail = rsItemData.getDouble("future_retail");
	                  
	 	  	          //
	 	  	          //if the item has retail pack > 1, then recalculate the cost and the margins based on retail pack.
	 	  	          retlPck = rsItemData.getInt("retail_pack");
	                  
	 	  	          if(retlPck > 1){
	 	  	         	 curSell = curSell/retlPck;
	 	  	         	 futSell = futSell/retlPck;
	 	  	          }
	                  
	 	  	          curMargin = (curRetail - curSell) /curRetail;
	 	  	          futMargin = (futRetail - futSell) /futRetail;
	                  
	 	  	          if(curSell != futSell){
	 	  	         	 addCell(rowNum, m_Col++, curSell, m_ShadeStyleDec);
	                }
	 	  	          else{
	 	  	         	 addCell(rowNum, m_Col++, curSell, m_StyleDec);
	                }
	                   
	                   
	 	  	          if(curRetail != futRetail){
	 	  	         	 addCell(rowNum, m_Col++, curRetail, m_ShadeStyleDec);
	                }
	 	  	          else{
	 	  	         	 addCell(rowNum, m_Col++, curRetail, m_StyleDec);
	                }
	             		
	 	  	          addCell(rowNum, m_Col++, curMargin, m_StylePct);
	                	                   
	 	  	          if(curSell != futSell){
	 	  	         	 addCell(rowNum, m_Col++, futSell, m_ShadeStyleDec);
	                }
	 	  	          else{
	 	  	         	 addCell(rowNum, m_Col++, futSell, m_StyleDec);
	                }
	                   	                    
	 	  	          if(curRetail != futRetail){
	 	  	         	 addCell(rowNum, m_Col++, futRetail, m_ShadeStyleDec);
	                }
	 	  	          else{
	 	  	         	 addCell(rowNum, m_Col++, futRetail, m_StyleDec);
	                }
	                   
	 	  	          addCell(rowNum, m_Col++, futMargin, m_StylePct);
	 	  	          addCell(rowNum, m_Col++, futMargin - curMargin, m_StylePct);
	                         	          
	 	  	          addCell(rowNum, m_Col++, unitsSoldLastRun( m_CustId, itemId ), m_StyleText);
	 	  	          addCell(rowNum, m_Col++, "", m_StyleText);
	      		 }
	      	}
				
				catch(Exception e){
					log.error("exception", e);
	      	}
				finally {
					DbUtils.closeDbConn(null, null, rsItemData);
	    		   rsItemData = null;
	    		}
	      		      		
	      	rowNum++;
	      }
	   }
		
		catch ( Exception e ) {
			log.error("exception", e);
		}
		
		finally {
		   DbUtils.closeDbConn(null, null, rsItemList);
		   rsItemList = null;
		}
		
		return rowNum;
	}
	
	/**
	    * Creates a row in the worksheet.
	    * @param rowNum The row number.
	    * @param colCnt The number of columns in the row.
	    * 
	    * @return The formatted row of the spreadsheet.
	    */
	   private int createItemDelListHdr(int rowNum)
	   {     
	      HSSFRow row = null;
	      HSSFCell cell = null;
	      HSSFCellStyle styleTitle1;  
	      HSSFFont fontTitle;   
	   	      
	      fontTitle = m_Wrkbk.createFont();
	      fontTitle.setFontHeightInPoints((short)9);
	      fontTitle.setFontName("Arial");
	      fontTitle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
	    
	      
	      styleTitle1 = m_Wrkbk.createCellStyle();
	      styleTitle1.setFont(fontTitle);
	      styleTitle1.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	      	      	      
	      row = addRow(rowNum);
	      row.setRowStyle(styleTitle1);
	     
	      //
	      // set the type and style of the cell.
	  	  cell = row.createCell(0);
	      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
	      cell.setCellStyle(styleTitle1);
	      cell.setCellValue(new HSSFRichTextString("Following items have been taken off the list from last run report:"));
	      
	      rowNum = rowNum + 2;
	      
	      return rowNum;
	
	   }
	
	    /**
	    * Creates the captions for the vendor filter.
	    * 
	    * @see SubRpt#createCaptions(int rowNum)
	    */
	   public void createRowCaptions(int rowNum)
	   {
	      HSSFRow row = null; 
	      HSSFRow row1 = null;
	      HSSFCellStyle styleCaptionsRow = null;
	      HSSFCellStyle styleCaptionsRow1 = null;
	      HSSFCellStyle styleCaptionsRow2 = null;
	      HSSFFont fontCaptionsRow = null;
	      HSSFFont fontCaptionsRow1 = null;
	      int col = 0;
	      ///int col1 = 0;
	              
	      fontCaptionsRow = m_Wrkbk.createFont();
	      fontCaptionsRow.setFontHeightInPoints((short)8);
	      fontCaptionsRow.setFontName("Arial");
	      fontCaptionsRow.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
	      
	      fontCaptionsRow1 = m_Wrkbk.createFont();
	      fontCaptionsRow1.setFontHeightInPoints((short)8);
	      fontCaptionsRow1.setFontName("Arial");
	      fontCaptionsRow1.setColor(HSSFColor.WHITE.index);
	      fontCaptionsRow1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
	      
	      styleCaptionsRow = m_Wrkbk.createCellStyle();
	      styleCaptionsRow.setFont(fontCaptionsRow);
	      styleCaptionsRow.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	      
	      //
	      //Shading
	      styleCaptionsRow.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
	      styleCaptionsRow.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	      
	      //
	      //Border
	      styleCaptionsRow.setBorderTop(HSSFCellStyle.BORDER_THIN);
	      styleCaptionsRow.setBorderBottom(HSSFCellStyle.BORDER_THIN);
	      styleCaptionsRow.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	      styleCaptionsRow.setBorderRight(HSSFCellStyle.BORDER_THIN);
	      
	      //
	      //Style for Today
	      styleCaptionsRow1 = m_Wrkbk.createCellStyle();
	      styleCaptionsRow1.setFont(fontCaptionsRow1);
	      styleCaptionsRow1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	      
	      //
	      //Shading
	      ///styleCaptionsRow1.setFillForegroundColor(HSSFColor.GREY_80_PERCENT.index);
	      styleCaptionsRow1.setFillForegroundColor(HSSFColor.PINK.index);
	      styleCaptionsRow1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	     	  	      
	      //
	      //Border
	      styleCaptionsRow1.setBorderTop(HSSFCellStyle.BORDER_THIN);
	      styleCaptionsRow1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
	      styleCaptionsRow1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	      styleCaptionsRow1.setBorderRight(HSSFCellStyle.BORDER_THIN);
	      
	      //
	      //Style for New
	      styleCaptionsRow2 = m_Wrkbk.createCellStyle();
	      styleCaptionsRow2.setFont(fontCaptionsRow1);
	      styleCaptionsRow2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	      
	      //
	      //Shading
	      styleCaptionsRow2.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
	      styleCaptionsRow2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	     	  	      
	      //
	      //Border
	      styleCaptionsRow2.setBorderTop(HSSFCellStyle.BORDER_THIN);// This is working
	      styleCaptionsRow2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
	      styleCaptionsRow2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	      styleCaptionsRow2.setBorderRight(HSSFCellStyle.BORDER_THIN);
	      	      
	      HSSFCellStyle styleWrapText; 
          
          //
          //Style for wrap text
          styleWrapText= m_Wrkbk.createCellStyle();
          styleWrapText.setFont(fontCaptionsRow);
          styleWrapText.setWrapText(true);
          styleWrapText.setAlignment(HSSFCellStyle.ALIGN_CENTER);
          //
          //Shading
          styleWrapText.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
          styleWrapText.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
          //
          //Border
          styleWrapText.setBorderTop(HSSFCellStyle.BORDER_THIN);// This is working
          styleWrapText.setBorderBottom(HSSFCellStyle.BORDER_THIN);
          styleWrapText.setBorderLeft(HSSFCellStyle.BORDER_THIN);
          styleWrapText.setBorderRight(HSSFCellStyle.BORDER_THIN);
	               
          row1 = m_Sheet.createRow(rowNum);
	      row1.setRowStyle(styleCaptionsRow);
	      createCaptionCell(row1, 1, "Best Brands Priced Right Changes Effective "+m_PriceDate,styleCaptionsRow);
	      m_Sheet.setColumnWidth(1, 12000);
	      createCaptionCell(row1, 11, "",styleCaptionsRow1);
	      m_Sheet.setColumnWidth(11, 2000);
	      createCaptionCell(row1, 12, "Today",styleCaptionsRow1);
	      m_Sheet.setColumnWidth(12, 2000);
	      createCaptionCell(row1, 13, "",styleCaptionsRow1);
	      m_Sheet.setColumnWidth(13, 2000);
	      	      
	      createCaptionCell(row1, 14, "",styleCaptionsRow2);
	      m_Sheet.setColumnWidth(14, 2500);
	      createCaptionCell(row1, 15, "New",styleCaptionsRow2);
	      m_Sheet.setColumnWidth(15, 2000);
	      createCaptionCell(row1, 16, "",styleCaptionsRow2);
	      m_Sheet.setColumnWidth(16, 2000);
	      createCaptionCell(row1, 17, "",styleCaptionsRow2);
	      m_Sheet.setColumnWidth(17, 2000);
	      
	      rowNum = rowNum + 1;
	    
	      row = m_Sheet.createRow(rowNum);
	      row.setRowStyle(styleCaptionsRow);
	          
	      createCaptionCell(row, col, "Department",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 3500);
	      createCaptionCell(row, col, "Vendor Name",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 6000);
	      createCaptionCell(row, col, "MFG#",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 1800);
	      createCaptionCell(row, col, "UPC",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 3000);
	      createCaptionCell(row, col, "Item Description",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 10000);
	      createCaptionCell(row, col, "UOM",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 1400);
	      createCaptionCell(row, col, "Stock Pk",styleWrapText);
	      m_Sheet.setColumnWidth(col++, 1500);
	      createCaptionCell(row, col, "NBC",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 1400);
	      createCaptionCell(row, col, "Top Seller",styleWrapText);
	      m_Sheet.setColumnWidth(col++, 1500);
	      createCaptionCell(row, col, "MADE IN USA",styleWrapText);
	      m_Sheet.setColumnWidth(col++, 1600);
	      createCaptionCell(row, col, "Item",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 2500);
	      createCaptionCell(row, col, "Cost",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 1800);
	      createCaptionCell(row, col, "Retail",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 1800);
	      createCaptionCell(row, col, "Margin",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 1800);
	      createCaptionCell(row, col, "Cost",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 1800);
	      createCaptionCell(row, col, "Retail",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 1800);
	      createCaptionCell(row, col, "Margin",styleCaptionsRow);
	      m_Sheet.setColumnWidth(col++, 1800);
	      createCaptionCell(row, col, "Margin %Chg",styleWrapText);
	      m_Sheet.setColumnWidth(col++, 2000);
	      createCaptionCell(row, col, "Purchased within 2yrs",styleWrapText);
	      m_Sheet.setColumnWidth(col++, 3000);
	      createCaptionCell(row, col, "NEW ITEM",styleWrapText);
	      m_Sheet.setColumnWidth(col++, 1800);
	         
	   }
	   
	   protected HSSFCell createCaptionCell(HSSFRow row, int col, String caption, HSSFCellStyle stylCaptions)
	   {
	      HSSFCell cell = null;
	      HSSFCellStyle m_CSCaption = null;
	      HSSFFont font = null;
	      
	      if ( row != null ) {
	      	 font = m_Wrkbk.createFont();
	    	 font.setFontHeightInPoints((short)8);
	         font.setFontName("Arial");
	         font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
	          
	         m_CSCaption = m_Wrkbk.createCellStyle();
	         m_CSCaption.setFont(font);
	         m_CSCaption.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	                    
	         //
	         //Shading
	         m_CSCaption.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
	         m_CSCaption.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	                 
	         cell = row.createCell(col);
	         cell.setCellType(HSSFCell.CELL_TYPE_STRING);
	         cell.setCellStyle(stylCaptions);
	         cell.setCellValue(new HSSFRichTextString(caption != null ? caption : ""));
	      }
	      
	      return cell;
	   }
	
	/**
	    * Creates a row in the worksheet.
	    * @param rowNum The row number.
	    * @param colCnt The number of columns in the row.
	    * 
	    * @return The formatted row of the spreadsheet.
	    */
	   private void createItmProgRow(int rowNum, int colCnt, String flcDesc)
	   {
	      HSSFRow row = null;
	      HSSFCell cell = null;
	      HSSFCellStyle styleTitle1;   
	      HSSFFont fontTitle;   
	   	      
	      fontTitle = m_Wrkbk.createFont();
	      fontTitle.setFontHeightInPoints((short)9);
	      fontTitle.setFontName("Arial");
	      fontTitle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
	      fontTitle.setColor(HSSFColor.WHITE.index);
	      
	      styleTitle1 = m_Wrkbk.createCellStyle();
	      styleTitle1.setFont(fontTitle);
	      styleTitle1.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	      
	      //
	      //Shading
	      styleTitle1.setFillForegroundColor(HSSFColor.GREY_80_PERCENT.index);
	      styleTitle1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	      
	      //
	      //Assign border for each cell of the row
	      styleTitle1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	      styleTitle1.setBorderRight(HSSFCellStyle.BORDER_THIN);
	   
	      row = addRow(rowNum);
	      row.setRowStyle(styleTitle1);
	     
	      //
	      // set the type and style of the cell.
	  	  cell = row.createCell(colCnt);
	      cell.setCellType(HSSFCell.CELL_TYPE_STRING);
	      cell.setCellStyle(styleTitle1);
	      cell.setCellValue(new HSSFRichTextString(flcDesc));
	
	   }
	   

	/* (non-Javadoc)
	 * @see com.emerywaterhouse.rpt.server.Report#createReport()
	 */
	@Override
	public boolean createReport()
	{
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         m_OraConn = m_RptProc.getOraConn();
         
         if ( m_OraConn != null && prepareStatements() )
         	created = buildOutputFile();            
      }
      
      catch ( Exception ex ) {
         log.fatal("exception:", ex);
      }
      
      finally {
         close(); 
         
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }
      
      return created;
	}
	
	/**
	 * Returns the name of the customer the spreadsheet it run for
	 * 
	 * @param custId String - the customer id
	 * @return String - the custome name
	 * @throws SQLException
	 */
	private String getCustName(String custId) throws SQLException
	{
		String name = "Invalid Customer#";
		Statement stmt = null;
		ResultSet rs = null;
		
		try {
			stmt = m_OraConn.createStatement();
			rs = stmt.executeQuery("select name from customer where customer_id = '" + custId + "'");
			
			if ( rs.next() )
				name = rs.getString("name");
		}
		
		finally {
			DbUtils.closeDbConn(null, stmt, rs);
			rs = null;
			stmt = null;
		}
		
		return name;
	}
	
	/**
	 * Prepares queries
	 * 
	 * @return boolean - true if successfull
	 * @throws Exception
	 */
	private boolean prepareStatements() throws Exception
	{
		StringBuffer sql = new StringBuffer();
		
		try {
    		sql.setLength(0);
			sql.append("select report_date, emery_dept.dept_num, vendor_item_num mfr_nbr,  ");
			sql.append("   emery_dept.name Department, vendor.name vendor_name, stock_pack, ");
			sql.append("   decode(broken_case.description, 'ALLOW BROKEN CASES', '', 'N') nbc, item.retail_pack, ");
			sql.append("   item.flc_id, wir.item_id as tsi, item_attribute.item_id as usa_item, item.item_id, item.description item_descr, retail_unit.unit, ");
			sql.append("   item_upc.upc_code, current_sell, current_retail, future_sell, future_retail, ");
			sql.append("   item_program, new_add ");
			sql.append("from bbpr_adv_change ");
			sql.append("join item on item.item_id = bbpr_adv_change.item_id ");
			sql.append("left outer join web_item_rank wir on bbpr_adv_change.item_id = wir.item_id ");
			sql.append("left outer join item_attribute on item.item_id = item_attribute.item_id  and "); 
        	sql.append("item_attribute.attribute_value_id in ( select attribute_value_id from attribute a , attribute_value av where a.attribute_id = av.attribute_id and av.value = 'MADE IN USA') ");
        	sql.append("join emery_dept on emery_dept.dept_id = item.dept_id ");
			sql.append("join retail_unit on retail_unit.unit_id = item.ret_unit_id ");
			sql.append("join vendor on vendor.vendor_id = item.vendor_id ");
			sql.append("join vendor_item_cross on vendor_item_cross.vendor_id = item.vendor_id and vendor_item_cross.item_id = item.item_id ");
			sql.append("join broken_case on item.broken_case_id = broken_case.broken_case_id ");
			sql.append("join item_upc on item_upc.item_id = item.item_id and ");
			sql.append("   item_upc.primary_upc = 1 ");
			sql.append("where customer_id = ? and ");
			sql.append("   report_date = ? ");
			sql.append("order by item_program, emery_dept.name, vendor.name,vendor_item_num ");
			m_ReportData = m_OraConn.prepareStatement(sql.toString());
			
			//
			//Get all the items that have been taken of the list w.r.to last run report
			//Last run report minus current month report
			sql.setLength(0);
			sql.append("select item_id,item_program ");
			sql.append("from bbpr_adv_change ");
			sql.append("where customer_id = ? and "); 
			sql.append("      report_date = ? ");
			sql.append("minus ");
			sql.append("select item_id,item_program from bbpr_adv_change ");  
			sql.append("where customer_id = ? and ");
			sql.append("      report_date = ? ");
			sql.append("order by item_program ");
			m_ItemDelList = m_OraConn.prepareStatement(sql.toString());
			
			sql.setLength(0);
			sql.append("select report_date, emery_dept.dept_num, vendor_item_num mfr_nbr,  ");
			sql.append("   emery_dept.name Department, vendor.name vendor_name, stock_pack, ");
			sql.append("   decode(broken_case.description, 'ALLOW BROKEN CASES', '', 'N') nbc, ");
			sql.append("   item.flc_id, item.item_id, item.description item_descr, retail_unit.unit, item.retail_pack, ");
			sql.append("   item.flc_id, wir.item_id as tsi, item_attribute.item_id as usa_item, ");
			sql.append("   item_upc.upc_code, current_sell, current_retail, future_sell, future_retail, ");
			sql.append("   item_program ");
			sql.append("from bbpr_adv_change ");
			sql.append("join item on item.item_id = bbpr_adv_change.item_id ");
			sql.append("left outer join web_item_rank wir on bbpr_adv_change.item_id = wir.item_id ");
			sql.append("left outer join item_attribute on item.item_id = item_attribute.item_id  and ");
			sql.append("item_attribute.attribute_value_id in ( select attribute_value_id from attribute a , attribute_value av where a.attribute_id = av.attribute_id and av.value = 'MADE IN USA') ");
			sql.append("join emery_dept on emery_dept.dept_id = item.dept_id ");
			sql.append("join retail_unit on retail_unit.unit_id = item.ret_unit_id ");
			sql.append("join vendor on vendor.vendor_id = item.vendor_id ");
			sql.append("left outer join vendor_item_cross on vendor_item_cross.item_id = item.item_id ");
			sql.append("join broken_case on item.broken_case_id = broken_case.broken_case_id ");
			sql.append("join item_upc on item_upc.item_id = item.item_id and ");
			sql.append("   item_upc.primary_upc = 1 ");
			sql.append("where customer_id = ? and "); 
			sql.append("   bbpr_adv_change.item_id = ? and ");
			sql.append("   item_program = ? and ");
			sql.append("   report_date = ? ");
			sql.append("order by item_program, emery_dept.name, item.flc_id, bbpr_adv_change.item_id ");
			m_ItemDelData = m_OraConn.prepareStatement(sql.toString());
				 			
			sql.setLength(0);
      	sql.append("select sum(qty_shipped) qty from inv_dtl ");
      	sql.append("where cust_nbr = ? and item_nbr = ? and ");
        	sql.append("invoice_date <= ? and ");
         sql.append("invoice_date >= ? -  ? ");
         m_PurchHist = m_OraConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select customer_id ");
         sql.append("from customer ");
        	sql.append("start with customer_id = ? ");
        	sql.append("connect by prior customer_id = parent_id ");
         m_FamilyTree = m_OraConn.prepareStatement(sql.toString());
   		
			return true;
		}
		
		finally {
			sql = null;
		}
	}
	
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
   	String name;
   	int yr, mo, dy;
   	String tmp;

      for ( int i = 0; i < params.size(); i++ ) {
      	name = params.get(i).name;
      	
      	if ( name.equals("custid")) { 
      		m_CustId = params.get(i).value.trim();
      		log.info("Running BBPR report for " + m_CustId);
      		continue;
      	}
      	
      	if ( name.equals("pardate") ) {   // format yyyymmdd
      		tmp = params.get(i).value.trim();
      		yr = Integer.valueOf(tmp.substring(0, 4));
      		mo = Integer.valueOf(tmp.substring(4, 6)) - 1;
      		dy = Integer.valueOf(tmp.substring(6));
      		
      		m_RunDate = new Date(new GregorianCalendar(yr, mo, dy).getTimeInMillis());
      		continue;
      	}
      	
      	if ( name.equals("pricedate") ) {   // format yyyymmdd
      		tmp = params.get(i).value.trim();
      		yr = Integer.valueOf(tmp.substring(0, 4));
      		mo = Integer.valueOf(tmp.substring(4, 6)) - 1;
      		dy = Integer.valueOf(tmp.substring(6));
      		
      		m_PriceDate = new Date(new GregorianCalendar(yr, mo, dy).getTimeInMillis());
      		continue;
      	}
      	
      	if ( name.equals("lastrundate") ) {   // format yyyymmdd
      		tmp = params.get(i).value.trim();
      		yr = Integer.valueOf(tmp.substring(0, 4));
      		mo = Integer.valueOf(tmp.substring(4, 6)) - 1;
      		dy = Integer.valueOf(tmp.substring(6));
      		
      		m_LastRunDate = new Date(new GregorianCalendar(yr, mo, dy).getTimeInMillis());
      		continue;
      	}
      	      	
      	if ( name.equals("historydays") ) {
      		m_Days = Integer.valueOf(params.get(i).value.trim() );
      		continue;
      	}
      }
   }
   
   /**
    * 
    * @param custid
    * @param itemid
    * @return Yes if the number of units sold > 0,else ""
    */
   private String unitsSold(String custid, String itemid)
   {
      ResultSet rs = null;
      int qty = 0;
      String result = "";

      try{
         for ( int j = 0; j < m_StoreId.size(); j++ ) {
            if(qty > 0){
               result = "Yes";		
               break;
            }
            m_PurchHist.setString(1, m_StoreId.get(j));
            m_PurchHist.setString(2, itemid);
            m_PurchHist.setDate(3, m_RunDate); 
            m_PurchHist.setDate(4, m_RunDate); 
            m_PurchHist.setInt(5, m_Days);

            rs = m_PurchHist.executeQuery();

            if ( rs.next() ){
               qty = rs.getInt("qty"); 
            }
         }
      }
      catch ( Exception e ) {
    	 qty = 0;
         log.error("exception",  e );
      }
      finally {
      	DbUtils.closeDbConn(null, null, rs);
        rs = null;
      }

      return result;
   }
   
   /**
    * 
    * @param custid
    * @param itemid
    * @return Yes if the number of units sold > 0,else ""
    */
   private String unitsSoldLastRun(String custid, String itemid)
   {
      ResultSet rs = null;
      int qty = 0;
      String result = "";
     
      try{
          for ( int j = 0; j < m_StoreId.size(); j++ ) {
             if(qty > 0){
                result = "Yes";		
                break;
             }
                  
             m_PurchHist.setString(1, m_StoreId.get(j));
             m_PurchHist.setString(2, itemid);
             m_PurchHist.setDate(3, m_LastRunDate); 
             m_PurchHist.setDate(4, m_LastRunDate); 
             m_PurchHist.setInt(5, m_Days);

             rs = m_PurchHist.executeQuery();

             if ( rs.next() ){
                qty = rs.getInt("qty"); 
             }
          }
      }
      catch ( Exception e ) {
    	 qty = 0;
         log.error("exception",  e );
      }
      finally {
      	DbUtils.closeDbConn(null, null, rs);
        rs = null;
      }

      return result;
   }

}

