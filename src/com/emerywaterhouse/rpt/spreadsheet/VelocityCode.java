/**
 * Title:			VelocityCode.java
 * Description:	ite sdales, inventory, orders by velocity code 
 * Company:			Emery-Waterhouse
 * @author			smurdock
 * @version			1.0
 * <p>
 * 
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;


import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class VelocityCode extends Report {

   private PreparedStatement m_Velocity;
   private PreparedStatement m_DefaultDates;
   
	 
   private XSSFWorkbook m_WrkBk;
   private XSSFSheet m_Sheet;
   private XSSFRow m_Row = null;
   private Header m_Header;
   
   private XSSFFont m_Font;
   private XSSFFont m_FontTitle;
   private XSSFFont m_FontBold;
   private XSSFFont m_FontData;

   private XSSFCellStyle m_StyleText;  		// Text left justified
   private XSSFCellStyle m_StyleTextRight;  	// Text right justified
   private XSSFCellStyle m_StyleTextCenter; 	// Text centered
   private XSSFCellStyle m_StyleTitle; 		// Bold, centered
   private XSSFCellStyle m_StyleBold;  		// Normal but bold
   private XSSFCellStyle m_StyleBoldRight; 	// Normal but bold & right aligned
   private XSSFCellStyle m_StyleBoldCenter; 	// Normal but bold & centered
   private XSSFCellStyle m_StyleDec;   		// Style with 1 decimals and % sign
   private XSSFCellStyle m_StyleDecBold;		// Style with 2 decimals, bold
   private XSSFCellStyle m_StyleHeader; 		// Bold, centered 12pt
   private XSSFCellStyle m_StyleInt;   		// Style with 0 decimals

   // Parameters
   private String m_BegDate;
   private String m_EndDate;
   private String m_Warehouse;
   
   private short m_RowNum = 0;
   
   /**
    * Builds the output file
    * @return boolean.  True if the file was created, false if not.
    * @throws FileNotFoundException
    */
   public boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream outFile = null;
      boolean result = true;
      ResultSet rs = null;
      int col;
      
      String lastDept = "begin";
      int ItemTotal = 0;
      int SalesTotal = 0;
      int MarginTotal = 0;
      int Inv$Total = 0;
      int InvUnitsTotal = 0;
      int OpenPO$Total = 0;
      int OpenPOUnitsTotal = 0;     
      int divizor = 0;
      
      m_FileNames.add(m_RptProc.getUid() + "VelocityCode" + getStartTime() + ".xlsx");      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      initReport();

      try {
         if ( m_BegDate == null || m_BegDate.length() == 0 || m_EndDate == null || m_EndDate.length() == 0 ){
            setDefaultDates();
         }
         m_RowNum = createCaptions();
         setCurAction( "Running the velocity report query" );
            m_Velocity.setString(1, m_BegDate); //next four are sales $
            m_Velocity.setString(2, m_EndDate);
            m_Velocity.setString(3, m_BegDate);
            m_Velocity.setString(4, m_EndDate);
            m_Velocity.setString(5, m_BegDate);  // next four are margin $
            m_Velocity.setString(6, m_EndDate);
            m_Velocity.setString(7, m_BegDate);
            m_Velocity.setString(8, m_EndDate);
            m_Velocity.setString(9, m_EndDate);
            m_Velocity.setString(10, m_EndDate);
            m_Velocity.setString(11, m_EndDate);
            m_Velocity.setString(12, m_EndDate);
            
          rs = m_Velocity.executeQuery();

         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            setCurAction( "Processing Velocity Report");            
            
            // Check for the start of a new dept
            if ( !rs.getString("dname").equals(lastDept) ) { 
                  lastDept = rs.getString("dname");
                  m_Row = m_Sheet.createRow(m_RowNum++);
                  m_Row = m_Sheet.createRow(m_RowNum++);
                  col = (short)0;
                  m_Sheet.setColumnWidth(col, 4000);
                  if (!lastDept.equals("EMERY")){ //if it's not EMERY it's a dept and we wnat the dept number too
                     createCell(m_Row, col++, rs.getString("dept")+ " " + lastDept, m_StyleHeader);
                  } else {
                     createCell(m_Row, col++, lastDept, m_StyleHeader);                     
                  }   
                  m_Sheet.setColumnWidth(col, 2200);
                  createCell(m_Row, col++, "Total", m_StyleHeader);
                  m_Sheet.setColumnWidth(col, 2200);
                  createCell(m_Row, col++, "A", m_StyleHeader);
                  m_Sheet.setColumnWidth(col, 2200);
                  createCell(m_Row, col++, "B", m_StyleHeader);
                  m_Sheet.setColumnWidth(col, 2200);
                  createCell(m_Row, col++, "C", m_StyleHeader);
                  m_Sheet.setColumnWidth(col, 2200);
                  createCell(m_Row, col++, "D", m_StyleHeader);
                  m_Sheet.setColumnWidth(col, 2200);
                  createCell(m_Row, col++, "E", m_StyleHeader);
                  m_Sheet.setColumnWidth(col, 2200);         
                  createCell(m_Row, col++, "I", m_StyleHeader);
                  m_Sheet.setColumnWidth(col, 2200);                 
                 
               }
                     
            if ( lastDept.equals("EMERY")){  //this is a gran totale of some sort

                  if (rs.getString("catto").substring(4).equals("Item Count")) {
                     ItemTotal = rs.getInt("Total");                    
                  } else if (rs.getString("catto").substring(4).equals("Sales $")){
                     SalesTotal = rs.getInt("Total");
                  } else if (rs.getString("catto").substring(4).equals("Margin $")){
                     MarginTotal = rs.getInt("Total");
                  } else if (rs.getString("catto").substring(4).equals("Inventory $")){
                     Inv$Total = rs.getInt("Total");
                  } else if (rs.getString("catto").substring(4).equals("Inventory Units")){
                     InvUnitsTotal = rs.getInt("Total");
                  } else if (rs.getString("catto").substring(4).equals("Open PO $")){
                     OpenPO$Total = rs.getInt("Total");
                  }  else if (rs.getString("catto").substring(4).equals("Open PO Units")){
                     OpenPOUnitsTotal = rs.getInt("Total");
                  }             }
               // Print a row of totals.  Then, print a row of percentages;            	
            col = (short)0;
            m_Row = m_Sheet.createRow(m_RowNum++);
            createCell(m_Row, col++, rs.getString("catto").substring(4), m_StyleBold);            
            createCell(m_Row, col++, rs.getInt("total"), m_StyleInt);
            createCell(m_Row, col++, rs.getInt("A"), m_StyleInt);
            createCell(m_Row, col++, rs.getInt("B"), m_StyleInt);
            createCell(m_Row, col++, rs.getInt("C"), m_StyleInt);
            
            createCell(m_Row, col++, rs.getInt("D"), m_StyleInt);
            createCell(m_Row, col++, rs.getInt("E"), m_StyleInt);
            createCell(m_Row, col++, rs.getInt("I"),m_StyleInt);
            
            col = (short)0;
            
            if (rs.getString("catto").substring(4).equals("Item Count")){
               divizor = ItemTotal;
            } else if (rs.getString("catto").substring(4).equals("Sales $")){
               divizor = SalesTotal;
            } else if (rs.getString("catto").substring(4).equals("Margin $")){
               divizor = MarginTotal;                              
            }else if (rs.getString("catto").substring(4).equals("Inventory $")){
               divizor = Inv$Total;              
            }else if (rs.getString("catto").substring(4).equals("Inventory Units")){
               divizor = InvUnitsTotal;              
            }else if (rs.getString("catto").substring(4).equals("Open PO $")){
               divizor = OpenPO$Total;              
            }else if (rs.getString("catto").substring(4).equals("Open PO Units")){
               divizor = OpenPOUnitsTotal;              
            }else {
               divizor = 1;                             
            }
            if (divizor == 0) {
              divizor = 1;
            }  
            m_Row = m_Sheet.createRow(m_RowNum++);
            createCell(m_Row, col++, "% of " + rs.getString("catto").substring(4), m_StyleText);
            createCell(m_Row, col++, (rs.getDouble("total") / divizor), m_StyleDec);
            createCell(m_Row, col++, (rs.getDouble("A") / divizor), m_StyleDec);
            createCell(m_Row, col++, (rs.getDouble("B") / divizor), m_StyleDec);
            createCell(m_Row, col++, (rs.getDouble("C") / divizor), m_StyleDec);            
            createCell(m_Row, col++, (rs.getDouble("D") / divizor), m_StyleDec);
            createCell(m_Row, col++, (rs.getDouble("E") / divizor), m_StyleDec);
            createCell(m_Row, col++, (rs.getDouble("I") / divizor),m_StyleDec);


         }
         m_WrkBk.write(outFile);
         result = true;
 
            
       }

      catch( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("The report had the following Error: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         
         result = false;
      }

      finally {
      	DbUtils.closeDbConn(null, null, rs);

         try {
            outFile.close();
            outFile = null;
         }
         catch ( Exception e ) {
            log.error( e );
         }         
      }
      
      return result;
   }


   /**
    * Resource cleanup
    */
   public void cleanup()
   {
   	DbUtils.closeDbConn(null, m_Velocity, null);
   	
   	m_Velocity = null;
    	
   	m_Header = null;
   	m_Font = null;
   	m_FontTitle = null;
   	m_FontBold = null;
   	m_FontData = null;
   	m_StyleText = null;
   	m_StyleTextRight = null;
   	m_StyleTextCenter = null;
   	m_StyleBold = null;
   	m_StyleBoldRight = null;
   	m_StyleBoldCenter = null;;
   	m_StyleDec = null;
   	m_StyleDecBold = null;
   	m_StyleHeader = null;
   	m_StyleInt = null;
   	m_BegDate = null;
   	m_EndDate = null;
   	m_Warehouse = null;   	
   	m_WrkBk = null;
   	m_Sheet = null;
   	m_Row = null;
   }
      
   /**
    * Creates a cell of type numeric
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return HSSFCell newly created cell
    */
   private XSSFCell createCell(XSSFRow row, int col, double val, XSSFCellStyle style)
   {
      XSSFCell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellValue(val);
      cell.setCellStyle(style);

      return cell;
   }

   /**
    * Creates a cell of type numeric
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return HSSFCell newly created cell
    */
   private XSSFCell createCell(XSSFRow row, int col, int val, XSSFCellStyle style)
   {
      XSSFCell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.NUMERIC);
      cell.setCellValue(val);
      cell.setCellStyle(style);

      return cell;
   }

   /**
    * Creates a cell of type String
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return HSSFCell newly created cell
    */
   private XSSFCell createCell(XSSFRow row, int col, String val, XSSFCellStyle style)
   {
      XSSFCell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(val));
      cell.setCellStyle(style);

      return cell;
   }
   
   /**
    * Creates the report title and the captions.
    */
   private short createCaptions()
   {
      XSSFRow row = null;
      XSSFCell cell = null;
      short rowNum = 0;
      StringBuffer caption = new StringBuffer("");
       
      if ( m_Sheet == null )
         return 0;      
      //
      // set the report title
      ++rowNum;
      row = m_Sheet.createRow(rowNum);
      cell = row.createCell(0); 
      cell.setCellType(CellType.STRING);
      cell.setCellStyle(m_StyleTitle);
      
      caption.append(m_BegDate);
      caption.append(" - ");
      caption.append(m_EndDate);
      caption.append("       ");

      if ( m_Warehouse != null && m_Warehouse.length() > 0 ){
         if (m_Warehouse.equals("01") || m_Warehouse.equals("02")){ // ok so i hard coded it sue me
            caption.append("Portland ");             
         }
         else 
            if (m_Warehouse.equals("04") ||m_Warehouse.equals("05")){
               caption.append("Pittston ");
         }   
      }
      else {
         caption.append("All Emery "); 
      }
      caption.append("Velocity Code Report");
      
      cell.setCellValue(new XSSFRichTextString(caption.toString()));
             
      return ++rowNum;
   }
   
   

   /**
    * Creates the report file.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {      
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         m_OraConn = m_RptProc.getOraConn();
         
         if ( prepareStatements() )
            created = buildOutputFile();            
      }
      
      catch ( Exception ex ) {
         log.fatal("exception:", ex);
      }
      
      finally {
        cleanup();
        
        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }
      
      return created;
   }

    /**
    * Creates the workbook and worksheet.  Creates any fonts and styles that
    * will be used.
    */
   private void initReport()
   {
      XSSFDataFormat df;

      try {
         m_WrkBk = new XSSFWorkbook();
         
         df = m_WrkBk.createDataFormat();

         //
         // Create the default font for this workbook
         m_Font = m_WrkBk.createFont();
         m_Font.setFontHeightInPoints((short) 8);
         m_Font.setFontName("Arial");

         //
         // Create a font for titles
         m_FontTitle = m_WrkBk.createFont();
         m_FontTitle.setFontHeightInPoints((short)10);
         m_FontTitle.setFontName("Arial");
         m_FontTitle.setBold(true);

         //
         // Create a font that is normal size & bold
         m_FontBold = m_WrkBk.createFont();
         m_FontBold.setFontHeightInPoints((short)8);
         m_FontBold.setFontName("Arial");
         m_FontBold.setBold(true);

         //
         // Create a font that is normal size & bold
         m_FontData = m_WrkBk.createFont();
         m_FontData.setFontHeightInPoints((short)8);
         m_FontData.setFontName("Arial");

         //
         // Create a font that is 12 pt & bold
         m_FontBold = m_WrkBk.createFont();
         m_FontBold.setFontHeightInPoints((short)8);
         m_FontBold.setFontName("Arial");
         m_FontBold.setBold(true);

         //
         // Setup the cell styles used in this report
         m_StyleText = m_WrkBk.createCellStyle();
         m_StyleText.setFont(m_FontData);
         m_StyleText.setAlignment(HorizontalAlignment.LEFT);

         m_StyleTextRight = m_WrkBk.createCellStyle();
         m_StyleTextRight.setFont(m_FontData);
         m_StyleTextRight.setAlignment(HorizontalAlignment.RIGHT);

         m_StyleTextCenter = m_WrkBk.createCellStyle();
         m_StyleTextCenter.setFont(m_FontData);
         m_StyleTextCenter.setAlignment(HorizontalAlignment.CENTER);

         // Style 8pt, left aligned, bold 
         m_StyleBold = m_WrkBk.createCellStyle();
         m_StyleBold.setFont(m_FontBold);
         m_StyleBold.setAlignment(HorizontalAlignment.LEFT);

         // Style 8pt, right aligned, bold 
         m_StyleBoldRight = m_WrkBk.createCellStyle();
         m_StyleBoldRight.setFont(m_FontBold);
         m_StyleBoldRight.setAlignment(HorizontalAlignment.RIGHT);

         // Style 8pt, centered, bold 
         m_StyleBoldCenter = m_WrkBk.createCellStyle();
         m_StyleBoldCenter.setFont(m_FontBold);
         m_StyleBoldCenter.setAlignment(HorizontalAlignment.CENTER);

         m_StyleDec = m_WrkBk.createCellStyle();
         m_StyleDec.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleDec.setFont(m_FontData);
         m_StyleDec.setDataFormat(df.getFormat("#,##0.0%"));

         m_StyleDecBold = m_WrkBk.createCellStyle();
         m_StyleDecBold.setFont(m_FontBold);
         m_StyleDecBold.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleDecBold.setDataFormat(df.getFormat("#,##0.0%"));

         m_StyleHeader = m_WrkBk.createCellStyle();
         m_StyleHeader.setFont(m_FontBold);
         m_StyleHeader.setAlignment(HorizontalAlignment.CENTER);
         m_StyleHeader.setFillPattern(FillPatternType.FINE_DOTS);
  
         m_StyleInt = m_WrkBk.createCellStyle();
         m_StyleInt.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleInt.setFont(m_FontBold);
         m_StyleInt.setDataFormat((short)3);

         m_StyleTitle = m_WrkBk.createCellStyle();
         m_StyleTitle.setFont(m_FontTitle);
         m_StyleTitle.setAlignment(HorizontalAlignment.LEFT);

         m_Sheet = m_WrkBk.createSheet();
         m_Sheet.setMargin(XSSFSheet.BottomMargin, .25);
         m_Sheet.getPrintSetup().setLandscape(true);
         m_Sheet.getPrintSetup().setPaperSize((short)5);
         
         m_Header = m_Sheet.getHeader();
         m_Header.setCenter(HSSFHeader.font("Arial", "Bold") + HSSFHeader.fontSize((short) 12) + "Daily Sales Order Cuts");
         m_Header.setLeft(HSSFHeader.font("Arial", "Bold") + HSSFHeader.fontSize((short) 12) + " " + m_BegDate + " thru " + m_EndDate);
         m_Header.setRight(HSSFHeader.font("Arial", "Bold") + HSSFHeader.fontSize((short) 12) + HSSFHeader.page());

         m_RowNum = 0;
         
         // Initialize the default column widths 
         for ( short i = 0; i < 8; i++ ) 
         	m_Sheet.setColumnWidth(i, 2000);  
                                
         m_Sheet.setColumnWidth(1, 2000);
         m_Sheet.setColumnWidth(2, 7000);
       }

      catch ( Exception e ) {
         log.error( e );
      }
   }
   
    

   private boolean prepareStatements() 
   {
      StringBuffer sql = new StringBuffer();
      ////for tracking only -- this query is long and Eclipse does not always show all of the string
      ////  so 'squalid' will get the back end of the query
      ///StringBuffer squalid = new StringBuffer(); 	
   	try {
   		sql.setLength(0);
         sql.append("select   \r\n");
         sql.append("      '2A: Item Count' catto,emery_dept.dept_num Dept, emery_dept.name Dname, count(item.item_id) \"Total\",  \r\n");
         sql.append("      sum(case when velocity = 'A' then 1 else 0 end) \"A\",  \r\n");
         sql.append("      sum(case when velocity = 'B' then 1 else 0 end) \"B\",  \r\n");
         sql.append("      sum(case when velocity = 'C' then 1 else 0 end) \"C\",  \r\n");
         sql.append("      sum(case when velocity = 'D' then 1 else 0 end) \"D\",  \r\n");
         sql.append("      sum(case when velocity = 'E' then 1 else 0 end) \"E\",  \r\n");
         sql.append("      sum(case when velocity = 'I' then 1 else 0 end) \"I\"  \r\n");
         sql.append(" from item  \r\n");
         sql.append(" join emery_dept on emery_dept.dept_id = item.dept_id  \r\n");
         sql.append(" join item_velocity on item_velocity.velocity_id = item.velocity_id  \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){         
            sql.append(" join item_warehouse iw on iw.item_id = item.item_id  and iw. active = 1  \r\n");
            sql.append("and iw.warehouse_id in  \r\n");
            sql.append("(select w.warehouse_id from warehouse w where w.fas_facility_id = ");
            sql.append(m_Warehouse);
            sql.append(") \r\n");           
         } else {
            sql.append("where item.item_id in \r\n");
            sql.append("(select distinct(iw.item_id) from item_warehouse iw \r\n");
            sql.append(" where iw.active = 1) \r\n");
         } 
         sql.append(" group by emery_dept.dept_num, emery_dept.name  \r\n");
         sql.append("   \r\n");
         sql.append(" union  \r\n");
         sql.append("   \r\n");
         sql.append(" select '1A: Item Count' catto, '00', 'EMERY', count(item.item_id) \"Total\",  \r\n");
         sql.append("      sum(case when velocity = 'A' then 1 else 0 end) \"A\",  \r\n");
         sql.append("      sum(case when velocity = 'B' then 1 else 0 end) \"B\",  \r\n");
         sql.append("      sum(case when velocity = 'C' then 1 else 0 end) \"C\",  \r\n");
         sql.append("      sum(case when velocity = 'D' then 1 else 0 end) \"D\",  \r\n");
         sql.append("      sum(case when velocity = 'E' then 1 else 0 end) \"E\",  \r\n");
         sql.append("      sum(case when velocity = 'I' then 1 else 0 end) \"I\"  \r\n");
         sql.append(" from item \r\n");
         sql.append(" join emery_dept on emery_dept.dept_id = item.dept_id  \r\n");
         sql.append(" join item_velocity on item_velocity.velocity_id = item.velocity_id   \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                  
            sql.append(" join item_warehouse iw on iw.item_id = item.item_id  and iw. active = 1  \r\n");
            sql.append("and iw.warehouse_id in  \r\n");
            sql.append("(select w.warehouse_id from warehouse w where w.fas_facility_id = ");
            sql.append(m_Warehouse);
            sql.append(") \r\n");            
         } else {
            sql.append("where item.item_id in \r\n");
            sql.append("(select distinct(iw.item_id) from item_warehouse iw \r\n");
            sql.append(" where iw.active = 1) \r\n");
         }    
         sql.append("   \r\n");
         
         sql.append(" union  \r\n");
         sql.append("   \r\n");
         sql.append(" select   \r\n");
         sql.append("      '2B: Sales $' catto, buying_dept Dept, emery_dept.name dname, sum(unit_sell * qty_shipped) \"Total\",  \r\n");
         sql.append("      sum(case when velocity = 'A' then unit_sell * qty_shipped else 0 end) \"A\",  \r\n");
         sql.append("      sum(case when velocity = 'B' then unit_sell * qty_shipped else 0 end) \"B\",  \r\n");
         sql.append("      sum(case when velocity = 'C' then unit_sell * qty_shipped else 0 end) \"C\",  \r\n");
         sql.append("      sum(case when velocity = 'D' then unit_sell * qty_shipped else 0 end) \"D\",  \r\n");
         sql.append("      sum(case when velocity = 'E' then unit_sell * qty_shipped else 0 end) \"E\",  \r\n");
         sql.append("      sum(case when velocity = 'I' then unit_sell * qty_shipped else 0 end) \"I\"  \r\n");
         sql.append(" from inv_dtl  \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                           
             sql.append("join inv_hdr on inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id  \r\n"); 
             sql.append("and inv_hdr.warehouse in  \r\n");
             sql.append("(select w.name from warehouse w where w.fas_facility_id = ");
             sql.append(m_Warehouse);
             sql.append(") \r\n");            
         }   
         sql.append(" join item on item.item_id = inv_dtl.item_nbr      \r\n");
         sql.append(" join item_velocity on item_velocity.velocity_id = item.velocity_id  \r\n");
         sql.append(" join emery_dept on emery_dept.DEPT_NUM = inv_dtl.BUYING_DEPT  \r\n");
         sql.append(" where inv_dtl.invoice_date between to_date(?,'mm/dd/yyyy') and to_date(?,'mm/dd/yyyy') and inv_dtl.sale_type = 'WAREHOUSE'  \r\n");
         sql.append(" group by inv_dtl.buying_dept, emery_dept.name  \r\n");
         sql.append("   \r\n");
         sql.append(" union  \r\n");
         sql.append("   \r\n");
         sql.append(" select   \r\n");
         sql.append("      '1B: Sales $' catto, '00','EMERY' dname, sum(unit_sell * qty_shipped) \"Total\",  \r\n");
         sql.append("      sum(case when velocity = 'A' then unit_sell * qty_shipped else 0 end) \"A\",  \r\n");
         sql.append("      sum(case when velocity = 'B' then unit_sell * qty_shipped else 0 end) \"B\",  \r\n");
         sql.append("      sum(case when velocity = 'C' then unit_sell * qty_shipped else 0 end) \"C\",  \r\n");
         sql.append("      sum(case when velocity = 'D' then unit_sell * qty_shipped else 0 end) \"D\",  \r\n");
         sql.append("      sum(case when velocity = 'E' then unit_sell * qty_shipped else 0 end) \"E\",  \r\n");
         sql.append("      sum(case when velocity = 'I' then unit_sell * qty_shipped else 0 end) \"I\"  \r\n");
         sql.append(" from inv_dtl  \r\n");
         sql.append(" join emery_dept on emery_dept.dept_num = buying_dept  \r\n");
         sql.append(" join item on item.item_id = item_nbr  \r\n");
         sql.append(" join item_velocity on item_velocity.velocity_id = item.velocity_id  \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                           
            sql.append("join inv_hdr on inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id  \r\n"); 
            sql.append("and inv_hdr.warehouse in  \r\n");
            sql.append("(select w.name from warehouse w where w.fas_facility_id = ");
            sql.append(m_Warehouse);
            sql.append(") \r\n");            
         }   
         sql.append(" where inv_dtl.invoice_date between to_date(?,'mm/dd/yyyy') and to_date(?,'mm/dd/yyyy') and inv_dtl.sale_type = 'WAREHOUSE'  \r\n");
         sql.append("   \r\n");
         
         sql.append(" union  \r\n");
         sql.append("   \r\n");
         sql.append(" select   \r\n");
         sql.append("      '2C: Margin $' catto, buying_dept Dept, emery_dept.name dname, sum(unit_sell * qty_shipped) - sum(unit_cost * qty_shipped) \"Total\",  \r\n");
         sql.append("      sum(case when velocity = 'A' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"A\",  \r\n");
         sql.append("      sum(case when velocity = 'B' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"B\",  \r\n");
         sql.append("      sum(case when velocity = 'C' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"C\",  \r\n");
         sql.append("      sum(case when velocity = 'D' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"D\",  \r\n");
         sql.append("      sum(case when velocity = 'E' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"E\",  \r\n");
         sql.append("      sum(case when velocity = 'I' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"I\"  \r\n");
         sql.append(" from inv_dtl  \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                           
             sql.append("join inv_hdr on inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id  \r\n"); 
             sql.append("and inv_hdr.warehouse in  \r\n");
             sql.append("(select w.name from warehouse w where w.fas_facility_id = ");
             sql.append(m_Warehouse);
             sql.append(") \r\n");            
         }   
         sql.append(" join item on item.item_id = inv_dtl.item_nbr      \r\n");
         sql.append(" join item_velocity on item_velocity.velocity_id = item.velocity_id  \r\n");
         sql.append(" join emery_dept on emery_dept.DEPT_NUM = inv_dtl.BUYING_DEPT  \r\n");
         sql.append(" where inv_dtl.invoice_date between to_date(?,'mm/dd/yyyy') and to_date(?,'mm/dd/yyyy') and inv_dtl.sale_type = 'WAREHOUSE'  \r\n");
         sql.append(" group by inv_dtl.buying_dept, emery_dept.name  \r\n");
         sql.append("   \r\n");
         sql.append(" union  \r\n");
         sql.append("   \r\n");
         sql.append(" select   \r\n");
         sql.append("      '1C: Margin $' catto, '00','EMERY' dname, sum(unit_sell * qty_shipped)- sum(unit_cost * qty_shipped) \"Total\",  \r\n");
         sql.append("      sum(case when velocity = 'A' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"A\",  \r\n");
         sql.append("      sum(case when velocity = 'B' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"B\",  \r\n");
         sql.append("      sum(case when velocity = 'C' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"C\",  \r\n");
         sql.append("      sum(case when velocity = 'D' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"D\",  \r\n");
         sql.append("      sum(case when velocity = 'E' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"E\",  \r\n");
         sql.append("      sum(case when velocity = 'I' then (unit_sell * qty_shipped) - (unit_cost * qty_shipped) else 0 end) \"I\"  \r\n");
         sql.append(" from inv_dtl  \r\n");
         sql.append(" join emery_dept on emery_dept.dept_num = buying_dept  \r\n");
         sql.append(" join item on item.item_id = item_nbr  \r\n");
         sql.append(" join item_velocity on item_velocity.velocity_id = item.velocity_id  \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                           
            sql.append("join inv_hdr on inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id  \r\n"); 
            sql.append("and inv_hdr.warehouse in  \r\n");
            sql.append("(select w.name from warehouse w where w.fas_facility_id = ");
            sql.append(m_Warehouse);
            sql.append(") \r\n");            
         }   
         sql.append(" where inv_dtl.invoice_date between to_date(?,'mm/dd/yyyy') and to_date(?,'mm/dd/yyyy') and inv_dtl.sale_type = 'WAREHOUSE'  \r\n");
         sql.append("   \r\n");
         
         sql.append(" union  \r\n");
         sql.append("   \r\n");
         sql.append(" select   \r\n");
         sql.append("     '2F: Open PO $' catto, dept_num Dept, emery_dept.name dname,  \r\n");
         sql.append("      sum((qty_ordered - qty_put_away) * unit_cost) \"Total\",  \r\n");
         sql.append("      sum(case when velocity_cd = 'A' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"A\",  \r\n");
         sql.append("      sum(case when velocity_cd = 'B' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"B\",  \r\n");
         sql.append("      sum(case when velocity_cd = 'C' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"C\",  \r\n");
         sql.append("      sum(case when velocity_cd = 'D' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"D\",  \r\n");
         sql.append("      sum(case when velocity_cd = 'E' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"E\",  \r\n");
         sql.append("     sum(case when velocity_cd = 'I' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"I\"  \r\n");
         sql.append(" from po_hdr  \r\n");
         sql.append(" join po_dtl on po_dtl.po_hdr_id = po_hdr.po_hdr_id  \r\n");
         sql.append(" join vendor_dept on vendor_dept.vendor_id = po_hdr.vendor_id  \r\n");
         sql.append(" join emery_dept on emery_dept.dept_id = vendor_dept.dept_id  \r\n");
         sql.append(" where  \r\n");
         sql.append("      po_hdr.status = 'OPEN' and  \r\n");
         sql.append("      po_hdr.po_nbr not like 'TR%'  \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                                    
            sql.append("     and po_hdr.warehouse = ");
            sql.append(m_Warehouse);
            sql.append("  \r\n");
          }                  
         sql.append(" group by   \r\n");
         sql.append("     dept_num, emery_dept.name  \r\n");
         sql.append("      \r\n");
         sql.append(" union  \r\n");
         sql.append("      \r\n");
         sql.append(" select '1F: Open PO $' catto, '00','EMERY' dname,  \r\n");
         sql.append("     sum((qty_ordered - qty_put_away) * unit_cost) \"Total\",  \r\n");
         sql.append("     sum(case when velocity_cd = 'A' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"A\",  \r\n");
         sql.append("     sum(case when velocity_cd = 'B' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"B\",  \r\n");
         sql.append("     sum(case when velocity_cd = 'C' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"C\",  \r\n");
         sql.append("     sum(case when velocity_cd = 'D' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"D\",  \r\n");
         sql.append("     sum(case when velocity_cd = 'E' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"E\",  \r\n");
         sql.append("     sum(case when velocity_cd = 'I' then (qty_ordered - qty_put_away) * unit_cost else 0 end) \"I\"  \r\n");
         sql.append(" from po_hdr  \r\n");
         sql.append(" join po_dtl on po_dtl.po_hdr_id = po_hdr.po_hdr_id  \r\n");
         sql.append(" join vendor_dept on vendor_dept.vendor_id = po_hdr.vendor_id  \r\n");
         sql.append(" join emery_dept on emery_dept.dept_id = vendor_dept.dept_id  \r\n");
         sql.append(" where  \r\n");
         sql.append("     po_hdr.status = 'OPEN' and  \r\n");
         sql.append("     po_hdr.po_nbr not like 'TR%'   \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                                    
            sql.append("     and po_hdr.warehouse = ");
                   sql.append(m_Warehouse);
            sql.append("  \r\n");
         }          
         sql.append("      \r\n");
         sql.append("    union  \r\n");
         sql.append("      \r\n");
         
         sql.append("     select    \r\n"); 
         sql.append("      '2G: Open PO Units' catto, dept_num Dept, emery_dept.name dname,   \r\n"); 
         sql.append("       sum((qty_ordered - qty_put_away) ) \"Total\",  \r\n");  
         sql.append("       sum(case when velocity_cd = 'A' then (qty_ordered - qty_put_away) else 0 end) \"A\",  \r\n");  
         sql.append("       sum(case when velocity_cd = 'B' then (qty_ordered - qty_put_away) else 0 end) \"B\",  \r\n");  
         sql.append("       sum(case when velocity_cd = 'C' then (qty_ordered - qty_put_away) else 0 end) \"C\",   \r\n"); 
         sql.append("       sum(case when velocity_cd = 'D' then (qty_ordered - qty_put_away) else 0 end) \"D\",   \r\n"); 
         sql.append("       sum(case when velocity_cd = 'E' then (qty_ordered - qty_put_away)  else 0 end) \"E\",  \r\n");  
         sql.append("      sum(case when velocity_cd = 'I' then (qty_ordered - qty_put_away) else 0 end) \"I\"   \r\n"); 
         sql.append("  from po_hdr  \r\n");  
         sql.append("  join po_dtl on po_dtl.po_hdr_id = po_hdr.po_hdr_id   \r\n"); 
         sql.append("  join vendor_dept on vendor_dept.vendor_id = po_hdr.vendor_id  \r\n");  
         sql.append("  join emery_dept on emery_dept.dept_id = vendor_dept.dept_id   \r\n"); 
         sql.append("  where  \r\n");  
         sql.append("       po_hdr.status = 'OPEN' and  \r\n");  
         sql.append("       po_hdr.po_nbr not like 'TR%'  \r\n");  
         sql.append("  group by   \r\n");  
         sql.append("      dept_num, emery_dept.name  \r\n");  
         sql.append("   \r\n");    
         sql.append("     union  \r\n");
         sql.append("    \r\n");   
         sql.append("  select '1G: Open PO Units' catto, '00','EMERY' dname,   \r\n"); 
         sql.append("      sum((qty_ordered - qty_put_away) ) \"Total\",  \r\n");  
         sql.append("      sum(case when velocity_cd = 'A' then (qty_ordered - qty_put_away)  else 0 end) \"A\",   \r\n"); 
         sql.append("      sum(case when velocity_cd = 'B' then (qty_ordered - qty_put_away)  else 0 end) \"B\",   \r\n"); 
         sql.append("      sum(case when velocity_cd = 'C' then (qty_ordered - qty_put_away)  else 0 end) \"C\",  \r\n");  
         sql.append("      sum(case when velocity_cd = 'D' then (qty_ordered - qty_put_away) else 0 end) \"D\",   \r\n"); 
         sql.append("      sum(case when velocity_cd = 'E' then (qty_ordered - qty_put_away) else 0 end) \"E\",   \r\n"); 
         sql.append("      sum(case when velocity_cd = 'I' then (qty_ordered - qty_put_away)  else 0 end) \"I\"   \r\n"); 
         sql.append("  from po_hdr  \r\n");  
         sql.append("  join po_dtl on po_dtl.po_hdr_id = po_hdr.po_hdr_id    \r\n");
         sql.append("  join vendor_dept on vendor_dept.vendor_id = po_hdr.vendor_id  \r\n");  
         sql.append("  join emery_dept on emery_dept.dept_id = vendor_dept.dept_id   \r\n"); 
         sql.append("  where   \r\n"); 
         sql.append("      po_hdr.status = 'OPEN' and   \r\n"); 
         sql.append("      po_hdr.po_nbr not like 'TR%'   \r\n");      
         
 
         sql.append("      \r\n");
         sql.append("    union  \r\n");
         sql.append("      \r\n");
         
         
         sql.append("  select '2D: Inventory $' catto, dept_num Dept, emery_dept.name DName,   \r\n");
         sql.append("     sum(total_cost) \"Total\",  \r\n");
         sql.append("     sum(case when velocity = 'A' then total_cost else 0 end) \"A\",  \r\n");
         sql.append("     sum(case when velocity = 'B' then total_cost else 0 end) \"B\",  \r\n");
         sql.append("     sum(case when velocity = 'C' then total_cost else 0 end) \"C\",  \r\n");
         sql.append("     sum(case when velocity = 'D' then total_cost else 0 end) \"D\",  \r\n");
         sql.append("     sum(case when velocity = 'E' then total_cost else 0 end) \"E\",  \r\n");
         sql.append("     sum(case when velocity = 'I' then total_cost else 0 end) \"I\"  \r\n");
         sql.append(" from item  \r\n");
         sql.append(" join emery_dept on emery_dept.dept_id = item.dept_id  \r\n");
         sql.append(" join item_velocity on item_velocity.velocity_id = item.velocity_id  \r\n");
         sql.append(" join item_inventory ii on ii.item_id = item.item_id and   \r\n");
         sql.append("    inventory_date = (  \r\n");
         sql.append("          select max(inventory_date)   \r\n");
         sql.append("          from item_inventory inv  \r\n");
         sql.append("          where   \r\n");
         sql.append("             inv.warehouse_id = ii.warehouse_id and  \r\n");
         sql.append("             inv.item_id = ii.item_id and  \r\n");
         sql.append("             inventory_date <= to_date(?,'mm/dd/yyyy')  \r\n");
         sql.append("    )   \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                           
            sql.append(" join warehouse w on ii.warehouse_id = w.warehouse_id and w.fas_facility_id = \r\n");
            sql.append(m_Warehouse);
            sql.append(" \r\n");            
         }   
         sql.append(" group by dept_num, emery_dept.name  \r\n");
         sql.append(" union  \r\n");
         sql.append("  select '1D: Inventory $' catto, '00' Dept, 'EMERY' DName,   \r\n");
         sql.append("     sum(total_cost) \"Total\",  \r\n");
         sql.append("     sum(case when velocity = 'A' then total_cost else 0 end) \"A\",  \r\n");
         sql.append("     sum(case when velocity = 'B' then total_cost else 0 end) \"B\",  \r\n");
         sql.append("     sum(case when velocity = 'C' then total_cost else 0 end) \"C\",  \r\n");
         sql.append("     sum(case when velocity = 'D' then total_cost else 0 end) \"D\",  \r\n");
         sql.append("     sum(case when velocity = 'E' then total_cost else 0 end) \"E\",  \r\n");
         sql.append("     sum(case when velocity = 'I' then total_cost else 0 end) \"I\"  \r\n");
         sql.append(" from item  \r\n");
         sql.append(" join emery_dept on emery_dept.dept_id = item.dept_id  \r\n");
         sql.append(" join item_velocity on item_velocity.velocity_id = item.velocity_id  \r\n");
         sql.append(" join item_inventory ii on ii.item_id = item.item_id and   \r\n");
         sql.append("      inventory_date = (  \r\n");
         sql.append("         select max(inventory_date)   \r\n");
         sql.append("         from item_inventory inv  \r\n");
         sql.append("         where   \r\n");
         sql.append("            inv.warehouse_id = ii.warehouse_id and  \r\n");
         sql.append("            inv.item_id = ii.item_id and  \r\n");
         sql.append("            inventory_date <= to_date(?,'mm/dd/yyyy')  \r\n");
         sql.append("      )   \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                           
            sql.append(" join warehouse w on ii.warehouse_id = w.warehouse_id and w.fas_facility_id = \r\n");
            sql.append(m_Warehouse);
            sql.append(" \r\n");            
         }   
         sql.append(" union  \r\n");
         sql.append(" select '2E: Inventory Units' catto, dept_num Dept, emery_dept.name Dname,   \r\n");
         sql.append("     sum(qty_on_hand) \"Total\",  \r\n");
         sql.append("     sum(case when velocity = 'A' then qty_on_hand else 0 end) \"A\",  \r\n");
         sql.append("     sum(case when velocity = 'B' then qty_on_hand else 0 end) \"B\",  \r\n");
         sql.append("     sum(case when velocity = 'C' then qty_on_hand else 0 end) \"C\",  \r\n");
         sql.append("     sum(case when velocity = 'D' then qty_on_hand else 0 end) \"D\",  \r\n");
         sql.append("     sum(case when velocity = 'E' then qty_on_hand else 0 end) \"E\",  \r\n");
         sql.append("     sum(case when velocity = 'I' then qty_on_hand else 0 end) \"I\"  \r\n");
         sql.append(" from item  \r\n");
         sql.append(" join emery_dept on emery_dept.dept_id = item.dept_id  \r\n");
         sql.append(" join item_velocity on item_velocity.velocity_id = item.velocity_id  \r\n");
         sql.append(" join item_inventory ii on ii.item_id = item.item_id and   \r\n");
         sql.append("     inventory_date = (  \r\n");
         sql.append("        select max(inventory_date)   \r\n");
         sql.append("        from item_inventory inv  \r\n");
         sql.append("        where   \r\n");
         sql.append("           inv.warehouse_id = ii.warehouse_id and  \r\n");
         sql.append("           inv.item_id = ii.item_id and  \r\n");
         sql.append("           inventory_date <= to_date(?,'mm/dd/yyyy')  \r\n");
         sql.append("    )   \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                  
            sql.append(" join warehouse w on ii.warehouse_id = w.warehouse_id and w.fas_facility_id = \r\n");
            sql.append(m_Warehouse);
            sql.append(" \r\n");            
         }   
         sql.append(" group by dept_num, emery_dept.name  \r\n");
         sql.append(" union  \r\n");
         sql.append("  select '1E: Inventory Units' catto, '00' Dept, 'EMERY' Dname,   \r\n");
         sql.append("     sum(qty_on_hand) \"Total\",  \r\n");
         sql.append("     sum(case when velocity = 'A' then qty_on_hand else 0 end) \"A\",  \r\n");
         sql.append("     sum(case when velocity = 'B' then qty_on_hand else 0 end) \"B\",  \r\n");
         sql.append("     sum(case when velocity = 'C' then qty_on_hand else 0 end) \"C\",  \r\n");
         sql.append("     sum(case when velocity = 'D' then qty_on_hand else 0 end) \"D\",  \r\n");
         sql.append("     sum(case when velocity = 'E' then qty_on_hand else 0 end) \"E\",  \r\n");
         sql.append("     sum(case when velocity = 'I' then qty_on_hand else 0 end) \"I\"  \r\n");
         sql.append(" from item  \r\n");
         sql.append(" join emery_dept on emery_dept.dept_id = item.dept_id  \r\n");
         sql.append(" join item_velocity on item_velocity.velocity_id = item.velocity_id  \r\n");
         sql.append(" join item_inventory ii on ii.item_id = item.item_id and   \r\n");
         sql.append("    ii.inventory_date = (  \r\n");
         sql.append("        select max(inventory_date)   \r\n");
         sql.append("        from item_inventory inv  \r\n");
         sql.append("        where   \r\n");
         sql.append("           inv.warehouse_id = ii.warehouse_id and  \r\n");
         sql.append("           inv.item_id = ii.item_id and  \r\n");
         sql.append("           inv.inventory_date <= to_date(?,'mm/dd/yyyy')  \r\n");
         sql.append("    )   \r\n");
         if ( m_Warehouse != null && m_Warehouse.length() > 0 ){                           
            sql.append(" join warehouse w on ii.warehouse_id = w.warehouse_id and w.fas_facility_id = \r\n");
            sql.append(m_Warehouse);
            sql.append(" \r\n");            
         }
         sql.append("   \r\n");
         sql.append(" order by dept, catto  \r\n");
        // left in for tracking of query only -- see just above start of query build
        /// squalid.append(sql.substring(1000));
         m_Velocity = m_OraConn.prepareStatement(sql.toString());
         
         return true;
      }
   	
   	catch ( Exception e ) {
   		log.error("exception", e);
   		return false;
   	}
   	
   	finally {
   		sql = null;
   	}
   }
   
/* if we don't get a date range use a default
 * Default is Sunday through Saturday of the week preceding the date of the request
 */
   public void setDefaultDates()
   {
      StringBuffer sqlthis = new StringBuffer();
      ResultSet rs2 = null;
      
      try {
         sqlthis.setLength(0);
         sqlthis.append("select to_char((trunc(sysdate - 7, 'DAY')),'mm/dd/yyyy') sun_start, ");         
         sqlthis.append("to_char((trunc(sysdate - 7, 'DAY')+6),'mm/dd/yyyy') sat_end  from dual");         
         m_DefaultDates = m_OraConn.prepareStatement(sqlthis.toString()); 
         
         rs2 = m_DefaultDates.executeQuery();

         while ( rs2.next() && m_Status == RptServer.RUNNING ) {
         m_BegDate = rs2.getString("sun_start");            
         m_EndDate = rs2.getString("sat_end");                
         }
      }
      catch( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("The report had the following Error: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
     
      }

      finally {
         DbUtils.closeDbConn(null, null, rs2);
      }     
   }

   /**
    * Sets the parameters for the report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
    {
      int pcount = params.size();
      Param param = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);                            
         if ( param.name.equals("dc") )
             m_Warehouse = param.value;         
         if ( param.name.equals("begdate") )
             m_BegDate = param.value;
         if ( param.name.equals("enddate") )
            m_EndDate = param.value;        
      }
    }  
}

