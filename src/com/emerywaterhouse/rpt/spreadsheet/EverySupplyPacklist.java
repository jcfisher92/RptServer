package com.emerywaterhouse.rpt.spreadsheet;

/**
 * Title:         EverySupplyPacklist.java
 * Description:   Generates a spreadsheet packlist for Every Supply, because they order 1 item per invoice and end up with
 *                dozens and dozens of invoices each week, which is too large for the Crystal Report which promptly explodes.
 *                Almost every Wednesday, Every Supply needs this packlist generated, which is why this is hardcoded.
 *                Could be expanded out to be more generic (passing in customer id, trip / stop ids, ...) if we ever need to do this
 *                for another customer in the future, but highly unlikely.
 *                
 *                Because there may be multple shipments on a given trip, an input date will be taken instead of a shipment id.
 *                
 * Company:       Emery-Waterhouse
 * @author        Stephen Martel
 */

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class EverySupplyPacklist extends Report {

   // Workbook and style stuff
   private XSSFWorkbook m_WrkBk;
   private XSSFSheet m_SummarySheet;
   private XSSFSheet m_InvoiceSheet;
   private XSSFSheet m_DetailSheet;
   private XSSFRow m_Row = null;

   private XSSFFont m_FontData;
   private XSSFFont m_FontDataSmall;
   private XSSFFont m_FontBold;
   private XSSFFont m_FontBoldSmall;
   private XSSFCellStyle m_StyleText;
   private XSSFCellStyle m_StyleTextSmall;
   private XSSFCellStyle m_StyleBold;
   private XSSFCellStyle m_StyleBoldSmall;
   
   // Query variables:   
   private PreparedStatement m_GetSummary; // tab 1 query
   private PreparedStatement m_GetInvoices; // tab 2 query
   private PreparedStatement m_GetDetails; // tab 3 query
   private String m_TripDate;
   private short m_SummRowNum = 0;
   private short m_InvRowNum = 0;
   private short m_DtlRowNum = 0;

   // Finals for retrieving fields from the ResultSets
   // Tab 1:
   private final String SHIP_IDS = "SHIP_IDS";
   private final String CUSTOMER_NAME = "CUSTOMER_NAME";
   private final String SHIP_DATE = "SHIP_DATE";
   private final String CUSTOMER_ID = "CUSTOMER_ID";
   private final String ROUTE = "ROUTE";
   private final String STOP = "STOP";
   private final String CARRIER_NAME = "CARRIER_NAME";
   // Tab 2:
   private final String ORDER_ID = "ORDER_ID";
   private final String INVOICE_NUM = "INVOICE_NUM";
   private final String PO_NUM = "PO_NUM";
   // Tab 3:
   private final String MU_TYPE = "MU_TYPE";
   private final String MU_ID = "MU_ID";
   private final String ITEM_ID = "ITEM_ID";
   private final String DESCRIPTION = "DESCRIPTION";
   private final String QTY_ORDERED = "QTY_ORDERED";
   private final String UNIT = "UNIT";
   private final String MU_QTY = "MU_QTY";
   private final String QTY_SHIPPED = "QTY_SHIPPED";
   private final String RETAIL_PRICE = "RETAIL_PRICE";
   private final String UPC_CODE = "UPC_CODE";
   private final String SELL_PRICE = "SELL_PRICE";

   /**
    * default constructor
    */
   public EverySupplyPacklist()
   {
      super();
     
      m_MaxRunTime = RptServer.HOUR * 2;
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    */
   @Override
   public void finalize() throws Throwable
   {      
   	m_GetSummary = null;
   	m_GetInvoices = null;
   	m_GetDetails = null;
      m_TripDate = null;
   	
      super.finalize();
   }
   
   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_GetSummary);
      closeStmt(m_GetInvoices);
      closeStmt(m_GetDetails);
   }
   
   /**
    * Creates the report file.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
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
         log.fatal("[EverySupplyPacklist#createReport] ", ex);
      }

      finally {
        closeStatements();

        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }

      return created;
   }
   
   
   private boolean prepareStatements() throws SQLException
   {
      StringBuffer sql = new StringBuffer();
      
      if ( m_OraConn == null )
         return false;
      
      sql.setLength(0);
      sql.append("select wm_concat(packlist.ship_id) as ship_ids, ");
      sql.append("	packlist.customer_name, ");
      sql.append("	packlist.ship_date, ");
      sql.append("	packlist.customer_id, ");
      sql.append("	packlist.route, ");
      sql.append("	packlist.stop, ");
      sql.append("	packlist.carrier_name ");
      sql.append("from packlist ");
      sql.append("join shipment on shipment.ship_id = packlist.ship_id ");
      sql.append("where ship_date = ? and ");
      sql.append("	   packlist.customer_id = 424404 ");
      sql.append("group by packlist.customer_name, packlist.ship_date, packlist.customer_id, packlist.route, packlist.stop, packlist.carrier_name ");
      m_GetSummary = m_OraConn.prepareStatement(sql.toString());
      
      sql.setLength(0);
      sql.append("select distinct order_line.order_id, order_line.invoice_num, po_num from order_line ");        
      sql.append("join order_header on order_line.order_id = order_header.order_id ");   
      sql.append("where order_line.invoice_num in ( ");   
      sql.append("   select invoice_num from packlist_invoice where ship_id in ( ");
      sql.append("      select ship_id from packlist where customer_id = 424404 and ship_date = ? ");  
      sql.append("   ) ");
      sql.append(") ");   
      sql.append("order by order_line.invoice_num ");
      m_GetInvoices = m_OraConn.prepareStatement(sql.toString());

      sql.setLength(0);
      sql.append("select mu_type, mu_id, item_id, description, qty_ordered, unit, mu_qty, qty_shipped, retail_price, upc_code, sell_price ");
      sql.append("from packlist_detail ");
      sql.append("where ship_id in ( ");
      sql.append("      select ship_id from packlist where customer_id = 424404 and ship_date = ? ");  
      sql.append(") order by item_id ");
      m_GetDetails = m_OraConn.prepareStatement(sql.toString());
      
      return true;
   }
   
   /**
    * Executes the queries and builds the output file
    * 
    * @return true if the report was successfully built
    * @throws FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      SimpleDateFormat dateFormat = new SimpleDateFormat("M/dd/yyyy");
      Row row = null;
      int colNum = 0;
      FileOutputStream outFile = null;
      ResultSet summaryData = null;
      ResultSet invoiceData = null;
      ResultSet detailData = null;
      boolean result = false;
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      initReport();
      
      try {
      	m_GetSummary.setDate(1,new java.sql.Date(dateFormat.parse(m_TripDate).getTime()));   
         summaryData = m_GetSummary.executeQuery();
         
         m_CurAction = "Building output file";
         XSSFCellStyle currentStyle = m_StyleText;
         
         if ( summaryData.next() && getStatus() != RptServer.STOPPED ) { // Should only be 1 record returned
            row = m_SummarySheet.createRow(m_SummRowNum++);

            // Initialization and data population at the same time because of its presentation format.
            
            createCell(row, 0, "Shipment#:", m_StyleBold);
            createCell(row, 1, summaryData.getString(SHIP_IDS), currentStyle);
            row = m_SummarySheet.createRow(m_SummRowNum++);
            
            createCell(row, 0, "Sold to:", m_StyleBold);
            createCell(row, 1, summaryData.getString(CUSTOMER_NAME), currentStyle);
            row = m_SummarySheet.createRow(m_SummRowNum++);

            createCell(row, 0, "Ship Date:", m_StyleBold);
            createCell(row, 1, summaryData.getString(SHIP_DATE), currentStyle);
            row = m_SummarySheet.createRow(m_SummRowNum++);

            createCell(row, 0, "Customer:", m_StyleBold);
            createCell(row, 1, summaryData.getString(CUSTOMER_ID), currentStyle);
            row = m_SummarySheet.createRow(m_SummRowNum++);

            createCell(row, 0, "Route:", m_StyleBold);
            createCell(row, 1, summaryData.getString(ROUTE), currentStyle);
            row = m_SummarySheet.createRow(m_SummRowNum++);

            createCell(row, 0, "Stop:", m_StyleBold);
            createCell(row, 1, summaryData.getString(STOP), currentStyle);
            row = m_SummarySheet.createRow(m_SummRowNum++);

            createCell(row, 0, "Carrier:", m_StyleBold);
            createCell(row, 1, summaryData.getString(CARRIER_NAME), currentStyle);
            row = m_SummarySheet.createRow(m_SummRowNum++);

         } else {
            log.fatal("[EverySupplyPacklist#buildOutputFile] - No shipment found for date " + m_TripDate);
         	return false;
         }
         
         // Tab 2 - Invoices / POs

      	m_GetInvoices.setDate(1,new java.sql.Date(dateFormat.parse(m_TripDate).getTime()));
      	invoiceData = m_GetInvoices.executeQuery();
         while ( invoiceData.next() && getStatus() != RptServer.STOPPED ) {
            row = m_InvoiceSheet.createRow(m_InvRowNum++);
            colNum = 0;
            
            createCell(row, colNum++, invoiceData.getString(ORDER_ID), currentStyle);
            createCell(row, colNum++, invoiceData.getString(INVOICE_NUM), currentStyle);
            createCell(row, colNum++, invoiceData.getString(PO_NUM), currentStyle);

         }
         
         
         // Tab 3 - Details

      	m_GetDetails.setDate(1,new java.sql.Date(dateFormat.parse(m_TripDate).getTime()));   
      	detailData = m_GetDetails.executeQuery();
         while ( detailData.next() && getStatus() != RptServer.STOPPED ) {
            row = m_DetailSheet.createRow(m_DtlRowNum++);
            colNum = 0;
            
            createCell(row, colNum++, detailData.getString(MU_TYPE), m_StyleTextSmall);
            createCell(row, colNum++, detailData.getString(MU_ID), m_StyleTextSmall);
            createCell(row, colNum++, detailData.getString(ITEM_ID), m_StyleTextSmall);
            createCell(row, colNum++, detailData.getString(DESCRIPTION), m_StyleTextSmall);
            createCell(row, colNum++, detailData.getString(QTY_ORDERED), m_StyleTextSmall);
            createCell(row, colNum++, detailData.getString(UNIT), m_StyleTextSmall);
            createCell(row, colNum++, detailData.getString(MU_QTY), m_StyleTextSmall);
            createCell(row, colNum++, detailData.getString(QTY_SHIPPED), m_StyleTextSmall);
            createCell(row, colNum++, detailData.getString(RETAIL_PRICE), m_StyleTextSmall);
            createCell(row, colNum++, detailData.getString(UPC_CODE), m_StyleTextSmall);
            createCell(row, colNum++, detailData.getString(SELL_PRICE), m_StyleTextSmall);

         }
         
         m_WrkBk.write(outFile);
         
         result = true;
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[EverySupplyPacklist#buildOutputFile] ", ex);
      }

      finally {         

         closeRSet(summaryData);
         closeRSet(invoiceData);
         closeRSet(detailData);
         
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[EverySupplyPacklist#buildOutputFile] " , e);
         }

         outFile = null;
      }

      return result;
   }
   
   
   /**
    * Creates a cell of type String
    * @param row Spreadsheet row cell is being added to
    * @param col Column of cell
    * @param val numeric value of cell
    * @return HSSFCell newly created cell
    */
   private Cell createCell(Row row, int col, String val, CellStyle style)
   {
      Cell cell = null;

      cell = row.createCell(col);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(new XSSFRichTextString(val));
      cell.setCellStyle(style);

      return cell;
   }
   
   
   /**
    * Creates the workbook and worksheet.  Creates any fonts and styles that
    * will be used.
    */
   private void initReport()
   {
      short col = 0;
      m_SummRowNum = 0;
      m_InvRowNum = 0;
      m_DtlRowNum = 0;

      try {
         m_WrkBk = new XSSFWorkbook();
         
         //
         // Create a font that is normal size
         m_FontData = m_WrkBk.createFont();
         m_FontData.setFontHeightInPoints((short)11);
         m_FontData.setFontName("Calibri");

         // "" but smaller
         m_FontDataSmall = m_WrkBk.createFont();
         m_FontDataSmall .setFontHeightInPoints((short)8);
         m_FontDataSmall.setFontName("Calibri");

         //
         // Create a font that is normal size & bold
         m_FontBold = m_WrkBk.createFont();
         m_FontBold.setFontHeightInPoints((short)11);
         m_FontBold.setFontName("Calibri");
         m_FontBold.setBold(true);
         
         // "" but smaller
         m_FontBoldSmall = m_WrkBk.createFont();
         m_FontBoldSmall.setFontHeightInPoints((short)8);
         m_FontBoldSmall.setFontName("Calibri");
         m_FontBoldSmall.setBold(true);
         
         //
         // Setup the cell styles used in this report
         m_StyleText = m_WrkBk.createCellStyle();
         m_StyleText.setFont(m_FontData);
         m_StyleText.setAlignment(HorizontalAlignment.LEFT);

         m_StyleTextSmall = m_WrkBk.createCellStyle();
         m_StyleTextSmall.setFont(m_FontDataSmall);
         m_StyleTextSmall.setAlignment(HorizontalAlignment.LEFT);
         
         // Style 11pt, left aligned, bold
         m_StyleBold = m_WrkBk.createCellStyle();
         m_StyleBold.setFont(m_FontBold);
         m_StyleBold.setAlignment(HorizontalAlignment.LEFT);

         m_StyleBoldSmall = m_WrkBk.createCellStyle();
         m_StyleBoldSmall.setFont(m_FontBoldSmall);
         m_StyleBoldSmall.setAlignment(HorizontalAlignment.LEFT);
         
         m_SummarySheet = m_WrkBk.createSheet("Summary");
         m_SummarySheet.setMargin(XSSFSheet.BottomMargin, .25);
         m_SummarySheet.getPrintSetup().setLandscape(true);
         m_SummarySheet.getPrintSetup().setPaperSize((short)5);

         m_InvoiceSheet = m_WrkBk.createSheet("Invoices");
         m_InvoiceSheet.setMargin(XSSFSheet.BottomMargin, .25);
         m_InvoiceSheet.getPrintSetup().setLandscape(true);
         m_InvoiceSheet.getPrintSetup().setPaperSize((short)5);
         
         m_DetailSheet = m_WrkBk.createSheet("Details");
         m_DetailSheet.setMargin(XSSFSheet.BottomMargin, .25);
         m_DetailSheet.getPrintSetup().setLandscape(true);
         m_DetailSheet.getPrintSetup().setPaperSize((short)5);

         
         // Create the column headings

         // FIRST PAGE - THE SUMMARY
         // The first page is a bit backwards - the way Cindy uses it, the rows and the columns are switched, ie, column 0 has all the column names, and column 1 has the values.
         // So, we'll wait to create the field headers until we're actually populating it with data.
         col = 0;
         m_Row = m_SummarySheet.createRow(m_SummRowNum);
         
         m_SummarySheet.setColumnWidth(col, 2844);
         m_SummarySheet.setColumnWidth(col+1, 4925);
         
         
         
         // SECOND PAGE - POs / INVOICES
         col = 0;
         m_Row = m_InvoiceSheet.createRow(m_InvRowNum);
         
         m_InvoiceSheet.setColumnWidth(col, 2300);
         createCell(m_Row, col++, "Order ID:", m_StyleBold);
         m_InvoiceSheet.setColumnWidth(col, 3193);
         createCell(m_Row, col++, "Invoice Num:", m_StyleBold);
         m_InvoiceSheet.setColumnWidth(col, 6352);
         createCell(m_Row, col++, "Customer PO References:", m_StyleBold);
         m_InvRowNum++;
         

         // THIRD PAGE - DETAILS
         col = 0;
         m_Row = m_DetailSheet.createRow(m_DtlRowNum);
         
         m_DetailSheet.setColumnWidth(col, 1686);
         createCell(m_Row, col++, "MU Type:", m_StyleBoldSmall);
         m_DetailSheet.setColumnWidth(col, 1997);
         createCell(m_Row, col++, "MU#:", m_StyleBoldSmall);
         m_DetailSheet.setColumnWidth(col, 1765);
         createCell(m_Row, col++, "Emery#:", m_StyleBoldSmall);
         m_DetailSheet.setColumnWidth(col, 16295);
         createCell(m_Row, col++, "Item Description:", m_StyleBoldSmall);
         m_DetailSheet.setColumnWidth(col, 1492);
         createCell(m_Row, col++, "Ord Qty:", m_StyleBoldSmall);
         m_DetailSheet.setColumnWidth(col, 985);
         createCell(m_Row, col++, "UoM:", m_StyleBoldSmall);
         m_DetailSheet.setColumnWidth(col, 1451);
         createCell(m_Row, col++, "MU Qty:", m_StyleBoldSmall);
         m_DetailSheet.setColumnWidth(col, 1607);
         createCell(m_Row, col++, "Qty Ship:", m_StyleBoldSmall);
         m_DetailSheet.setColumnWidth(col, 2152);
         createCell(m_Row, col++, "Sugg. Retail:", m_StyleBoldSmall);
         m_DetailSheet.setColumnWidth(col, 2848);
         createCell(m_Row, col++, "Retail UPC:", m_StyleBoldSmall);
         m_DetailSheet.setColumnWidth(col, 1765);
         createCell(m_Row, col++, "Cost:", m_StyleBoldSmall);
         m_DtlRowNum++;
         
      }

      catch ( Exception e ) {
         log.error("[EverySupplyPacklist#initReport] ", e );
      }
   }

   /**
    * Sets the parameters for the report.
    *    param(0) = date of the EverySupply trip with the packlist you want to create.
    *    
    * @param params ArrayList<Param> - list of report parameters.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fname = new StringBuffer();
      int pcount = params.size();                     
      Param param = null;          
      SimpleDateFormat dateFormat = new SimpleDateFormat("M/dd/yyyy");                   
      SimpleDateFormat formatter = new SimpleDateFormat ("yyyyMMddHHmmss");
      Date day = new Date();
       
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
         if ( param.name.equals("tripdate") ) {
         	m_TripDate = param.value;
         }
      }
      
      if (m_TripDate == null || (m_TripDate.length() == 0)) {
          // No date supplied - use today as the default date.
      	m_TripDate = dateFormat.format(day);
      } else {
	      try {
	      	Date tripDate = new SimpleDateFormat("M/dd/yyyy").parse(m_TripDate);
	      	m_TripDate = dateFormat.format(tripDate);
	      } catch (Exception e) {
	         // Invalid date supplied - use today as the default date? TODO
	      	m_TripDate = dateFormat.format(day);
	      }
      }
      
      //
      // Build the file name.
      fname.append(formatter.format( day ));
      fname.append("-");
      fname.append("EverySupplyPacklist.xlsx");
      m_FileNames.add(fname.toString());
   }
}
