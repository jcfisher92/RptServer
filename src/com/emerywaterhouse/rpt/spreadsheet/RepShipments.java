package com.emerywaterhouse.rpt.spreadsheet;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Header;
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

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class RepShipments extends Report
{
   //
   private String m_LoadDate;
   private String m_Rep;
   private String m_RepNoSpaces;
   private String m_RptDate;
   private short m_RowNum = 0;

   private PreparedStatement m_RepShipments;
   private PreparedStatement m_RepSales;

   // workbook entries.
   private XSSFWorkbook m_WrkBk;
   private XSSFSheet m_Sheet;
   private XSSFRow m_Row = null;
   private Header m_Header;

   private XSSFFont m_Font;
   private XSSFFont m_FontTitle;
   private XSSFFont m_FontBold;
   private XSSFFont m_FontData;

   private XSSFCellStyle m_StyleText;        // Text left justified
   private XSSFCellStyle m_StyleTextRight;   // Text right justified
   private XSSFCellStyle m_StyleTextCenter;  // Text centered
   private XSSFCellStyle m_StyleTitle;       // Bold, centered
   private XSSFCellStyle m_StyleBold;        // Normal but bold
   private XSSFCellStyle m_StyleBoldRight;   // Normal but bold & right aligned
   private XSSFCellStyle m_StyleBoldCenter;  // Normal but bold & centered
   private XSSFCellStyle m_StyleDec;         // Style with 2 decimals
   private XSSFCellStyle m_StyleDecBold;     // Style with 2 decimals, bold
   private XSSFCellStyle m_StyleHeader;      // Bold, centered 12pt
   private XSSFCellStyle m_StyleInt;         // Style with 0 decimals

   //
   // Log4j logger
   private final Logger m_Log;


   /**
    * default constructor
    */
   public RepShipments()
   {
      super();
      m_Log = Logger.getLogger(RptServer.class);
      m_WrkBk = new XSSFWorkbook();
      m_Sheet = m_WrkBk.createSheet();
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {

      m_Sheet = null;
      m_WrkBk = null;

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
      FileOutputStream outFile = null;
      ResultSet RepShipments = null;
      ResultSet RepSales = null;
      boolean result = false;
      String LastCust = "begin";
      double CustSalesTotal = 0;
      double TotalSales = 0;
      double TotalCredits = 0;
      double LineTotal = 0;
      double LostSales = 0;
      double TotalLostSales = 0;
      double CustLostSalesTotal = 0;
      double UnitSell = 0;
      int LostQty = 0;
      int col;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      initReport();

      try {
         //
         // Note - the load date is set to the report date if it wasn't sent in.  We can assume
         // that if a report date was passed the load date is set as well.
         if (m_RptDate != null && m_RptDate.length() > 0) {
            m_RepShipments.setString(1,m_RepNoSpaces);
            m_RepShipments.setString(2,m_RptDate);
            m_RepShipments.setString(3,m_LoadDate);
            m_RepShipments.setString(4,m_RepNoSpaces);
         } else {
            m_RepShipments.setString(1,m_RepNoSpaces);
            m_RepShipments.setString(2,m_RepNoSpaces);
         }

         RepShipments = m_RepShipments.executeQuery();

         while ( RepShipments.next() && m_Status == RptServer.RUNNING ) {
            if ( !RepShipments.getString("Cust_nbr").equals(LastCust) ) {

               if ( !LastCust.equals("begin"))
                  custTrailer(LastCust, CustSalesTotal, CustLostSalesTotal);

               CustSalesTotal = 0;
               CustLostSalesTotal = 0;
            }
            
            LineTotal = RepShipments.getDouble("Ext_Price");
            if ( LineTotal > 0 )
               TotalSales += LineTotal;
            else
               TotalCredits += LineTotal;
            
            UnitSell = RepShipments.getDouble("Unit_Sell");
            LostQty = RepShipments.getInt("Lost_Qty");
            
            if ( LostQty > 0 ) {
               LostSales = LostQty * UnitSell;
               TotalLostSales += LostSales;
            }               

            m_Row = m_Sheet.createRow(m_RowNum++);
            col = (short)0;
            createCell(m_Row, col++, RepShipments.getString("Sales_Rep"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getString("Cust_Nbr"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getString("Customer_Name"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getString("Invoice_Nbr"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getString("Cust_PO_Nbr"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getString("Item_id"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getString("Item_Description"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getString("vendor_name"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getString("UOM"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getInt("Qty_Ord"), m_StyleInt);
            createCell(m_Row, col++, RepShipments.getInt("Qty_Ship"), m_StyleInt);
            createCell(m_Row, col++, LostQty, m_StyleInt);
            createCell(m_Row, col++, RepShipments.getString("Comments"), m_StyleText);
            createCell(m_Row, col++, UnitSell, m_StyleDec);
            createCell(m_Row, col++, RepShipments.getDouble("Reg_Sell"), m_StyleDec);
            createCell(m_Row, col++, RepShipments.getDouble("Mgn_Pct"), m_StyleDec);
            createCell(m_Row, col++, LineTotal, m_StyleDec);
            createCell(m_Row, col++, LostSales, m_StyleDec);
            createCell(m_Row, col++, RepShipments.getString("Price_Method"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getString("Best_Price"), m_StyleText);
            createCell(m_Row, col++, RepShipments.getString("Brk_Pck_Chg"), m_StyleText);

            LastCust = RepShipments.getString("Cust_Nbr");
            
            CustSalesTotal += LineTotal;
            CustLostSalesTotal += LostSales;
            LostSales = 0;
            
         }

         custTrailer(LastCust, CustSalesTotal, CustLostSalesTotal);

         if ( m_RptDate != null && m_RptDate.length() > 0 )
            rptTrailer(m_RptDate, TotalSales, TotalCredits, TotalLostSales);
         else
            rptTrailer(getDateYesterday(), TotalSales, TotalCredits, TotalLostSales);

         //and now for the MTD and YTD sales stuff
         try {
            m_RepSales.setString(1,m_Rep);
            RepSales = m_RepSales.executeQuery();

            m_Row = m_Sheet.createRow(m_RowNum++);
            col = (short)1;
            createCell(m_Row, col++, "Preliminary Rolling Sales Totals for " + m_Rep, m_StyleBold);
            m_Row = m_Sheet.createRow(m_RowNum++);
            createCell(m_Row, col, "by invoice date, including credits, later dates subject to revision", m_StyleDec);

            m_Row = m_Sheet.createRow(m_RowNum++);
            m_Row = m_Sheet.createRow(m_RowNum++);
            col = (short)1;

            createCell(m_Row, col++, "Period", m_StyleBold);
            createCell(m_Row, col++, "Sales", m_StyleBold);
            createCell(m_Row, col++, "MaxDate", m_StyleBold);

            while ( RepSales.next() && m_Status == RptServer.RUNNING ) {
               m_Row = m_Sheet.createRow(m_RowNum++);
               col = (short)1;
               createCell(m_Row, col++, RepSales.getString("YTD_Per"), m_StyleText);
               createCell(m_Row, col++, RepSales.getDouble("YTD"), m_StyleDec);
               createCell(m_Row, col++, RepSales.getString("YTD_Date"), m_StyleText);

               m_Row = m_Sheet.createRow(m_RowNum++);
               col = (short)1;
               if (RepSales.getString("LYTD_Per") != null){
                  createCell(m_Row, col++, RepSales.getString("LYTD_Per"), m_StyleText);
                  createCell(m_Row, col++, RepSales.getDouble("LYTD"), m_StyleDec);
                  createCell(m_Row, col++, RepSales.getString("LYTD_Date"), m_StyleText);
               }

               m_Row = m_Sheet.createRow(m_RowNum++);
               col = (short)1;
               createCell(m_Row, col++, RepSales.getString("MTD_Per"), m_StyleText);
               createCell(m_Row, col++, RepSales.getDouble("MTD"), m_StyleDec);
               createCell(m_Row, col++, RepSales.getString("MTD_Date"), m_StyleText);

               m_Row = m_Sheet.createRow(m_RowNum++);
               col = (short)1;
               if (RepSales.getString("LYMTD_Per") != null){
                  createCell(m_Row, col++, RepSales.getString("LYMTD_Per"), m_StyleText);
                  createCell(m_Row, col++, RepSales.getDouble("LYMTD"), m_StyleDec);
                  createCell(m_Row, col++, RepSales.getString("LYMTD_Date"), m_StyleText);
               }
            }

         }
         catch ( Exception ex ) {
            m_ErrMsg.append(ex.getClass().getName() + "\r\n");
            m_ErrMsg.append(ex.getMessage());
            m_Log.error("exception", ex);
         }
         finally{
            closeRSet(RepSales);
         }

         m_WrkBk.write(outFile);
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         m_Log.error("exception", ex);
      }

      finally {
         closeRSet(RepShipments);
         try {
            outFile.close();
         }

         catch( Exception e ) {
            m_Log.error("exception:", e);
         }

         outFile = null;
      }

      return result;
   }

   /**
    * Builds the sql based on the type of filter requested by the user.
    * @return A complete sql statement.
    */
   private String buildSql()
   {
      StringBuffer sql = new StringBuffer(256);

      sql.append("select cust_rep_div_view.RepName as Sales_Rep, \r\n");
      sql.append("   inv_dtl.cust_nbr as Cust_Nbr, \r\n");
      sql.append("   inv_hdr.cust_name as Customer_Name, \r\n");
      sql.append("   inv_dtl.invoice_nbr as Invoice_Nbr, \r\n");
      sql.append("   to_char(inv_dtl.invoice_date,'mm/dd/yyyy') as dato, \r\n"); //mostly for future stuff when they want more than one date
      sql.append("   inv_hdr.cust_po_nbr as Cust_PO_Nbr, \r\n");
      sql.append("   inv_dtl.item_nbr as Item_Id, \r\n");
      sql.append("   inv_dtl.item_descr as Item_Description, \r\n");
      sql.append("   inv_dtl.ship_unit as UOM, \r\n");
      sql.append("   inv_dtl.qty_ordered as Qty_Ord, \r\n");
      sql.append("   inv_dtl.qty_shipped as Qty_Ship, \r\n");
      sql.append("   (inv_dtl.qty_ordered - inv_dtl.qty_shipped) as Lost_Qty, \r\n");
      sql.append("   inv_dtl.action as Comments, \r\n");
      sql.append("   inv_dtl.unit_sell as Unit_Sell, \r\n");
      sql.append("   case ejd_cust_procs.cust_can_price_item(inv_dtl.cust_nbr, inv_dtl.item_ea_id) \r\n");
      sql.append("      when 0 then 0 \r\n");
      sql.append("      else (select price from ejd_cust_procs.get_sell_price(inv_dtl.cust_nbr, inv_dtl.item_ea_id)) \r\n ");
      sql.append("   end Reg_Sell, \r\n");
      sql.append("   (select margin_pct from ejd_price_procs.get_margin_price_pct(inv_dtl.cust_nbr, inv_dtl.item_ea_id)) as Mgn_Pct, \r\n ");
      sql.append("   inv_dtl.ext_sell as Ext_Price, \r\n");
      sql.append("   inv_dtl.sell_source as Price_Method, \r\n");
      sql.append("   best_price.calculation as Best_Price, \r\n");
      sql.append("   round(break_pack.percent / 100 * ext_sell,3) as Brk_Pck_Chg, \r\n");
      sql.append("   inv_dtl.vendor_name ");
      sql.append("from inv_dtl \r\n");
      sql.append("join inv_hdr on inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id  \r\n");
      sql.append("join cust_rep_div_view on cust_rep_div_view.customer_id = inv_dtl.cust_nbr and \r\n");
      sql.append("     replace(cust_rep_div_view.RepName,' ','') = ? and cust_rep_div_view.rep_type = 'SALES REP'\r\n");
      sql.append("left outer join sa.inv_price_calc best_price on best_price.inv_dtl_id = inv_dtl.inv_dtl_id and \r\n");
      sql.append("     best_price.calculation = 'BEST PRICE' \r\n");
      sql.append("left outer join sa.inv_price_calc break_pack on break_pack.inv_dtl_id = inv_dtl.inv_dtl_id and \r\n");
      sql.append("     break_pack.calculation = 'Brk Pck Chg'  \r\n");

      if (m_RptDate != null && m_RptDate.length() > 0)
         sql.append("where inv_dtl.invoice_date = to_date(?,'mm/dd/yyyy') and \r\n");
      else
         sql.append("where inv_dtl.invoice_date = trunc(now()) - 1 and \r\n");

      sql.append("inv_dtl.tran_type = 'CREDIT' and inv_dtl.sale_type = 'WAREHOUSE'\r\n");

      sql.append("union \r\n");

      sql.append("select cust_rep_div_view.RepName as Sales_Rep, \r\n");
      sql.append("   inv_dtl.cust_nbr as Cust_Nbr, \r\n");
      sql.append("   inv_hdr.cust_name as Customer_Name, \r\n");
      sql.append("   inv_dtl.invoice_nbr as Invoice_Nbr, \r\n");
      sql.append("   to_char(inv_dtl.invoice_date,'mm/dd/yyyy') as dato, \r\n"); //mostly for future stuff when they want more than one date
      sql.append("   inv_hdr.cust_po_nbr as Cust_PO_Nbr, \r\n");
      sql.append("   inv_dtl.item_nbr as Item_Id, \r\n");
      sql.append("   inv_dtl.item_descr as Item_Description, \r\n");
      sql.append("   inv_dtl.ship_unit as UOM, \r\n");
      sql.append("   inv_dtl.qty_ordered as Qty_Ord, \r\n");
      sql.append("   inv_dtl.qty_shipped as Qty_Ship, \r\n");
      sql.append("   (inv_dtl.qty_ordered - inv_dtl.qty_shipped) as Lost_Qty, \r\n");
      sql.append("   inv_dtl.action as Comments, \r\n");
      sql.append("   inv_dtl.unit_sell as Unit_Sell, \r\n");
      sql.append("   case ejd_cust_procs.cust_can_price_item(inv_dtl.cust_nbr, inv_dtl.item_ea_id) \r\n");
      sql.append("      when 0 then 0 \r\n");
      sql.append("      else (select price from ejd_cust_procs.get_sell_price(inv_dtl.cust_nbr, inv_dtl.item_ea_id)) \r\n ");
      sql.append("   end Reg_Sell, \r\n");
      sql.append("   (select margin_pct from ejd_price_procs.get_margin_price_pct(inv_dtl.cust_nbr, inv_dtl.item_ea_id)) as Mgn_Pct, \r\n ");
      sql.append("   inv_dtl.ext_sell as Ext_Price, \r\n");
      sql.append("   inv_dtl.sell_source as Price_Method, \r\n");
      sql.append("   best_price.calculation as Best_Price, \r\n");
      sql.append("   round(break_pack.percent / 100 * ext_sell,3) as Brk_Pck_Chg, ");
      sql.append("   inv_dtl.vendor_name ");
      sql.append("from inv_dtl \r\n");
      sql.append("join inv_hdr  on inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id \r\n");
      sql.append("join invoice i on to_char(i.invoice_num) = inv_hdr.invoice_nbr \r\n");
      sql.append("join trip_stop ts on ts.trip_stop_id = i.trip_stop_id \r\n");
      sql.append("join trip t on t.trip_id = ts.trip_id and \r\n");

      if (m_LoadDate != null && m_LoadDate.length() > 0)
         sql.append("t.load_date = to_date(?,'mm/dd/yyyy') \r\n");
      else
         sql.append("t.load_date = trunc(now()) - 1 \r\n");

      sql.append("join cust_rep_div_view on cust_rep_div_view.customer_id = inv_dtl.cust_nbr and \r\n");
      sql.append("     replace(cust_rep_div_view.RepName,' ','') = ? and cust_rep_div_view.rep_type = 'SALES REP'\r\n");
      sql.append("left outer join sa.inv_price_calc best_price on best_price.inv_dtl_id = inv_dtl.inv_dtl_id and \r\n");
      sql.append("     best_price.calculation = 'BEST PRICE' \r\n");
      sql.append("left outer join sa.inv_price_calc break_pack on break_pack.inv_dtl_id = inv_dtl.inv_dtl_id and \r\n");
      sql.append("     break_pack.calculation = 'Brk Pck Chg'  \r\n");
      sql.append("where inv_dtl.tran_type = 'SALE' and inv_dtl.sale_type = 'WAREHOUSE'\r\n");
      sql.append("order by cust_nbr, cust_po_nbr, item_id  \r\n");

      return sql.toString();
   }

   //
   // Creates the sales history query
   private String buildRepSalesSql()
   {
      StringBuilder sql = new StringBuilder();

      sql.append("select ytd_dollars as YTD,  ");
      sql.append("to_char(last_updated, 'mm/dd/yyyy') as YTD_Date,  ");
      sql.append("to_char(date_trunc('year', last_updated), 'yyyy') as YTD_PER,  ");
      sql.append("mtd_dollars as MTD,  ");
      sql.append("to_char(last_updated, 'mm/dd/yyyy') as MTD_Date,  ");
      sql.append("to_char(date_trunc('month', last_updated), 'MON yyyy') as MTD_PER,    ");
      sql.append("prev_ytd_dollars as LYTD,  ");
      sql.append("to_char(last_updated-365, 'mm/dd/yyyy') as LYTD_Date,  ");
      sql.append("to_char(date_trunc('year', last_updated-365), 'yyyy') as LYTD_PER,    ");
      sql.append("prev_mtd_dollars as LYMTD,  ");
      sql.append("to_char(last_updated-365, 'mm/dd/yyyy') as LYMTD_Date,  ");
      sql.append("to_char(date_trunc('month', last_updated-365), 'MON yyyy') as LYMTD_PER      ");
      sql.append("from sa.tm_summary ");
      sql.append("join emery_rep on emery_rep.er_id = tm_summary.er_id ");
      sql.append("where first || ' ' || last = ? ");
	  
      return sql.toString();
   }


   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_RepShipments);
   }


   private void initReport()
   {
      XSSFDataFormat df;
      short col = 0;

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
         m_StyleDec.setDataFormat(df.getFormat("#,##0.00"));

         m_StyleDecBold = m_WrkBk.createCellStyle();
         m_StyleDecBold.setFont(m_FontBold);
         m_StyleDecBold.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleDecBold.setDataFormat(df.getFormat("#,##0.00"));

         m_StyleHeader = m_WrkBk.createCellStyle();
         m_StyleHeader.setFont(m_FontTitle);
         m_StyleHeader.setAlignment(HorizontalAlignment.CENTER);

         m_StyleInt = m_WrkBk.createCellStyle();
         m_StyleInt.setAlignment(HorizontalAlignment.RIGHT);
         m_StyleInt.setFont(m_FontData);
         m_StyleInt.setDataFormat((short)3);

         m_StyleTitle = m_WrkBk.createCellStyle();
         m_StyleTitle.setFont(m_FontTitle);
         m_StyleTitle.setAlignment(HorizontalAlignment.LEFT);

         m_Sheet = m_WrkBk.createSheet();
         m_Sheet.setMargin(XSSFSheet.BottomMargin, .25);
         m_Sheet.getPrintSetup().setLandscape(true);
         m_Sheet.getPrintSetup().setPaperSize((short)5);

         m_Header = m_Sheet.getHeader();
         m_Header.setCenter(HSSFHeader.font("Arial", "Bold") + HSSFHeader.fontSize((short) 12) + "REPREP Report");

         m_RowNum = 0;

         if (m_RptDate != null && m_RptDate.length() > 0)
            rptHeader(m_Rep, m_RptDate, m_LoadDate);
         else
            rptHeader(m_Rep, getDateYesterday(), getDateYesterday());

         // Initialize the default column widths
         // Create the column headings
         m_Row = m_Sheet.createRow(m_RowNum);
         col = (short)0;
         m_Sheet.setColumnWidth(col, 4000);
         createCell(m_Row, col++, "Sales Rep", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2000);
         createCell(m_Row, col++, "Cust Nbr", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 10000);
         createCell(m_Row, col++, "Customer Name", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2000);
         createCell(m_Row, col++, "Inv Nbr", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 5000);
         createCell(m_Row, col++, "Cust PO Nbr", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2000);
         createCell(m_Row, col++, "Item ID", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 10000);
         createCell(m_Row, col++, "Item Description", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 10000);
         createCell(m_Row, col++, "Vendor Name", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2000);
         createCell(m_Row, col++, "UOM", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2000);
         createCell(m_Row, col++, "Qty Ord", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2000);
         createCell(m_Row, col++, "Qty Ship", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2000);
         createCell(m_Row, col++, "Lost Qty", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 7000);
         createCell(m_Row, col++, "Comments", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2500);
         createCell(m_Row, col++, "Unit Sell", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2500);
         createCell(m_Row, col++, "Reg Sell", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 3000);
         createCell(m_Row, col++, "Mgn Pct", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2500);
         createCell(m_Row, col++, "Ext Price", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2500);
         createCell(m_Row, col++, "Lost Sales", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 2500);
         createCell(m_Row, col++, "Price Method", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 3000);
         createCell(m_Row, col++, "Best Price", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 3000);
         createCell(m_Row, col++, "Brk Pck Chg", m_StyleBoldCenter);
         m_Sheet.setColumnWidth(col, 3000);

         // Set the column heading row to repeat on each page         
         m_Sheet.setRepeatingRows(CellRangeAddress.valueOf("1:1"));
         m_Sheet.setRepeatingColumns(CellRangeAddress.valueOf("A:C"));
         m_RowNum+= 2;
      }

      catch ( Exception e ) {
         log.error( e );
      }
   }


   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   @Override
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;
      setupWorkbook();

      try {
      	m_EdbConn = m_RptProc.getEdbConn();
         prepareStatements();
         created = buildOutputFile();
      }

      catch ( Exception ex ) {
         m_Log.fatal("exception:", ex);
      }

      finally {
         closeStatements();

         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Prepares the sql queries for execution.
    *
    */
   private void prepareStatements() throws Exception
   {
      if ( m_EdbConn != null ) {

         // our main big honking query gets put together in buildsql
         m_RepShipments = m_EdbConn.prepareStatement(buildSql());
         m_RepSales = m_EdbConn.prepareStatement(buildRepSalesSql());
      }
   }

   private String getDateYesterday()
   {
      Calendar cal = Calendar.getInstance();
      cal.add(Calendar.DATE, -1);
      DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
      return dateFormat.format(cal.getTime());
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    *
    * Because it's possible that this report can be called from some other system, the
    * best way to deal with params is to not go by the order, but by the name.
    *
    */
   @Override
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fname = new StringBuffer();
      String tm = Long.toString(System.currentTimeMillis()).substring(3);
      int pcount = params.size();
      Param param = null;

      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);

         if ( param.name.equals("Rep") ) {
            m_Rep = param.value;
            m_RepNoSpaces = m_Rep.replace(" ","");
         }

         if ( param.name.equals("RptDate") )
            m_RptDate = param.value;

         if ( param.name.equals("LoadDate") )
            m_LoadDate = param.value;
      }

      if ( m_LoadDate == null || m_LoadDate.length() == 0 )
         m_LoadDate = m_RptDate;

      fname.append(String.format("repship-%s-%s.xlsx", tm.substring(tm.length()-5, tm.length()), m_Rep.replace(' ', '-' ).toLowerCase()));

      m_FileNames.add(fname.toString());
   }

   /**
    * Sets up the styles for the cells based on the column data.  Does any other inititialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {
      XSSFCellStyle styleText;      // Text right justified
      XSSFCellStyle styleInt;       // Style with 0 decimals
      XSSFCellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00)
      XSSFCellStyle stylePct;       // Style with 0 decimals + %

      styleText = m_WrkBk.createCellStyle();
      styleText.setAlignment(HorizontalAlignment.LEFT);

      styleInt = m_WrkBk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);

      styleMoney = m_WrkBk.createCellStyle();
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short)8);

      stylePct = m_WrkBk.createCellStyle();
      stylePct.setAlignment(HorizontalAlignment.RIGHT);
      stylePct.setDataFormat((short)9);


   }

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

   private void custTrailer(String LastCust, double CustSales, double CustLostSales)
   {
      short col = 1;

      m_Row = m_Sheet.createRow(m_RowNum++);
      createCell(m_Row, col++, LastCust,  m_StyleBold);
      createCell(m_Row, col++, "Customer Total",  m_StyleBold);
      col+= 13;
      createCell(m_Row, col++, CustSales, m_StyleDecBold);
      createCell(m_Row, col++, CustLostSales, m_StyleDecBold);
      m_RowNum++;
   }

   private void rptTrailer(String DateString, double TotalSales, double TotalCredits, double TotalLostSales)
   {
      short col = 1;

      m_Row = m_Sheet.createRow(m_RowNum++);
      createCell(m_Row, col++, DateString, m_StyleDecBold);
      createCell(m_Row, col++, "Total Sales",  m_StyleBold);
      col+= 13;
      createCell(m_Row, col++, TotalSales, m_StyleDecBold);
      m_Row = m_Sheet.createRow(m_RowNum++);
      col = 1;
      createCell(m_Row, col++, DateString, m_StyleDecBold);
      createCell(m_Row, col++, "Total Credits",  m_StyleBold);
      col+= 13;
      createCell(m_Row, col++, TotalCredits, m_StyleDecBold);
      m_Row = m_Sheet.createRow(m_RowNum++);
      col = 1;
      createCell(m_Row, col++, DateString, m_StyleDecBold);
      createCell(m_Row, col++, "Total Lost Sales",  m_StyleBold);
      col += 14;
      createCell(m_Row, col++, TotalLostSales, m_StyleDecBold);
      m_RowNum++;
   }

   private void rptHeader(String RepString, String rptDate, String loadDate)
   {
      short col = 2;

      m_Row = m_Sheet.createRow(m_RowNum++);
      createCell(m_Row, col++, "Ship Report for " + RepString + "   Report Date " + rptDate + " Load Date " + loadDate, m_StyleDecBold);
      m_RowNum++;
   }
   
}

