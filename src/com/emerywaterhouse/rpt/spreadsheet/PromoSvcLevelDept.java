/**
 *
 * @author Sam Gillis
 * $Revision: 1.3 $
 *
 * Last Update: $Id: PromoSvcLevelDept.java,v 1.3 2014/05/02 13:21:10 sgillis Exp $
 *
 * History
 *    $Log: PromoSvcLevelDept.java,v $
 *    Revision 1.3  2014/05/02 13:21:10  sgillis
 *    comment header added
 *
 */

package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;


public class PromoSvcLevelDept extends Report {

   private PreparedStatement m_PromoSvcDept;
   private PreparedStatement m_PromoSvcDeptSubs;
   private PreparedStatement m_PromoSvcDeptPacketTotals;
   
   
   private static final short MAX_COLS = 15;
   private static final short MAX_COLS_SUBT = 6;
   //
   // The cell styles for each of the base columns in the spreadsheet.
   private CellStyle[] m_CellStyles;

   
   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private Sheet m_Sheet;  
   private Sheet m_Sheet2;
   private Font m_FontBold;
   private Font m_FontNormal;
   
   // Parameter member variables
   private String m_BegDate;
   private String m_EndDate;
   private String m_Packets;
   
   //subtotals
   /*
   private HashMap<String, Integer> units_ordered_dept_ptld = new HashMap<String, Integer>();
   private HashMap<String, Integer> units_ordered_dept_pitt = new HashMap<String, Integer>();
   private HashMap<String, Integer> units_invoiced_dept_ptld = new HashMap<String, Integer>();
   private HashMap<String, Integer> units_invoiced_dept_pitt = new HashMap<String, Integer>();
   private HashMap<String, Integer> dollars_invoiced_dept_ptld = new HashMap<String, Integer>();
   private HashMap<String, Integer> dollars_invoiced_dept_pitt = new HashMap<String, Integer>();
   private HashMap<String, Integer> emery_on_hand_dept_ptld = new HashMap<String, Integer>();
   private HashMap<String, Integer> emery_on_hand_dept_pitt = new HashMap<String, Integer>();
   */ //actually, nope, just make oracle do this subtotal crap.
   //calc fill rate
   
   /**
    * 
    */
   public PromoSvcLevelDept()
   {
      super();
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet("Data");
      m_Sheet2 = m_Wrkbk.createSheet("Subtotals");
      setupWorkbook();      
   }
   private void setupWorkbook()
   {      
      CellStyle styleText;      // Text left justified
      CellStyle styleInt;       // Style with 0 decimals
      CellStyle styleDec;       // 2 decimal positions
      
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
      
      m_CellStyles = new CellStyle[] {
         styleText,     // col 0 packet id
         styleText,     // col 1 Title
         styleText,     // col 2 Promo id
         styleText,     // col 3 Item id
         styleText,     // col 4 description
         styleText,     // col 5 vendor
         styleInt,     // col 6 units ordered
         styleInt,      // col 7 units invoice
         styleDec,      // col 8 fill rate %
         styleDec,     // col 9 dollars invoiced
         styleInt,     // col 10 emery on order
         styleInt,     // col 11 buying dept
         styleText,     // col 12 warehouse
         styleText,     // col 13 dsb date
         styleText,     // col 14 dsa date
      };
      
      styleText = null;
      styleInt = null;
      styleDec = null;
   }
   
   @Override
   public boolean createReport() {
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         m_EdbConn = m_RptProc.getEdbConn();  
         //java.util.Properties connProps = new java.util.Properties();
         //connProps.put("user", "sgillis");
         //connProps.put("password", "ucb0JVLwake");
         
         //m_EdbConn = java.sql.DriverManager.getConnection(
         //      "jdbc:oracle:thin:@10.128.0.9:1521:GROK",connProps);
         
         if ( prepareStatements() )
            created = buildOutputFile();            
      }
      
      catch ( Exception ex ) {
         System.out.println(ex);
         log.fatal("exception:", ex);
      }
      
      finally {
         closeStatements(); 
         
         if ( m_Status == RptServer.RUNNING )
            m_Status = RptServer.STOPPED;
      }
      
      return created;
   }
   
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;
      
      if ( m_EdbConn != null ) {
         try { 
         
         sql.append("select  ");
         sql.append("   packet_id, title, promo_id, item_id, description,  ");
         sql.append("   vendor_name, units_ordered, units_invoiced, fill_rate_pct, dollars_invoiced, ");
         sql.append("   decode(warehouse, 'PORTLAND', pos_util.get_on_order(item_id, '01'),  ");
         sql.append("      pos_util.get_on_order(item_id, '02')) emery_on_order,   ");
         sql.append("   buying_dept, warehouse, dsb_date, dsa_date ");
         sql.append("from ( ");
         sql.append("   select  ");
         sql.append("      packet.packet_id, packet.title, promo_item.promo_id, promo_item.item_id, ");
         sql.append("      item.description, inv_dtl.vendor_name, sum(inv_dtl.qty_ordered) units_ordered, ");
         sql.append("      sum(inv_dtl.qty_shipped) units_invoiced, sum(inv_dtl.ext_sell) dollars_invoiced,  ");
         sql.append("      round(sum(inv_dtl.qty_shipped)/sum(qty_ordered) * 100, 1) fill_rate_pct, ");
         sql.append("      inv_dtl.buying_dept, inv_dtl.warehouse, promotion.dsb_date, promotion.dsa_date  ");
         sql.append("   from packet ");
         sql.append("   join promotion on promotion.packet_id = packet.packet_id   ");
         sql.append("   join promo_item on promo_item.promo_id = promotion.promo_id ");
         sql.append("   join item on item.item_id = promo_item.item_id ");
         sql.append("   join inv_dtl on inv_dtl.item_nbr = promo_item.item_id and inv_dtl.promo_nbr = promo_item.promo_id and ");
         sql.append("      inv_dtl.tran_type = 'SALE' "); 
         sql.append("   and inv_dtl.invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') ");
         //not sure if date needed at this point, request indicates just by packet is fine
         sql.append("   where ");
         sql.append("      InStr(?,packet.packet_id) > 0 ");
         sql.append("   group by  ");
         sql.append("      packet.packet_id, packet.title, inv_dtl.buying_dept, inv_dtl.warehouse, ");
         sql.append("      promo_item.promo_id, promo_item.item_id, item.description, inv_dtl.vendor_name, ");
         sql.append("      promotion.dsb_date, promotion.dsa_date ");
         sql.append("   order by  ");
         sql.append("      packet.packet_id, inv_dtl.buying_dept, warehouse, promo_item.item_id ");
         sql.append(") packet_svc_lvl ");
         m_PromoSvcDept = m_EdbConn.prepareStatement(sql.toString());  
         
         sql.setLength(0);
         sql.append("select buying_dept, warehouse, sum(units_ordered) units_ordered, ");
         sql.append("sum(units_invoiced) units_invoiced, sum(dollars_invoiced) dollars_invoiced from ");
         sql.append("(select  ");
         sql.append("   packet_id, title, promo_id, item_id, description,  ");
         sql.append("   vendor_name, units_ordered, units_invoiced, fill_rate_pct, dollars_invoiced, ");
         sql.append("   decode(warehouse, 'PORTLAND', pos_util.get_on_order(item_id, '01'),  ");
         sql.append("      pos_util.get_on_order(item_id, '02')) emery_on_order,   ");
         sql.append("   buying_dept, warehouse, dsb_date, dsa_date ");
         sql.append("from ( ");
         sql.append("   select  ");
         sql.append("      packet.packet_id, packet.title, promo_item.promo_id, promo_item.item_id, ");
         sql.append("      item.description, inv_dtl.vendor_name, sum(inv_dtl.qty_ordered) units_ordered, ");
         sql.append("      sum(inv_dtl.qty_shipped) units_invoiced, sum(inv_dtl.ext_sell) dollars_invoiced,  ");
         sql.append("      round(sum(inv_dtl.qty_shipped)/sum(qty_ordered) * 100, 1) fill_rate_pct, ");
         sql.append("      inv_dtl.buying_dept, inv_dtl.warehouse, promotion.dsb_date, promotion.dsa_date  ");
         sql.append("   from packet ");
         sql.append("   join promotion on promotion.packet_id = packet.packet_id ");
         sql.append("   join promo_item on promo_item.promo_id = promotion.promo_id ");
         sql.append("   join item on item.item_id = promo_item.item_id ");
         sql.append("   join inv_dtl on inv_dtl.item_nbr = promo_item.item_id and inv_dtl.promo_nbr = promo_item.promo_id and ");
         sql.append("      inv_dtl.tran_type = 'SALE' "); 
         sql.append(" and inv_dtl.invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') ");
         //not sure if date needed at this point, request indicates just by packet is fine
         sql.append("   where ");
         sql.append("      InStr(?,packet.packet_id)  > 0 ");
         sql.append("   group by  ");
         sql.append("      packet.packet_id, packet.title, inv_dtl.buying_dept, inv_dtl.warehouse, ");
         sql.append("      promo_item.promo_id, promo_item.item_id, item.description, inv_dtl.vendor_name, ");
         sql.append("      promotion.dsb_date, promotion.dsa_date ");
         sql.append("   order by  ");
         sql.append("      packet.packet_id, inv_dtl.buying_dept, warehouse, promo_item.item_id ");
         sql.append(") packet_svc_lvl) ");
         sql.append(" group by buying_dept, warehouse ");
         sql.append(" order by warehouse, buying_dept ");
         m_PromoSvcDeptSubs = m_EdbConn.prepareStatement(sql.toString()); 
         
         sql.setLength(0);
         sql.append("select packet_id, warehouse, sum(units_ordered) units_ordered, ");
         sql.append("sum(units_invoiced) units_invoiced, sum(dollars_invoiced) dollars_invoiced from ");
         sql.append("(select  ");
         sql.append("   packet_id, title, promo_id, item_id, description,  ");
         sql.append("   vendor_name, units_ordered, units_invoiced, fill_rate_pct, dollars_invoiced, ");
         sql.append("   decode(warehouse, 'PORTLAND', pos_util.get_on_order(item_id, '01'),  ");
         sql.append("      pos_util.get_on_order(item_id, '02')) emery_on_order,   ");
         sql.append("   buying_dept, warehouse, dsb_date, dsa_date ");
         sql.append("from ( ");
         sql.append("   select  ");
         sql.append("      packet.packet_id, packet.title, promo_item.promo_id, promo_item.item_id, ");
         sql.append("      item.description, inv_dtl.vendor_name, sum(inv_dtl.qty_ordered) units_ordered, ");
         sql.append("      sum(inv_dtl.qty_shipped) units_invoiced, sum(inv_dtl.ext_sell) dollars_invoiced,  ");
         sql.append("      round(sum(inv_dtl.qty_shipped)/sum(qty_ordered) * 100, 1) fill_rate_pct, ");
         sql.append("      inv_dtl.buying_dept, inv_dtl.warehouse, promotion.dsb_date, promotion.dsa_date  ");
         sql.append("   from packet ");
         sql.append("   join promotion on promotion.packet_id = packet.packet_id ");
         sql.append("   join promo_item on promo_item.promo_id = promotion.promo_id ");
         sql.append("   join item on item.item_id = promo_item.item_id ");
         sql.append("   join inv_dtl on inv_dtl.item_nbr = promo_item.item_id and inv_dtl.promo_nbr = promo_item.promo_id and ");
         sql.append("      inv_dtl.tran_type = 'SALE' "); 
         sql.append(" and inv_dtl.invoice_date between to_date(?, 'mm/dd/yyyy') and to_date(?, 'mm/dd/yyyy') ");
         //not sure if date needed at this point, request indicates just by packet is fine
         sql.append("   where ");
         sql.append("      InStr(?,packet.packet_id) > 0 ");
         sql.append("   group by  ");
         sql.append("      packet.packet_id, packet.title, inv_dtl.buying_dept, inv_dtl.warehouse, ");
         sql.append("      promo_item.promo_id, promo_item.item_id, item.description, inv_dtl.vendor_name, ");
         sql.append("      promotion.dsb_date, promotion.dsa_date ");
         sql.append("   order by  ");
         sql.append("      packet.packet_id, inv_dtl.buying_dept, warehouse, promo_item.item_id ");
         sql.append(") packet_svc_lvl ) ");
         sql.append(" group by packet_id, warehouse ");
         sql.append(" order by warehouse, packet_id ");
         m_PromoSvcDeptPacketTotals = m_EdbConn.prepareStatement(sql.toString()); 
         
         isPrepared = true;
         }catch (SQLException e) {
            log.error("exception:", e);
         }
      } else {
         log.error("PromoSvcLevelDept.prepareStatements - null Edb connection");
      }
         
      return isPrepared;
   }
   
   private void closeStatements()
   {
      closeStmt(m_PromoSvcDept);
      closeStmt(m_PromoSvcDeptSubs);
      closeStmt(m_PromoSvcDeptPacketTotals);
   }
   
   /**
    * Executes the queries and builds the output file
    * 
    * @return true if the report was successfully built
    * @throws FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {   
      Row row = null;
      int rowNum = 0;
      int colNum = 0;
      FileOutputStream outFile = null;
      ResultSet rs = null;
      boolean result = false;
      
      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
      
      try {
         rowNum = createCaptions();
         m_PromoSvcDept.setString(1, m_BegDate);
         m_PromoSvcDept.setString(2, m_EndDate);
         m_PromoSvcDept.setString(3, m_Packets);
         rs = m_PromoSvcDept.executeQuery();
         m_CurAction = "Building output file";
         rowNum = createCaptions(); //why are there two??? WHAT DOES THIS CODE DO??
         
         while ( rs.next() && getStatus() != RptServer.STOPPED ) {
            row = createRowSheetOne(rowNum++, MAX_COLS);
            colNum = 0;
            //TODO
            //le psuedocode
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("packet_id")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("title")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("promo_id")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("item_id")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("description")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("vendor_name")));
            row.getCell(colNum++).setCellValue(rs.getInt("units_ordered"));
            row.getCell(colNum++).setCellValue(rs.getInt("units_invoiced"));
            row.getCell(colNum++).setCellValue(rs.getDouble("fill_rate_pct"));
            row.getCell(colNum++).setCellValue(rs.getDouble("dollars_invoiced"));
            row.getCell(colNum++).setCellValue(rs.getInt("emery_on_order"));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("buying_dept")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("warehouse")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("dsb_date")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("dsa_date")));
            
         }

         //add whitespace
         rs.close();
         m_PromoSvcDeptSubs.setString(1, m_BegDate);
         m_PromoSvcDeptSubs.setString(2, m_EndDate);
         m_PromoSvcDeptSubs.setString(3, m_Packets);
         rs = m_PromoSvcDeptSubs.executeQuery();
         rowNum = 0; //reset row num, sheet 2
         //structure from the query is pitt/by dept, then port/by dept
         rowNum = createCaptionsDeptSubtotals();
         while (rs.next()) {
           colNum = 0;
           row = createRowSheetTwo(rowNum++, MAX_COLS_SUBT);
           row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("warehouse")));
           row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("buying_dept")));
           row.getCell(colNum++).setCellValue(rs.getInt("units_ordered"));
           row.getCell(colNum++).setCellValue(rs.getInt("units_invoiced"));
           row.getCell(colNum++).setCellValue(rs.getDouble("dollars_invoiced"));
           row.getCell(colNum++).setCellValue(
                 new DecimalFormat("#.#").format((((double)rs.getInt("units_invoiced") / (double)rs.getInt("units_ordered"))*100D)));
         }
         
         rowNum++;
         rs.close();
         rowNum = createCaptionsPacketSubtotals(rowNum);
       
         //pltd-pckt
         //pitt-pckt
         m_PromoSvcDeptPacketTotals.setString(1, m_BegDate);
         m_PromoSvcDeptPacketTotals.setString(2, m_EndDate);
         m_PromoSvcDeptPacketTotals.setString(3, m_Packets);
         rs = m_PromoSvcDeptPacketTotals.executeQuery();
         
         while (rs.next()) {
            colNum = 0;
            row = createRowSheetTwo(rowNum++, MAX_COLS_SUBT);
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("warehouse")));
            row.getCell(colNum++).setCellValue(new XSSFRichTextString(rs.getString("packet_id")));
            row.getCell(colNum++).setCellValue(rs.getInt("units_ordered"));
            row.getCell(colNum++).setCellValue(rs.getInt("units_invoiced"));
            row.getCell(colNum++).setCellValue(rs.getDouble("dollars_invoiced"));
            row.getCell(colNum++).setCellValue(
                  new DecimalFormat("#.#").format((((double)rs.getInt("units_invoiced") / (double)rs.getInt("units_ordered"))*100D)));
          }
         rs.close();
         
         
         //print
         m_Wrkbk.write(outFile);
         result = true;
         
      } catch (SQLException e) {
         System.out.println(e);
         log.error(e);
      } catch (IOException e) {
         log.error(e);
      } catch (Exception e) {
         System.out.println(e);
         //log.error(e);
      } finally {         
         closeRSet(rs);
         rs = null;
         
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
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();      
      String tmp = Long.toString(System.currentTimeMillis());
                  
      fileName.append("promo_svc_dept");      
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      m_FileNames.add(fileName.toString());
      
      m_Packets = params.get(0).value.trim();
      m_BegDate = params.get(1).value.trim();
      m_EndDate = params.get(2).value.trim();
   }
   
   /**
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    * @param colCnt The number of columns in the row.
    * 
    * @return The formatted row of the spreadsheet.
    */
   private Row createRowSheetOne(int rowNum, short colCnt)
   {
      Row row = null;
      Cell cell = null;
      
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
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    * @param colCnt The number of columns in the row.
    * 
    * @return The formatted row of the spreadsheet.
    */
   private Row createRowSheetTwo(int rowNum, short colCnt)
   {
      Row row = null;
      Cell cell = null;
      
      if ( m_Sheet2 == null )
         return row;

      row = m_Sheet2.createRow(rowNum);

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
    * Creates the report and column headings
    * 
    * @return short - the next available row#
    */
   private int createCaptions()
   {
      Row row = null;
      Cell cell = null;
      int colCnt = 20;
      int col = 0;
      int rw = 0;

      //
      row = m_Sheet.createRow(rw++);
      
      cell = row.createCell( 0);
      cell.setCellType(CellType.STRING);
      cell.getCellStyle().setFont(m_FontBold);
      cell.setCellValue(new XSSFRichTextString("Promo Svc Level -Dept"));


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

         //these column widths are in 1/256 of a character
         //so to make it easier to read, i converted the widths to length_in_chars*256
         m_Sheet.setColumnWidth(col, 6*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Packet ID"));

         m_Sheet.setColumnWidth(col, 30*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Title"));
         
         m_Sheet.setColumnWidth(col, 5*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Promo ID"));
         
         m_Sheet.setColumnWidth(col, 8*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Item ID"));
         
         m_Sheet.setColumnWidth(col, 40*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Description"));

         m_Sheet.setColumnWidth(col, 18*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Vendor"));

         m_Sheet.setColumnWidth(col, 10*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Units Ordered"));

         m_Sheet.setColumnWidth(col, 10*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Units Invoiced"));

         m_Sheet.setColumnWidth(col, 8*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Fill Rate %"));

         m_Sheet.setColumnWidth(col, 12*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Dollars Invoiced"));

         m_Sheet.setColumnWidth(col, 12*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Emery On-Order"));

         m_Sheet.setColumnWidth(col, 4*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Dept"));

         m_Sheet.setColumnWidth(col, 10*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Warehouse"));

         m_Sheet.setColumnWidth(col, 7*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("DSA Date"));

         m_Sheet.setColumnWidth(col, 7*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("DSB Date"));

      }
      
      return rw;
   }

   private int createCaptionsDeptSubtotals()
   {
      Row row = null;
      Cell cell = null;
      int colCnt = 20;
      int col = 0;
      int rw = 0;

      //
      row = m_Sheet2.createRow(rw++);
      
      cell = row.createCell( 0);
      cell.setCellType(CellType.STRING);
      cell.getCellStyle().setFont(m_FontBold);
      cell.setCellValue(new XSSFRichTextString("Promo Svc Level -Dept"));

      //
      // Show the current date
      row = m_Sheet2.createRow(rw++);
      cell = row.createCell( 0);
      cell.setCellType(CellType.STRING);
      cell.setCellValue(
         new XSSFRichTextString(new SimpleDateFormat("yyyy/MM/dd").format(new java.util.Date()))
      );

      //
      // Build the column headings      
      row = m_Sheet2.createRow(rw++);

      if ( row != null ) {
         for ( int i = 0; i < colCnt; i++ ) {
            cell = row.createCell(i);
            cell.setCellType(CellType.STRING);
            cell.getCellStyle().setFont(m_FontBold);
         }

         //these column widths are in 1/256 of a character
         //so to make it easier to read, i converted the widths to length_in_chars*256
         m_Sheet2.setColumnWidth(col, 12*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Warehouse"));

         m_Sheet2.setColumnWidth(col, 6*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Dept #"));
         
         m_Sheet2.setColumnWidth(col, 15*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Units Ordered"));
         
         m_Sheet2.setColumnWidth(col, 15*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Units Invoiced"));
         
         m_Sheet2.setColumnWidth(col, 15*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Dollars Invoiced"));

         m_Sheet2.setColumnWidth(col, 12*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Fill Rate %"));


      }
      
      return rw;
   }
   
   private int createCaptionsPacketSubtotals(int rownum)
   {
      Row row = null;
      Cell cell = null;
      int colCnt = 20;
      int col = 0;
      int rw = rownum;


      //
      // Build the column headings      
      row = m_Sheet2.createRow(rw++);

      if ( row != null ) {
         for ( int i = 0; i < colCnt; i++ ) {
            cell = row.createCell(i);
            cell.setCellType(CellType.STRING);
            cell.getCellStyle().setFont(m_FontBold);
         }

         //these column widths are in 1/256 of a character
         //so to make it easier to read, i converted the widths to length_in_chars*256
         m_Sheet2.setColumnWidth(col, 12*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Warehouse"));

         m_Sheet2.setColumnWidth(col, 6*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Packet #"));
         
         m_Sheet2.setColumnWidth(col, 15*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Units Ordered"));
         
         m_Sheet2.setColumnWidth(col, 15*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Units Invoiced"));
         
         m_Sheet2.setColumnWidth(col, 15*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Dollars Invoiced"));

         m_Sheet2.setColumnWidth(col, 12*256);
         row.getCell(col++).setCellValue(new XSSFRichTextString("Fill Rate %"));


      }
      
      return rw;
   }
   
   
   public static void main(String args[]) {
      PromoSvcLevelDept psld = new PromoSvcLevelDept();
      psld.m_BegDate = "01/01/2014";
      psld.m_EndDate = "04/21/2014";
      psld.m_Packets = "800,801,812";
      StringBuffer fileName = new StringBuffer();      
      String tmp = Long.toString(System.currentTimeMillis());
      fileName.append("promo_svc_dept");      
      fileName.append("-");
      fileName.append(tmp.substring(tmp.length()-5, tmp.length()));
      fileName.append(".xlsx");
      psld.m_FileNames.add(fileName.toString());
      
      psld.m_FilePath = "C:\\exp\\";
      
      psld.createReport();
   }
}
