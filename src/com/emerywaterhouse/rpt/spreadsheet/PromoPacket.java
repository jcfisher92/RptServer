/**
 * File: PromoPacket.java
 * Description: Promo sales report based on packet
 *
 * @author Seth Murdock
 * $Revision: 1.18 $
 *
 * Create Date: 07/18/2006
 * Last Update: $Id: PromoPacket.java,v 1.18 2013/02/07 15:46:05 jfisher Exp $
 *
 * History
 *    $Log: PromoPacket.java,v $
 *    Revision 1.18  2013/02/07 15:46:05  jfisher
 *    Switched to xlsx to get the extra rows.  Fixed some logging and functions as well.
 *
 *    Revision 1.17  2013/02/06 21:18:17  epearson
 *    change rownum variable to int from short
 *
 *    Revision 1.16  2009/03/25 18:36:29  pdavidson
 *    Added additional cell styles for new columns
 *
 *    Revision 1.15  2009/03/25 18:24:37  pdavidson
 *    Fixed broup by clause to include new dia_date column
 *
 *    Revision 1.14  2009/03/25 18:19:02  pdavidson
 *    Added column for Emery cost (todays buy)
 *
 *    Revision 1.13  2009/03/25 17:28:50  pdavidson
 *    Added column for Emery cost (todays buy)
 *
 *    Revision 1.12  2009/03/25 17:14:09  pdavidson
 *    Added column for Emery cost (todays buy)
 *
 *    Revision 1.11  2009/03/25 16:42:45  pdavidson
 *    Added additional columns for units on order and dollars on order
 *
 *    Revision 1.10  2009/03/25 00:38:02  pdavidson
 *    Modified main query to get units and dollars on order (not yet invoiced)
 *
 *    Revision 1.9  2009/03/23 21:48:51  pdavidson
 *    Improved efficiency pf main query by adding line in where caluse that joined
 *    subquery with outer query.  This prevented a merge cartesian join.
 *    Added new columns for special cost and order deadline date.
 *
 *    Revision 1.8  2009/02/18 16:53:10  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 *    Revision 1.7  2008/10/29 21:44:58  jfisher
 *    Fixed potential null warnings.
 *
 *    Revision 1.6  2007/07/11 22:49:48  jfisher
 *    Fixed a bunch of warnings, removed unused code, added header comments that the original author left out.
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class PromoPacket extends Report
{
   private static int BASE_COLS = 21;

   private String m_PromoId;
   private String m_PacketId;
   private String m_DSBDate;
   private PreparedStatement m_GetDSBDate;
   private PreparedStatement m_PromoPacket;

   //
   // The cell styles for each of the base columns in the spreadsheet.
   private CellStyle[] m_CellStyles;

   //
   // workbook entries.
   private Workbook m_Wrkbk;
   private Sheet m_Sheet;

   //
   // Log4j logger
   private Logger m_Log;

   /**
    * default constructor
    */
   public PromoPacket()
   {
      super();
      m_Log = Logger.getLogger(RptServer.class);
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
      setupWorkbook();
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
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
    * @return true if the file was built, false if not.
    * @throws FileNotFoundException
    */

   private boolean buildOutputFile() throws FileNotFoundException
   {
      Row row = null;
      FileOutputStream outFile = null;
      ResultSet promopacket = null;
      ResultSet DontShipBefore = null;
      int colCnt = BASE_COLS;
      int rowNum = 1;
      boolean result = false;
      String Plist[];
      String OnePromo;
      double promo_margin;
      double total_margin;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         rowNum = createCaptions();
         Plist = m_PromoId.split(",");     // if all we got was one cust ID, what the hell we'll split it anyway

         for (int i = 0 ; i < Plist.length ; i++) {  // for each customer
            OnePromo = Plist[i];
            m_GetDSBDate.setString(1,OnePromo);
            DontShipBefore = m_GetDSBDate.executeQuery();

            while (DontShipBefore.next() && m_Status == RptServer.RUNNING ) {
               m_DSBDate = (DontShipBefore.getString("dsb_date"));
            }
            m_PromoPacket.setString(1, m_DSBDate);
            m_PromoPacket.setString(2, OnePromo);
            m_PromoPacket.setString(3, OnePromo);
            m_PromoPacket.setString(4, OnePromo);
            promopacket = m_PromoPacket.executeQuery();

            while ( promopacket.next() && m_Status == RptServer.RUNNING ) {
               // we must calculate margins outside of SQl due to divided by zero
               if (promopacket.getDouble("promo_base")!= 0)
                  promo_margin = 100 * ((promopacket.getDouble("promo_base") - promopacket.getDouble("promo_cost")) / promopacket.getDouble("promo_base"));
               else
                  promo_margin = 0;

               if (promopacket.getDouble("dolla_sold")!= 0)
                  total_margin = 100 * ((promopacket.getDouble("dolla_sold") - promopacket.getDouble("dolla_cost")) / promopacket.getDouble("dolla_sold"));
               else
                  total_margin = 0;

               row = createRow(rowNum, colCnt);
               row.getCell(0).setCellValue(new XSSFRichTextString(promopacket.getString("packet_id")));
               row.getCell(1).setCellValue(new XSSFRichTextString(promopacket.getString("promo_id")));
               row.getCell(2).setCellValue(new XSSFRichTextString(promopacket.getString("title")));
               row.getCell(3).setCellValue(new XSSFRichTextString(promopacket.getString("dept_num")));
               row.getCell(4).setCellValue(new XSSFRichTextString(promopacket.getString("bname")));
               row.getCell(5).setCellValue(new XSSFRichTextString(promopacket.getString("vendor_id")));
               row.getCell(6).setCellValue(new XSSFRichTextString(promopacket.getString("vname")));
               row.getCell(7).setCellValue(new XSSFRichTextString(promopacket.getString("item_id")));
               row.getCell(8).setCellValue(new XSSFRichTextString(promopacket.getString("description")));
               row.getCell(9).setCellValue(new XSSFRichTextString(promopacket.getString("todays_buy")));
               row.getCell(10).setCellValue(new XSSFRichTextString(promopacket.getString("promo_cost")));
               row.getCell(11).setCellValue(promopacket.getInt("units_on_ord"));
               row.getCell(12).setCellValue(promopacket.getDouble("dollars_on_ord"));
               row.getCell(13).setCellValue(promopacket.getInt("units_sold"));
               row.getCell(14).setCellValue(promopacket.getDouble("dolla_sold"));
               row.getCell(15).setCellValue(promopacket.getInt("number_lines"));
               row.getCell(16).setCellValue(promo_margin);
               row.getCell(17).setCellValue(total_margin);
               row.getCell(18).setCellValue(new XSSFRichTextString(promopacket.getString("promo_ship_date")));
               row.getCell(19).setCellValue(new XSSFRichTextString(promopacket.getString("terms_date")));
               row.getCell(20).setCellValue(new XSSFRichTextString(promopacket.getString("promo_dia_date")));

               rowNum++;
            }
         }

         m_Wrkbk.write(outFile);
         result = true;
      }

      catch ( Exception ex ) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         m_Log.error("[PromoPacket]", ex);
      }

      finally {
         row = null;
         closeRSet(promopacket);
         closeRSet(DontShipBefore);

         try {
            outFile.close();
         }

         catch( Exception e ) {
            m_Log.error("[PromoPacket]", e);
         }

         outFile = null;
      }

      return result;
   }

   /**
    * Builds the sql based on the type of filter requested by the user.
    * @return A complete sql statement.
    * SHM 08/30/2006:
    * This thing gets all items for a promo, looks in inv_dtl to get totals for sold
    * items (see sums and counts in the inline "invd" select), then adds those totals to
    * other item data ( see "max(decode" in the outer select).
    *
    * PD 3/24/09
    *    Modified to use ansi join syntax and also fixed merge cartisian join to inner inv_dtl query.
    *
    * Report also displays margins but those are calulated when the spreadsheet is built.
    *
    * SQL below returns data for ONE promo, code in buildOutputFile calls it as many times
    * as necessary for a given packet, based on which promos the user selected to report on
    * in the app.
    */
   private String buildSql()
   {
      StringBuffer sql = new StringBuffer();
      sql.append("select ");
      sql.append("   item_entity_attr.item_id, packet_id, promotion.promo_id, title, ");
      sql.append("   item_entity_attr.description, vendor.vendor_id, vendor.name as vname, ");
      sql.append("   emery_dept.dept_num, buyer.name as bname, ");
      sql.append("   nvl(max(invd.total_dollars_sold), 0) as dolla_sold,  "); // using aggregates, by item, from the inline select below
      sql.append("   nvl(max(invd.total_dollars_cost), 0) as dolla_cost, ");
      sql.append("   nvl(max(invd.qty), 0) as units_sold, ");
      sql.append("   max(invd.orders) as number_lines, ");
      sql.append("   nvl(max(ordd.units_on_order), 0) as units_on_ord, ");
      sql.append("   nvl(max(ordd.dollars_on_order), 0) as dollars_on_ord, ");
      sql.append("   ejd_item_price.buy as todays_buy, ");
      sql.append("   promo_base, promo_cost, ");
      sql.append("   to_char(ship_date, 'mm/dd/yyyy') as promo_ship_date, ");
      sql.append("   to_char(dia_date, 'mm/dd/yyyy') as promo_dia_date, ");
      sql.append("   to_char(ejd.terms_procs.get_date(promotion.term_id, null), 'mm/dd/yyyy') as terms_date ");
      sql.append("from promo_item ");
      sql.append("join promotion on promo_item.promo_id = promotion.promo_id ");
      sql.append("join item_entity_attr on item_entity_attr.item_ea_id = promo_item.item_ea_id ");
      sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
      sql.append("join ejd_item_price on ejd_item_price.ejd_item_id = ejd_item.ejd_item_id and ejd_item_price.warehouse_id = promotion.warehouse_id ");
      sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
      sql.append("join emery_dept on emery_dept.dept_id = ejd_item.dept_id ");
      sql.append("left outer join buyer on emery_dept.buyer_id = buyer.buyer_id ");
      sql.append("join ( ");
      sql.append("   select ");
      sql.append("   distinct(inv_dtl.item_ea_id) as item_ea_id, ");
      sql.append("   sum(ext_sell) over (partition by inv_dtl.item_ea_id) as total_dollars_sold, ");
      sql.append("   sum(ext_cost) over (partition by inv_dtl.item_ea_id) as total_dollars_cost, ");
      sql.append("   sum(qty_shipped) over (partition by inv_dtl.item_ea_id) as qty, ");
      sql.append("   count(inv_dtl_id) over (partition by inv_dtl.item_ea_id) as orders ");
      sql.append("   from promo_item ");
      sql.append("   left outer join inv_dtl on inv_dtl.promo_nbr = promo_item.promo_id and inv_dtl.item_ea_id = promo_item.item_ea_id and ");
      sql.append("         inv_dtl.invoice_date >= to_date(?, 'mm/dd/yyyy') ");
      sql.append("   where promo_item.promo_id = ? ");
      sql.append(") invd on promo_item.item_ea_id = invd.item_ea_id ");
      sql.append("left outer join ( ");
      sql.append("   select item_ea_id, sum(qty_ordered) as units_on_order, sum(qty_ordered * sell_price) as dollars_on_order ");
      sql.append("   from order_line ");
      sql.append("   join order_header on order_header.order_id = order_line.order_id ");
      sql.append("   join order_status oh_status on order_header.order_status_id = oh_status.order_status_id and oh_status.description not in ('CANCELLED', 'COMPLETE') ");
      sql.append("   join order_status ol_status on order_line.order_status_id = ol_status.order_status_id and ol_status.description <> 'CANCELLED' ");
      sql.append("   where order_line.invoice_num is null and order_line.promo_id = ? ");
      sql.append("   group by item_ea_id ");
      sql.append(") ordd on promo_item.item_ea_id = ordd.item_ea_id ");
      sql.append("where ");
      sql.append("promo_item.promo_id = ? ");
      sql.append("group by ");
      sql.append("   item_entity_attr.item_id, packet_id, promotion.promo_id, ");
      sql.append("   title, description, vendor.vendor_id, vendor.name, ");
      sql.append("   to_char(ship_date, 'mm/dd/yyyy'), ");
      sql.append("   to_char(dia_date, 'mm/dd/yyyy'), ");
      sql.append("   to_char(ejd.terms_procs.get_date(promotion.term_id, null), 'mm/dd/yyyy'), ");
      sql.append("   emery_dept.dept_num, buyer.name, promo_base, promo_cost, ejd_item_price.buy ");
      sql.append("order by promo_id, emery_dept.dept_num, vendor.name, item_entity_attr.item_id ");
            
      return sql.toString();
   }

   /**
    *  Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_PromoPacket);
   }

   /**
    * Creates the report title and the captions.
    */
   private int createCaptions()
   {
      Font fontTitle;
      CellStyle styleTitle;   // Bold, centered
      CellStyle styleTitleLeft;   // Bold, Left Justified
      Row row = null;
      Cell cell = null;
      int rowNum = 0;
      StringBuffer caption = new StringBuffer("Promotion Sales Report: ");

      if ( m_Sheet != null ) {
         fontTitle = m_Wrkbk.createFont();
         fontTitle.setFontHeightInPoints((short) 10);
         fontTitle.setFontName("Arial");
         fontTitle.setBold(true);

         styleTitle = m_Wrkbk.createCellStyle();
         styleTitle.setFont(fontTitle);
         styleTitle.setAlignment(HorizontalAlignment.CENTER);

         styleTitleLeft = m_Wrkbk.createCellStyle();
         styleTitleLeft.setFont(fontTitle);
         styleTitleLeft.setAlignment(HorizontalAlignment.LEFT);

         //
         // set the report title
         row = m_Sheet.createRow(rowNum);
         cell = row.createCell(0);
         cell.setCellType(CellType.STRING);
         cell.setCellStyle(styleTitleLeft);

         if ( m_PacketId != null && m_PacketId.length() > 0){
            caption.append("Packet ");
            caption.append(m_PacketId);
         }
         else
            caption.append(m_PromoId);

         cell.setCellValue(new XSSFRichTextString(caption.toString()));

         rowNum = 2;
         row = m_Sheet.createRow(rowNum);

         try {
            if ( row != null ) {
               for ( int i = 0; i < BASE_COLS; i++ ) {
                  cell = row.createCell(i);
                  cell.setCellStyle(styleTitleLeft);
               }

               row.getCell(0).setCellValue(new XSSFRichTextString("Packet #"));
               row.getCell(1).setCellValue(new XSSFRichTextString("Promo #"));
               row.getCell(2).setCellValue(new XSSFRichTextString("Promo Title"));
               m_Sheet.setColumnWidth(2, 14000);
               row.getCell(3).setCellValue(new XSSFRichTextString("Dept #"));
               row.getCell(4).setCellValue(new XSSFRichTextString("Buyer Name"));
               m_Sheet.setColumnWidth(4, 7000);
               row.getCell(5).setCellValue(new XSSFRichTextString("Vendor Number"));
               m_Sheet.setColumnWidth(5, 2000);
               row.getCell(6).setCellValue(new XSSFRichTextString("Vendor Name"));
               m_Sheet.setColumnWidth(6, 7000);
               row.getCell(7).setCellValue(new XSSFRichTextString("Item ID"));
               m_Sheet.setColumnWidth(7, 3000);
               row.getCell(8).setCellValue(new XSSFRichTextString("Item Description"));
               m_Sheet.setColumnWidth(8, 14000);
               row.getCell(9).setCellValue(new XSSFRichTextString("Emery Cost"));
               m_Sheet.setColumnWidth(9, 2000);
               row.getCell(10).setCellValue(new XSSFRichTextString("Special Cost"));
               m_Sheet.setColumnWidth(10, 2000);
               row.getCell(11).setCellValue(new XSSFRichTextString("Units on Order"));
               m_Sheet.setColumnWidth(11, 2000);
               row.getCell(12).setCellValue(new XSSFRichTextString("$ on Order"));
               m_Sheet.setColumnWidth(12, 3000);
               row.getCell(13).setCellValue(new XSSFRichTextString("Units Sold"));
               m_Sheet.setColumnWidth(13, 2000);
               row.getCell(14).setCellValue(new XSSFRichTextString("$ Sold"));
               m_Sheet.setColumnWidth(14, 3000);
               row.getCell(15).setCellValue(new XSSFRichTextString("Lines"));
               m_Sheet.setColumnWidth(15, 2000);
               row.getCell(16).setCellValue(new XSSFRichTextString("GM %"));
               m_Sheet.setColumnWidth(16, 3000);
               row.getCell(17).setCellValue(new XSSFRichTextString("Total GM %"));
               m_Sheet.setColumnWidth(17, 3000);
               row.getCell(18).setCellValue(new XSSFRichTextString("Approx. Ship Date"));
               m_Sheet.setColumnWidth(18, 3000);
               row.getCell(19).setCellValue(new XSSFRichTextString("Terms Due Date"));
               m_Sheet.setColumnWidth(19, 3000);
               row.getCell(20).setCellValue(new XSSFRichTextString("Order Deadline Date"));
               m_Sheet.setColumnWidth(20, 3000);
            }
         }

         finally {
            row = null;
            cell = null;
            fontTitle = null;
            styleTitle = null;
            caption = null;
         }
      }

      return ++rowNum;
   }

   /**
    * Creates a row in the worksheet.
    * @param rowNum The row number.
    * @param colCnt The number of columns in the row.
    *
    * @return The formatted row of the spreadsheet.
    */
   private Row createRow(int rowNum, int colCnt)
   {
      Row row = null;
      Cell cell = null;

      if ( m_Sheet != null ) {
         row = m_Sheet.createRow(rowNum);

         //
         // set the type and style of the cell.
         if ( row != null ) {
            for ( int i = 0; i < colCnt; i++ ) {
               cell = row.createCell(i);
               cell.setCellStyle(m_CellStyles[i]);
            }
         }
      }

      return row;
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;

      try {                  
         prepareStatements();
         created = buildOutputFile();
      }

      catch ( Exception ex ) {
         m_Log.fatal("[PromoPacket]", ex);
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
         m_PromoPacket = m_EdbConn.prepareStatement(buildSql());
         m_GetDSBDate = m_EdbConn.prepareStatement("select to_char(dsb_date, 'mm/dd/yyyy') dsb_date from promotion where promo_id = ?");
      }
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    *
    * Because it's possible that this report can be called from some other system, the
    * best way to deal with params is to not go by the order, but by the name.
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fname = new StringBuffer();
      String tm = Long.toString(System.currentTimeMillis()).substring(3);
      int pcount = params.size();
      Param param = null;

      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);

         if ( param.name.equals("promo") )
            m_PromoId = param.value;

         if ( param.name.equals("packet") )
            m_PacketId = param.value;
      }

      //
      // Build the file name.
      fname.append(tm);
      fname.append("-");
      fname.append(m_RptProc.getUid());
      fname.append("pp.xlsx");

      m_FileNames.add(fname.toString());
   }

   /**
    * Sets up the styles for the cells based on the column data.  Does any other initialization
    * needed by the workbook.
    */
   private void setupWorkbook()
   {
      CellStyle styleText;      // Text right justified
      CellStyle styleInt;       // Style with 0 decimals
      CellStyle styleMoney;     // Money ($#,##0.00_);[Red]($#,##0.00)
      CellStyle stylePct;       // Style with 0 decimals + %

      styleText = m_Wrkbk.createCellStyle();
      styleText.setAlignment(HorizontalAlignment.LEFT);

      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short)3);

      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short)8);

      stylePct = m_Wrkbk.createCellStyle();
      stylePct.setAlignment(HorizontalAlignment.RIGHT);
      stylePct.setDataFormat((short)2);

      m_CellStyles = new CellStyle[] {
            styleText,    // col 0 packet id
            styleText,    // col 1 promo id
            styleText,    // col 2 title
            styleText,    // col 3 dept number
            styleText,    // col 4 buyer name
            styleText,    // col 5 vendor id
            styleText,    // col 6 vendor name
            styleText,    // col 7 item id
            styleText,    // col 8 item desc
            styleMoney,   // col 9 Emery cost
            styleMoney,   // col 10 Promo cost
            styleInt,     // col 11 units on order
            styleMoney,   // col 12 dollars on order
            styleInt,     // col 13 units sold
            styleMoney,   // col 14 dollars sold
            styleInt,     // col 15 number of lines
            stylePct,     // col 16 gross margin at promor prices
            stylePct,     // col 17 gross margin for actual promo cost and base
            styleText,    // col 18 ship date
            styleText,    // col 19 terms due date
            styleText,    // col 20 order deadline date
      };

      styleText = null;
      styleInt = null;
      styleMoney = null;
      stylePct = null;
   }

}
