/**
 * File: CustSales.java
 * Description: The customer sales excel spreadsheet.
 *
 * @author Seth Murdock
 *
 * Create Date: 03/21/2006
 * Last Update: $Id: CustSales.java,v 1.13 2010/03/20 18:36:58 smurdock Exp $
 * 
 * History
 *    $Log: CustSales.java,v $
 *    Revision 1.13  2010/03/20 18:36:58  smurdock
 *    oops query from item_price was getting future price dates -- not now
 *
 *    Revision 1.12  2010/03/20 16:18:31  smurdock
 *    added sell_date of most recent price update from table item_price
 *
 *    Revision 1.11  2009/02/18 14:20:42  jfisher
 *    Fixed depricated methods after poi upgrade
 *
 */
package com.emerywaterhouse.rpt.spreadsheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class CustSales extends Report {
   private static short BASE_COLS = 19;
   private int colCnt = BASE_COLS;
   private Boolean ONE_CUSTOMER;
   private Boolean ACCOUNT_TOTAL_ONLY;

   private String m_CustId;
   private String m_BegDate;
   private String m_EndDate;
   private String m_ConsolidatedCustID;
   private String m_CustAccount;
   private String m_CustList;
   private String m_MerchClass;
   private String m_FlcId;
   private String m_NrhaId;
   private String m_VndId;
   private String m_Warehouse; // FASCOR id , sez Jeff. From Delphi.
   private String m_Accpac_Warehouse;
   private PreparedStatement m_CustSales;
   private PreparedStatement m_GetConsolidatedCustID;
   private PreparedStatement m_GetCustNames;
   private PreparedStatement m_GetChildAccounts;
   private PreparedStatement m_GetAccPacWarehouse;
   private ArrayList<String> m_AccountList = new ArrayList<String>();

   //
   // The cell styles for each of the base columns in the spreadsheet.
   private ArrayList<XSSFCellStyle> m_CellStyles = new ArrayList<XSSFCellStyle>();

   //
   // workbook entries.
   private XSSFWorkbook m_Wrkbk;
   private XSSFSheet m_Sheet;

   //
   // Log4j logger
   private Logger m_Log;

   /**
    * default constructor
    */
   public CustSales() {
      super();
      m_Log = Logger.getLogger(RptServer.class);
      m_Wrkbk = new XSSFWorkbook();
      m_Sheet = m_Wrkbk.createSheet();
   }

   /**
    * This report changes its column setup based on whether it is for one
    * customer or several; we have to determine which report type it is before
    * we can do jack.
    * 
    * colCnt, in a report by account, can vary depending on the number
    * requested so we set that here too.
    * 
    */
   private void determineCustomers() {
      String SplitCustList[];
      ResultSet ChildAccounts = null;

      ONE_CUSTOMER = ((m_CustList.length() == 0) && (m_CustAccount.length() == 0));
      ACCOUNT_TOTAL_ONLY = ((m_CustAccount.length() > 0) && (m_CustList.length() == 0));

      if (ONE_CUSTOMER) {
         SplitCustList = m_CustId.split(","); // all we got was one cust
         // ID,what the hell we'll
         // split it anyway
         m_AccountList.add(SplitCustList[0]);
         colCnt = 19;
      } else {
         if (ACCOUNT_TOTAL_ONLY) {
            try {
               m_GetChildAccounts.setString(1, m_CustAccount);
               m_GetChildAccounts.setString(2, m_CustAccount);
               ChildAccounts = m_GetChildAccounts.executeQuery();

               while (ChildAccounts.next() && m_Status == RptServer.RUNNING) {
                  m_AccountList.add(ChildAccounts.getString("customer_id"));
               }

               colCnt = 17;
            }

            catch (Exception ex) {
               m_ErrMsg.append(ex.getClass().getName() + "\r\n");
               m_ErrMsg.append(ex.getMessage());
               m_Log.error("exception", ex);
            }

            finally {
               closeRSet(ChildAccounts);
            }
         } else {
            colCnt = 15;
            SplitCustList = m_CustList.split(","); // we are splitting a
            // comma
            // delimited list of
            // customers
            for (int i = 0; i < SplitCustList.length; i++) { // for each
               // customer
               m_AccountList.add(SplitCustList[i]); // building an
               // ArrayList of
               // customers
               colCnt++;
            }
         }
      }
   }

   /**
    * Cleanup any allocated resources.
    * 
    * @throws Throwable
    */
   public void finalize() throws Throwable {
      if (m_CellStyles.size() > 0) {
         m_CellStyles.clear();
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
   private boolean buildOutputFile() throws FileNotFoundException {
      XSSFRow row = null;
      FileOutputStream outFile = null;
      ResultSet custSales = null;
      ResultSet ConsolidatedID = null;
      int colNum = 0;
      int rowNum = 1;
      boolean result = false;
      double cost;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         rowNum = createCaptions();
         m_GetConsolidatedCustID.setString(1, m_AccountList.get(0));
         ConsolidatedID = m_GetConsolidatedCustID.executeQuery();
         m_GetAccPacWarehouse.setString(1, m_Warehouse);

         while (ConsolidatedID.next() && m_Status == RptServer.RUNNING) {
            if (ConsolidatedID.getString("cust_cons_id") != null)
               m_ConsolidatedCustID = ConsolidatedID.getString("cust_cons_id");
            else
               m_ConsolidatedCustID = m_AccountList.get(0);
         }

         if (ONE_CUSTOMER) {
            m_CustSales.setString(1, m_CustId);
            m_CustSales.setString(2, m_CustId);
            m_CustSales.setString(3, m_CustId);
            m_CustSales.setString(4, m_BegDate);
            m_CustSales.setString(5, m_EndDate);
            m_CustSales.setString(6, m_ConsolidatedCustID);
         } else { // set the params for customers by account
            m_CustSales.setString(1, m_BegDate);
            m_CustSales.setString(2, m_EndDate);
            m_CustSales.setString(3, m_ConsolidatedCustID);
         }

         custSales = m_CustSales.executeQuery();

         if (ONE_CUSTOMER) {
            while (custSales.next() && m_Status == RptServer.RUNNING) {
               if (custSales.getDouble("avgcost") != 0)
                  cost = custSales.getDouble("avgcost");
               else if (custSales.getDouble("lastcost") != 0)
                  cost = custSales.getDouble("lastcost");
               else
                  cost = custSales.getDouble("buy");

               row = createRow(rowNum);
               row.getCell(0).setCellValue(new XSSFRichTextString(custSales.getString("cust_nbr")));
               row.getCell(1).setCellValue(new XSSFRichTextString(custSales.getString("customer")));
               row.getCell(2).setCellValue(new XSSFRichTextString(custSales.getString("item_nbr")));
               row.getCell(3).setCellValue(new XSSFRichTextString(custSales.getString("customer_sku")));
               row.getCell(4).setCellValue(new XSSFRichTextString(custSales.getString("vendor_id")));
               row.getCell(5).setCellValue(new XSSFRichTextString(custSales.getString("vendor")));
               row.getCell(6).setCellValue(new XSSFRichTextString(custSales.getString("nrha_id")));
               row.getCell(7).setCellValue(new XSSFRichTextString(custSales.getString("mdc_id")));
               row.getCell(8).setCellValue(new XSSFRichTextString(custSales.getString("flc_id")));
               row.getCell(9).setCellValue(new XSSFRichTextString(custSales.getString("upc_code")));
               row.getCell(10).setCellValue(new XSSFRichTextString(custSales.getString("vendor_item_num")));
               row.getCell(11).setCellValue(new XSSFRichTextString(custSales.getString("description")));
               row.getCell(12).setCellValue(new XSSFRichTextString(custSales.getString("sell_date")));

               if (custSales.getDouble("sell_price") != 0)
                  row.getCell(13).setCellValue(custSales.getDouble("sell_price"));

               if (custSales.getDouble("retail_price") != 0)
                  row.getCell(14).setCellValue(custSales.getDouble("retail_price"));

               row.getCell(15).setCellValue(cost);
               row.getCell(16).setCellValue(custSales.getDouble("sell"));
               row.getCell(17).setCellValue(custSales.getDouble("qty_shipped"));
               row.getCell(18).setCellValue(custSales.getDouble("num_orders"));

               rowNum++;
            }
         } else if (ACCOUNT_TOTAL_ONLY) { // process for customers by account
            while (custSales.next() && m_Status == RptServer.RUNNING) {
               if (custSales.getDouble("avgcost") != 0)
                  cost = custSales.getDouble("avgcost");
               else if (custSales.getDouble("lastcost") != 0)
                  cost = custSales.getDouble("lastcost");
               else
                  cost = custSales.getDouble("buy");

               row = createRow(rowNum);
               row.getCell(0).setCellValue(new XSSFRichTextString(custSales.getString("item_nbr")));
               row.getCell(1).setCellValue(new XSSFRichTextString(custSales.getString("customer_sku")));
               row.getCell(2).setCellValue(new XSSFRichTextString(custSales.getString("vendor_id")));
               row.getCell(3).setCellValue(new XSSFRichTextString(custSales.getString("vendor")));
               row.getCell(4).setCellValue(new XSSFRichTextString(custSales.getString("nrha_id")));
               row.getCell(5).setCellValue(new XSSFRichTextString(custSales.getString("mdc_id")));
               row.getCell(6).setCellValue(new XSSFRichTextString(custSales.getString("flc_id")));
               row.getCell(7).setCellValue(new XSSFRichTextString(custSales.getString("upc_code")));
               row.getCell(8).setCellValue(new XSSFRichTextString(custSales.getString("vendor_item_num")));
               row.getCell(9).setCellValue(new XSSFRichTextString(custSales.getString("description")));
               row.getCell(10).setCellValue(new XSSFRichTextString(custSales.getString("sell_date")));

               if (custSales.getDouble("sell_price") != 0)
                  row.getCell(11).setCellValue(custSales.getDouble("sell_price"));

               if (custSales.getDouble("retail_price") != 0)
                  row.getCell(12).setCellValue(custSales.getDouble("retail_price"));

               row.getCell(13).setCellValue(cost);
               row.getCell(14).setCellValue(custSales.getDouble("sell"));
               row.getCell(15).setCellValue(custSales.getDouble("qty_shipped"));
               row.getCell(16).setCellValue(custSales.getDouble("num_orders"));

               rowNum++;
            }
         }

         else { // process for customers by account
            while (custSales.next() && m_Status == RptServer.RUNNING) {
               if (custSales.getDouble("avgcost") != 0)
                  cost = custSales.getDouble("avgcost");
               else if (custSales.getDouble("lastcost") != 0)
                  cost = custSales.getDouble("lastcost");
               else
                  cost = custSales.getDouble("buy");

               row = createRow(rowNum);
               row.getCell(0).setCellValue(new XSSFRichTextString(custSales.getString("item_nbr")));
               row.getCell(1).setCellValue(new XSSFRichTextString(custSales.getString("customer_sku")));
               row.getCell(2).setCellValue(new XSSFRichTextString(custSales.getString("vendor_id")));
               row.getCell(3).setCellValue(new XSSFRichTextString(custSales.getString("vendor")));
               row.getCell(4).setCellValue(new XSSFRichTextString(custSales.getString("nrha_id")));
               row.getCell(5).setCellValue(new XSSFRichTextString(custSales.getString("mdc_id")));
               row.getCell(6).setCellValue(new XSSFRichTextString(custSales.getString("flc_id")));
               row.getCell(7).setCellValue(new XSSFRichTextString(custSales.getString("upc_code")));
               row.getCell(8).setCellValue(new XSSFRichTextString(custSales.getString("vendor_item_num")));
               row.getCell(9).setCellValue(new XSSFRichTextString(custSales.getString("description")));
               row.getCell(10).setCellValue(new XSSFRichTextString(custSales.getString("sell_date")));

               if (custSales.getDouble("sell_price") != 0)
                  row.getCell(11).setCellValue(custSales.getDouble("sell_price"));

               if (custSales.getDouble("retail_price") != 0)
                  row.getCell(12).setCellValue(custSales.getDouble("retail_price"));

               row.getCell(13).setCellValue(cost);
               row.getCell(14).setCellValue(custSales.getDouble("sell"));
               colNum = 15;

               for (int i = 0; i < (m_AccountList.size()); i++) {
                  row.getCell(colNum).setCellValue(custSales.getDouble("Cust" + m_AccountList.get(i)));
                  colNum++;
               }

               rowNum++;
            }
         }

         m_Wrkbk.write(outFile);
         result = true;
      }

      catch (Exception ex) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         m_Log.error("exception", ex);
      }

      finally {
         row = null;
         closeRSet(custSales);
         closeRSet(ConsolidatedID);
         try {
            outFile.close();
         }

         catch (Exception e) {
            m_Log.error("exception:", e);
         }

         outFile = null;
      }

      return result;
   }

   private String buildSql() {
      StringBuffer sql = new StringBuffer(256);
      StringBuffer sql_start = new StringBuffer(256);
      StringBuffer sql_from = new StringBuffer(256);
      StringBuffer sql_where = new StringBuffer(256);
      StringBuffer sql_inline = new StringBuffer(256);
      StringBuffer sql_group = new StringBuffer(256);

      if (ONE_CUSTOMER)
         sql_start.append("select invd.cust_nbr, cust.name as customer, invd.item_nbr,   \r\n");
      else
         sql_start.append("select invd.item_nbr,   \r\n");
      sql_start.append("   item_ea_cross.customer_sku,   \r\n");
      sql_start.append("   vend.vendor_id, vend.name as vendor,   \r\n");
      sql_start.append("   nrha.nrha_id, mdc.mdc_id, flc.flc_id, ");
      sql_start.append("   ejd_item_whs_upc.upc_code,   \r\n");
      sql_start.append("   vendor_item_ea_cross.vendor_item_num,   \r\n");
      sql_start.append("   item_entity_attr.description,   \r\n");
      sql_start.append("   to_char(ejd_item_price.update_time,'YYYY-MM-DD') as sell_date,   \r\n");
      if (ONE_CUSTOMER) {
         sql_start.append("   max(decode(cust_nbr,?,qty,null)) as qty_shipped,   \r\n");
         sql_start.append("   max(decode(cust_nbr,?,numb,null)) as num_orders,   \r\n");
      }
      if (ACCOUNT_TOTAL_ONLY) {
         sql_start.append("   max(decode(item_nbr,item_nbr,qty,null)) as qty_shipped,   \r\n");
         sql_start.append("   max(decode(item_nbr,item_nbr,numb,null)) as num_orders,   \r\n");
      } else
         // select qty-shipped for each customer in our account list, call
         // each
         // total "CustXXXXXX"
         for (int i = 0; i < m_AccountList.size(); i++) {
            sql_start.append("   max(decode(cust_nbr,");
            sql_start.append(m_AccountList.get(i));
            sql_start.append(",qty,null)) as Cust");
            sql_start.append(m_AccountList.get(i));
            sql_start.append(",   \r\n");
         }
      sql_start.append("   (select price from ejd_cust_procs.get_sell_price(invd.cust_nbr, invd.item_ea_id)) as sell_price,   \r\n");
      sql_start.append("   ejd_price_procs.get_retail_price(invd.cust_nbr, invd.item_ea_id) as retail_price,   \r\n");
      sql_start.append("   ejd_item_price.buy,   \r\n");
      sql_start.append("   ejd_item_price.sell,   \r\n");
      sql_start.append("   totalcost, qtyonhand, lastcost, decode(qtyonhand, 0, 0, round((totalcost / qtyonhand)::numeric, 3)) as avgcost   \r\n");

      sql_inline.append("from  \r\n");
      sql_inline.append("   (select distinct item_ea_id, item_nbr, cust_nbr, \r\n");

      if (ONE_CUSTOMER)
         sql_inline.append("   count(inv_dtl_id) over (partition by cust_nbr, item_nbr, item_ea_id) as numb, \r\n");

      if (ACCOUNT_TOTAL_ONLY) {
         sql_inline.append("   sum(qty_shipped) over (partition by item_nbr) as qty, \r\n");
         sql_inline.append("   count(inv_dtl_id) over (partition by item_nbr) as numb \r\n");
      } else
         sql_inline.append("   sum(qty_shipped) over (partition by cust_nbr, item_nbr) as qty \r\n");

      sql_inline.append("   from inv_dtl   \r\n");

      //
      // we join by DC to exclude the other DC (if we want the report by DC,
      // that is)
      if (m_Warehouse != null && m_Warehouse.length() > 0) {
         sql_inline.append("join warehouse on inv_dtl.warehouse = warehouse.name and warehouse.fas_facility_id = '")
         .append(m_Warehouse).append("' \r\n");
      }
      sql_inline.append("   where exists (\r\n");
      sql_inline.append("      select inv_hdr_id\r\n");
      sql_inline.append("      from inv_hdr\r\n");

      if (ONE_CUSTOMER)
         sql_inline.append("      where cust_nbr = ? and \r\n");
      else { // select on all of our multiple customers
         sql_inline.append("      where cust_nbr in ( ");

         for (int i = 0; i < m_AccountList.size(); i++) {
            if (i == m_AccountList.size() - 1)
               sql_inline.append(String.format("'%s') and ", m_AccountList.get(i)));
            else
               sql_inline.append(String.format("'%s', ", m_AccountList.get(i)));
         }
      }

      sql_inline.append("(invoice_date between to_date(?,'mm/dd/yyyy') and to_date(?,'mm/dd/yyyy')) and \r\n");
      sql_inline.append("inv_dtl.inv_hdr_id = inv_hdr.inv_hdr_id)) invd, \r\n");

      if (ONE_CUSTOMER)
         sql_from.append("   customer cust, ");

      sql_from.append("   vendor vend, item_entity_attr   \r\n");
      sql_from.append("left outer join item_ea_cross on item_entity_attr.item_ea_id = item_ea_cross.item_ea_id and item_ea_cross.customer_id = ? \r\n");
      sql_from.append("left outer join vendor_item_ea_cross on item_entity_attr.item_ea_id = vendor_item_ea_cross.item_ea_id and item_entity_attr.vendor_id = vendor_item_ea_cross.vendor_id \r\n");
      sql_from.append("left outer join ejd_item_whs_upc on item_entity_attr.ejd_item_id = ejd_item_whs_upc.ejd_item_id and primary_upc = 1 ");

      if ( m_Warehouse != null && m_Warehouse.length() > 0 )
         sql_from.append("and ejd_item_whs_upc.warehouse_id = " + m_Warehouse + " \r\n");
      else 
         sql_from.append("and ejd_item_whs_upc.warehouse_id in (1,2) \r\n");

      sql_from.append("join ejd_item_price on ejd_item_price.ejd_item_id = item_entity_attr.ejd_item_id ");

      if (m_Warehouse != null && m_Warehouse.length() > 0)
         sql_from.append("and ejd_item_price.warehouse_id = " + m_Warehouse + " \r\n");
      else 
         sql_from.append("and ejd_item_price.warehouse_id in (1,2) \r\n");
      

      sql_from.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");

      if (m_Accpac_Warehouse != null && m_Accpac_Warehouse.length() > 0) {
      	sql_from.append("join ejd.sage300_iciloc_mv iciloc on iciloc.itemno = item_entity_attr.item_id and iciloc.location = '").append(m_Accpac_Warehouse).append("',\r\n");
      } 
      else { // get both DC's
         sql_from.append("join (select itemno, sum(totalcost) as totalcost, sum(qtyonhand) as qtyonhand, \r\n");
         sql_from.append("   max(lastcost) as lastcost from iciloc group by itemno) iciloc on iciloc.itemno = item_entity_attr.item_id, \r\n");
      }

      sql_from.append(" nrha, mdc, flc \r\n");
      sql_where.append("where invd.item_ea_id = item_entity_attr.item_ea_id   \r\n");

      if (ONE_CUSTOMER)
         sql_where.append("      and invd.cust_nbr = cust.customer_id   \r\n");

      sql_where.append("      and item_entity_attr.vendor_id = vend.vendor_id   \r\n");
      sql_where.append("      and ejd_item.flc_id = flc.flc_id   \r\n");
      sql_where.append("      and flc.mdc_id = mdc.mdc_id   \r\n");
      sql_where.append("      and mdc.nrha_id = nrha.nrha_id   \r\n");

      if (m_VndId != null && m_VndId.length() > 0) {
         sql_where.append("      and item_entity_attr.vendor_id = ");
         sql_where.append(m_VndId);
      }

      if (m_FlcId != null && m_FlcId.length() > 0) {
         sql_where.append(" and ejd_item.flc_id = ");
         sql_where.append(m_FlcId);
      }

      if (m_NrhaId != null && m_NrhaId.length() > 0) {
         sql_where.append(" and mdc.nrha_id = ");
         sql_where.append(m_NrhaId);
      }

      if (m_MerchClass != null && m_MerchClass.length() > 0) {
         sql_where.append(" and flc.mdc_id = ");
         sql_where.append(m_MerchClass);
      }

      sql_where.append("   and iciloc.itemno = item_entity_attr.item_id   \r\n");
      sql_group.append("group by   \r\n");

      if (ONE_CUSTOMER)
         sql_group.append("   invd.cust_nbr, cust.name, ");

      sql_group.append("   item_nbr, vend.vendor_id, vend.name,   \r\n");
      sql_group.append("   nrha.nrha_id, mdc.mdc_id, flc.flc_id,   \r\n");
      sql_group.append("   (select price from ejd_cust_procs.get_sell_price(invd.cust_nbr, invd.item_ea_id)),   \r\n");
      sql_group.append("   ejd_price_procs.get_retail_price(invd.cust_nbr,invd.item_ea_id),   \r\n");
      sql_group.append("   ejd_item_price.buy,   \r\n");
      sql_group.append("   ejd_item_price.sell,   \r\n");
      sql_group.append("   item_entity_attr.description, ejd_item_price.update_time, customer_sku, vendor_item_ea_cross.vendor_item_num,   \r\n");
      sql_group.append("   item_nbr, ejd_item_whs_upc.upc_code,   \r\n");
      sql_group.append("   totalcost, lastcost, qtyonhand, invd.cust_nbr , invd.item_ea_id");

      sql.append(sql_start);
      sql.append(sql_inline);
      sql.append(sql_from);
      sql.append(sql_where);
      sql.append(sql_group);

      return sql.toString();
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements() {
      closeStmt(m_CustSales);
      closeStmt(m_GetConsolidatedCustID);
   }

   private String getCustomerName(String custo) {
      ResultSet CustName = null;
      String ThisCust = null;

      try {
         m_GetCustNames.setString(1, custo);
         CustName = m_GetCustNames.executeQuery();
         while (CustName.next() && m_Status == RptServer.RUNNING) {
            if (CustName.getString("name") != null)
               ThisCust = CustName.getString("name");
            else
               ThisCust = "";
         }
      } catch (Exception ex) {
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
         m_Log.error("exception", ex);
      }

      finally {
         closeRSet(CustName);
      }

      return (ThisCust);

   }

   /**
    * Creates the report title and the captions.
    */
   private int createCaptions() {
      XSSFFont fontTitle;
      XSSFCellStyle styleTitle; // Bold, centered
      XSSFCellStyle styleTitleLeft; // Bold, Left Justified
      XSSFRow row = null;
      XSSFCell cell = null;
      int rowNum = 0;
      int colNum;
      StringBuffer caption = new StringBuffer("Customer Sales Report: ");
      String CustNameString = "";

      if (m_Sheet == null)
         return 0;

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

      if (m_CustAccount != null && m_CustAccount.length() > 0) {
         caption.append("Account ");
         caption.append(m_CustAccount);
      } else
         caption.append(m_CustId);

      caption.append(" ");
      caption.append(m_BegDate);
      caption.append(" - ");
      caption.append(m_EndDate);

      if (m_Warehouse != null && m_Warehouse.length() > 0) {
         caption.append(" ");
         caption.append(m_Warehouse);
      }

      cell.setCellValue(new XSSFRichTextString(caption.toString()));

      if (ACCOUNT_TOTAL_ONLY) {
         caption.setLength(0);
         rowNum++;
         row = m_Sheet.createRow(rowNum);
         cell = row.createCell(0);
         cell.setCellType(CellType.STRING);
         cell.setCellStyle(styleTitleLeft);
         caption.append(getCustomerName(m_CustAccount));
         cell.setCellValue(new XSSFRichTextString(caption.toString()));
      }

      if ((m_VndId != null && m_VndId.length() > 0) || (m_FlcId != null && m_FlcId.length() > 0)
            || (m_MerchClass != null && m_MerchClass.length() > 0) || (m_NrhaId != null && m_NrhaId.length() > 0)) {
         rowNum++;
         caption.setLength(0);
         row = m_Sheet.createRow(rowNum);
         cell = row.createCell(0);
         cell.setCellType(CellType.STRING);
         cell.setCellStyle(styleTitleLeft);

         if (m_VndId != null && m_VndId.length() > 0) {
            caption.append(" Vendor: ");
            caption.append(m_VndId);
         }

         if (m_FlcId != null && m_FlcId.length() > 0) {
            caption.append(" FLC: ");
            caption.append(m_FlcId);
         }

         if (m_MerchClass != null && m_MerchClass.length() > 0) {
            caption.append(" MDC: ");
            caption.append(m_MerchClass);
         }

         if (m_NrhaId != null && m_NrhaId.length() > 0) {
            caption.append(" NRHA: ");
            caption.append(m_NrhaId);
         }

         cell.setCellValue(new XSSFRichTextString(caption.toString()));
      }

      rowNum++;
      rowNum++;
      row = m_Sheet.createRow(rowNum);

      try {
         if (row != null) {
            for (int i = 0; i < colCnt; i++) {
               cell = row.createCell(i);
               cell.setCellStyle(styleTitleLeft);
            }

            if (ONE_CUSTOMER) {
               row.getCell(0).setCellValue(new XSSFRichTextString("Cust ID"));
               row.getCell(1).setCellValue(new XSSFRichTextString("Customer Name"));
               m_Sheet.setColumnWidth(1, 14000);
               row.getCell(2).setCellValue(new XSSFRichTextString("Item #"));
               row.getCell(3).setCellValue(new XSSFRichTextString("Cust SKU"));
               row.getCell(4).setCellValue(new XSSFRichTextString("Vendor ID"));
               row.getCell(5).setCellValue(new XSSFRichTextString("Vendor Name"));
               m_Sheet.setColumnWidth(5, 14000);
               row.getCell(6).setCellValue(new XSSFRichTextString("NRHA"));
               m_Sheet.setColumnWidth(6, 2000);
               row.getCell(7).setCellValue(new XSSFRichTextString("MDC"));
               m_Sheet.setColumnWidth(7, 2000);
               row.getCell(8).setCellValue(new XSSFRichTextString("FLC"));
               m_Sheet.setColumnWidth(8, 2000);
               row.getCell(9).setCellValue(new XSSFRichTextString("UPC Primary"));
               m_Sheet.setColumnWidth(9, 5000);
               row.getCell(10).setCellValue(new XSSFRichTextString("Mfr. Part"));
               m_Sheet.setColumnWidth(10, 7000);
               row.getCell(11).setCellValue(new XSSFRichTextString("Item Description"));
               m_Sheet.setColumnWidth(11, 14000);
               row.getCell(12).setCellValue(new XSSFRichTextString("Price Date"));
               m_Sheet.setColumnWidth(12, 3000);
               row.getCell(13).setCellValue(new XSSFRichTextString("Cust Cost"));
               m_Sheet.setColumnWidth(13, 3000);
               row.getCell(14).setCellValue(new XSSFRichTextString("Cust Retail"));
               m_Sheet.setColumnWidth(14, 3000);
               row.getCell(15).setCellValue(new XSSFRichTextString("Emery Cost"));
               m_Sheet.setColumnWidth(15, 3000);
               row.getCell(16).setCellValue(new XSSFRichTextString("Base Cost"));
               m_Sheet.setColumnWidth(16, 3000);
               row.getCell(17).setCellValue(new XSSFRichTextString("Qty Purch"));
               m_Sheet.setColumnWidth(17, 3000);
               row.getCell(18).setCellValue(new XSSFRichTextString("# of Orders"));
               m_Sheet.setColumnWidth(18, 3000);
            } else if (ACCOUNT_TOTAL_ONLY) {
               row.getCell(0).setCellValue(new XSSFRichTextString("Item #"));
               row.getCell(1).setCellValue(new XSSFRichTextString("Cust SKU"));
               row.getCell(2).setCellValue(new XSSFRichTextString("Vendor ID"));
               row.getCell(3).setCellValue(new XSSFRichTextString("Vendor Name"));
               m_Sheet.setColumnWidth(3, 14000);
               row.getCell(4).setCellValue(new XSSFRichTextString("NRHA"));
               m_Sheet.setColumnWidth(4, 2000);
               row.getCell(5).setCellValue(new XSSFRichTextString("MDC"));
               m_Sheet.setColumnWidth(5, 2000);
               row.getCell(6).setCellValue(new XSSFRichTextString("FLC"));
               m_Sheet.setColumnWidth(6, 2000);
               row.getCell(7).setCellValue(new XSSFRichTextString("UPC Primary"));
               m_Sheet.setColumnWidth(7, 5000);
               row.getCell(8).setCellValue(new XSSFRichTextString("Mfr. Part"));
               m_Sheet.setColumnWidth(8, 7000);
               row.getCell(9).setCellValue(new XSSFRichTextString("Item Description"));
               m_Sheet.setColumnWidth(9, 14000);
               row.getCell(10).setCellValue(new XSSFRichTextString("Price Date"));
               m_Sheet.setColumnWidth(10, 3000);
               row.getCell(11).setCellValue(new XSSFRichTextString("Cust Cost"));
               m_Sheet.setColumnWidth(11, 3000);
               row.getCell(12).setCellValue(new XSSFRichTextString("Cust Retail"));
               m_Sheet.setColumnWidth(12, 3000);
               row.getCell(13).setCellValue(new XSSFRichTextString("Emery Cost"));
               m_Sheet.setColumnWidth(13, 3000);
               row.getCell(14).setCellValue(new XSSFRichTextString("Base Cost"));
               m_Sheet.setColumnWidth(14, 3000);
               row.getCell(15).setCellValue(new XSSFRichTextString("Qty Purch"));
               m_Sheet.setColumnWidth(15, 3000);
               row.getCell(16).setCellValue(new XSSFRichTextString("# of Orders"));
               m_Sheet.setColumnWidth(16, 3000);
            } else { // multiple customers
               row.getCell(0).setCellValue(new XSSFRichTextString("Item #"));
               row.getCell(1).setCellValue(new XSSFRichTextString("Cust SKU"));
               row.getCell(2).setCellValue(new XSSFRichTextString("Vendor ID"));
               row.getCell(3).setCellValue(new XSSFRichTextString("Vendor Name"));
               m_Sheet.setColumnWidth(3, 14000);
               row.getCell(4).setCellValue(new XSSFRichTextString("NRHA"));
               m_Sheet.setColumnWidth(4, 2000);
               row.getCell(5).setCellValue(new XSSFRichTextString("MDC"));
               m_Sheet.setColumnWidth(5, 2000);
               row.getCell(6).setCellValue(new XSSFRichTextString("FLC"));
               m_Sheet.setColumnWidth(6, 2000);
               row.getCell(7).setCellValue(new XSSFRichTextString("UPC Primary"));
               m_Sheet.setColumnWidth(7, 5000);
               row.getCell(8).setCellValue(new XSSFRichTextString("Mfr. Part"));
               m_Sheet.setColumnWidth(8, 7000);
               row.getCell(9).setCellValue(new XSSFRichTextString("Item Description"));
               m_Sheet.setColumnWidth(9, 14000);
               row.getCell(10).setCellValue(new XSSFRichTextString("Price Date"));
               m_Sheet.setColumnWidth(10, 3000);
               row.getCell(11).setCellValue(new XSSFRichTextString("Cust Cost"));
               m_Sheet.setColumnWidth(11, 3000);
               row.getCell(12).setCellValue(new XSSFRichTextString("Cust Retail"));
               m_Sheet.setColumnWidth(12, 3000);
               row.getCell(13).setCellValue(new XSSFRichTextString("Emery Cost"));
               m_Sheet.setColumnWidth(13, 3000);
               row.getCell(14).setCellValue(new XSSFRichTextString("Base Cost"));
               m_Sheet.setColumnWidth(14, 3000);
               colNum = 15;

               for (int i = 0; i < (m_AccountList.size()); i++) {
                  CustNameString = getCustomerName(m_AccountList.get(i));
                  row.getCell(colNum)
                  .setCellValue(new XSSFRichTextString(m_AccountList.get(i) + " " + CustNameString));
                  m_Sheet.setColumnWidth(colNum, 3000);
                  colNum++;
               }
            }
         }
      }

      finally {
         row = null;
         cell = null;
         fontTitle = null;
         styleTitle = null;
         caption = null;
      }

      return ++rowNum;
   }

   /**
    * Creates a row in the worksheet.
    * 
    * @param rowNum
    *            The row number.
    * 
    * @return The formatted row of the spreadsheet.
    */
   private XSSFRow createRow(int rowNum) {
      XSSFRow row = null;
      XSSFCell cell = null;

      if (m_Sheet == null)
         return row;

      row = m_Sheet.createRow(rowNum);

      //
      // set the type and style of the cell.
      if (row != null) {
         for (int i = 0; i < colCnt; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(m_CellStyles.get(i));
         }
      }

      return row;
   }

   public boolean createReport() {
      boolean created = false;
      m_Status = RptServer.RUNNING;

      try {
         m_EdbConn = m_RptProc.getEdbConn();
         prepareStatements();
         setupWorkbook();

         created = buildOutputFile();
      }

      catch (Exception ex) {
         m_Log.fatal("[CustSales]", ex);
      }

      finally {
         closeStatements();

         if (m_Status == RptServer.RUNNING)
            m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Prepares the sql queries for execution.
    * 
    */
   private void prepareStatements() throws Exception {
      StringBuffer sql = new StringBuffer(256);
      ResultSet AccPacJac = null;

      if (m_EdbConn != null) {
         //
         // Has to be before determineCustomers, this is used there.
         sql.setLength(0);
         sql.append("select customer_id from customer where customer_id = ? or parent_id = ?");
         m_GetChildAccounts = m_EdbConn.prepareStatement(sql.toString());

         //
         // we need Accpac warehouse id and we need it NOW
         sql.setLength(0);
         sql.append("select accpac_wh_id from warehouse where fas_facility_id = ?");
         m_GetAccPacWarehouse = m_EdbConn.prepareStatement(sql.toString());

         //
         // once per run, find the consolidated account ID for customer sku
         // lookup
         sql.setLength(0);
         sql.append("select cust_cons_id from cust_consolidate ");
         sql.append("where customer_id =  ?");
         sql.append(" and cons_type_id in ");
         sql.append("(select cons_type_id from consolidate_type ");
         sql.append("where description = 'ITEM XREF')");
         m_GetConsolidatedCustID = m_EdbConn.prepareStatement(sql.toString());

         //
         // Has to be before "ONE_CUSTOMER" usage because it gets set in this
         // method. Thank you Seth.
         determineCustomers();

         //
         // for multiple customer reports, we need to get the cust names for
         // the
         // column headers, once per run
         if (!ONE_CUSTOMER) {
            sql.setLength(0);
            sql.append("select name from customer where customer_id = ?");
            m_GetCustNames = m_EdbConn.prepareStatement(sql.toString());
         }

         try {
            m_GetAccPacWarehouse.setString(1, m_Warehouse);
            AccPacJac = m_GetAccPacWarehouse.executeQuery();

            while (AccPacJac.next() && m_Status == RptServer.RUNNING) {
               m_Accpac_Warehouse = AccPacJac.getString("accpac_wh_id");
            }
         } catch (Exception ex) {
            m_ErrMsg.append(ex.getClass().getName() + "\r\n");
            m_ErrMsg.append(ex.getMessage());
            m_Log.error("exception", ex);
         }

         finally {
            closeRSet(AccPacJac);
         }

         // our main big honking query gets put together in buildsql
         sql.setLength(0);
         m_CustSales = m_EdbConn.prepareStatement(buildSql());
      }
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    * 
    *      Because it's possible that this report can be called from some other
    *      system, the best way to deal with params is to not go by the order,
    *      but by the name.
    * 
    *      07/06 shm: did away with "rolling twelve months" stuff. Send it a
    *      begin and end date and let the app call it whatever name it wants.
    * 
    *      This program expects either a customer or a custlist (a comma
    *      delimited string of ID's), if it gets both it uses custlist and
    *      customer is ignored. (Custlist, if it exists, gets put into an
    *      ArrayList for processing.)
    * 
    *      Mandatory params are customer OR custlist, begdate and enddate.
    *      Program does not check for these, just goes to hell or something.
    * 
    *      There are no checks on date ranges.
    * 
    * 
    */
   public void setParams(ArrayList<Param> params) {
      StringBuffer fname = new StringBuffer();
      String tm = Long.toString(System.currentTimeMillis()).substring(3);
      int pcount = params.size();
      Param param = null;

      for (int i = 0; i < pcount; i++) {
         param = params.get(i);

         if (param.name.equals("dc"))
            m_Warehouse = param.value;

         if (param.name.equals("customer"))
            m_CustId = param.value;

         if (param.name.equals("account"))
            m_CustAccount = param.value;

         if (param.name.equals("merchclass"))
            m_MerchClass = param.value;

         if (param.name.equals("nrha"))
            m_NrhaId = param.value;

         if (param.name.equals("flc"))
            m_FlcId = param.value;

         if (param.name.equals("vendor") && param.value.trim().length() > 0)
            m_VndId = (param.value);

         if (param.name.equals("custlist"))
            m_CustList = param.value;

         if (param.name.equals("begdate"))
            m_BegDate = param.value;

         if (param.name.equals("enddate"))
            m_EndDate = param.value;
      }

      //
      // Build the file name.
      fname.append(tm);
      fname.append("-");
      fname.append(m_RptProc.getUid());
      fname.append("cs.xlsx");
      m_FileNames.add(fname.toString());
   }

   /**
    * Sets up the styles for the cells based on the column data. Does any other
    * inititialization needed by the workbook.
    */
   private void setupWorkbook() 
   {
      XSSFCellStyle styleText; // Text right justified
      XSSFCellStyle styleInt; // Style with 0 decimals
      XSSFCellStyle styleMoney; // Money ($#,##0.00_);[Red]($#,##0.00)
      XSSFCellStyle stylePct; // Style with 0 decimals + %

      styleText = m_Wrkbk.createCellStyle();
      // styleText.setFont(m_FontData);
      styleText.setAlignment(HorizontalAlignment.LEFT);

      styleInt = m_Wrkbk.createCellStyle();
      styleInt.setAlignment(HorizontalAlignment.RIGHT);
      styleInt.setDataFormat((short) 3);

      styleMoney = m_Wrkbk.createCellStyle();
      styleMoney.setAlignment(HorizontalAlignment.RIGHT);
      styleMoney.setDataFormat((short) 8);

      stylePct = m_Wrkbk.createCellStyle();
      stylePct.setAlignment(HorizontalAlignment.RIGHT);
      stylePct.setDataFormat((short) 9);

      if (ONE_CUSTOMER) {
         m_CellStyles.add(styleText); // col 0 customer id
         m_CellStyles.add(styleText); // col 1 customer name
         m_CellStyles.add(styleText); // col 2 item id
         m_CellStyles.add(styleText); // col 3 cust SKU
         m_CellStyles.add(styleText); // col 4 vendor id
         m_CellStyles.add(styleText); // col 5 vendor name
         m_CellStyles.add(styleText); // col 6 nrha dept
         m_CellStyles.add(styleText); // col 7 mdse class
         m_CellStyles.add(styleText); // col 8 flc
         m_CellStyles.add(styleText); // col 9 upc primary
         m_CellStyles.add(styleText); // col 10 mfr part
         m_CellStyles.add(styleText); // col 11 item description
         m_CellStyles.add(styleText); // col 12 price change date
         m_CellStyles.add(styleMoney); // col 13 cust cost
         m_CellStyles.add(styleMoney); // col 14 cust retail
         m_CellStyles.add(styleMoney); // col 15 emery cost
         m_CellStyles.add(styleMoney); // col 16 base cost
         m_CellStyles.add(styleInt); // col 17 qty shipped
         m_CellStyles.add(styleInt); // col 18 qty ordered
      } else {
         if (ACCOUNT_TOTAL_ONLY) {
            m_CellStyles.add(styleText); // col 0 item id
            m_CellStyles.add(styleText); // col 1 cust SKU
            m_CellStyles.add(styleText); // col 2 vendor id
            m_CellStyles.add(styleText); // col 3 vendor name
            m_CellStyles.add(styleText); // col 4 nrha dept
            m_CellStyles.add(styleText); // col 5 mdse class
            m_CellStyles.add(styleText); // col 6 flc
            m_CellStyles.add(styleText); // col 7 upc primary
            m_CellStyles.add(styleText); // col 8 mfr part
            m_CellStyles.add(styleText); // col 9 item description
            m_CellStyles.add(styleText); // col 10 price change date
            m_CellStyles.add(styleMoney); // col 11 cust cost
            m_CellStyles.add(styleMoney); // col 12 cust retail
            m_CellStyles.add(styleMoney); // col 13 emery cost
            m_CellStyles.add(styleMoney); // col 14 base cost
            m_CellStyles.add(styleInt); // col 15 qty shipped
            m_CellStyles.add(styleInt); // col 16 qty ordered
         } else { // multiple customers
            m_CellStyles.add(styleText); // col 0 item id
            m_CellStyles.add(styleText); // col 1 cust SKU
            m_CellStyles.add(styleText); // col 2 vendor id
            m_CellStyles.add(styleText); // col 3 vendor name
            m_CellStyles.add(styleText); // col 4 nrha dept
            m_CellStyles.add(styleText); // col 5 mdse class
            m_CellStyles.add(styleText); // col 6 flc
            m_CellStyles.add(styleText); // col 7 upc primary
            m_CellStyles.add(styleText); // col 8 mfr part
            m_CellStyles.add(styleText); // col 9 item description
            m_CellStyles.add(styleText); // col 10 price change date
            m_CellStyles.add(styleMoney); // col 11 cust cost
            m_CellStyles.add(styleMoney); // col 12 cust retail
            m_CellStyles.add(styleMoney); // col 13 emery cost
            m_CellStyles.add(styleMoney); // col 14 base cost

            for (int i = 0; i < (m_AccountList.size()); i++) {
               m_CellStyles.add(styleInt); // col 15 and up: qty shipped to
               // each
               // customer, # of customers can
               // vary
            }
         }
      }

      styleText = null;
      styleInt = null;
      styleMoney = null;
      stylePct = null; 
   }


   /**
    * Main method for testing the Rep Shipment output.
    * Can supply a LogDate here if desired for testing the queries on a specific date.
    * @param args
    *
   public static void main(String args[]) {
      CustSales cs = new CustSales();

      Param p1 = new Param();
      p1.name = "dc";
      p1.value = "01";
      Param p2 = new Param();
      p2.name = "account";
      p2.value = "001732";
      Param p3 = new Param();
      p3.name = "custlist";
      p3.value = "001716,001732,001741,031925,031933,059510,063941,068403";
      Param p4 = new Param();
      p4.name = "begdate";
      p4.value = "06/26/2017";
      Param p5 = new Param();
      p5.name = "enddate";
      p5.value = "07/26/2017";

      ArrayList<Param> params = new ArrayList<Param>();
      //params.add(p1);
      params.add(p2);
      params.add(p3);
      params.add(p4);
      params.add(p5);

      cs.m_FilePath = "C:\\Exp\\";

   	java.util.Properties connProps = new java.util.Properties();
   	connProps.put("user", "ejd");
   	connProps.put("password", "boxer");
   	try {
   		cs.m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33:5444/emery_jensen",connProps);

   		cs.setParams(params);
   		cs.createReport();
   	} catch (Exception e) {
   		e.printStackTrace();
   	}
   }*/

}
