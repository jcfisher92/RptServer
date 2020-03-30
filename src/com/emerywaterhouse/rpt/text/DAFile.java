/**
 * File: DAFile.java
 * Description: Month end file build for DA.  Needs to be emailed as and attachment.
 *    The report is a fixed length file based on predefined file layout.  This is the reworked
 *    report class from the original mt server.
 *
 * @author Jeffrey Fisher
 *
 * Create Data: 04/05/2005
 * Last Update: $Id: DAFile.java,v 1.12 2012/07/11 19:15:37 jfisher Exp $
 *
 * History
 *    $Log: DAFile.java,v $
 *    Revision 1.12  2012/07/11 19:15:37  jfisher
 *    in_catalog modification
 *
 *    Revision 1.11  2009/06/29 18:47:06  npasnur
 *    Replaced catalog_item with bmi_item
 *
 *    Revision 1.10  2008/10/29 21:33:51  jfisher
 *    Fixed some warnings and added cvs logging
 *
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Arrays;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;


public class DAFile extends Report
{
   private PreparedStatement m_CatData;
   private PreparedStatement m_ItemData;
   private PreparedStatement m_PMItemSales;
   private PreparedStatement m_UnitsSold;
   private PreparedStatement m_UpcData;
   private PreparedStatement m_SenCode;

   private java.util.Date m_Date;

   /**
    * default constructor
    */
   public DAFile()
   {
      super();

      m_FileNames.add("dafile.txt");
      m_Date = new java.util.Date();
      m_MaxRunTime = RptServer.HOUR * 12;              // This takes a long time to run, we'll start at 7 hours.
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not guaranteed to know when finalization occurs.
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      m_Date = null;
      m_ItemData = null;
      m_PMItemSales = null;
      m_UnitsSold = null;
      m_UpcData = null;
      m_CatData = null;
      m_SenCode = null;

      super.finalize();
   }

   /**
    * Builds the output file based on the query selection criteria.
    *
    * @return  boolean
    *    true if the file was created.
    *    false if there was some sort of error.
    */
   private boolean buildOutputFile()
   {
      StringBuffer Line = new StringBuffer(1024);
      char[] Filler = new char[266];
      FileOutputStream OutFile = null;
      SimpleDateFormat df = null;
      ResultSet ItemData = null;
      String ItemId;
      String ItemDesc;
      String RetC;
      String Sell;
      String Buy;
      String VndName;
      String FileDate;
      int BcId;
      int StockPack;
      String Tmp;
      boolean Result = false;

      try {
         OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
         df = new SimpleDateFormat("MM/dd/yy");
         FileDate = df.format(m_Date);
         ItemData = m_ItemData.executeQuery();

         //
         // Prefill the with spaces.
         Arrays.fill(Filler, ' ');

         while ( ItemData.next() && m_Status == RptServer.RUNNING ) {
            Line.setLength(0);
            Line.append(Filler);

            ItemId = ItemData.getString(1);
            ItemDesc = ItemData.getString(6);
            setCurAction("processing item: " + ItemId);

            RetC = Double.toString(ItemData.getDouble("retail"));
            if ( RetC.length() > 8 )
               RetC = RetC.substring(0, 8);

            Sell = Double.toString(ItemData.getDouble("sell"));
            if ( Sell.length() > 8 )
               Sell = Sell.substring(0, 8);

            Buy = Double.toString(ItemData.getDouble("buy"));
            if ( Buy.length() > 8 )
               Buy = RetC.substring(0, 8);

            VndName = ItemData.getString("vnd_name");
            BcId = ItemData.getInt("broken_case_id");
            StockPack = ItemData.getInt("stock_pack");

            if ( ItemDesc.length() > 40 )
               ItemDesc = ItemDesc.substring(0, 40);

            if ( VndName.length() > 30 )
               VndName = VndName.substring(0, 30);

            Line.replace(0, 5, "EMERY");
            Line.replace(5, 5 + ItemId.length(), ItemId);

            Tmp = ItemData.getString("vendor_item_num");
            if ( Tmp.length() > 14 )
               Tmp = Tmp.substring(0, 14);

            Line.replace(19, 19 + Tmp.length(), Tmp);
            Tmp = getUpc(ItemId);
            Line.replace(33, 33 + Tmp.length(), Tmp);
            Tmp = ItemData.getString(3);
            Line.replace(45, 45 + Tmp.length(), Tmp);          // Nrha
            Tmp = ItemData.getString(4);
            Line.replace(47, 47 + Tmp.length(), Tmp);          // Mdc
            Tmp = ItemData.getString(5);
            Line.replace(50, 50 + Tmp.length(), Tmp);          // flc
            Line.replace(56, 56 + ItemDesc.length(), ItemDesc);
            Tmp = getCatPage(ItemId);
            Line.replace(96, 96 + Tmp.length(), Tmp);         // Catalog Page
            Line.replace(102, 102 + RetC.length(), RetC);
            Line.replace(118, 118 + Sell.length(), Sell);
            Tmp = ItemData.getString(11);
            Line.replace(158, 158 + Tmp.length(), Tmp);
            Line.replace(162, 162 + Tmp.length(), Tmp);

            //
            // Min qty.  1 unless nbc, then the stock pack
            Tmp = (BcId == 1) ? String.valueOf(StockPack) : "1";
            Line.replace(166, 166 + Tmp.length(), Tmp);

            Line.replace(171, 171 + VndName.length(), VndName);
            Tmp = getPMUnitsSold(ItemId);
            Line.replace(201, 201 + Tmp.length(), Tmp);
            Tmp = getYTDUnitsSold(ItemId);
            Line.replace(210, 210 + Tmp.length(), Tmp);
            Tmp = getSenCode(ItemId);
            Line.replace(217, 217 + Tmp.length(), Tmp);

            //
            // Handle the broken case designator.  If no broken cases then we
            // have to use NB + the stock pack.  Shouldn't really use the id, but it's quick.
            if ( BcId != 1 ) {
               Tmp = "NB" + StockPack;
               Line.replace(219, 219 + Tmp.length(), Tmp);
            }

            Line.replace(226, 226 + FileDate.length(), FileDate);
            Line.replace(250, 250 + Buy.length(), Buy);

            //
            // Chop off anything that might have gone past our line limit.
            if (Line.length() > 266 )
               Line.setLength(266);

            Line.append("\r\n");
            OutFile.write(Line.toString().getBytes());
         }

         Result = true;
      }

      catch ( Exception ex ) {
         log.error("exception:", ex);
      }

      return Result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_ItemData);
      closeStmt(m_PMItemSales);
      closeStmt(m_UnitsSold);
      closeStmt(m_UpcData);
      closeStmt(m_CatData);
      closeStmt(m_SenCode);
   }

   /**
    * Implements the base class abstract method.  Creates a connection to Oracle for data.
    * Then builds the output file.
    *
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    * @see com.emerywaterhouse.rpt.text.DAFile#buildOutputFile()
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
         log.fatal("exception:", ex);
      }

      finally {
        closeStatements();

        if ( m_Status == RptServer.RUNNING )
           m_Status = RptServer.STOPPED;
      }

      return created;
   }

   /**
    * Returns the catalog page number if one exists for the item.
    *
    * @param   itemId - The id of the item that the catalog page will be retrieved.
    *
    * @return  The catalog page for the item specified in itemId.
    */
   private String getCatPage(String itemId)
   {
      String Page = null;
      ResultSet Rs = null;

      if ( itemId != null && m_CatData != null) {
         try {
            m_CatData.setString(1, itemId);
            Rs = m_CatData.executeQuery();

            if ( Rs.next() )
               Page = Rs.getString(1);

            if ( Page == null )
               Page = "";
            else {
               if ( Page.length() > 6 )
                  Page = Page.substring(0, 6);
            }
         }

         catch ( Exception ex ) {
            log.error("exception:", ex);
         }

         finally {
            if ( Rs != null ) {
               try {
                  Rs.close();
               }

               catch ( Exception e ) {};
               Rs = null;
            }
         }
      }

      return Page;
   }

   /**
    * Gets the sensitivity code for and item.
    *
    * @param   itemId - The id of the item that the sen code will be retrieved.
    * @return  String - The sensitivity code for the item in itemId.
    */
   private String getSenCode(String itemId)
   {
      String SenCode = null;
      ResultSet Rs = null;

      if ( itemId != null && m_SenCode != null) {
         try {
            m_SenCode.setString(1, itemId);
            Rs = m_SenCode.executeQuery();

            if ( Rs != null && Rs.next() )
               SenCode = String.valueOf(Rs.getInt(1));

            if ( SenCode == null )
               SenCode = " ";
            else {
               if ( SenCode.length() > 2 )
                  SenCode = SenCode.substring(0, 2);
            }
         }

         catch ( Exception ex ) {
            log.error("exception:", ex);
         }

         finally {
            if ( Rs != null ) {
               try {
                  Rs.close();
               }

               catch ( Exception e) {};
               Rs = null;
            }
         }
      }

      return SenCode;
   }

   /**
    * Returns the upc code for an item.  The length of the upc returned can be a max
    * of 12 chars.
    *
    * @param   itemId - The id of the item to locate the upc for.
    * @return  String - the upc for the item in itemId
    */
   private String getUpc(String itemId)
   {
      String Upc = null;
      ResultSet Rs = null;

      if ( itemId != null && m_UpcData != null) {
         try {
            m_UpcData.setString(1, itemId);
            Rs = m_UpcData.executeQuery();

            if ( Rs != null && Rs.next() )
               Upc = Rs.getString(1);

            if ( Upc == null )
               Upc = "";
            else {
               if ( Upc.length() > 12 )
                  Upc = Upc.substring(0, 12);
            }
         }

         catch ( Exception ex ) {
            log.error("exception:", ex);
         }

         finally {
            if ( Rs != null ) {
               try {
                  Rs.close();
               }

               catch ( Exception e) {};
               Rs = null;
            }
         }
      }

      return Upc;
   }

   /**
    * Returns the prior months units sold.
    * @param   itemId -  the id of the item we are trying to get the units sold.
    * @return  The number of units sold as a string.
    */
   private String getPMUnitsSold(String itemId)
   {
      String Qty = null;
      ResultSet Rs = null;

      if ( itemId != null && m_PMItemSales != null) {
         try {
            m_PMItemSales.setString(1, itemId);
            Rs = m_PMItemSales.executeQuery();

            if ( Rs != null && Rs.next() )
               Qty = String.valueOf(((int)Rs.getDouble(1)));

            if ( Qty == null )
               Qty = "0";
            else {
               if ( Qty.length() > 8 )
                  Qty = Qty.substring(0, 8);
            }
         }

         catch ( Exception ex ) {
            log.error("exception:", ex);
         }

         finally {
            if ( Rs != null ) {
               try {
                  Rs.close();
               }

               catch ( Exception e) {};
               Rs = null;
            }
         }
      }

      return Qty;
   }

   /**
    * Returns the last 12 months of the units sold for the specified item.
    *
    * @param   itemId - the emerey item number to calculate the units sold
    *
    * @return  String representing the number of units sold.  The return value
    *    is converted to a string since the output has to go into a delimited file.
    */
   private String getYTDUnitsSold(String itemId)
   {
      String Qty = null;
      ResultSet Rs = null;
      if ( itemId != null && m_UnitsSold != null) {
         try {
            m_UnitsSold.setString(1, itemId);
            Rs = m_UnitsSold.executeQuery();

            if ( Rs != null && Rs.next() )
               Qty = String.valueOf(((int)Rs.getDouble(1)));

            if ( Qty == null )
               Qty = "0";
            else {
               if ( Qty.length() > 8 )
                  Qty = Qty.substring(0, 8);
            }
         }

         catch ( Exception ex ) {
            log.error("exception:", ex);
         }

         finally {
            if ( Rs != null ) {
               try {
                  Rs.close();
               }

               catch ( Exception e) {};
               Rs = null;
            }
         }
      }

      return Qty;
   }

   /**
    * Prepares the sql queries for execution.
    * @throws  SQLException
    */
   private boolean prepareStatements() throws SQLException
   {
      boolean prepared = false;
      StringBuffer sql = new StringBuffer();

      if ( m_OraConn != null ) {
         sql.append("select ");
         sql.append("   item.item_id, vendor_item_num, mdc.nrha_id, flc.mdc_id, ");
         sql.append("   item.flc_id, item.description, ");
         sql.append("   item_price_procs.todays_retailc(item.item_id) retail, ");
         sql.append("   item_price_procs.todays_sell(item.item_id) sell, ");
         sql.append("   item_price_procs.todays_buy(item.item_id) buy, vendor.name vnd_name, ship_unit.unit uom, ");
         sql.append("    broken_case_id, stock_pack ");
         sql.append("from ");
         sql.append("   item ");
         sql.append("join item_warehouse on item_warehouse.item_id = item.item_id and item_warehouse.in_catalog = 1 ");
         sql.append("join warehouse on warehouse.warehouse_id = item_warehouse.warehouse_id and warehouse.name = 'PORTLAND' ");
         sql.append("join vendor on vendor.vendor_id = item.vendor_id ");
         sql.append("join vendor_item_cross vic on vic.item_id = item.item_id and vic.vendor_id = item.vendor_id ");
         sql.append("join flc on flc.flc_id = item.flc_id ");
         sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
         sql.append("join ship_unit on ship_unit.unit_id = item.ship_unit_id ");
         sql.append("where ");
         sql.append("   item.disp_id = 1");

         m_ItemData = m_OraConn.prepareStatement(sql.toString());

         m_UpcData = m_OraConn.prepareStatement(
            "select upc_code from item_upc where item_id = ? and primary_upc = 1"
         );

         m_CatData = m_OraConn.prepareStatement(
            "select page from bmi_item where item_id = ?"
         );

         m_PMItemSales = m_OraConn.prepareStatement(
            "select sum(qty_shipped) qty from itemsales where item_nbr = ? and " +
            "invoice_date > add_months(sysdate,-2) and invoice_date < add_months(sysdate,-1)"
         );

         m_UnitsSold = m_OraConn.prepareStatement(
            "select sum(qty_shipped) qty from itemsales where item_nbr = ? and " +
            "invoice_date >= add_months(trunc(sysdate),-12)"
         );

         m_SenCode = m_OraConn.prepareStatement(
            "select sen_code_id from item_price " +
            "where price_id = item_price_procs.todays_sell_id(?)"
         );

         prepared = true;
      }

      sql = null;

      return prepared;
   }
}
