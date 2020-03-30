/**
 * File: TelxonFile.java
 * Description: class descendant that generates a Telxon file.
 *    Telxon file has a fixed column format with each line = 108 bytes
 *    Rewritten to work with the new report server.
 *    Original author was Paul Davidson
 *
 * @author Paul Davidson
 * @author Jeffrey Fisher
 *
 * Create Date: 05/20/2005
 * Last Update: $Id: TelxonFile.java,v 1.10 2012/07/11 19:15:53 jfisher Exp $
 *
 * History
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Arrays;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;


public class TelxonFile extends Report
{
   private PreparedStatement m_ItemData;

   /**
    * default constructor
    */
   public TelxonFile()
   {
      super();

      //
      // Set the time stamp for this file build.  The time stamp will be appended to
      // the end of the file name and will prevent concurrent writes to the same file
      m_FileNames.add("telxon_" + Long.toString(System.currentTimeMillis()) + ".txt");
   }

   /**
    * Builds the fixed column width Telxon file.
    * Record definition: ( total character width = 108 )
    *    item nbr         -  7 bytes
    *    upc              -  12
    *    bytesspaces      -  2 bytes
    *    vendor item nbr  -  15 bytes
    *    item desc        -  35 bytes
    *    ship unit        -  3 bytes
    *    stock pack       -  6 bytes
    *    retail pack      -  6 bytes
    *    nbc              -  1 byte
    *    fine line class  -  4 bytes
    *    spaces           -  2 bytes
    *    nrha             -  2 bytes
    *    spaces           -  4 bytes
    *    retail price     -  9 bytes ( inc. decimal point )
    *
    * @return true if the file was created, false if there was an error.
    */
   private boolean buildOutputFile()
   {
      StringBuffer Line = new StringBuffer();
      char[] Filler = new char[108];
      FileOutputStream OutFile = null;
      ResultSet itemData = null;
      String itemId;
      String upc;
      String vndItemNum;
      String itemDesc;
      String shipUnit;
      String stockPack;
      String retailPack;
      String nbc;
      String flc;
      String nrha;
      String retailC;
      double retlRnd;
      boolean result = false;

      try
      {
         setCurAction("building telxon file");
         Arrays.fill(Filler, ' ');
         OutFile = new FileOutputStream(m_FilePath + m_FileNames.get(0) , false);

         if ( m_Status == RptServer.RUNNING ) {
            itemData = m_ItemData.executeQuery();

            while ( itemData.next() && m_Status == RptServer.RUNNING ) {
               Line.setLength(0);
               Line.append(Filler);

               //
               // get the field values from the current record buffer
               itemId = itemData.getString(1);

               upc = itemData.getString(2);
               if ( upc == null )
                  upc = "";
               else {
                  if ( upc.length() > 12 )
                     upc = upc.substring(0, 12);
               }

               vndItemNum = itemData.getString(3);
               if ( vndItemNum == null )
                  vndItemNum = "";
               else {
                  if ( vndItemNum.length() > 15 )
                     vndItemNum = vndItemNum.substring(0, 15);
               }

               itemDesc = itemData.getString(4);
               if ( itemDesc.length() > 35 )
                  itemDesc = itemDesc.substring(0, 35);

               shipUnit = itemData.getString(5);

               stockPack = String.valueOf( itemData.getInt(6) );
               if ( stockPack.length() > 6 )
                  stockPack = stockPack.substring(0, 6);

               retailPack = String.valueOf( itemData.getInt(7) );
               if ( retailPack.length() > 6 )
                  retailPack = retailPack.substring(0, 6);

               nbc = itemData.getString(8);
               flc = itemData.getString(9);

               nrha = itemData.getString(10);
               //
               // convert to int and back to String so that any leading zeroes are removed
               nrha = String.valueOf( Integer.parseInt(nrha) );

               //
               // make sure retail price is rounded to 2 decimal places
               retlRnd = Math.floor( itemData.getDouble(11) * 100 + .5 ) / 100;
               retailC = Double.toString( retlRnd );
               if ( retailC.length() > 9 )
                  retailC = retailC.substring(0, 9);

               //
               // place the field values into the Line string buffer at the right positions
               Line.replace(0, itemId.length(), itemId);
               Line.replace(7, 7 + upc.length(), upc);
               Line.replace(19, 21, "  ");
               Line.replace(21, 21 + vndItemNum.length(), vndItemNum);
               Line.replace(36, 36 + itemDesc.length(), itemDesc);
               Line.replace(71, 71 + shipUnit.length(), shipUnit);
               Line.replace(80 - stockPack.length(), 80, stockPack);  // right justified
               Line.replace(86 - retailPack.length(), 86, retailPack); // right justified
               Line.replace(86, 86 + nbc.length(), nbc);
               Line.replace(87, 87 + flc.length(), flc);
               Line.replace(91, 93, "  ");
               Line.replace(95 - nrha.length(), 95, nrha); // right justified
               Line.replace(95, 99, "    ");
               Line.replace(108 - retailC.length(), 108, retailC); // right justified

               //
               // Chop off anything that might have gone past our line limit.
               if ( Line.length() > 108 )
                  Line.setLength(108);

               Line.append("\r\n");
               OutFile.write(Line.toString().getBytes());

               setCurAction("building telxon file - item# " + itemId);
            }

            result = true;
         }
      }

      catch ( Exception ex )
      {
         log.error("TelxonFile: buildOutputFile: " + ex.getClass().getName() +
            " "  + ex.getMessage());
      }

      finally {
         if ( OutFile != null ) {
            try {
               OutFile.close();
            }
            catch ( IOException iex ) {
               log.error("TelxonFile: buildOutputFile: " + iex.getClass().getName() +
                  " "  + iex.getMessage());
            }
         }

         OutFile = null;

         if ( itemData != null ) {
            try {
               itemData.close();
               itemData = null;
            }
            catch ( Exception exc ) {
               log.error("TelxonFile: buildOutputFile: " + exc.getClass().getName() +
                  " "  + exc.getMessage());
            }
         }
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      if ( m_ItemData != null ) {
         try {
            m_ItemData.close();
            m_ItemData = null;
         }

         catch ( Exception ex )  {

         }
      }
   }

   /**
    * Create the report.
    *
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
    * Prepares the sql queries for execution.
    *
    */
   private boolean prepareStatements()
   {
      boolean isPrepared = false;
      StringBuffer sql = new StringBuffer();

      if ( m_OraConn != null ) {
         sql.append("select");
         sql.append("   item.item_id, upc_code, vendor_item_num, ");
         sql.append("   item.description, ship_unit.unit as uom, ");
         sql.append("   stock_pack, retail_pack, ");
         sql.append("   decode( broken_case.description, 'ALLOW BROKEN CASES', ' ', 'S' ) as nbc, ");
         sql.append("   item.flc_id, mdc.nrha_id, item_price_procs.todays_retailc(item.item_id) as retailc ");
         sql.append("from ");
         sql.append("   item ");
         sql.append("join item_warehouse on item_warehouse.item_id = item.item_id and item_warehouse.in_catalog = 1 ");
         sql.append("join warehouse on warehouse.warehouse_id = item_warehouse.warehouse_id and warehouse.name = 'PORTLAND' ");
         sql.append("left outer join vendor_item_cross on vendor_item_cross.item_id = item.item_id and vendor_item_cross.vendor_id = item.vendor_id ");
         sql.append("join flc on flc.flc_id = item.flc_id ");
         sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
         sql.append("join broken_case on broken_case.broken_case_id = item.broken_case_id ");
         sql.append("join ship_unit on ship_unit.unit_id = item.ship_unit_id ");
         sql.append("join item_disp on item_disp.disp_id = item.disp_id and item_disp.disposition in ( 'BUY-SELL', 'NOBUY' ) ");
         sql.append("left outer join item_upc on item_upc.item_id = item.item_id and item_upc.primary_upc = 1 ");
         sql.append("where ");
         sql.append("    item.flc_id < '9994' ");
         sql.append("order by item.item_id");

         try {
            m_ItemData = m_OraConn.prepareStatement(sql.toString());
            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.fatal("exception:", ex);
         }

         finally {
            sql = null;
         }
      }

      return isPrepared;
   }
}
