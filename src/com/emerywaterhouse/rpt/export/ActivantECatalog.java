/**
 * File: ActivantCatalog.java
 * Description: Export file for the Activant ECatalog.
 *
 * @author Jeff Fisher, Seth Murdock
 *
 * Create Date: 12/01/2009
 * Last Update: $Id: ActivantECatalog.java,v 1.13 2013/01/16 19:54:10 jfisher Exp $
 *
 * History:
 *    $Log: ActivantECatalog.java,v $
 *    Revision 1.13  2013/01/16 19:54:10  jfisher
 *    Null check for retail c
 *
 *    Revision 1.12  2012/09/20 18:31:53  jfisher
 *    Added a check for a null sell variable.
 *
 *    Revision 1.11  2012/07/05 01:18:03  npasnur
 *    in_catalog database field is migrated from item to item_warehouse table
 *
 *    Revision 1.10  2011/05/14 07:39:30  jfisher
 *    Added customer specific pricing method.
 *
 *    Revision 1.9  2010/03/28 17:01:29  jfisher
 *    Fixed some warnings and fixed some formatting issues.
 *
 *
 */
package com.emerywaterhouse.rpt.export;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Arrays;

import org.apache.log4j.AsyncAppender;
import org.apache.log4j.Logger;
import org.apache.log4j.xml.DOMConfigurator;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class ActivantECatalog extends Report
{
   private static final int itemFileLen  = 300;
   private static final int descFileLen  = 1264;
   private static final int deptFileLen  = 40;
   private static final int classFileLen = 40;
   private static final int flcFileLen   = 40;
   private static final int vndFileLen   = 286;

   //
   // Export file types
   private static final int itemFileExp = 0;
   private static final int itemDescExp = 1;
   private static final int deptExp     = 2;
   private static final int classExp    = 3;
   private static final int flcExp      = 4;
   private static final int vndExp      = 5;
   private static final int custItemExp = 6;

   private PreparedStatement m_ItemData;
   private PreparedStatement m_ItemPriceData;
   private PreparedStatement m_CustItemData;
   private PreparedStatement m_CustItemPriceData;
   private PreparedStatement m_ItemDesc;
   private PreparedStatement m_ClassData;
   private PreparedStatement m_DeptData;
   private PreparedStatement m_FlcData;
   private PreparedStatement m_VndData;

   //
   // Params
   private String m_CustId;      // The customer number for the report data to be run against.
   private int m_RptType;        // The type of report file to create

   /**
    * Default constructor.
    */
   public ActivantECatalog()
   {
      super();

      m_CustId = "";
   }

   /**
    * Cleanup any allocated resources.
    * @throws Throwable
    */
   @Override
   public void finalize() throws Throwable
   {
      m_CustId = null;

      super.finalize();
   }

   /**
    * Activant has trouble with non-ASCII chars.  This deletes them.
    * Uses convenient fact that all ASCCi are between " " (space) and "~" (tilde).
    *
    * @param inputter The input string.
    * @return A string with the non ascii chars removed.
    */
   public String Strippo(String inputter)
   {
      StringBuffer Outputter = new StringBuffer();
      int max = inputter.length();

      for (int i = 0; i < max; i++) {
         if (inputter.substring(i,i+1).matches("[ -~]")) {
            Outputter = Outputter.append(inputter.substring(i,i+1));
         }
      }
      
      return Outputter.toString();
   }

   /**
    * Builds the class(MDC) export in the Activant file format.
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildClassFile(FileOutputStream outFile) throws Exception
   {
      boolean result = false;
      StringBuffer line = new StringBuffer();
      char[] filler = new char[classFileLen];
      ResultSet rs = null;
      String mdc = null;
      String desc = null;

      Arrays.fill(filler, ' ');
      rs = m_ClassData.executeQuery();

      try {
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            line.append(filler);

            mdc = rs.getString("mdc_id");
            setCurAction(String.format("processing mdc %s ", mdc));
            desc = rs.getString("description");

            //
            // The spec calls for 32 characters for the description.
            if ( desc.length() > 32 )
               desc = desc.substring(0, 32);

            line.replace(0, mdc.length(), mdc);           //  3
            line.replace(3, 3 + desc.length(), desc);     // 32
            line.replace(35, 36, "C");                    //  1
            line.append("\r\n");

            outFile.write(line.toString().getBytes());
         }

         result = true;
      }

      catch( Exception e ) {
         log.error("[Activant] " + e);
      }

      finally {
         closeRSet(rs);
         rs = null;

         mdc = null;
         desc = null;

         outFile.close();
         outFile = null;
      }

      return result;
   }

   /**
    * Builds the item export in the Activant file format for a specific customer.  This is a copy of
    * the buildItemFile method but with customer specific pricing.
    * @see buildItemFile
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildCustItemFile(FileOutputStream outFile) throws Exception
   {
      boolean result = false;
      StringBuffer line = new StringBuffer();
      char[] filler = new char[itemFileLen];
      ResultSet itemData = null;
      ResultSet priceData = null;
      boolean isASCIIRangeOnly;
      String itemId = null;
      int itemEaId = 0;
      String mdc = null;
      String flc = null;
      String dept = null;
      String desc = null;
      String upc = null;
      String vndSku = null;
      String vndId = null;
      String weight = null;
      String stocking_unit = null;
      String retail_unit = null;
      String stock_pack = null;
      String yesPortland = null;
      String yesPittston = null;
      String rtlc = null;
      String sell = null;
      String whsCode = " ";  //special activant DC code 1 = Portland 2 = Pittston blank = both

      itemData = m_CustItemData.executeQuery();

      try {
         Arrays.fill(filler, ' ');

         while ( itemData.next() && m_Status == RptServer.RUNNING ) {
            itemId = itemData.getString("item_id");
            itemEaId = itemData.getInt("item_ea_id");
            
            line.setLength(0);
            line.append(filler);
            setCurAction(String.format("processing item %s ", itemId));

            // Try to get pricing (rtlc, sell). If we can't, continue to the next item
            try {
               m_CustItemPriceData.setString(1, m_CustId);
               m_CustItemPriceData.setInt(2, itemEaId);
               m_CustItemPriceData.setString(3, m_CustId);
               m_CustItemPriceData.setInt(4, itemEaId);
               priceData = m_CustItemPriceData.executeQuery();

               if (!priceData.next()) {
                  continue;
               }
            } 

            catch (SQLException e) {               
               continue;
            }

            mdc = itemData.getString("mdc_id");
            dept = itemData.getString("dept");
            flc = itemData.getString("flc_id");
            retail_unit = itemData.getString("ru");
            stocking_unit = itemData.getString("su");
            isASCIIRangeOnly = itemData.getString("description").matches("[ -~]+"); //all ASCII are between space and tilde.

            if ( isASCIIRangeOnly )
               desc = itemData.getString("description");
            else //Activant has trouble with non-ASCII chars, so get rid of them
               desc = Strippo(itemData.getString("description"));

            weight = String.format("%05d", (int)(itemData.getDouble("weight") * 100));

            upc = itemData.getString("portland_upc_code");
            if (upc == null || upc.isEmpty()) {
               upc = itemData.getString("pittston_upc_code"); // likely still null but there's a slim chance it won't be
            }

            vndId = String.format("%05d", itemData.getInt("vendor_id") - 700000);
            vndSku = itemData.getString("vendor_item_num");

            stock_pack = itemData.getString("port_stock_pack");
            if (stock_pack.equals("00000")) {
               stock_pack = itemData.getString("pitt_stock_pack");
            }

            rtlc = priceData.getString("rtlc");
            sell = priceData.getString("sell");

            //
            // special code for Activant for Portland and/or Pittston
            yesPortland = itemData.getString("Portland");
            yesPittston = itemData.getString("Pittston");  // activant:  Portland = 1  Pittston = 2 Both = blank

            if ( yesPortland != null ) {  //it's in portland
               if ( yesPittston == null )
                  whsCode = "1";    // it's in portland only
               else
                  whsCode = " ";    //it's in both
            }
            else {
               if ( yesPittston != null ) //it's NOT in portland
                  whsCode= "2";     //it's in pittston only
               else
                  whsCode = " ";    // it should never get here, but if it does, wtf just call it both
            }

            //
            // Trim the fields down to the max size per file spec.
            if ( desc.length() > 32 )
               desc = desc.substring(0, 32);

            if ( vndSku.length() > 14 )
               vndSku = vndSku.substring(0, 14);

            if ( upc != null ) {
               if ( upc.length() > 13 )
                  upc = upc.substring(0, 13);
            }
            else
               upc = " ";

            line.replace(0, itemId.length(), itemId);             // 14
            line.replace(20, 20 + mdc.length(), mdc);             // 3
            line.replace(23, 23 + dept.length(), dept);           // 2
            line.replace(25, 30,"+0000");                          // 4 signed desired gross profit
            line.replace(30, 30 + flc.length(), flc);             // 6
            line.replace(36, 36 + desc.length(), desc);           // 32
            line.replace(68, 69 + weight.length(), "+" + weight);       // 5
            line.replace(74, 82, "+0000000");                      // 7 signed list price (not out sell cost?
            line.replace(87, 87 + vndId.length(), vndId);         // 5
            line.replace(98, 98 + vndSku.length(), vndSku);       // 14
            line.replace(112, 120, "+0000000");                      // 7 signed mfg vend cost
            line.replace(120, 128, "+0000000");                      // 7 signed mfg vend sugg retail
            //line_replace(128, 128 + om.length(), "+" + om);          // 5 signed  order multiple
            line.replace(134, 140,"+00000");                            // 5 signed  order point
            line.replace(140, 144, "EMERY");                      // 5
            line.replace(145, 145 + retail_unit.length(), retail_unit);   // purchase unit of measure (used retail unit(?)                        // 2 purchasing unit of measure (used Emery retail unit)
            //replacement cost -- is this sell cost???  yes!!  01/05/2010
            line.replace(147,147,"+");
            line.replace(148,149 + sell.length(), sell);
            //retail price    //S9(4)V999
            line.replace(155,155,"+");
            line.replace(156,157 + rtlc.length(), rtlc);
            line.replace(164, 165 + stock_pack.length(), "+" + stock_pack);   // stocking unit of measure                       // 2 purchasing unit of measure (used Emery retail unit)
            line.replace(170, 171 + stocking_unit.length(), stocking_unit);   // stocking unit of measure                       // 2 purchasing unit of measure (used Emery retail unit)
            line.replace(175, 175, whsCode);                      // 4  this is a single char place in fourth position of 172-175
            line.replace(272, 272 + upc.length(), upc);           // 13
            line.append("\r\n");

            outFile.write(line.toString().getBytes());
         }

         result = true;
      }

      catch( Exception e ) {
         log.error("[Activant]", e);
      }

      finally {
         closeRSet(itemData);
         itemData = null;
         closeRSet(priceData);
         priceData = null;

         itemId = null;
         dept = null;
         desc = null;
         upc = null;
         vndSku = null;
         vndId = null;
         weight = null;
         flc = null;
         mdc = null;

         outFile.close();
         outFile = null;
      }

      return result;
   }

   /**
    * Builds the catalog export in the Activant file format.
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildDeptFile(FileOutputStream outFile) throws Exception
   {
      boolean result = false;
      StringBuffer line = new StringBuffer();
      char[] filler = new char[deptFileLen];
      ResultSet rs = null;
      String dept = null;
      String name = null;

      Arrays.fill(filler, ' ');
      rs = m_DeptData.executeQuery();

      try {
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            line.append(filler);

            dept = rs.getString("dept_num");
            setCurAction(String.format("processing dept %s ", dept));
            name = rs.getString("name");

            line.replace(0, dept.length(), dept);         //  2
            line.replace(2, 2 + name.length(), name);     // 32
            line.replace(34, 35, "D");                    //  1
            line.append("\r\n");

            outFile.write(line.toString().getBytes());
         }

         result = true;
      }

      catch( Exception e ) {
         log.error("[Activant]", e);
      }

      finally {
         closeRSet(rs);
         rs = null;

         dept = null;
         name = null;

         outFile.close();
         outFile = null;
      }

      return result;
   }

   /**
    * Builds the export file in the Activant long item description format.
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildDescFile(FileOutputStream outFile) throws Exception
   {
      boolean result = false;
      StringBuffer line = new StringBuffer();
      char[] filler = new char[descFileLen];
      ResultSet itemData = null;
      String itemId = null;
      String desc = null;

      itemData = m_ItemDesc.executeQuery();

      try {
         Arrays.fill(filler, ' ');

         while ( itemData.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            line.append(filler);

            itemId = itemData.getString("item_id");
            setCurAction(String.format("processing item %s ", itemId));
            desc = itemData.getString("description");

            line.replace(0, itemId.length(), itemId);          //  14
            line.replace(14, 14 + desc.length(), desc);        //  1250
            line.append("\r\n");

            outFile.write(line.toString().getBytes());
         }

         outFile.write(line.toString().getBytes());
         result = true;
      }

      catch( Exception e ) {
         log.error("[Activant]", e);
      }

      finally {
         closeRSet(itemData);
         itemData = null;

         itemId = null;
         desc = null;

         outFile.close();
         outFile = null;
      }

      return result;
   }

   /**
    * Builds the catalog export in the Activant file format.
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildFlcFile(FileOutputStream outFile) throws Exception
   {
      boolean result = false;
      StringBuffer line = new StringBuffer();
      char[] filler = new char[flcFileLen];
      ResultSet rs = null;
      String flc = null;
      String desc = null;

      Arrays.fill(filler, ' ');
      rs = m_FlcData.executeQuery();

      try {
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            line.append(filler);

            flc = rs.getString("flc_id");
            setCurAction(String.format("processing flc %s ", flc));
            desc = rs.getString("description");

            //
            // The spec calls for 32 characters for the description.
            if ( desc.length() > 32 )
               desc = desc.substring(0, 32);

            line.replace(0, flc.length(), flc);           //  6
            line.replace(6, 6 + desc.length(), desc);     // 32
            line.replace(38, 39, "F");                    //  1
            line.append("\r\n");

            outFile.write(line.toString().getBytes());
         }

         result = true;
      }

      catch( Exception e ) {
         log.error("[Activant]", e);
      }

      finally {
         closeRSet(rs);
         rs = null;

         flc = null;
         desc = null;

         outFile.close();
         outFile = null;
      }

      return result;
   }

   /**
    * Builds the item export in the Activant file format.
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildItemFile(FileOutputStream outFile) throws Exception
   {
      boolean result = false;
      StringBuffer line = new StringBuffer();
      char[] filler = new char[itemFileLen];
      ResultSet itemData = null;
      ResultSet priceData = null;
      boolean isASCIIRangeOnly;
      String itemId = null;      
      int ejdItemId = 0;
      int whsId = 0;
      String mdc = null;
      String flc = null;
      String dept = null;
      String desc = null;
      String upc = null;
      String vndSku = null;
      String vndId = null;
      String weight = null;
      String stocking_unit = null;
      String retail_unit = null;
      String stock_pack = null;
      String rtlc = null;
      String sell = null;
      String whsCode = " ";  //special activant DC code 1 = Portland 2 = Pittston blank = both -- always blank to accommodate ACE

      itemData = m_ItemData.executeQuery();

      try {
         Arrays.fill(filler, ' ');

         while ( itemData.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            line.append(filler);
            itemId = itemData.getString("item_id");            
            ejdItemId = itemData.getInt("ejd_item_id");
            whsId = itemData.getInt("warehouse_id");

            setCurAction(String.format("processing item %s ", itemId));

            // Try to get pricing (rtlc, sell). If we can't, continue to the next item
            try {
               m_ItemPriceData.setInt(1, ejdItemId);
               m_ItemPriceData.setInt(2, whsId);               
               priceData = m_ItemPriceData.executeQuery();

               if ( !priceData.next() ) {
                  continue;
               }
            }
            
            catch(SQLException e) {               
               m_EdbConn.rollback();
               continue;
            }

            mdc = itemData.getString("mdc_id");
            dept = itemData.getString("dept");
            flc = itemData.getString("flc_id");
            retail_unit = itemData.getString("ru");
            stocking_unit = itemData.getString("su");
            isASCIIRangeOnly = itemData.getString("description").matches("[ -~]+"); //all ASCII are between space and tilde.

            if (isASCIIRangeOnly)
               desc = itemData.getString("description");
            else //Activant has trouble with non-ASCII chars, so get rid of them
               desc = Strippo(itemData.getString("description"));

            weight = String.format("%05d", (int)(itemData.getDouble("weight") * 100));
            upc = itemData.getString("upc_code");
            vndId = String.format("%05d", itemData.getInt("vendor_id") - 700000);
            vndSku = itemData.getString("vendor_item_num");
            stock_pack = itemData.getString("stock_pack");

            rtlc = priceData.getString("rtlc");
            if (rtlc == null )
               rtlc = " ";

            sell = priceData.getString("sell");
            if ( sell == null )
               sell = " ";

            // Trim the fields down to the max size per file spec.
            if ( desc.length() > 32 )
               desc = desc.substring(0, 32);

            if ( vndSku.length() > 14 )
               vndSku = vndSku.substring(0, 14);

            if ( upc != null ) {
               if ( upc.length() > 13 )
                  upc = upc.substring(0, 13);
            }
            else
               upc = " ";

            line.replace(0, itemId.length(), itemId);             // 14
            line.replace(20, 20 + mdc.length(), mdc);             // 3
            line.replace(23, 23 + dept.length(), dept);           // 2
            line.replace(25, 30,"+0000");                          // 4 signed desired gross profit
            line.replace(30, 30 + flc.length(), flc);             // 6
            line.replace(36, 36 + desc.length(), desc);           // 32
            line.replace(68, 69 + weight.length(), "+" + weight);       // 5
            line.replace(74, 82, "+0000000");                      // 7 signed list price (not out sell cost?
            line.replace(87, 87 + vndId.length(), vndId);         // 5
            line.replace(98, 98 + vndSku.length(), vndSku);       // 14
            line.replace(112, 120, "+0000000");                      // 7 signed mfg vend cost
            line.replace(120, 128, "+0000000");                      // 7 signed mfg vend sugg retail
            //line_replace(128, 128 + om.length(), "+" + om);          // 5 signed  order multiple
            line.replace(134, 140,"+00000");                            // 5 signed  order point
            line.replace(140, 144, "EMERY");                      // 5
            line.replace(145, 145 + retail_unit.length(), retail_unit);   // purchase unit of measure (used retail unit(?)                        // 2 purchasing unit of measure (used Emery retail unit)
            //replacement cost -- is this sell cost???  yes!!  01/05/2010
            line.replace(147,147,"+");
            line.replace(148,149 + sell.length(), sell);
            //retail price    //S9(4)V999
            line.replace(155,155,"+");
            line.replace(156,157 + rtlc.length(), rtlc);
            line.replace(164, 165 + stock_pack.length(), "+" + stock_pack);   // stocking unit of measure                       // 2 purchasing unit of measure (used Emery retail unit)
            line.replace(170, 171 + stocking_unit.length(), stocking_unit);   // stocking unit of measure                       // 2 purchasing unit of measure (used Emery retail unit)
            line.replace(175, 175, whsCode);                      // 4  this is a single char place in fourth position of 172-175
            line.replace(272, 272 + upc.length(), upc);           // 13
            line.append("\r\n");

            outFile.write(line.toString().getBytes());
         }

         result = true;
      }

      catch( Exception e ) {
         log.error("[Activant]", e);
      }

      finally {
         closeRSet(itemData);
         itemData = null;
         closeRSet(priceData);
         priceData = null;

         itemId = null;
         dept = null;
         desc = null;
         upc = null;
         vndSku = null;
         vndId = null;
         weight = null;
         flc = null;
         mdc = null;

         outFile.close();
         outFile = null;
      }

      return result;
   }

   /**
    * Executes the queries and builds the output file
    *
    * @throws java.io.FileNotFoundException
    */
   private boolean buildOutputFile() throws FileNotFoundException
   {
      FileOutputStream outFile = null;
      boolean result = false;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         switch ( m_RptType ) {
            case itemFileExp: {
               result = buildItemFile(outFile);
               break;
            }
   
            case itemDescExp: {
               result = buildDescFile(outFile);
               break;
            }
   
            case deptExp: {
               result = buildDeptFile(outFile);
               break;
            }
   
            case classExp: {
               result = buildClassFile(outFile);
               break;
            }
   
            case flcExp: {
               result = buildFlcFile(outFile);
               break;
            }
   
            case vndExp: {
               result = buildVndFile(outFile);
               break;
            }
   
            case custItemExp: {
               result = buildCustItemFile(outFile);
            }
         }
      }

      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("[Activant]", ex);
      }

      finally {
         try {
            outFile.close();
         }

         catch( Exception e ) {
            log.error("[Activant]", e);
         }

         outFile = null;
      }

      return result;
   }

   /**
    * Builds the catalog export in the Activant file format.
    *
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    *
    * @throws Exception on errors.
    */
   private boolean buildVndFile(FileOutputStream outFile) throws Exception
   {
      boolean result = false;
      StringBuffer line = new StringBuffer();
      char[] filler = new char[vndFileLen];
      ResultSet rs = null;
      int vndId;
      int prevVndId = 0;
      String id = null;
      String name = null;
      String addr1 = null;
      String addr2 = null;
      String city = null;
      String state = null;
      String zip = null;

      Arrays.fill(filler, ' ');
      rs = m_VndData.executeQuery();

      try {
         while ( rs.next() && m_Status == RptServer.RUNNING ) {
            line.setLength(0);
            line.append(filler);

            vndId = rs.getInt("vendor_id");

            //
            // There are duplicates in the data so screen for them here.
            if ( vndId != prevVndId ) {
               prevVndId = vndId;
               setCurAction(String.format("processing vendor %d ", vndId));

               //
               // Strip the leading 7 off of the id so that it fits within the
               // limits of the file format.
               id = String.format("%05d", vndId - 700000);
               name = rs.getString("name");
               addr1 = rs.getString("addr1");
               addr2 = rs.getString("addr2");
               city = rs.getString("city");
               state = rs.getString("state");
               zip = rs.getString("postal_code");

               //
               // Shorten the fields to the same max size on the spec sheet.
               if ( name.length() > 30 )
                  name = name.substring(0, 30);

               if ( addr1.length() > 30 )
                  addr1 = addr1.substring(0, 30);

               if ( addr2 != null ) {
                  if ( addr2.length() > 30 )
                     addr1 = addr2.substring(0, 30);
               }
               else
                  addr2 = "";

               if ( city.length() > 23 )
                  city = city.substring(0, 23);

               if ( zip.length() > 5 )
                  zip = zip.substring(0, 5);

               line.replace(1, 6, id);                            //  5
               line.replace(6, 11, "EMERY");                      //  5
               line.replace(11, 11 + name.length(), name);        //  30
               line.replace(41, 41 + addr1.length(), addr1);      //  30
               line.replace(71, 71 + addr2.length(), addr2);      //  30
               line.replace(101, 101 + city.length(), city);      //  23
               line.replace(124, 124 + state.length(), state);    //  2
               line.replace(126, 126 + zip.length(), zip);        //  5
               line.append("\r\n");

               outFile.write(line.toString().getBytes());
            }
         }

         result = true;
      }

      catch( Exception e ) {
         log.error("[Activant]", e);
      }

      finally {
         closeRSet(rs);
         rs = null;

         id = null;
         name = null;
         addr1 = null;
         addr2 = null;
         city = null;
         state = null;
         zip = null;

         outFile.close();
         outFile = null;
      }

      return result;
   }

   /**
    * Closes all the sql statements so they release the db cursors.
    */
   private void closeStatements()
   {
      closeStmt(m_ItemData);
      closeStmt(m_ItemPriceData);
      closeStmt(m_ItemDesc);
      closeStmt(m_ClassData);
      closeStmt(m_DeptData);
      closeStmt(m_FlcData);
      closeStmt(m_VndData);
      closeStmt(m_CustItemData);
      closeStmt(m_CustItemPriceData);

      m_ItemData = null;
      m_ItemPriceData = null;
      m_ItemDesc = null;
      m_ClassData = null;
      m_DeptData = null;
      m_FlcData = null;
      m_VndData = null;
      m_CustItemData = null;
      m_CustItemPriceData = null;
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
         if (m_EdbConn == null )
            m_EdbConn = m_RptProc.getEdbConn();
         
         if ( prepareStatements() )
            created = buildOutputFile();
      }

      catch ( Exception ex ) {
         log.fatal("[Activant]", ex);
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
    */
   private boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;

      if ( m_EdbConn != null ) {
         try {
            sql.setLength(0);
            sql.append("select item_id, mdc_id, dept, flc_id, description, weight, upc_code, vendor_id, vendor_item_num, su, ru, stock_pack, ");
            sql.append("warehouse_id, item_ea_id, ejd_item_id ");
            sql.append("from ( ");
            sql.append("select ");
            sql.append("   item_entity_attr.item_ea_id, item_entity_attr.item_id, mdc.mdc_id, emery_dept.dept_num dept, ejd_item.flc_id, ");
            sql.append("   item_entity_attr.description, weight, upc_code, vendor.vendor_id, viec.vendor_item_num, ");
            sql.append("   substr(su.unit,1,2) as su, ");
            sql.append("   substr(ru.unit,1,2) as ru, ");
            sql.append("   lpad(to_char(ejd_item_warehouse.stock_pack),5,'0') as stock_pack, ");
            sql.append("   ejd_item_warehouse.warehouse_id, ejd_item.ejd_item_id ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join vendor_dept on vendor_dept.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join vendor_item_ea_cross viec on viec.vendor_id = item_entity_attr.vendor_id and viec.item_ea_id = item_entity_attr.item_ea_id ");   
            sql.append("join emery_dept on emery_dept.dept_id = vendor_dept.dept_id ");
            sql.append("join flc on flc.flc_id = ejd_item.flc_id ");
            sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
            sql.append("join retail_unit su on su.unit_id = item_entity_attr.ship_unit_id ");
            sql.append("join retail_unit ru on ru.unit_id = item_entity_attr.ret_unit_id ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.warehouse_id = 1 and ejd_item_warehouse.in_catalog = 1 ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = ejd_item.ejd_item_id and ejd_item_whs_upc.primary_upc = 1 and ");
            sql.append("   ejd_item_whs_upc.warehouse_id = 1 ");
            sql.append("where item_entity_attr.item_type_id = 1 ");
            sql.append("union ");
            sql.append("select ");
            sql.append("   item_entity_attr.item_ea_id, item_entity_attr.item_id, mdc.mdc_id, emery_dept.dept_num dept, ejd_item.flc_id, ");
            sql.append("   item_entity_attr.description, weight, upc_code, vendor.vendor_id, viec.vendor_item_num, ");
            sql.append("   substr(su.unit,1,2) as su, ");
            sql.append("   substr(ru.unit,1,2) as ru, ");
            sql.append("   lpad(to_char(ejd_item_warehouse.stock_pack),5,'0') as stock_pack, ");
            sql.append("   ejd_item_warehouse.warehouse_id, ejd_item.ejd_item_id ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join vendor_dept on vendor_dept.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join vendor_item_ea_cross viec on viec.vendor_id = item_entity_attr.vendor_id and viec.item_ea_id = item_entity_attr.item_ea_id ");   
            sql.append("join emery_dept on emery_dept.dept_id = vendor_dept.dept_id ");
            sql.append("join flc on flc.flc_id = ejd_item.flc_id ");
            sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
            sql.append("join retail_unit su on su.unit_id = item_entity_attr.ship_unit_id ");
            sql.append("join retail_unit ru on ru.unit_id = item_entity_attr.ret_unit_id ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.warehouse_id = 2 and ejd_item_warehouse.in_catalog = 1 ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = ejd_item.ejd_item_id and ejd_item_whs_upc.primary_upc = 1 and ");
            sql.append("   ejd_item_whs_upc.warehouse_id = 2 ");
            sql.append("where item_entity_attr.item_type_id = 1 ");
            sql.append("union ");
            sql.append("select ");
            sql.append("   item_entity_attr.item_ea_id, item_entity_attr.item_id, mdc.mdc_id, emery_dept.dept_num dept, ejd_item.flc_id, ");
            sql.append("   item_entity_attr.description, weight, upc_code, vendor.vendor_id, viec.vendor_item_num, ");
            sql.append("   substr(su.unit,1,2) as su, ");
            sql.append("   substr(ru.unit,1,2) as ru, ");
            sql.append("   lpad(to_char(ejd_item_warehouse.stock_pack),5,'0') as stock_pack, ");
            sql.append("   ejd_item_warehouse.warehouse_id, ejd_item.ejd_item_id ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join vendor_dept on vendor_dept.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join vendor_item_ea_cross viec on viec.vendor_id = item_entity_attr.vendor_id and viec.item_ea_id = item_entity_attr.item_ea_id ");   
            sql.append("join emery_dept on emery_dept.dept_id = vendor_dept.dept_id ");
            sql.append("join flc on flc.flc_id = ejd_item.flc_id ");
            sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
            sql.append("join retail_unit su on su.unit_id = item_entity_attr.ship_unit_id ");
            sql.append("join retail_unit ru on ru.unit_id = item_entity_attr.ret_unit_id ");
            sql.append("join ejd_item_warehouse on ejd_item_warehouse.ejd_item_id = item_entity_attr.ejd_item_id and ejd_item_warehouse.warehouse_id = 11 and ejd_item_warehouse.in_catalog = 1 ");
            sql.append("left outer join ejd_item_whs_upc on ejd_item_whs_upc.ejd_item_id = ejd_item.ejd_item_id and ejd_item_whs_upc.primary_upc = 1 and ");
            sql.append("   ejd_item_whs_upc.warehouse_id = 11 ");
            sql.append("where item_entity_attr.item_type_id = 9 ");
            sql.append(") all_items ");
            sql.append("group by item_id, warehouse_id, mdc_id, dept, flc_id, description, weight, upc_code, vendor_id, vendor_item_num, su, ru, stock_pack, ejd_item_id, item_ea_id ");
            sql.append("order by item_id ");
            m_ItemData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select ");
            sql.append("   lpad(to_char((retail_c * 1000)::integer),7,'0') as rtlc, ");
            sql.append("   lpad(to_char((sell * 1000)::integer),7,'0') as sell ");
            sql.append("from ejd_item_price ");
            sql.append("where ejd_item_id = ? and warehouse_id = ? ");
            m_ItemPriceData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select ");
            sql.append("   item_entity_attr.item_id, mdc.mdc_id, emery_dept.dept_num dept, ejd_item.flc_id, item_entity_attr.description, ");
            sql.append("   weight, portupc.upc_code as portland_upc_code, pittupc.upc_code as pittston_upc_code, vendor.vendor_id, vic.vendor_item_num, ");
            sql.append("   iwm.warehouse_id as Portland,");
            sql.append("   iwp.warehouse_id as Pittston,");
            sql.append("   substr(su.unit,1,2) as su, ");
            sql.append("   substr(ru.unit,1,2) as ru, ");
            sql.append("   lpad(to_char(iwm.stock_pack),5,'0') as port_stock_pack, ");
            sql.append("   lpad(to_char(iwp.stock_pack),5,'0') as pitt_stock_pack, ");
            sql.append("   item_entity_attr.item_ea_id ");
            sql.append("from item_entity_attr ");
            sql.append("join ejd_item on ejd_item.ejd_item_id = item_entity_attr.ejd_item_id ");
            sql.append("join vendor on vendor.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join vendor_dept on vendor_dept.vendor_id = item_entity_attr.vendor_id ");
            sql.append("join vendor_item_ea_cross vic on vic.vendor_id = item_entity_attr.vendor_id and vic.item_ea_id = item_entity_attr.item_ea_id ");
            sql.append("join emery_dept on emery_dept.dept_id = vendor_dept.dept_id ");
            sql.append("join flc on flc.flc_id = ejd_item.flc_id ");
            sql.append("join mdc on mdc.mdc_id = flc.mdc_id ");
            sql.append("join retail_unit su on su.unit_id = item_entity_attr.ship_unit_id ");
            sql.append("join retail_unit ru on ru.unit_id = item_entity_attr.ret_unit_id ");
            sql.append("left outer join ejd_item_warehouse iwm on iwm.ejd_item_id = ejd_item.ejd_item_id and iwm.warehouse_id = 1 ");
            sql.append("left outer join ejd_item_warehouse iwp on iwp.ejd_item_id = ejd_item.ejd_item_id and iwp.warehouse_id = 2 ");
            sql.append("left outer join ejd_item_whs_upc portupc on portupc.ejd_item_id = ejd_item.ejd_item_id and portupc.primary_upc = 1 and portupc.warehouse_id = 1");
            sql.append("left outer join ejd_item_whs_upc pittupc on pittupc.ejd_item_id = ejd_item.ejd_item_id and pittupc.primary_upc = 1 and pittupc.warehouse_id = 2 ");
            sql.append("where item_entity_attr.item_type_id = 1 and nvl(iwm.in_catalog, 0) = 1 ");
            sql.append("order by item_entity_attr.item_id");
            m_CustItemData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select ");
            sql.append("   lpad(to_char((ejd_price_procs.get_retail_price(?, ?)*1000)::integer),7,'0') as rtlc, ");
            sql.append("   (select lpad(to_char((price*1000)::integer),7,'0') from ejd_cust_procs.get_sell_price(?, ?)) as sell"); 
            m_CustItemPriceData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select distinct item_id, description ");
            sql.append("from item_entity_attr ");
            sql.append("order by item_id ");
            m_ItemDesc = m_EdbConn.prepareStatement(sql.toString()); 

            sql.setLength(0);
            sql.append("select mdc_id, description ");
            sql.append("from mdc ");
            sql.append("order by mdc_id");
            m_ClassData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select dept_num, name ");
            sql.append("from emery_dept ");
            sql.append("order by dept_num");
            m_DeptData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select flc_id, description ");
            sql.append("from flc ");
            sql.append("order by flc_id");
            m_FlcData = m_EdbConn.prepareStatement(sql.toString());

            sql.setLength(0);
            sql.append("select vendor.vendor_id, vendor.name, addr1, addr2, city, state, postal_code ");
            sql.append("from vendor ");
            sql.append("join vendor_address va on va.vendor_id = vendor.vendor_id ");
            sql.append("join vendor_address_type vat on vat.vat_id = va.vat_id and type in ('SALES', 'SHIPPING', 'DROPSHIP') ");
            sql.append("where active = 1 ");
            sql.append("group by vendor.vendor_id, name, addr1, addr2, city, state, postal_code ");
            sql.append("order by vendor.vendor_id ");
            m_VndData = m_EdbConn.prepareStatement(sql.toString());

            isPrepared = true;
         }

         catch ( SQLException ex ) {
            log.error("[Activant]", ex);
         }

         finally {
            sql = null;
         }
      }
      else
         log.error("[Activant].prepareStatements - null enterprisedb connection");

      return isPrepared;
   }

   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();
      int pcount = params.size();
      Param param = null;
      String tmp = null;

      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);

         if ( param.name.equals("cust") )
            m_CustId = param.value;

         if ( param.name.equals("type") ) {
            tmp = param.value;

            if ( tmp.equalsIgnoreCase("item") ) {
               m_RptType = itemFileExp;
               fileName.append("emery_i.dat");
            }

            if ( tmp.equalsIgnoreCase("citem") ) {
               m_RptType = custItemExp;
               fileName.append("emery_i_%s.dat");
            }

            if ( tmp.equalsIgnoreCase("desc") ) {
               m_RptType = itemDescExp;
               fileName.append("emery_ldesc.dat");
            }

            if ( tmp.equalsIgnoreCase("dept") ) {
               m_RptType = deptExp;
               fileName.append("emery_de.dat");
            }

            if ( tmp.equalsIgnoreCase("class") ) {
               m_RptType = classExp;
               fileName.append("emery_cl.dat");
            }

            if ( tmp.equalsIgnoreCase("flc") ) {
               m_RptType = flcExp;
               fileName.append("emery_fi.dat");
            }

            if ( tmp.equalsIgnoreCase("vnd") ) {
               m_RptType = vndExp;
               fileName.append("emery_vm.dat");
            }
         }
      }

      if ( m_RptType == custItemExp )
         m_FileNames.add(String.format(fileName.toString(), m_CustId));
      else
         m_FileNames.add(fileName.toString());
   }


   public static void main(String[] args) {
      AsyncAppender asyncAppender = null;

      try {         
         DOMConfigurator.configure("logcfg.xml");
         //log.setAdditivity(false);
         asyncAppender = (AsyncAppender)Logger.getRootLogger().getAppender("ASYNC");
         asyncAppender.setBufferSize(15);
         asyncAppender.setLocationInfo(true);
      }

      catch( Exception ex ) {

      }
      
      ActivantECatalog cat = new ActivantECatalog();

      Param p1 = new Param();
      p1.name = "cust";
      p1.value = "137502";
      
      Param p2 = new Param();
      p2.name = "type";
      p2.value = "item";
      //p2.value = "citem";
      ArrayList<Param> params = new ArrayList<Param>();
      params.add(p1);
      params.add(p2);

      cat.m_FilePath = "C:\\Users\\JFisher\\temp\\";

      java.util.Properties connProps = new java.util.Properties();
      connProps.put("user", "ejd");
      connProps.put("password", "boxer");
      try {
         cat.m_EdbConn = java.sql.DriverManager.getConnection("jdbc:edb://172.30.1.33:5444/emery_jensen",connProps);
         cat.m_EdbConn.setAutoCommit(false);
      } 
      
      catch (Exception e) {
         e.printStackTrace();
      }

      cat.setParams(params);
      cat.createReport();
   }


}
