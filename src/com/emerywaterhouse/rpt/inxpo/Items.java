/**
 * File: Items.java
 * Description: Extracts the item data.
 *    
 * @author Jeffrey Fisher
 * 
 * Create Date: 05/31/2006
 * Last Update: $Id: Items.java,v 1.8 2014/03/17 18:37:50 epearson Exp $
 * 
 * History:
 *     
 */
package com.emerywaterhouse.rpt.inxpo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.StringTokenizer;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class Items extends Report
{
   private static final String inxpoImgUrl = "<center><img src=http://presentations.inxpo.com/Shows/Emery/ImageIcon.gif></center>";
   private static final String greenImgUrl = "http://www.emeryonline.com/shared/images/green20.png";
   private PreparedStatement m_BrandName;
   private PreparedStatement m_CatalogImage;
   private PreparedStatement m_Desc;
   private PreparedStatement m_GrpCode;
   private PreparedStatement m_GrpDesc;
   private PreparedStatement m_ImageItem;
   private PreparedStatement m_ItemDtl;
   private PreparedStatement m_Items;
   private PreparedStatement m_Price;   
   private PreparedStatement m_QtyDisc;
   private PreparedStatement m_VndSku;
   
   private String m_ItemId;
   private short m_ItemOpt;
   private String m_Packet;
   @SuppressWarnings("unused")
   private short m_SelectOpt;
   private int m_VndId;
   
   /**
    * 
    */
   public Items()
   {
      super();
      
      m_MaxRunTime = RptServer.HOUR * 12;
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
      StringBuffer line = new StringBuffer(1024);
      FileOutputStream outFile = null;
      boolean result = false;
      ResultSet items = null;
      ResultSet itemDtl = null;
      ResultSet qtyDisc = null;
      String brandName = null;
      String dtlDesc = null;
      String imageUrl = null;
      String itemDesc = null;
      String itemId = null;
      String stockNbc = null;
      String upc = null;
      String vndSku = null;
      String imageItemUrl = null;
      String greenItemUrl = null;
      PromoFlag pf = null;
      int minQty = 0;
      int vndId;
      double discPct = 0.0;
      double qtyBase = 0.0;
      double promoBase = 0.0;

      try {
         setCurAction("creating/opening output file " + m_FilePath + m_FileNames.get(0));
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

         setCurAction("starting item detail export");         
         m_Items.setString(1, m_Packet);
         items = m_Items.executeQuery();
         
         //
         // Iterate through the records and create a line of delimited text based on the Inxpo
         // file layout.
         while ( items.next() && m_Status == RptServer.RUNNING ) {
            itemId = items.getString(1);
            setCurAction("getting data for item: " + itemId);
            
            m_ItemDtl.setString(1, itemId);
            itemDtl = m_ItemDtl.executeQuery();
                        
            if ( itemDtl.next() ) {
               pf = getPromoFlags(itemId);
               itemDesc = getDesc(itemId);
               vndSku = getVndSku(itemId);
               promoBase = getPrice(itemId);
               imageUrl = getCatalogImageId(itemId);
               brandName = getBrandName(itemId);
               vndId = itemDtl.getInt("vendor_id");
               upc = itemDtl.getString("upc");
               dtlDesc = itemDtl.getString("detail");
               stockNbc = itemDtl.getString("stock_pack");
               
               if ( !(itemDtl.getString("NBC").equals("ALLOW BROKEN CASES")) )
                  stockNbc += "N";
               
               //
               // make sure any strings that might be null are set to empty so the word null
               // doesn't show up.               
               if ( upc == null )
                  upc = "";
               
               //
               // Inxpo can't handle quotes, or imbeded line breaks on their import program.
               if ( dtlDesc != null ) {
                  dtlDesc = dtlDesc.replaceAll("\"", "&quot;");
                  dtlDesc = dtlDesc.replaceAll("\r", "");
                  dtlDesc = dtlDesc.replaceAll("\n", "");
               }
               else
                  dtlDesc = "";
               
               imageItemUrl = (pf.Image ? inxpoImgUrl : "");
               greenItemUrl = (pf.Green ? greenImgUrl : "");
               
               setCurAction("writing data for item: " + itemId);
               line.append("Emery\t");                            // distributor name
               line.append(vndId + "\t");                         // vendor name
               line.append(itemDesc + "\t");                      // item description
               line.append(getItemGroup(pf, false) + "\t");       // inventory item grouping
               line.append(itemId + "\t");                        // distributor product id
               line.append(vndSku + "\t");                        // vendor product id
               line.append(stockNbc + "\t");                      // unit of measure
               line.append(promoBase + "\t");                     // standard price
               line.append("0\t");                                // current promo discount
               line.append("0\t");                                // show price discount 1
               line.append("0\t");                                // show price discount 2
               line.append("0\t");                                // show price discount 3
               line.append("0\t");                                // show price discount 4
               line.append("0\t");                                // discount pct 1
               line.append("0\t");                                // discount pct 2
               line.append("0\t");                                // discount pct 3
               line.append("0\t");                                // discount pct 4
               line.append("\t");                                 // qty available
               line.append("0\t");                                // MaxQtyPerBuyer
               line.append(upc + "\t");                           // UPC Code
               line.append("1\t");                                // active
               line.append(pf.HotBuy ? "1\t":"0\t");              // hot deal flag
               line.append(pf.NewItem ? "1\t":"0\t");             // new flag
               line.append("0\t");                                // closeout flag
               line.append("0\t");                                // limited qty flag
               line.append("0\t");                                // qty discount flag
               line.append("0\t");                                // HasDeliveryDates
               line.append("0\t");                                // AllowFractionalQty
               line.append("\t");                                 // product detail url
               line.append(dtlDesc + "\t");                       // detail description
               line.append(imageUrl + "\t");                      // product image url
               line.append(imageItemUrl + "\t");                  // user field 1 (image item graphic url)               
               line.append(brandName + "\t");                     // user field 2 (brand name)
               line.append(greenItemUrl + "\t");                  // user field 3 (url for green graphic)
               line.append("\t");                                 // user field 4 (not used)
               line.append("\r\n");                               // external key
                           
               outFile.write(line.toString().getBytes());
               line.delete(0, line.length());
               
               //
               // Create separate lines for any quantity discounts
               m_QtyDisc.setString(1, itemId);
               m_QtyDisc.setString(2, m_Packet);
               qtyDisc = m_QtyDisc.executeQuery();
               
               if ( qtyDisc.next() ) {
                  minQty = qtyDisc.getInt(1);
                  discPct = qtyDisc.getDouble(2);
                  qtyBase = promoBase * (1 - (discPct / 100));
      
                  line.append("Emery\t");                         // distributor name
                  line.append(vndId + "\t");                      // vendor name
                  line.append(itemDesc + "\t");                   // item description
                  line.append(getItemGroup(pf, true) + "\t");     // inventory item grouping
                  line.append(itemId + "\t");                     // distributor product id
                  line.append(vndSku + "\t");                     // vendor product id
                  line.append(minQty + "N\t");                    // unit of measure
                  line.append(String.format("%1.2f\t", qtyBase)); // standard price
                  line.append("0\t");                             // current promo discount
                  line.append("0\t");                             // show price discount 1
                  line.append("0\t");                             // show price discount 2
                  line.append("0\t");                             // show price discount 3
                  line.append("0\t");                             // show price discount 4
                  line.append("0\t");                             // discount pct 1
                  line.append("0\t");                             // discount pct 2
                  line.append("0\t");                             // discount pct 3
                  line.append("0\t");                             // discount pct 4
                  line.append("\t");                              // qty available
                  line.append("0\t");                             // MaxQtyPerBuyer
                  line.append(upc + "\t");                        // UPC Code
                  line.append("1\t");                             // active
                  line.append(pf.HotBuy ? "1\t":"0\t");           // hot deal flag
                  line.append(pf.NewItem ? "1\t":"0\t");          // new flag
                  line.append("0\t");                             // closeout flag
                  line.append("0\t");                             // limited qty flag
                  line.append("1\t");                             // qty discount flag
                  line.append("0\t");                             // HasDeliveryDates
                  line.append("0\t");                             // AllowFractionalQty
                  line.append("\t");                              // product detail url
                  line.append(dtlDesc + "\t");                    // detail description
                  line.append(imageUrl + "\t");                   // product image url                  
                  line.append(imageItemUrl + "\t");               // user field 1 (image item graphic url)                  
                  line.append(brandName + "\t");                  // user field 2 (brand name)
                  line.append(greenItemUrl + "\t");               // user field 3 (url for green graphic)
                  line.append("\t");                              // user field 4 (not used)
                  line.append("\r\n");                            // external key
                              
                  outFile.write(line.toString().getBytes());
                  line.delete(0, line.length());
               }
               
               closeRSet(qtyDisc);
            }
            
            closeRSet(itemDtl);
            pf = null;
         }
         
         setCurAction("finished exporting item groups");
         closeRSet(items);
         closeRSet(itemDtl);
         closeRSet(qtyDisc);
         result = true;
      }

      catch ( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }
      
      finally {
         if ( outFile != null ) {
            try {
               outFile.close();
            }
            catch ( IOException iex ) {
               ;
            }
         }

         outFile = null;
      }
      
      return result;
   }
   
   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not garunteed to know when finalization occurs.
    */
   protected void cleanup()
   {
      closeStmt(m_BrandName);
      closeStmt(m_CatalogImage);
      closeStmt(m_Desc);
      closeStmt(m_GrpCode);
      closeStmt(m_GrpDesc);
      closeStmt(m_ImageItem);
      closeStmt(m_ItemDtl);
      closeStmt(m_Items);
      closeStmt(m_Price);      
      closeStmt(m_QtyDisc);
      closeStmt(m_VndSku);
             
      m_BrandName = null;
      m_CatalogImage = null;
      m_Desc = null;
      m_GrpCode = null;
      m_GrpDesc = null;
      m_ImageItem = null;
      m_ItemDtl = null;
      m_Items = null;
      m_Price = null;
      //m_PromoMsg = null;
      m_QtyDisc = null;
      m_VndSku = null; 
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
         m_OraConn = m_RptProc.getOraConn();
         m_EdbConn = m_RptProc.getEdbConn();
         
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
    * Returns the item's brand name from the catalog
    * 
    * @param itemId The item to get information for
    * @return The catalog brand name.
    */
   private String getBrandName(String itemId)
   {
      String name = "";
      ResultSet rset = null;
            
      if ( itemId != null ) {
         try {
            m_BrandName.setString(1, itemId);            
            rset = m_BrandName.executeQuery();
            
            if ( rset.next() ) {
               name = rset.getString(1);
               
               if ( name == null )                  
                  name = "";
               else {
                  
                  name = name.replaceAll("~", "");
                  name = name.replaceAll("`", "");
                  name = name.replaceAll("\\|", "");
               }
            }
         }
         
         catch ( SQLException ex ) {
            log.error("exception", ex);
         }
         
         finally {
            closeRSet(rset);
         }
      }
      
      return name;
   }
   
   /**
    * Returns the id of the image used in the catalog
    * 
    * @param itemId The item to lookup information on.
    * @return The url of the item's image if it has one.
    */
   private String getCatalogImageId(String itemId)
   {
      String url = "http://www.emeryonline.com/shared/images/catalog/%s.gif";
      String imageId = "";
      String imageUrl = "";
      ResultSet rset = null;
      
      if ( itemId != null ) {
         try {
            m_CatalogImage.setString(1, itemId);            
            rset = m_CatalogImage.executeQuery();
            
            if ( rset.next() ) {
               imageId = rset.getString(1);
               
               if ( imageId != null )
                  imageUrl = String.format(url, imageId);
            }
         }
         
         catch ( SQLException ex ) {
            log.error("exception", ex);
         }
         
         finally {
            closeRSet(rset);
         }
      }
      
      return imageUrl;
   }
   
   /**
    * Gets the promo message description as it would appear on the promo item.
    * 
    *  @param itemId The emery sku to lookup information on.
    *  @returns The description.
    */
   private String getDesc(String itemId)
   {
      String desc = "";
      ResultSet rset = null;
            
      if ( itemId != null ) {
         try {
            m_Desc.setString(1, itemId);
            m_Desc.setString(2, m_Packet);
            rset = m_Desc.executeQuery();
            
            if ( rset.next() ) {
               desc = rset.getString(1);
               
               if ( desc == null )
                  desc = "";
            }
         }
         
         catch ( SQLException ex ) {
            log.error("exception", ex);
         }
         
         finally {
            closeRSet(rset);
         }
      }
      
      return desc.replaceAll("\"", "&quot;");
   }
     
   /**
    * Gest the inventory item group description.  This designates how items will be grouped.
    * 
    * @param pf A reference to a PromoFlag object that contains the flags to check for grouping.
    * @param bb A flag to determine if the bulk buy flag can be used.  Items may be bulk buy grouped only
    *    if they are in a bulk buy pricing record.
    * @return The item group description
    */
   private String getItemGroup(PromoFlag pf, boolean bb)
   {
      String code = null;
      
      if ( pf.EnergyFallWinter )
         code = PromoFlag.ENERGY_FALL_WINTER;
      else {
         if ( pf.NewItem )
            code = PromoFlag.NEW_ITEM;
         else {
            if ( pf.QuantityBuy )
               code = PromoFlag.QUANTITY_BUY;
            else {
               if ( pf.Green )
                  code = PromoFlag.GREEN;
            }
         }
      }
      
      if ( code == null || code.length() == 0 || (pf.QuantityBuy && !bb) )
         code = "SHOW ITEMS";
         
      return code;
   }
   
   /**
    * Gets the promo base price for the item and is used as the show price.
    * 
    * @param itemId The item to looup the price for.
    * @return The price for the item.
    */
   private double getPrice(String itemId) 
   {
      double price = 0.0;
      ResultSet rset = null;
      
      if ( itemId != null ) {
         try {
            m_Price.setString(1, m_Packet);
            m_Price.setString(2, itemId);            
            rset = m_Price.executeQuery();
            
            if ( rset.next() )
               price = rset.getDouble(1);
         }
         
         catch ( SQLException ex ) {
            log.error("exception", ex);
         }
         
         finally {
            closeRSet(rset);
         }
      }
      
      return price;
   }
   
   /**
    * Gets the list of promo flags for the item.
    * 
    * @param itemId The emery sku to lookup information on.
    * @returns A reference to a PromoFlag object.
    */
   private PromoFlag getPromoFlags(String itemId)
   {
      String code = "";
      ResultSet rset = null;
      PromoFlag pf = new PromoFlag();
      StringTokenizer st = null;
      String tmp = null;
            
      if ( itemId != null ) {
         try {
            m_GrpCode.setString(1, m_Packet);
            m_GrpCode.setString(2, itemId);            
            rset = m_GrpCode.executeQuery();
            
            if ( rset.next() ) {
               code = rset.getString(1);
               
               if ( code == null )
                  code = "";
            }
         }
         
         catch ( SQLException ex ) {
            log.error("exception", ex);
         }
         
         finally {
            closeRSet(rset);
         }
         
         //
         // Parse out each flag.
         if ( code.length() > 0 ) {
            st = new StringTokenizer(code);
            
            while ( st.hasMoreTokens() ) {
              tmp = st.nextToken();
              
              if ( !pf.EnergyFallWinter )
                 pf.EnergyFallWinter = tmp.equals(PromoFlag.ENERGY_FALL_WINTER);
              
              if ( !pf.Green )
                 pf.Green = tmp.equals(PromoFlag.GREEN);
              
              if ( !pf.HotBuy )
                 pf.HotBuy = tmp.equals(PromoFlag.HOT_BUY);
              
              if ( !pf.Image )
                 pf.Image = tmp.equals(PromoFlag.IMAGE);
              
              if (!pf.NewItem )
                 pf.NewItem = tmp.equals(PromoFlag.NEW_ITEM);
              
              if ( !pf.QuantityBuy )
                 pf.QuantityBuy = tmp.equals(PromoFlag.QUANTITY_BUY);
              
              if ( !pf.AtShowOnly )
                 pf.AtShowOnly = tmp.equals(PromoFlag.AT_SHOW_ONLY);
              
              //
              // Warechouse assortment is currently whs ast.  This might have to change.
              // for now we can just look for the whs and ignore the next token
              if ( !pf.WarehouseAsst )
                 pf.WarehouseAsst = tmp.equals("WHS");
              
              if ( !pf.Assortment )
                 pf.Assortment = tmp.equals(PromoFlag.ASSORTMENT);
            }
         }
      }
      
      return pf;
   }
   
   /**
    * Returns the vendor sku or part number
    * 
    * @param itemId The item to cross reference
    * @return The vendor's sku number for the emery item.
    */
   private String getVndSku(String itemId)
   {
      String sku = "";
      ResultSet rset = null;
      
      if ( itemId != null ) {
         try {            
            m_VndSku.setString(1, itemId);            
            rset = m_VndSku.executeQuery();
            
            if ( rset.next() ) {
               sku = rset.getString(1);
               
               if ( sku == null )
                  sku = "";
            }
         }
         
         catch ( SQLException ex ) {
            log.error("exception", ex);
         }
         
         finally {
            closeRSet(rset);
         }
      }
   
      return sku;
   }  
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements() throws Exception
   {      
      StringBuffer sql = new StringBuffer();
      boolean prepared = false;

      if ( m_OraConn != null ) {
         sql.setLength(0);
         sql.append("select distinct(item.item_id) ");
         sql.append("from \"promotion\"@edbprod, \"promo_item\"@edbprod, item ");
         sql.append("where ");
         sql.append("  \"packet_id\" = ? and ");
         sql.append("  \"promotion\".\"promo_id\" = \"promo_item\".\"promo_id\" and ");
         sql.append("   item.item_id = \"promo_item\".\"item_id\" ");
         switch ( m_ItemOpt ) {
            case 1: // eoSelectedItem
               sql.append(String.format(" and \"promo_item\".\"item_id\" = '%s' ", m_ItemId));
            break;
            
            case 2: // eoSelectedVendor
               sql.append(String.format(" and item.vendor_id = %d ", m_VndId));
            break;
         }
         sql.append("order by item.item_id");
         m_Items = m_OraConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select bmi_item.brand_name ");
         sql.append("from catalog_location, bmi_item ");
         sql.append("where ");
         sql.append("   item_id = ? and ");
         sql.append("   catalog_location.location = bmi_item.location");
         m_BrandName = m_OraConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select ");
         sql.append("   item.vendor_id, ");
         sql.append("   item.item_id, ");
         sql.append("   catalogData(item.item_id) as detail, ");
         sql.append("   upc_procs.get_upc(item.item_id) as UPC, ");
         sql.append("   item.stock_pack, ");
         sql.append("   broken_case.description as NBC ");
         sql.append("from ");
         sql.append("   item, broken_case ");
         sql.append("where ");
         sql.append("   item.item_id = ? and ");
         sql.append("   broken_case.broken_case_id = item.broken_case_id");
         m_ItemDtl = m_OraConn.prepareStatement(sql.toString());
                 
         sql.setLength(0);
         sql.append("select alt_item_desc ");
         sql.append("from preprint_item, \"promo_item\"@edbprod, \"promotion\"@edbprod ");
         sql.append("where ");
         sql.append("   \"promo_item\".\"item_id\" = ? and ");
         sql.append("   \"promotion\".\"packet_id\" = ? and ");
         sql.append("   preprint_item.promo_item_id = \"promo_item\".\"promo_item_id\" and ");
         sql.append("   \"promo_item\".\"promo_id\" = \"promotion\".\"promo_id\" ");
         m_Desc = m_OraConn.prepareStatement(sql.toString());

         sql.setLength(0);
         sql.append("select price_method ");
         sql.append("from item_price_method, item_price ");
         sql.append("where ");        
         sql.append("   item_price.price_id = item_price_procs.todays_sell_id(?) and ");
         sql.append("   item_price_method.method_id = item_price.method_id ");
         m_ImageItem = m_OraConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select round(promo_base, 2) price ");
         sql.append("from promo_item, promotion ");
         sql.append("where ");
         sql.append("   promo_item.promo_id = promotion.promo_id and ");
         sql.append("   packet_id = ? and item_id = ? ");
         m_Price = m_EdbConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select min_qty, percent ");
         sql.append("from item_qty_discount ");
         sql.append("where item_id = ? and packet_id = ? ");
         sql.append("order by min_qty ");
         m_QtyDisc = m_OraConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select vendor_item_num ");
         sql.append("from vendor_item_cross, item ");
         sql.append("where ");
         sql.append("   item.item_id = ? and ");
         sql.append("   vendor_item_cross.vendor_id = item.vendor_id and ");
         sql.append("   vendor_item_cross.item_id = item.item_id");
         m_VndSku = m_OraConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select main_image as picture_id ");
         sql.append("from bmi_item, catalog_location ");
         sql.append("where ");
         sql.append("   item_id = ? and ");
         sql.append("   catalog_location.location = bmi_item.location ");
         m_CatalogImage = m_OraConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select message as code ");
         sql.append("from \"promotion\"@edbprod, \"promo_item\"@edbprod, preprint_item ");
         sql.append("where ");
         sql.append("   \"promotion\".\"packet_id\" = ? and ");
         sql.append("   \"promo_item\".\"item_id\" = ? and ");
         sql.append("   \"promo_item\".\"promo_id\" = \"promotion\".\"promo_id\" and ");
         sql.append("   preprint_item.promo_item_id = \"promo_item\".\"promo_item_id\" ");
         m_GrpCode = m_OraConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select description ");
         sql.append("from inxpo.inventory_item_grouping ");
         sql.append("where grouping_code = :code ");
         m_GrpDesc = m_OraConn.prepareStatement(sql.toString());
         
         prepared = true;
      }
            

      return prepared;
   }

   /**
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    * 
    * Because it's possible that this report can be called from some other system, the
    * best way to deal with params is to not go by the order, but by the name.
    */
   public void setParams(ArrayList<Param> params)
   {      
      String tm = Long.toString(System.currentTimeMillis()).substring(3);
      StringBuffer fname = new StringBuffer();
      int pcount = params.size();
      Param param = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
         
         if ( param.name.equals("item") )
            m_ItemId = param.value;
         
         if ( param.name.equals("itemOpt") )
            m_ItemOpt = Short.parseShort(param.value);
         
         if ( param.name.equals("packet") )
            m_Packet = param.value;        
         
         if ( param.name.equals("selectOpt") )
            m_SelectOpt = Short.parseShort(param.value);
         
         if ( param.name.equals("vendor") ) {
            if ( param.value.trim().length() > 0 )
               m_VndId = Integer.parseInt(param.value.trim());
         }
      }
      
      //
      // Build the file name.
      fname.append(tm);
      
      if ( m_VndId > 0 ) {
         fname.append("-");
         fname.append(m_VndId);
      }
      else {
         if (m_ItemId != null && m_ItemId.length() > 0 ) {
            fname.append("-");
            fname.append(m_ItemId);            
         }
      }
      
      fname.append("-itemdtl.txt");
      m_FileNames.add(fname.toString());
   }

}
