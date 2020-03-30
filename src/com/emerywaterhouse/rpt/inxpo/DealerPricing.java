/**
 * File: DealerPricing.java
 * Description: Extracts the dealer(customer) pricing data for inxpo.
 *    
 * @author Jeffrey Fisher
 * 
 * Create Date: 05/08/2006
 * Last Update: $Id: DealerPricing.java,v 1.4 2008/10/30 15:47:21 jfisher Exp $
 * 
 * History:
 */
package com.emerywaterhouse.rpt.inxpo;

import java.io.FileOutputStream;
import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Types;
import java.util.ArrayList;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

/**
 *
 */
public class DealerPricing extends Report
{
   private PreparedStatement m_BaseCost;
   private PreparedStatement m_DealerItem;
   private PreparedStatement m_SalesHist;
   private CallableStatement m_SellPrice;
   
   private String m_CustId;
   private String m_Date;
   private short m_DealerOpt;
   private String m_ItemId;
   private short m_ItemOpt;
   private String m_Packet;
   private short m_SelectOpt;
   private String m_ShowName;
   private int m_VndId;
         
   /**
    * Default constructor
    */
   public DealerPricing()
   {
      super();
            
      m_MaxRunTime = RptServer.HOUR * 24;
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
      ResultSet dealerItem = null;
      ResultSet salesHist = null;
      String itemId;
      String custId;
      String promoId;
      int qty = 0;
      double amt = 0.0;
      double regPrice = 0.0;
      double disc = 0.0;
            
      try {
         setCurAction("creating/opening output file " + m_FilePath + m_FileNames.get(0));
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

         setCurAction("starting dealer pricing export");
         m_DealerItem.setString(1, m_ShowName);
         m_DealerItem.setString(2, m_Packet);
         dealerItem = m_DealerItem.executeQuery();
         
         //
         // Iterate through the records and create a line of delimited text based on the Inxpo
         // file layout.
         while ( dealerItem.next() && m_Status == RptServer.RUNNING ) {
            custId = dealerItem.getString(1);
            itemId = dealerItem.getString(2);
            promoId = dealerItem.getString(3);
            setCurAction("getting data for cust: " + custId + " item: " + itemId);
            
            //
            // Get the sales history data
            setCurAction("getting history data for cust: " + custId + " item: " + itemId);
            m_SalesHist.setString(1, custId);
            m_SalesHist.setString(2, itemId);
            //m_SalesHist.setString(3, m_Packet);
            salesHist = m_SalesHist.executeQuery();
            
            if ( salesHist.next() ) {
               qty = salesHist.getInt(1);
               amt = salesHist.getDouble(2);
            }
            
            closeRSet(salesHist);
                        
            //
            // Get the sell price and discount
            setCurAction("getting price data for cust: " + custId + " item: " + itemId);
            m_SellPrice.setString(1, custId);
            m_SellPrice.setString(2, itemId);
            m_SellPrice.setString(3, promoId);
            m_SellPrice.execute();
            regPrice = m_SellPrice.getDouble(4);            
            disc = m_SellPrice.getDouble(6);
            
            if ( regPrice == 0.0 )
               regPrice = getBaseCost(itemId);
            
            setCurAction("writing data for cust: " + custId + " item: " + itemId);
            line.append("Emery\t");                         // distributor name
            line.append(itemId + "\t");                     // distributer product id
            line.append("\t");                              // uom            
            line.append(custId + "\t");                     // customer id
            line.append("\t");                              // bill to id
            line.append("\t");                              // ship to id
            line.append(regPrice + "\t");                   // price
            line.append(disc + "\t");                       // Buyerfield 1
            line.append("\t");                              // buyer field 2
            line.append("\t");                              // buyer field 3
            line.append("\t");                              // buyer field 4
            line.append(qty + "\t");                        // purchase history quantity
            line.append(String.format("%1.2f\t", amt));     // purchase history amount
            line.append("\t");                              // amount discount index
            line.append("\r\n");                            // percent discount index
            
            outFile.write(line.toString().getBytes());
            line.delete(0, line.length());
         }
         
         setCurAction("finished exporting item groups");
         closeRSet(dealerItem);
         result = true;
      }

      catch ( Exception ex ) {
         log.error("exception", ex);
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());
      }
      
      return result;
   }

   /**
    * Creates the where clause for the main query.  Builds in selection criteria based
    * on the index of the selection params in the sending application.
    * 
    * @return The where clause for the main pricing query.
    */
   private String buildWhere()
   {
      StringBuffer sql = new StringBuffer();
      sql.append("where show.name = ? and ");
      sql.append("promotion.packet_id = ? and ");
      sql.append("dealer_show.show_id = show.show_id and ");
      sql.append("dealer.customer_id = dealer_show.customer_id and ");
      sql.append("promo_item.promo_id = promotion.promo_id ");
      
      //
      // 
      if ( m_SelectOpt == 1 ) {         
         sql.append(getCustSql());
         sql.append(getItemSql());
      }
      
      return sql.toString();
   }
   
   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not garunteed to know when finalization occurs.
    */
   protected void cleanup()
   {
      closeStmt(m_BaseCost);
      closeStmt(m_DealerItem);
      closeStmt(m_SalesHist);
      closeStmt(m_SellPrice);
      
      m_BaseCost = null;
      m_DealerItem = null;
      m_SalesHist = null;
      m_SellPrice = null;
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
    * Gets the base cost of an item.
    * 
    * @param itemId The item to get the cost of.
    * @return The base cost (todays_sell) for the item.
    */
   private double getBaseCost(String itemId)
   {
      double base = 0.0;
      ResultSet rs = null;
      
      if ( itemId != null ) {
         try {
            m_BaseCost.setString(1, itemId);
            rs = m_BaseCost.executeQuery();
            
            if ( rs.next() )
               base = rs.getDouble(1);
         }
         
         catch ( SQLException ex ) {
            ex.printStackTrace();
         }
         
         finally {
            closeRSet(rs);
         }
      }
      
      return base;
   }
   
   /**
    * Creates a piece of the where clause based on the item option parameter.  Taken from the Delphi
    * extraction program.
    * 
    * @return a piece of sql code for the customer selection.
    */
   private String getCustSql()
   {
      StringBuffer sql = new StringBuffer();
      
      //
      // Currently, this code only gets called when the select option is something other than
      // all customers and items.  0 means no specific dealer export, but selected dealer price export.
      if ( m_CustId != null && m_CustId.length() == 6 ) {      
         switch ( m_DealerOpt ) {
            case 0:
            case 1:  // eoSelected 
               sql.append(String.format(" and dealer.customer_id = '%s' ", m_CustId));
            break;
            
            case 2: // eoNew
               sql.append(" and dealer.last_extract is null ");            
            break;
            
            case 3:  // eoDate
               sql.append(String.format(" and dealer.last_extract >= to_date('%s', 'mm/dd/yyyy')", m_Date));            
            break;      
         };
      }
           
      return sql.toString();
   }
   
   /**
    * Creates a piece of the where clause based on the item option parameter.
    * 
    * @return a piece of sql code for the item selection of the query.
    */
   private String getItemSql()
   {
      StringBuffer sql = new StringBuffer();
     
      if ( m_ItemId != null && m_ItemId.length() == 7 ) {
         switch ( m_ItemOpt ) {
            case 1: // eoSelectedItem
               sql.append(String.format(" and promo_item.item_id = '%s' ", m_ItemId));
            break;
            
            case 2: // eoSelectedVendor
               sql.append(String.format("  and item.vendor_id = %d ", m_VndId));
            break;
         }
      }

      return sql.toString();
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements() throws Exception
   {      
      StringBuffer sql = new StringBuffer();
      boolean prepared = false;

      if ( m_OraConn != null ) {
         sql.append("select round(item_price_procs.todays_sell(?), 2) as base_cost from dual");
         m_BaseCost = m_OraConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         sql.append("select dealer.customer_id, item_id, promotion.promo_id ");
         sql.append("from inxpo.dealer, inxpo.dealer_show, inxpo.show, promotion, promo_item ");
         sql.append(buildWhere());
         sql.append("order by customer_id, item_id");
         m_DealerItem = m_OraConn.prepareStatement(sql.toString());
                  
         sql.setLength(0);
         sql.append("select nvl(sum(qty_shipped),0) qty, nvl(sum(qty_shipped * unit_sell),0) amt ");
         sql.append("from inv_dtl ");
         sql.append("where cust_nbr = ? and ");
         sql.append("item_nbr = ? and ");
         sql.append("sale_type = 'WAREHOUSE' and ");         
         sql.append("invoice_date <= sysdate and invoice_date >= sysdate - 365");
         m_SalesHist = m_OraConn.prepareStatement(sql.toString());
                  
         sql.setLength(0);
         sql.append("declare ");
         sql.append("   custId varchar2(6); ");
         sql.append("   itemId varchar2(7); ");
         sql.append("   promoId varchar2(6); ");
         sql.append("   method varchar2(40); ");
         sql.append("   regprice number; ");
         sql.append("   promoprice number; ");
         sql.append("   discount number; ");
         sql.append("begin ");
         sql.append("   custId := ?; ");
         sql.append("   itemId := ?; ");
         sql.append("   promoId := ?; ");
         sql.append("   begin ");
         sql.append("      regprice := cust_procs.GetSellPrice(custId, itemId); ");
         sql.append("      method := cust_procs.GetPriceMethod; ");
         sql.append("   exception ");
         sql.append("      when others then ");
         sql.append("         regprice := 0; "); 
         sql.append("         method := 'NONE'; ");
         sql.append("   end; ");
         
         sql.append("   if method = 'CONTRACT' then ");
         sql.append("      select promo_base into promoprice from promo_item where promo_id = promoId and item_id = itemId; "); 
         sql.append("   else ");
         sql.append("      begin ");
         sql.append("         promoprice := cust_procs.GetSellPrice(custId, itemId, promoId); ");
         sql.append("       exception "); 
         sql.append("          when others then ");
         sql.append("             select promo_base into promoprice from promo_item where promo_id = promoId and item_id = itemId; ");
         sql.append("      end; ");
         sql.append("   end if; ");
         
         sql.append("   if regprice = 0 then ");
         sql.append("      discount := 0; ");         
         sql.append("   else ");
         sql.append("      discount := round((1 - (promoprice / regprice)) * 100, 2); "); 
         sql.append("   end if; ");
         sql.append("   ? := round(regprice, 2); ");
         sql.append("   ? := round(promoprice, 2); ");
         sql.append("   ? := discount; ");
         sql.append("end;");
         
         m_SellPrice = m_OraConn.prepareCall(sql.toString());
         m_SellPrice.registerOutParameter(4, Types.DOUBLE);
         m_SellPrice.registerOutParameter(5, Types.DOUBLE);
         m_SellPrice.registerOutParameter(6, Types.DOUBLE);
         
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
         
         if ( param.name.equals("cust") )
            m_CustId = param.value;
         
         if ( param.name.equals("date") )
            m_Date = param.value;
         
         if ( param.name.equals("dealerOpt") )
            m_DealerOpt = Short.parseShort(param.value);
         
         if ( param.name.equals("item") )
            m_ItemId = param.value;
         
         if ( param.name.equals("itemOpt") )
            m_ItemOpt = Short.parseShort(param.value);
         
         if ( param.name.equals("packet") )
            m_Packet = param.value;        
         
         if ( param.name.equals("show") )
            m_ShowName = param.value;
         
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
      
      if ( m_CustId != null && m_CustId.length() > 0 ) {
         fname.append("-");
         fname.append(m_CustId);         
      }
      else {
         if (m_ItemId != null && m_ItemId.length() > 0 ) {
            fname.append("-");
            fname.append(m_ItemId);            
         }
      }
      
      fname.append("-dealer_prices.txt");
      m_FileNames.add(fname.toString());
   }
}
