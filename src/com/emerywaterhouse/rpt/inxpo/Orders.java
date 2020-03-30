/**
 * File: Orders.java
 * Description: Extracts the order data for pre show orders.
 *    
 * @author Jeffrey Fisher
 * 
 * Create Date: 06/13/2006
 * Last Update: $Id: Orders.java,v 1.4 2008/10/30 15:47:21 jfisher Exp $
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

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

public class Orders extends Report
{
   private PreparedStatement m_Orders;
   private PreparedStatement m_CustReg;
   
   private String m_CustId;
   private String m_Packet;
   private int m_ShowId;
   
   /**
    * Default constructor.  Setup base elements.
    */
   public Orders()
   {
      super();
      
      m_CustId = null;
            
      m_FileNames.add("orders.txt");
      m_FileNames.add("orders.err");
      m_MaxRunTime = RptServer.HOUR * 3;
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
      FileOutputStream errFile = null;
      boolean result = false;
      boolean registered = false;
      ResultSet orders = null;
      String custId = null;      
      String itemId = null;
      String lastCustId = "";
      String poNum = null;
      int qty = 0;
      double retail = 0;
      
      try {
         setCurAction("creating/opening output file " + m_FilePath + m_FileNames.get(0));
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);
         errFile = new FileOutputStream(m_FilePath + m_FileNames.get(1), false);

         setCurAction("starting item detail export");         
         m_Orders.setString(1, m_Packet);
         
         if ( m_CustId.length() > 0 )
            m_Orders.setString(2, m_CustId);
         
         orders = m_Orders.executeQuery();
         
         //
         // Iterate through the records and create a line of delimited text based on the Inxpo
         // file layout.
         while ( orders.next() && m_Status == RptServer.RUNNING ) {
            custId = orders.getString(1);
            poNum = orders.getString(2);            
            itemId = orders.getString(3);
            retail = orders.getDouble(4);            
            qty =  orders.getInt(5);
                            
            if ( poNum == null )
               poNum = "onlineshow";
                           
            setCurAction("writing data for customer: " + custId);
            line.append("Emery\t");                         // distributor name
            line.append(custId + "\t");                     // customer id
            line.append("\t");                              // bill to id
            line.append("\t");                              // ship to id
            line.append(poNum + "\t");                      // order title (PO Number)            
            line.append(itemId + "\t");                     // distributor product id
            line.append("\t");                              // unit of measure            
            line.append(qty + "\t");                        // order qty
            line.append("\t");                              // order date
            
            //
            // Only add the retail override if there was one.  
            // Null values come back as 0.
            // field = notes (retail override)
            line.append(
               (retail > 0 ? Double.toString(retail) : "") + "\r\n"
            );  
            
            //
            // Check to see if the customer is registered for the current show.
            // Only check on new customers.
            if ( !lastCustId.equals(custId) ) {
               registered = custRegistered(custId);
               lastCustId = custId;
            }
            
            if ( registered )
               outFile.write(line.toString().getBytes());
            else
               errFile.write(line.toString().getBytes());

            line.delete(0, line.length());
         }
         
         setCurAction("finished exporting orders");
         closeRSet(orders);
         
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
         
         if ( errFile != null ) {
            try {
               errFile.close();
            }
            catch ( IOException iex ) {
               ;
            }
         }

         outFile = null;
         errFile = null;         
      }
            
      return result;
   }
   
   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not garunteed to know when finalization occurs.
    */
   protected void cleanup()
   {
      closeStmt(m_Orders);
      
      m_Orders = null; 
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
    * Determines if a customer has registered for the online show.
    * @param custId
    * @return boolean True if the customer was registered, false if not.
    */
   private boolean custRegistered(String custId)
   {
      boolean registered = false;
      ResultSet rset = null;
      
      if ( custId != null ) {
         try {
            m_CustReg.setInt(1, m_ShowId);
            m_CustReg.setString(2, custId);            
            rset = m_CustReg.executeQuery();
            
            if ( rset.next() ) {               
               registered = (rset.getString(1) != null);
            }
         }
         
         catch ( SQLException ex ) {
            log.error("exception", ex);
         }
         
         finally {
            closeRSet(rset);
         }
      }
      
      return registered;
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements() throws Exception
   {      
      StringBuffer sql = new StringBuffer();
      boolean prepared = false;

      if ( m_OraConn != null ) {
         //
         // Customer service will manually release orders even though they are supposed to be pre show only.
         // We have to filter out those orders that are released or shipped.  Backorder is OK.
         sql.setLength(0);
         sql.append("select customer_id, po_num, item_id, ");
         sql.append("decode(retail_override_user, null, null, retail_price) retail, ");
         sql.append("sum(qty_ordered) qty ");
         sql.append("from order_header, order_line, \"promotion\"@edbprod, order_status ");
         sql.append("where ");
         sql.append("   \"promotion\".\"packet_id\" = ? and ");
         sql.append("   description in ('NEW', 'BACKORDERED') and ");
         
         if ( m_CustId.length() > 0 )
            sql.append("   customer_id = ? and ");
         
         sql.append("   \"promotion\".\"promo_id\" = order_line.promo_id and ");
         sql.append("   order_header.order_id = order_line.order_id and ");
         sql.append("   order_line.order_status_id = order_status.order_status_id ");
         sql.append("group by customer_id, po_num, item_id, ");
         sql.append("   decode(retail_override_user, null, null, retail_price) ");         
         sql.append("order by customer_id, item_id, po_num");
         m_Orders = m_OraConn.prepareStatement(sql.toString());
                  
         sql.setLength(0);
         sql.append("select customer_id ");
         sql.append("from inxpo.dealer_show ");
         sql.append("where show_id = ? and customer_id = ? ");
         m_CustReg = m_OraConn.prepareStatement(sql.toString());
         
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
      int pcount = params.size();
      Param param = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
         
         if ( param.name.equals("packet") )
            m_Packet = param.value;
         
         if ( param.name.equals("cust") )
            m_CustId = param.value;
         
         if ( param.name.equals("showid") ) {
            if ( param.value != null && param.value.length() > 0 )
               m_ShowId = Integer.parseInt(param.value);
         }
      }
   }
}
