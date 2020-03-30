package com.emerywaterhouse.rpt.export;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Types;
import java.text.DecimalFormat;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;

import com.emerywaterhouse.oag.build.noun.Invoice;
import com.emerywaterhouse.oag.build.noun.Charge;
import com.emerywaterhouse.oag.build.bod.ShowInvoice;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DataSender;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class RetailOrderChgNotif extends Report {
	
   private PreparedStatement m_Orders;
   private PreparedStatement m_FreightCharges;
   private PreparedStatement m_ShipCutStmt;
   private PreparedStatement m_ShipCutItems;
   private PreparedStatement m_BestPriceEligb;
   private PreparedStatement m_ItemDesc;  
   private CallableStatement m_GetSell;
   private CallableStatement m_GetCustSell;
   private CallableStatement m_GetCustSellPromo;
   private CallableStatement m_GetCustQtyBuySell;
   
   private String m_DataFmt;     
   private String m_CustId = "";
   private long m_OrderId;
   private final int LINE_LEN = 155;
   private final String CRLF = "\r\n";
   
   
   
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
         if ( m_DataFmt.equals("xml") )
            result = buildXml(outFile);
      }
      
      catch ( Exception ex ) {
         m_ErrMsg.append("Your report had the following errors: \r\n");
         m_ErrMsg.append(ex.getClass().getName() + "\r\n");
         m_ErrMsg.append(ex.getMessage());

         log.fatal("exception:", ex);
      }

      finally {         
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
    * Builds the customer retail order export in XML format.
    * 
    * @param outFile The file to write to.
    * @return True if the file was written to successfully, false if not.
    * 
    * @throws Exception on errors.
    */
  
   private boolean buildXml(FileOutputStream outFile) throws Exception
   {
      ShowInvoice doc = new ShowInvoice();
      Invoice invoice = null;
      Invoice.Header header = null;
      Invoice.Line line = null;
      Charge chg = null;
      ResultSet rs = null;   
      ResultSet rsFreight = null;
      boolean result = false;
      int count = 1;

      invoice = doc.addInvoice();
      header = invoice.getHeader();
      chg = header.addCharge(null, "AdditionalCharge");
      rs = m_Orders.executeQuery();
      
      try {
      	 while ( rs.next() ) {
      		      		 
      		if( count == 1 ){ 
      		   //
      	       //Freight Charges
      	       m_FreightCharges.setString(1,rs.getString("order_id"));
      	       rsFreight = m_FreightCharges.executeQuery();
      	      
      	       if( rsFreight.next() )
       	          chg.setAmount(invoice.getPrefix(), Double.toString(rsFreight.getDouble("freight")));
      	    
               header.setPoNbr(invoice.getPrefix(),rs.getString("order_id"));
               count++;
      		}
            
      		line = invoice.addLine();
      	    line.setItemId(invoice.getPrefix(), rs.getString("item_id"));
      		line.setQtyOrd(invoice.getPrefix(), rs.getString("qty_ordered"));
      		line.setQtyShipped(invoice.getPrefix(), rs.getString("qty_shipped"));
      	 }
      	 
      	 ordChangeNotifEmail(); 
    	  
      	 outFile.write(doc.toString().getBytes());
      	 result = true;
      }
      catch(Exception e){
    	  log.error("RetailOrderChgNotif.buildXml: exception while trying to build to retail web order change xml: "+e.getMessage());
      }
      finally {
         setCurAction(String.format("finished processing retail order data"));
         DbUtils.closeDbConn(null, null, rs);
         DbUtils.closeDbConn(null, null, rsFreight);
         rs = null;
         rsFreight = null;
      }
      
      return result;
   }
   
   /**
    * Private overloaded method which sends an email notification to the appropriate recipients
    * 
    * @param conn Connection - a jdbc connection reference.
    * @param emailText String - email text containing order details.
    * @throws Exception - if something went wrong whilst building the email text.
    */
   private void sendChangeNoticeEmail(Connection conn, String emailText) throws Exception
   {
      String from = "customerservice@emeryonline.com";
      String subject = "";
      subject = "Change Notice for Order "+ m_OrderId;
      log.info("sending Change Notice email to Customer:"+m_CustId+" for Order "+ m_OrderId);
      if ( conn != null ) {
    	  try {
             String[] recipList = getRecipientList(conn, m_CustId); 
             if ( emailText != null && recipList != null ) {
            	//
            	//Web service call through jar file is causing issue in the production environment, and not in test.
            	//So, use instead DataSender class 
                DataSender.smtp(from, recipList, subject, emailText);
             }
    	  }
    	  catch(Exception e){
    		  log.info("RetailOrderChgNotif.sendChangeNoticeEmail: error while sending order change notification email for order id: "+m_OrderId);
    	  }
          finally {
             //
          }
      }
   }
   
   
   /**
    * Returns a String array of email recipients for the given system, component,  
    * and notification level
    * 
    * @param conn Connection - a jdbc connection reference.
    * @param custId String - customer to whom the email is to be sent.
    * @return String[]  - email recipient list for the customer.
    */
   private String[] getRecipientList(Connection conn, String custId)
   {
      ResultSet rset = null;
      ArrayList<String> recipients = new ArrayList<String>();
      StringBuffer sql = new StringBuffer();
      String[] rlist = null;
      PreparedStatement m_EMailList = null; // gets the email recipient list
      
      try {
         sql.setLength(0);
         sql.append("select email1, email2 ");
         sql.append("from emery_contact, cust_contact, cust_contact_type ");
         sql.append("where cust_contact_type.description = 'ORDER CONFIRMATION' and ");
         sql.append("   cust_contact.cct_id = cust_contact_type.cct_id and ");
         sql.append("   emery_contact.ec_id = cust_contact.ec_id and ");
         sql.append("   cust_contact.CUSTOMER_ID = ? ");
         m_EMailList = conn.prepareStatement(sql.toString());
         m_EMailList.setString(1, custId); 
                
         rset = m_EMailList.executeQuery();
         
         if ( rset.next() ) {
        	   if( rset.getString("email1") != null && !rset.getString("email1").equals("") ) 
               recipients.add(rset.getString("email1"));
        	   if( rset.getString("email2") != null && !rset.getString("email2").equals("") )
               recipients.add(rset.getString("email2"));
         }
                  
         if ( recipients.size() == 0 )
            return null;
         
         rlist = new String[recipients.size()];
         
         for ( int i = 0; i < recipients.size(); i++ ) {
            rlist[i] = recipients.get(i);
         }
      }
      
      catch ( Exception e ) {
          e.printStackTrace();
      }
      
      finally {
         //
         // Close result set
         if ( rset != null ) {
            try {
               rset.close();
            }
            catch ( SQLException se )
            {}
            
            rset = null;
         }
         
         //
         // Close prepared statement
         if ( m_EMailList != null ) {
            try {
               m_EMailList.close();
            }
            catch ( SQLException se )
            {}
            
            m_EMailList = null;
         }
         
         sql = null;
         recipients.clear();
         recipients = null;
      }
      
      return rlist;
   }
   
   
   private boolean ordChangeNotifEmail()
   {
	  ResultSet rsShipCut = null;
	  ResultSet rsShipCutItems = null;
	  ResultSet rsBestPriceEligb = null;
	  String shipId = null;
	  long orderId;
	  boolean result = false;
	  boolean bestPriceElib = false;
	  
      try {
         //
         // Get the list of shipments that are completely cut
    	 rsShipCut = m_ShipCutStmt.executeQuery();

         while( rsShipCut.next() ) {
    	    try {
               shipId = rsShipCut.getString("ship_id");
               m_ShipCutItems.setString(1, shipId);
               rsShipCutItems = m_ShipCutItems.executeQuery();
               log.info("Order Changes found for shipment id:"+shipId);
               while( rsShipCutItems.next() ){ 
               	  orderId = (long)rsShipCutItems.getDouble("order_id");
            	  m_CustId = rsShipCutItems.getString("customer_id");
            	  log.info("Order Changes found for order id:"+orderId +"cust id: "+m_CustId);
            	  m_BestPriceEligb.setString(1,m_CustId);
            	  rsBestPriceEligb = m_BestPriceEligb.executeQuery();
            	  if(rsBestPriceEligb.next())
            	     bestPriceElib = true;            		  
            	  sendChangeNoticeEmail(m_OraConn, buildChangeNoticeEmail(m_OraConn,orderId,bestPriceElib));
               }
            }
            catch ( Exception e) {
               log.error("exception while trying to send order change notification email for Ship Id: "+ shipId+ " "+e.getMessage());
            }
            finally {
           	   DbUtils.closeDbConn(null, null, rsShipCutItems);
               rsShipCutItems = null;
            }
         }
      }
      catch ( Exception e ) {
         log.error("exception while trying to send order change notification email "+e.getMessage());
      }
      finally {
         DbUtils.closeDbConn(null, null, rsShipCut);
         rsShipCut = null;
      }
   //}
  return result;
}
  
   
   /**
    * Resource cleanup
    */
   private void closeStatements()
   {
      DbUtils.closeDbConn(m_OraConn, m_Orders, null);
   	  DbUtils.closeDbConn(null, m_FreightCharges, null);
   	  DbUtils.closeDbConn(null, m_ShipCutStmt, null);
   	  DbUtils.closeDbConn(null, m_ShipCutItems, null);
   	  DbUtils.closeDbConn(null, m_BestPriceEligb, null);
   	  DbUtils.closeDbConn(null, m_ItemDesc, null);
   	  DbUtils.closeDbConn(null, m_GetSell, null);
   	  DbUtils.closeDbConn(null, m_GetCustSell, null);
   	  DbUtils.closeDbConn(null, m_GetCustSellPromo, null);
   	  DbUtils.closeDbConn(null, m_GetCustQtyBuySell, null);
   	  m_Orders = null;
   	  m_OraConn = null;
   	  m_FreightCharges = null;
   	  m_ShipCutStmt = null;
   	  m_ShipCutItems = null;
      m_BestPriceEligb = null;
      m_ItemDesc = null;
      m_GetSell = null;
      m_GetCustSell = null;
      m_GetCustSellPromo = null;
      m_GetCustQtyBuySell = null;
   	
   }

	/* (non-Javadoc)
	 * @see com.emerywaterhouse.rpt.server.Report#createReport()
	 */
	@Override
	public boolean createReport()
	{
      boolean created = false;
      boolean prepareCreated = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         m_OraConn = m_RptProc.getOraConn();         
         prepareCreated = prepareStatements();
         if(prepareCreated){
            created = buildOutputFile();
         }
         
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
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer();
      boolean isPrepared = false;
           
      if ( m_OraConn != null ) {
         try {
           
           	sql.setLength(0);
        	sql.append("select ");
        	sql.append("oh.order_id,ol.item_id,ol.qty_ordered,ol.qty_shipped,ol.qty_cut ");
        	sql.append("from  ");      
        	sql.append("order_header oh ");
        	sql.append("join order_line ol on oh.order_id = ol.order_id and ");
        	sql.append("ol.qty_cut > 0 ");
        	sql.append("join order_method om on oh.order_method_id = om.order_method_id and ");
        	sql.append("om.description = 'RETAIL WEB' ");
        	sql.append("join trip_stop on trip_stop.trip_stop_id = ol.trip_stop_id  ");
        	sql.append("join trip on trip.trip_id = trip_stop.trip_id and ");
        	sql.append("trip.load_date = trunc(sysdate) "); 
        	sql.append("order by oh.order_id ");
           	m_Orders = m_OraConn.prepareStatement(sql.toString());
        	
        	sql.setLength(0);
            sql.append("select ");
            sql.append("   sum(amount) as freight ");
            sql.append("from   ");     
            sql.append("   invoice_adder ");
            sql.append("where  ");
            sql.append("   invoice_num in (select distinct(invoice_num) from order_line where order_id = ?)");
            m_FreightCharges = m_OraConn.prepareStatement(sql.toString());
        	
        	m_ItemDesc = m_OraConn.prepareStatement(
    		         "select description from item where item_id = ?"
    		);

    		//
    		// Prepare statement that gets today's sell price.
    		m_GetSell = m_OraConn.prepareCall(
    		   "begin " +
    		      "? := item_price_procs.todays_sell(?); " +
    		   "end;"
    		);
    		m_GetSell.registerOutParameter(1, Types.DOUBLE);


    		//
    		// Prepare statement that gets customer-specific sell price.
    		m_GetCustSell = m_OraConn.prepareCall(
    		           "begin " +
    		               "? := cust_procs.getsellprice(?, ?); " +
    		            "end;"
    		);
    		m_GetCustSell.registerOutParameter(1, Types.DOUBLE);

    		//
    		// Prepare statement that gets customer-specific promo sell price.
    		m_GetCustSellPromo = m_OraConn.prepareCall(
    		   "begin " +
    		      "? := cust_procs.getsellprice(?, ?, ?); " +
    		   "end;"
    		);
    		m_GetCustSellPromo.registerOutParameter(1, Types.DOUBLE);

            //
            // Prepare statement that gets qty buy pricing (with promo, if included)
            sql.setLength(0);
            sql.append("begin ");
            sql.append("? := cust_procs.getsellprice(?, ?, ?, null, ?); ");
            sql.append("end;");
            m_GetCustQtyBuySell = m_OraConn.prepareCall(sql.toString());
            m_GetCustQtyBuySell.registerOutParameter(1, Types.DOUBLE);
    		    
    	    //
            //Get shipments which have at least one item cut.
            sql.setLength(0);
            sql.append("select distinct shipment_item.ship_id ");
            sql.append("from shipment_item ");
            sql.append("join shipment on shipment.ship_id = shipment_item.ship_id ");
            sql.append("join trip_stop on trip_stop.trip_stop_id = shipment.trip_stop_id "); 
            sql.append("join trip on trip.trip_id = trip_stop.trip_id and ");
            sql.append("trip.load_date = trunc(sysdate) ");
            sql.append("join carrier on carrier.carrier_id = trip.carrier_id and ");
            sql.append("carrier.is_tms = 1 ");
            sql.append("group by shipment_item.ship_id ,qty_cut ");
            sql.append("having  qty_cut > 0 ");
            sql.append("order by ship_id ");
            m_ShipCutStmt = m_OraConn.prepareStatement(sql.toString());
      
            //
            //Get list of shipment items which are completely cut.
            sql.setLength(0);
            sql.append("select distinct(order_line.order_id),order_header.customer_id ");
           	sql.append("from shipment_item ");
           	sql.append("join shipment on shipment.ship_id = shipment_item.ship_id and ");    
            sql.append("shipment.ship_id = ? ");
            sql.append("join trip_stop on trip_stop.trip_stop_id = shipment.trip_stop_id ");  
            sql.append("join trip on trip.trip_id = trip_stop.trip_id "); 
            sql.append("join order_line on order_line.trip_stop_id = trip_stop.trip_stop_id ");
            sql.append("and order_line.qty_cut > 0 ");
            sql.append("join order_header on order_header.order_id = order_line.order_id "); 
            sql.append("join order_method on order_header.order_method_id = order_method.order_method_id "); 
            sql.append("and order_method.description = 'EMERY API' ");
            sql.append("order by order_line.order_id ");
            m_ShipCutItems = m_OraConn.prepareStatement(sql.toString());   
        
            sql.setLength(0);
            sql.append("select cust_price_method.customer_id ");
       	    sql.append("from ");
       	    sql.append("   cust_price_method,price_method ");
       	    sql.append("where ");
       	    sql.append("price_method.price_method_id = cust_price_method.price_method_id and ");
       	    sql.append("cust_price_method.customer_id = ? and ");
       	    sql.append("price_method.description = 'ELIGIBLE FOR LOWEST PRICE' ");
       	    m_BestPriceEligb = m_OraConn.prepareStatement(sql.toString());
                 
            isPrepared = true;
        }
        catch ( SQLException ex ) {
           log.error("RetailOrderChgNotif.prepareStatements:", ex);
        }
         
        finally {
           sql = null;
        }         
      }
      else
         log.error("RetailOrderChgNotif.prepareStatements - null oracle connection");
      
      return isPrepared;
   }
   
   /**
    * Sets the parameters of this report.
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer fileName = new StringBuffer();      
      String tmp = Long.toString(System.currentTimeMillis());
      int pcount = params.size();
      Param param = null;
      
      for ( int i = 0; i < pcount; i++ ) {
         param = params.get(i);
         
         if ( param.name.equals("datafmt") )
            m_DataFmt = param.value;
      }
            
      fileName.append(tmp);
      fileName.append("-");
      
      fileName.append("RetailOrderChanges.xml");
      m_FileNames.add(fileName.toString());
   } 
   
   /**
    * Build the order change notification email text.
    *
    * @param conn Connection - a jdbc connection reference.
    * @param orderId long  - the order through which data is pulled from the database.
    * @param isEligBestPr boolean - whether or not the customer is eligible for Best Price.
    * @return String  - the email text containing the order details.
    * @throws Exception - if something went wrong whilst building the email text.
    */
   public String buildChangeNoticeEmail(Connection conn, long orderId, boolean isEligBestPr)
   {
	  StringBuffer msg = new StringBuffer(1024);
      StringBuffer line = new StringBuffer(LINE_LEN);
      char[] ch = new char[LINE_LEN];
      DecimalFormat fmt = new DecimalFormat("$0.00");
      DecimalFormat costfmt = new DecimalFormat("$0.000");
      String tmp = null;
      double sellPrice = 0.0;
      double ordTot = 0.0;
      double extCost = 0.0;
      String promo = null;
      String poNum = "";
      String orderDate = "";
      String itemId = "";
      StringBuffer ord = new StringBuffer();
      StringBuffer ordHdr = new StringBuffer();
      Statement ordStmt = null;
      Statement ordHdrStmt = null;
      ResultSet ordRset = null;
      ResultSet ordHdrRset = null;
      Format formatter = null;
      int j = 0;
       
      m_OrderId = orderId;
      
      if ( conn != null ) {
         try {
    	                 	
            //
            //check if the order submitted has any error lines
    	    ordHdr.setLength(0); 
    	    ordHdr.append("select order_date, po_num, customer_id ");
    	    ordHdr.append("from order_header ");
    	    ordHdr.append("where order_id = ");
    	    ordHdr.append(m_OrderId);
    	    ordHdrStmt = conn.createStatement();
            ordHdrRset = ordHdrStmt.executeQuery(ordHdr.toString());
         
            if( ordHdrRset.next() ) {
               formatter = new SimpleDateFormat("MM/dd/yyyy");
        	   if ( ordHdrRset.getString("customer_id") != null )   
                  m_CustId = ordHdrRset.getString("customer_id");
        	   if ( ordHdrRset.getDate("order_date") != null )   
                  orderDate = formatter.format(ordHdrRset.getDate("order_date"));
               if ( ordHdrRset.getString("po_num") != null )   
                  poNum = ordHdrRset.getString("po_num");
            }
         
            msg.append("Please note the following changes to order ");
            msg.append(m_OrderId);
            msg.append(" placed with Emery Waterhouse.\r\n\r\n");
            msg.append("Order Date  :");
            msg.append(orderDate);
            msg.append("\r\nOrder ID    :" + m_OrderId);
            msg.append("\r\nPO #        :" + poNum);
            msg.append(CRLF);
         
            Arrays.fill(ch, 0, LINE_LEN, ' ');
            line.setLength(0);
            line.append(ch);
            line.replace(0, 14, "Order details:");
            msg.append(CRLF);
            msg.append(CRLF);
            msg.append(line);
            msg.append(CRLF);
            
            //
            //build the order line(s) 
            ord.setLength(0);
            ord.append("select ol_id,item_id,promo_id,qty_ordered,qty_shipped,qty_cut,comments ");
            ord.append("from order_line ");
            ord.append("where order_id = ");
            ord.append(m_OrderId);
            ord.append(" ");
            ord.append("order by line_seq asc");
            ordStmt = conn.createStatement();
            ordRset = ordStmt.executeQuery(ord.toString());
         
         while ( ordRset.next() ) {
            
            if( j == 0 ) {
               line.setLength(0);
               line.append(ch);
                             
               line.replace(0, 3, "Item");
               line.replace(10, 25, "Description");
               line.replace(50, 60, "Qty Ord");
               line.replace(64, 74, "Qty Ship");
               line.replace(77, 87, "Qty Cut");
               line.replace(92, 111, "Cut Reason");
               line.replace(112, 120, "Cost");
               line.replace(122, 132, "Ext Cost");
                             
               msg.append(line);
               msg.append(CRLF);         
            
               j++;
            }
               
            
            line.setLength(0);
            line.append(ch);  
            //
            // clear previous value
            promo = "";
            itemId = "";
            
          
           
            if( ordRset.getString("item_id") != null ) {
               itemId = ordRset.getString("item_id");
               line.replace(0, 6, itemId);
            }
            
            //
            // Add line item description
            if( itemId != null && !itemId.equals("") ) {
               tmp = getItemDesc(itemId);
               line.replace(10, 10 + tmp.length(), tmp);
            }
            
            //Qty ordered
            tmp = Integer.toString(ordRset.getInt("qty_ordered"));
            line.replace(50, 50 + tmp.length(), tmp);
            
            //Qty shipped
            tmp = Integer.toString(ordRset.getInt("qty_shipped"));
            line.replace(64, 64 + tmp.length(), tmp);
            
            //Qty cut
            tmp = Integer.toString(ordRset.getInt("qty_cut"));
            line.replace(77, 77 + tmp.length(), tmp);
            
            //Cut Reason
            if( ordRset.getString("comments") != null ) {
               tmp = ordRset.getString("comments");
               line.replace(92, 92 + tmp.length(), tmp);
            }
            
            //
            // Get the promo id if present
            if ( ordRset.getString("promo_id") != null ) {
               promo = ordRset.getString("promo_id");
            }
                        
            //
            // Only show appropriate customers best price information
            if ( isEligBestPr ) {
                                      
               if ( promo == null || promo.equals("") ) {
                  sellPrice = getSellPrice(ordRset.getString("item_id"),ordRset.getInt("qty_shipped"));
                  tmp = costfmt.format(sellPrice);
                  extCost = sellPrice * ordRset.getInt("qty_shipped");
               }
               else {
                  sellPrice = getSellPrice(ordRset.getString("item_id"),ordRset.getInt("qty_shipped"),promo);
                  tmp = costfmt.format(sellPrice);
                  extCost = sellPrice * ordRset.getInt("qty_shipped");
               }
               
               line.replace(112, 112 + tmp.length(), tmp);
   
               tmp = fmt.format(extCost);
               line.replace(122, 122 + tmp.length(), tmp);
            }
            else {
               if (promo == null || promo.equals("") ) {
                  sellPrice = getSellPrice(ordRset.getString("item_id"),ordRset.getInt("qty_shipped"));
                  tmp = costfmt.format(sellPrice);
                  extCost = sellPrice * ordRset.getInt("qty_shipped");
               }
               else {
                  sellPrice = getSellPrice(ordRset.getString("item_id"),ordRset.getInt("qty_shipped"),promo);
                  tmp = costfmt.format(sellPrice);
                  extCost = sellPrice * ordRset.getInt("qty_shipped");
               }
               
               line.replace(112, 112 + tmp.length(), tmp);
   
               tmp = fmt.format(extCost);
               line.replace(122, 122 + tmp.length(), tmp);
            }

            ordTot = ordTot + extCost;

            msg.append(line);
            msg.append(CRLF);
                      
         }
         
         if( j > 0 ) { // execute this code only for successful order line(s).
            msg.append(CRLF);
            msg.append("Total Cost: ");
            msg.append(fmt.format(ordTot));
            msg.append(CRLF);
            msg.append("\r\nOrder subject to vendor price at time of shipment.\r\n");
         }
         
         msg.append("If you have any questions about your order, call customer service:\r\n");
         msg.append("(800) 283-0236 option 1\r\n");
         msg.append(CRLF);
      }
      catch ( Exception e) {
          log.error("RetailOrderChgNotif.buildChangeNoticeEmail: exception while trying to build change notification email for Order Id: "+ orderId+ " "+e.getMessage());
       }
         
      finally {
         
         if ( ordRset != null ) {
            try {
               ordRset.close();
            }
            catch ( Exception ex )
            {}

            ordRset = null;
         }
         
         if ( ordStmt != null ) {
            try {
               ordStmt.close();
            }
            catch ( Exception ex )
            {}

            ordStmt = null;
         }
                  
         ord = null;
         ordHdr = null;
         tmp = null;
         promo = null;
         
      }
   }
           
      return msg.toString();
   }
 
   /**
    * Returns the item description for a specific item.  Used by the buildEmailText() method.
    *
    * @param itemId String - the id number of the item to retrieve the description for.
    * @return String the description of the item in itemId if the item is found.  If the item is
    *    not found, then an empty string is returned.
    */
   protected String getItemDesc(String itemId)
   {
      String desc = "";
      ResultSet rset = null;

      try {
         m_ItemDesc.setString(1, itemId);
         rset = m_ItemDesc.executeQuery();

         if ( rset.next() ){
            desc = rset.getString(1);
            if (desc.length() >= 35)
               desc = desc.substring(0, 35);
         }
            
      }

      catch ( SQLException e) {
         desc = Integer.toString(e.getErrorCode());
      }

      catch ( Exception ex ) {
         desc = ex.getMessage();

         if ( desc != null ) {
            if ( desc.length() > 35 )
               desc = desc.substring(0, 35);
         }
         else
            desc = "exception";
      }

      finally {
         if( rset!= null ) {
            try {
               rset.close();
            }
            catch ( Exception ex )
            {}
            rset = null;
         }
      }

      return desc;
   }
      
   /**
    * Gets the sell price for the current item and customer.
    *
    * @param m_ItemId String - the id number of the item to calculate its sell price.
    * @param m_Quantity int - qty of the item ordered.
    * @return double - the sell price of the item
    */
   private double getSellPrice(String m_ItemId,int m_Quantity)
   {
      double price = 0.0;

      //
      // Get customer specific sell price.
      if ( m_ItemId != null && m_CustId != null ) {
         try {
            if ( m_Quantity > 0 ) { // Qty buy pricing change
               m_GetCustQtyBuySell.setString(2, m_CustId);
               m_GetCustQtyBuySell.setString(3, m_ItemId);
               m_GetCustQtyBuySell.setNull(4, Types.VARCHAR);
               m_GetCustQtyBuySell.setInt(5, m_Quantity);
               m_GetCustQtyBuySell.registerOutParameter(1, Types.DOUBLE);
               m_GetCustQtyBuySell.executeUpdate();
   
               price = m_GetCustQtyBuySell.getDouble(1);
            }
            else {
               m_GetCustSell.setString(2, m_CustId);
               m_GetCustSell.setString(3, m_ItemId);
               m_GetCustSell.registerOutParameter(1, Types.DOUBLE);
               m_GetCustSell.executeUpdate();
   
               price = m_GetCustSell.getDouble(1);
            }
         }
         catch ( Exception e ) {
            //
            // If any exception occurred, get the regular base cost
            try {
               m_GetSell.setString(2, m_ItemId);
               m_GetSell.registerOutParameter(1, Types.DOUBLE);
               m_GetSell.executeUpdate();

               price = m_GetSell.getDouble(1);
            }
            catch ( Exception e2 ) {
               price = 0.0;
            }
         }
               
      }

      return price;
   }

   /**
    * Gets the promotional sell price for the current item, customer and the
    * specified promotion identifier.
    *
    * @param m_ItemId String - the id number of the item to calculate its sell price.
    * @param m_Quantity int - qty of the item ordered.
    * @param promo String - available promotion for the item in question.
    * @return   the promotional sell price for the current customer and item.
    */
   private double getSellPrice(String m_ItemId,int m_Quantity,String promo)
   {
      double price = 0.0;

      if ( m_ItemId != null && m_CustId != null && promo != null ) {
         try {
            if ( m_Quantity > 0 ) { // Qty buy pricing change  
               m_GetCustQtyBuySell.setString(2, m_CustId);
               m_GetCustQtyBuySell.setString(3, m_ItemId);
               m_GetCustQtyBuySell.setString(4, promo);
               m_GetCustQtyBuySell.setInt(5, m_Quantity);
               m_GetCustQtyBuySell.registerOutParameter(1, Types.DOUBLE);
               m_GetCustQtyBuySell.executeUpdate();
   
               price = m_GetCustQtyBuySell.getDouble(1);
            }
            else {
               m_GetCustSellPromo.setString(2, m_CustId);
               m_GetCustSellPromo.setString(3, m_ItemId);
               m_GetCustSellPromo.setString(4, promo);
               m_GetCustSellPromo.registerOutParameter(1, Types.DOUBLE);
               m_GetCustSellPromo.executeUpdate();
   
               price = m_GetCustSellPromo.getDouble(1);
            }
         }
         catch ( Exception e ) {
            price = 0.0; 
         }
                  
      }

      return price;
   }
   
}

