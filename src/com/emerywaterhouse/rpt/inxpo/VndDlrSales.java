/**
 * File: VndDlrSales.java
 * Description: Extracts the vendor / dealer(customer) sales history data for inxpo.
 *    
 * @author Jeffrey Fisher
 * 
 * Create Date: 05/09/2006
 * Last Update: $Id: VndDlrSales.java,v 1.3 2006/08/09 14:50:50 jfisher Exp $
 * 
 * History:
 */

package com.emerywaterhouse.rpt.inxpo;

import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;

/**
 *
 */
public class VndDlrSales extends Report
{
   private PreparedStatement m_CustAddr;
   private PreparedStatement m_DealerData;
   private PreparedStatement m_VendorData;
   private PreparedStatement m_VndCustSales;
   
   private String m_CustId;
   private String m_Date;
   private short m_DealerOpt;
   private String m_Packet;
   private short m_SelectOpt;
   private String m_ShowName;
   private short m_VendorOpt;
   private int m_VndId;
   
   /**
    * Default constructor
    */
   public VndDlrSales()
   {      
      super();
            
      m_MaxRunTime = RptServer.HOUR * 12;
   }
   
   /**
    * Creates the where clause for the dealer data query.  Builds in selection criteria based
    * on the index of the selection params in the sending application.
    * 
    * @return The where clause for the dealer data query.
    */
   private String buildCustWhere()
   {
      StringBuffer sql = new StringBuffer();
      sql.append("where show.name = ? and ");
      
      if ( m_SelectOpt == 1 ) {
         if ( m_CustId != null && m_CustId.length() == 6 )
            sql.append(getCustSql());
      }
      
      sql.append("dealer_show.show_id = show.show_id and ");
      sql.append("dealer.customer_id = dealer_show.customer_id ");
            
      return sql.toString();
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
      ResultSet dealerData = null;
      ResultSet vendorData = null;
      ResultSet custAddr = null;
      String custId = null;
      String company = null;
      String city = null;
      String state = null;
      double amt = 0.0;      
      int vndId;
      
      try {
         setCurAction("creating/opening output file " + m_FilePath + m_FileNames.get(0));
         outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

         setCurAction("starting vendor/dealer sales export");
         m_DealerData.setString(1, m_ShowName);         
         dealerData = m_DealerData.executeQuery();
         
         //
         // Note - this gets exected only once.
         m_VendorData.setString(1, m_ShowName);
         vendorData = m_VendorData.executeQuery();
                  
         //
         // Iterate through the dealers and generate a list of all vendors this customer
         // has purchased from in the past year.
         while ( dealerData.next() && m_Status == RptServer.RUNNING ) {
            custId = dealerData.getString(1);
            company = dealerData.getString(2);
            setCurAction("getting data for cust: " + custId);
            
            //
            // other customer data
            m_CustAddr.setString(1, custId);
            custAddr = m_CustAddr.executeQuery();
            
            if ( custAddr.next() ) {
               city = custAddr.getString(1);
               state = custAddr.getString(2);
            }
                        
            //
            // Loop through each vendor and get the sales data and write out the
            // data to disk.
            while ( vendorData.next() ) {               
               vndId = vendorData.getInt(1);
               setCurAction("getting sales for cust: " + custId + " vendor " + vndId);
               amt = getSalesAmt(m_Packet, custId, vndId);
               
               setCurAction("writing data for cust: " + custId);
               line.append("Emery\t");                               // distributor name
               line.append(custId + "\t");                           // dealer name
               line.append((company != null ? company : "") + "\t"); // company name
               line.append((city != null ? city : "") + "\t");       // city
               line.append((state != null ? state : "") + "\t");     // state
               line.append(vndId + "\t");                            // vendor id                       
               line.append(String.format("%1.2f\r\n", amt));         // sales amount
               
               outFile.write(line.toString().getBytes());
               line.delete(0, line.length());
            }
            
            //
            // Move back to the first record of the resultset.  This keeps from haveing
            // to requery the exact same data.
            vendorData.beforeFirst();
            
            //
            // reset everything
            closeRSet(custAddr);
            city = null;
            state = null;
            company = null;
         }
         
         setCurAction("finished exporting sales data");
         closeRSet(vendorData);
         closeRSet(dealerData);
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
    * Creates the where clause for the vendor data query.  Builds in selection criteria based
    * on the index of the selection params in the sending application.
    * 
    * @return The where clause for the vendor data query.
    */
   private String buildVndWhere()
   {
      StringBuffer sql = new StringBuffer();
      sql.append("where show.name = ? and ");
      
      if ( m_SelectOpt == 1 ) {
         if ( m_VndId > 0 )
            sql.append(getVndSql());
      }
      
      sql.append("vendor_show.show_id = show.show_id and ");
      sql.append("vendor.vendor_id = vendor_show.vendor_id ");
            
      return sql.toString();
   }
   
   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not garunteed to know when finalization occurs.
    */
   protected void cleanup()
   {
      closeStmt(m_CustAddr);
      closeStmt(m_DealerData);
      closeStmt(m_VendorData);
      closeStmt(m_VndCustSales);
            
      m_CustAddr = null;
      m_DealerData = null;
      m_VendorData = null;
      m_VndCustSales = null;
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
    * Creates a piece of the where clause based on the customer option parameter.  Taken from the Delphi
    * extraction program.  If the selectOpt value is 1 then that means there is a filter.  Use 0 to indicate
    * that a selected customer is needed
    *    0 none
    *    1 selected customer
    *    2 new customer
    *    3 last extract date
    *    4 all
    * 
    * @return a piece of sql code for the customer selection.
    */
   private String getCustSql()
   {
      StringBuffer sql = new StringBuffer();
      
      switch ( m_DealerOpt ) {
         case 0:
         case 1: 
            sql.append(String.format(" dealer.customer_id = '%s' and ", m_CustId));
         break;
         
         case 2:
            sql.append(" dealer.last_extract is null and ");
         break;
         
         case 3:
            sql.append(String.format(" dealer.last_extract >= to_date('%s', 'mm/dd/yyyy') and ", m_Date));
         break;      
      };
      
      return sql.toString();
   }
   
   /**
    * Gets the sales for a specific customer and vendor for one year starting based on 
    * packet report date.
    * 
    * @param pktId The packet id to report from
    * @param custId The customer
    * @param vndId The vendor
    * 
    * @return The amount of sales if any for the customer and vendor during the period.
    */
   private double getSalesAmt(String pktId, String custId, int vndId)
   {
      ResultSet sales = null; 
      double amt = 0.0;
      
      try {
         //m_VndCustSales.setString(1, pktId);
         m_VndCustSales.setString(1, custId);
         m_VndCustSales.setInt(2, vndId);
         
         sales = m_VndCustSales.executeQuery();
         
         if ( sales.next() )
            amt = sales.getDouble(1);
      }
      
      catch ( SQLException ex ) {         
         ex.printStackTrace();
      }
      
      finally {
         closeRSet(sales);
      }
      
      return amt;
   }
   
   /**
    * Creates a piece of the where clause based on the vendor selection option parameter.  Taken from the 
    * Delphi extraction program. 
    *    0 = none
    *    1 = selected vendor
    *    2 = all vendors
    * 
    * @return a piece of sql code for the vendor selection.
    */
   private String getVndSql()
   {  
      return m_VendorOpt == 1 ? String.format(" vendor.vendor_id = %d and ", m_VndId) : "";
   }
   
   /**
    * Prepares the sql queries for execution.
    */
   private boolean prepareStatements() throws Exception
   {      
      StringBuffer sql = new StringBuffer();
      boolean prepared = false;

      if ( m_OraConn != null ) {
         sql.append("select dealer.customer_id, company_name ");
         sql.append("from inxpo.dealer, inxpo.dealer_show, inxpo.show ");
         sql.append(buildCustWhere());
         sql.append("order by dealer.customer_id");
         m_DealerData = m_OraConn.prepareStatement(sql.toString());
                  
         sql.setLength(0);
         sql.append("select vendor.vendor_id ");         
         sql.append("from inxpo.vendor, inxpo.vendor_show, inxpo.show ");
         sql.append(buildVndWhere());
         sql.append("order by vendor_id");         
         m_VendorData = m_OraConn.prepareCall(sql.toString(), 
                  ResultSet.TYPE_SCROLL_INSENSITIVE, 
                  ResultSet.CONCUR_READ_ONLY);
                  
         sql.setLength(0);
         sql.append("select city, state ");
         sql.append("from cust_address_view ");
         sql.append("where customer_id = ? and addrtype = 'SHIPPING' ");         
         m_CustAddr = m_OraConn.prepareStatement(sql.toString());
         
         sql.setLength(0);
         //sql.append("select /*+rule*/sum(ext_sell) sales ");
         /*sql.append("from inv_dtl, packet, inxpo.vendor ");
         sql.append("where packet.packet_id = ? and ");
         sql.append("cust_nbr = ? and ");
         sql.append("vendor_id = ? and ");
         sql.append("sale_type = 'WAREHOUSE' and ");
         sql.append("invoice_date > packet.report_begin_date - 365 and ");
         sql.append("invoice_date <= packet.report_begin_date and ");
         sql.append("to_number(inv_dtl.vendor_nbr) = vendor.vendor_id ");
         sql.append("group by vendor_nbr");*/
         
         
         sql.append("select sum(ext_sell) sales ");
         sql.append("from inv_dtl, inxpo.vendor ");
         sql.append("where exists (");
         sql.append("   select inv_hdr_id ");
         sql.append("   from inv_hdr ");
         sql.append("   where (");
         sql.append("      invoice_date >= (sysdate - 365) and ");
         sql.append("      invoice_date <= (sysdate) and ");
         sql.append("      cust_nbr = ? and ");
         sql.append("      inv_hdr.inv_hdr_id = inv_dtl.inv_hdr_id ");
         sql.append("   )");
         sql.append(")");
         sql.append("and vendor_id = ? ");
         sql.append("and sale_type = 'WAREHOUSE' ");
         sql.append("and inv_dtl.vendor_nbr = inxpo.vendor.vendor_id ");
         sql.append("group by vendor_nbr ");
         m_VndCustSales = m_OraConn.prepareStatement(sql.toString());
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
         
         if ( param.name.equals("packet") )
            m_Packet = param.value;        
         
         if ( param.name.equals("show") )
            m_ShowName = param.value;
         
         if ( param.name.equals("cust") )
            m_CustId = param.value;
         
         if ( param.name.equals("vendor") ) {
            if ( param.value.trim().length() > 0 )
               m_VndId = Integer.parseInt(param.value.trim());
         }
         
         if ( param.name.equals("date") )
            m_Date = param.value;
         
         if ( param.name.equals("selectOpt") )
            m_SelectOpt = Short.parseShort(param.value);
         
         if ( param.name.equals("dealerOpt") )
            m_DealerOpt = Short.parseShort(param.value);
         
         if ( param.name.equals("vendorOpt") )
            m_VendorOpt = Short.parseShort(param.value);
      }
      
      //
      // Build the file name.
      fname.append(tm);
      
      if ( m_CustId != null && m_CustId.length() > 0 ) {
         fname.append("-");
         fname.append(m_CustId);
         fname.append("-");
      }
      else {
         if (m_VndId > 0 ) {
            fname.append("-");
            fname.append(m_VndId);            
         }
      }
                  
      fname.append("-vnd_dlr_sales.txt");
      m_FileNames.add(fname.toString());
   }

}
