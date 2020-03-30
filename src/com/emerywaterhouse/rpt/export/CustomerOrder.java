/**
 * Title:			CustomerOrder.java
 * Description:
 * Company:			Emery-Waterhouse
 * @author			naresh pasnur
 * <p>
 * Create Date:	08/18/2010
 * <p>
 * History:
 *		$Log: CustomerOrder.java,v $
 *		Revision 1.12  2010/09/29 23:20:43  npasnur
 *		dev commit
 *
 *		Revision 1.11  2010/09/27 16:06:54  npasnur
 *		dev commit
 *
 *		Revision 1.10  2010/09/25 19:41:23  npasnur
 *		Reverted back some changes.
 *
 *		Revision 1.8  2010/09/25 18:25:52  npasnur
 *		dev commit
 *
 *		Revision 1.7  2010/09/24 13:21:35  npasnur
 *		dev commit
 *
 *		Revision 1.6  2010/09/23 10:11:20  npasnur
 *		dev commit
 *
 *		Revision 1.5  2010/09/20 14:20:09  npasnur
 *		dev commit
 *
 *		Revision 1.4  2010/09/18 16:15:26  npasnur
 *		dev commit
 *
 *		Revision 1.3  2010/09/13 06:19:20  npasnur
 *		dev commit
 *
 *		Revision 1.2  2010/09/05 08:54:26  npasnur
 *		dev commit
 *
 *		Revision 1.1  2010/08/22 20:34:57  npasnur
 *		initial commit
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
import com.emerywaterhouse.oag.build.noun.Invoice;
import com.emerywaterhouse.oag.build.noun.Charge;
import com.emerywaterhouse.oag.build.bod.ShowInvoice;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.utils.DbUtils;
import com.emerywaterhouse.websvc.Param;

public class CustomerOrder extends Report {
	
   private String m_DataFmt;     // The output format, xml
   
   private PreparedStatement m_Orders;
   private PreparedStatement m_FreightCharges;
   
   
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

      System.out.println("Building xml file ");      
      
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
      	 outFile.write(doc.toString().getBytes());
      	 result = true;
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
    * Resource cleanup
    */
   private void closeStatements()
   {
   	DbUtils.closeDbConn(m_OraConn, m_Orders, null);
   	DbUtils.closeDbConn(m_OraConn, m_FreightCharges, null);
   	m_Orders = null;
   	m_FreightCharges = null;
   	m_OraConn = null;
   }

	/* (non-Javadoc)
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
    */
   private boolean prepareStatements()
   {      
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;
      
      if ( m_OraConn != null ) {
         try {
           
           	sql.setLength(0);
            sql.append("select ");
            sql.append("   oh.order_id,ol.item_id,ol.qty_ordered,ol.qty_shipped,ol.qty_cut ");
            sql.append("from   ");     
            sql.append("   order_header oh, order_line ol ");
            sql.append("where ");
            sql.append("   oh.order_id = ol.order_id and ");
            sql.append("   oh.order_method_id = 23 and ");
            sql.append("   trunc(oh.date_entered) >= trunc(sysdate) - 1 ");
            sql.append("   and ol.qty_cut > 0 ");
            sql.append("   order by oh.order_date desc ");
            
            if( m_OraConn != null ){
            	log.info("true");
            }
            else
            	log.info("false");
                        
        	m_Orders = m_OraConn.prepareStatement(sql.toString());
        	
        	sql.setLength(0);
            sql.append("select ");
            sql.append("   sum(amount) as freight ");
            sql.append("from   ");     
            sql.append("   invoice_adder ");
            sql.append("where  ");
            sql.append("   invoice_num in (select distinct(invoice_num) from order_line where order_id = ?)");
                    	
        	m_FreightCharges = m_OraConn.prepareStatement(sql.toString());
                 
            isPrepared = true;
         }
         
         catch ( SQLException ex ) {
            log.error("exception:", ex);
         }
         
         finally {
            sql = null;
         }         
      }
      else
         log.error("CustomerOrder.prepareStatements - null oracle connection");
      
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
      
      fileName.append("Emeryorders.xml");
      m_FileNames.add(fileName.toString());
   } 
      
   
}

