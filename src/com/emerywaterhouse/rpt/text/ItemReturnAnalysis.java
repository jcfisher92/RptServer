/**
 * File: ItemReturnAnalysis.java
 * Description: ItemReturnAnalysis produces a tab delimited file, one line per
 *    customer, with the following columns:
 *       customer id
 *       customer name
 *       status
 *       sales rep
 *       substitutions accepted?
 *       total sales - net of credits
 *       # of lines shipped
 *       # of lines returned
 *       total handling fees for returns
 *       total credits
 *       plus columns breaking down # of lines and $ of returns by reason/disposition
 *       
 *    The original Author was Peggy Richter.  It's been rewritten to work with the new report server.
 *
 * @author Peggy Richter
 * @author Jeffrey Fisher
 *
 * Create Date: 05/13/2005
 * Last Update: $Id: ItemReturnAnalysis.java,v 1.13 2008/10/30 16:53:53 jfisher Exp $
 * 
 * History
 *    03/25/2005 - Added log4j logging. jcf
 * 
 *    05/03/2004 - Removed the setting of the m_DistList variable when sending the report notification.  This variable
 *       will get cleaned up before it can be used by the webservice. - jcf
 *
 *    04/07/2004 - Applied Email class changes. - jcf
 *
 *    04/05/2004 - Include all 4 digits of year in file names.  The last 2 digits of the year were being chopped off.  pjr
 *
 *    12/23/2003 - Modified the email and params to handle the xml request protocol and running this report as
 *       a webservice.  Also changed the sql from concatenated strings to a stringbuffer. - jcf
 *
 *    08/11/2003 - Change the query so it uses items acutally shipped. - jcf
 *
 *    03/19/2002 - Removed the use of the connection pool connections and replaced it with the call
 *       to getOracleConn().  Also added the cleanup() method. - jcf
 *
 *    11/7/2001   Changed logging to use getProcName() - jcf
 */
package com.emerywaterhouse.rpt.text;

import java.io.FileOutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.emerywaterhouse.rpt.server.Report;
import com.emerywaterhouse.rpt.server.RptServer;
import com.emerywaterhouse.websvc.Param;
import com.emerywaterhouse.fascor.Facility;
import com.emerywaterhouse.rpt.helper.ReasonDispCode;


public class ItemReturnAnalysis extends Report
{
   private String m_BegDate;
   private String m_EndDate;
   private PreparedStatement m_Cust;
   private PreparedStatement m_CustStatus;
   private PreparedStatement m_Sales;
   private PreparedStatement m_LinesShip;
   private PreparedStatement m_LinesRet;
   private PreparedStatement m_Handling;
   private PreparedStatement m_CreditAmt;
   private PreparedStatement m_LinesByReason;
   private PreparedStatement m_AmtByReason;
   private PreparedStatement m_ReasonDisp;
   private List<Facility> m_FacilityList;             // List of all whse facilities
   private List<ReasonDispCode> m_ReasonDispCodes;    // List of all reason, disp & Facility     
   private String m_FacilityCondition;                // SQL facility condition fragment.
   private String m_WhseId;                           // Facility input parameter   
     
   /**
    * default constructor
    */
   public ItemReturnAnalysis()
   {
      super();
      
      m_MaxRunTime = RptServer.HOUR * 8;
      m_FacilityList = new ArrayList<Facility>();   
      m_ReasonDispCodes = new ArrayList<ReasonDispCode>();
   }
   
   /**
    * Builds a facility sql condition snippet for use in the prep. statements, 
    * of the form: ('PORTLAND', 'PITSTON').
    * 
    * @return String - the facility sql condition snippet.
    * @throws SQLException
    */
   private String buildFacilityCondition() 
   {
      StringBuffer tmp = new StringBuffer();
      int i = 1;      
      
      try {
         tmp.append("(");
         for (Facility f : this.getFacilityList()){
            tmp.append("'");
            tmp.append(f.getName());
            tmp.append("'");
            tmp.append(i != this.getFacilityList().size() ? ", " : "");
            i++;            
         }
         
         tmp.append(")");

      }
      catch(Exception e) {
         log.fatal("ItemReturnAnalysis.buildFacilityCondition ", e);
         m_ErrMsg.append("The report had the following Error(s) in: \r\n");
         m_ErrMsg.append(e.getClass().getName() + "\r\n" + e.getMessage());
      }
      return tmp.toString();
   }   

   /**
    * Builds the output file.
    * @return boolean
    *    true if the file was built.<p>
    *    false if there was a problem.<p>
    * @throws Exception if unable to open output file or other.
    */
   public boolean buildOutputFile() throws Exception
   {      
      StringBuffer line = new StringBuffer(1024);
      FileOutputStream outFile = null;
      boolean result = true;
      String custId = null;
      String status = null;
      String salesRep = null;
      String canSub = null;
      String custName = null;
      ReasonDispCode rdCode = null;     //Temporary Reason & Disposition holder.
      
      //
      //Some maps and list to facilitate drawing variable length columnar data
      //without a bunch of multiple dimensional arrays.
      Map<Facility, Double> sales = new HashMap<Facility, Double>();
      Map<Facility, Integer> linesShipped = new HashMap<Facility, Integer>();
      Map<Facility, Integer> linesRet = new HashMap<Facility,Integer>();
      Map<Facility, Double> handling = new HashMap<Facility,Double>();
      Map<Facility, Double> creditAmt = new HashMap<Facility,Double>();
      Map<ReasonDispCode, Integer> linesByCust = new HashMap<ReasonDispCode, Integer>();
      Map<ReasonDispCode, Double> amtByCust = new HashMap<ReasonDispCode, Double>();      

      ResultSet custData = null;
      ResultSet statusData = null;
      ResultSet salesData = null;
      ResultSet linesShipData = null;
      ResultSet linesRetData = null;
      ResultSet handlingData = null;
      ResultSet creditAmtData = null;
      ResultSet linesByReasonData = null;
      ResultSet amtByReasonData = null;
      ResultSet reasonDispData = null;

      outFile = new FileOutputStream(m_FilePath + m_FileNames.get(0), false);

      try {
         
         //
         // Add the date range of the report to the header in the customer name column
         line.append("\tFrom: " + m_BegDate + "  to: " + m_EndDate + "\r\n");

         // Build the Captions
         line.append("Customer ID\tName\tStatus\tSale Rep\tSub?\t");
         
         //
         //This column is per facility
         for (Facility f : this.getFacilityList()){
            line.append(f.getName() + " Sales\t");
         }
         
         //
         //This column is per facility
         for (Facility f : this.getFacilityList()){
            line.append(f.getName() + " Lines Ship\t");
         }     
         
         //
         //This column per facility
         for (Facility f : this.getFacilityList()){
            line.append(f.getName() + " Lines Ret\t");
         }   
         
         //
         //This column per facility
         for (Facility f : this.getFacilityList()){
            line.append(f.getName() + " Handling\t");
         }   
         
         //
         //This column per facility
         for (Facility f : this.getFacilityList()){
            line.append(f.getName() + "Total Credits\t");
         }          
         
         try {
            reasonDispData = m_ReasonDisp.executeQuery();
            
            //
            //reasonDispData contains every combination of reason/dispossition
            //codes found within the selected time frame.  These become
            //the column headings for the lines/$$$ by reason/disp section
            //of the report
            while ( reasonDispData.next() && m_Status == RptServer.RUNNING ) {
               rdCode = new ReasonDispCode(this.getFacilityByName(reasonDispData.getString("warehouse")), 
                     reasonDispData.getString( "return_reason_cd" ), 
                     reasonDispData.getString( "return_disposition_cd" ));
               this.getReasonDispCodes().add(rdCode);

               //
               //Add this combination, one column for lines & one for amt
               //to the column headings of the report
               line.append( rdCode.getFacility().getName() + " Lines " + rdCode.getReason() + "-" + rdCode.getDisp() + "\t" );
               line.append( rdCode.getFacility().getName() + " Amt " + rdCode.getReason() + "-" + rdCode.getDisp() + "\t" );
            }
         }
         catch( Exception ex ) {
            log.fatal(".buildOutputFile1: ", ex);
            m_ErrMsg.append("The report had the following Error in column build: \r\n");
            m_ErrMsg.append(ex.getClass().getName() + "\r\n" + ex.getMessage());
            
            result = false;
         }
         finally {
            closeResultSet( reasonDispData );
         }

         line.append("\r\n");
         
         try {
            setCurAction( "Executing Customer Query" );
            custData = m_Cust.executeQuery();

            while ( custData.next() && m_Status == RptServer.RUNNING ) {
               //
               //Initialize reused variables:
               sales.clear();
               linesShipped.clear();
               linesRet.clear();
               handling.clear();
               creditAmt.clear();
               linesByCust.clear();
               amtByCust.clear();    
               
               custId = custData.getString( "cust_nbr" );
               setCurAction( "Processing " + custId );

               m_CustStatus.setString( 1, custId );
               statusData = m_CustStatus.executeQuery();

               if ( statusData.next() ) {
                  status = statusData.getString( "status" );
                  salesRep = statusData.getString( "Sales_Rep" );
                  canSub = statusData.getString( "can_substitute" );
                  custName = statusData.getString( "name" );
               }
               
               //
               //Add these fields to the report.
               line.append(custId + "\t");
               line.append(custName + "\t");
               line.append(status + "\t");
               line.append(salesRep + "\t");
               line.append(canSub + "\t");               

               m_Sales.setString( 1, custId );
               salesData = m_Sales.executeQuery();
               
               //
               //Gather sales by Facility data.
               while ( salesData.next() ) {
                  sales.put(this.getFacilityByName(salesData.getString("warehouse")), salesData.getDouble( "sales" ));
               }
               
               //
               //Put sales data in file by Facility.  This must be done this way to ensure column ordering
               //even when data is absent.
               for (Facility f : this.getFacilityList()){
                  line.append(sales.get(f) == null ? "0" : sales.get(f).doubleValue());
                  line.append("\t");
               }               
               
               m_LinesShip.setString( 1, custId );
               linesShipData = m_LinesShip.executeQuery();

               //
               //Gather lines shipped by Facility data.
               while ( linesShipData.next() ) {
                  linesShipped.put(this.getFacilityByName(linesShipData.getString("warehouse")), linesShipData.getInt( "linesship" ));
               }
               
               //
               //Put lines shipped data in file by Facility.  
               //This must be done this way to ensure column ordering even when data is absent.
               for (Facility f : this.getFacilityList()){
                  line.append(linesShipped.get(f) == null ? "0" : linesShipped.get(f).intValue());
                  line.append("\t");
               }                

               m_LinesRet.setString( 1, custId );
               linesRetData = m_LinesRet.executeQuery();
               
               //
               //Gather lines returned data by Facility.
               while ( linesRetData.next() ) {
                  linesRet.put(this.getFacilityByName(linesRetData.getString("warehouse")), linesRetData.getInt( "linesret" ));
               }
               
               //
               //Put lines returned data in file by Facility.  
               //This must be done this way to ensure column ordering even when data is absent.
               for (Facility f : this.getFacilityList()){
                  line.append(linesRet.get(f) == null ? "0" : linesRet.get(f).intValue());
                  line.append("\t");
               }                

               m_Handling.setString( 1, custId );
               handlingData = m_Handling.executeQuery();

               //
               //Gather handling data by Facility.
               while ( handlingData.next() ) {
                  handling.put(this.getFacilityByName(handlingData.getString("warehouse")), handlingData.getDouble( "handling" ));
               }
               
               //
               //Put handling data in file by Facility.  
               //This must be done this way to ensure column ordering even when data is absent.
               for (Facility f : this.getFacilityList()){
                  line.append(handling.get(f) == null ? "0" : handling.get(f).doubleValue());
                  line.append("\t");
               }                

               m_CreditAmt.setString( 1, custId );
               creditAmtData = m_CreditAmt.executeQuery();

               //
               //Gather credit data by Facility.
               while ( creditAmtData.next() ) {
                  creditAmt.put(this.getFacilityByName(creditAmtData.getString("warehouse")), creditAmtData.getDouble( "creditamt" ));
               }
               
               //
               //Put credit data in file by Facility.  
               //This must be done this way to ensure column ordering even when data is absent.
               for (Facility f : this.getFacilityList()){
                  line.append(creditAmt.get(f) == null ? "0" : creditAmt.get(f).doubleValue());
                  line.append("\t");
               }                 

               m_LinesByReason.setString( 1, custId );
               linesByReasonData = m_LinesByReason.executeQuery();
               
               //
               //Gather and store lines by facility, reason, disposition.
               while ( linesByReasonData.next() ) {
                  //
                  //Make sure to get reference to existing reasondispcode object so we can match it to the 
                  //correct column later.
                  linesByCust.put(this.getReasonDispCode(linesByReasonData.getString( "return_reason_cd" ), 
                        linesByReasonData.getString( "return_disposition_cd" ), 
                        this.getFacilityByName(linesByReasonData.getString("warehouse"))), 
                        linesByReasonData.getInt( "lines" ));
               }
               
               //
               //Put return lines in file by Facility, return & disposition.  
               //This must be done this way to ensure column ordering even when data is absent.
               for (ReasonDispCode rd : this.getReasonDispCodes()){
                  line.append(linesByCust.get(rd) == null ? "0" : linesByCust.get(rd).intValue());
                  line.append("\t");
               }                

               m_AmtByReason.setString( 1, custId );
               amtByReasonData = m_AmtByReason.executeQuery();
               
               //
               //Gather and store amount by facility, reason, disposition.
               while ( amtByReasonData.next() ) {
                  //
                  //Make sure to get reference to existing reasondispcode object so we can match it to the 
                  //correct column later.              
                  amtByCust.put(this.getReasonDispCode(amtByReasonData.getString( "return_reason_cd" ), 
                        amtByReasonData.getString( "return_disposition_cd" ), 
                        this.getFacilityByName(amtByReasonData.getString("warehouse"))), 
                        amtByReasonData.getDouble( "amt" ));                  
               }

               //
               //Put return amount in file by Facility, return & disposition.  
               //This must be done this way to ensure column ordering even when data is absent.
               for (ReasonDispCode rd : this.getReasonDispCodes()){
                  line.append(amtByCust.get(rd) == null ? "0" : amtByCust.get(rd).doubleValue());
                  line.append("\t");
               }

               line.append("\r\n");

               // close the result sets
               closeResultSet( statusData );
               closeResultSet( salesData );
               closeResultSet( linesShipData );
               closeResultSet( linesRetData );
               closeResultSet( handlingData );
               closeResultSet( creditAmtData );
               closeResultSet( linesByReasonData );
               closeResultSet( amtByReasonData );
            } // while

            outFile.write(line.toString().getBytes());
            line.delete(0, line.length());
            closeResultSet( custData );
            setCurAction( "Complete" );
         }

         catch( Exception ex ) {
            log.fatal("buildOutputFile: ", ex);
            m_ErrMsg.append("The report had the following Error while building customer data: " + custId + " \r\n");
            m_ErrMsg.append(ex.getClass().getName() + "\r\n" + ex.getMessage());
            result = false;
            setCurAction( "Complete with errors" );
         }
      }

      finally {
         line = null;        
         custId = null;
         status = null;
         salesRep = null;
         canSub = null;
         custData = null;
         statusData = null;
         salesData = null;
         linesShipData = null;
         linesRetData = null;
         handlingData = null;
         creditAmtData = null;
         linesByReasonData = null;
         amtByReasonData = null;
         reasonDispData = null;

        try {
            outFile.close();
            outFile = null;
         }
         
         catch( Exception e ) {
            log.error(e);
         }         
      }
      
      return result;
   }

   /**
    * Performs any cleanup we need to handle.  All variables are closed and set to null
    * here since we are not garunteed to know when finalization occurs.
    */
   protected void cleanup()
   {
      closeStatement( m_Cust );
      closeStatement( m_CustStatus );
      closeStatement( m_Sales );
      closeStatement( m_LinesShip );
      closeStatement( m_LinesRet );
      closeStatement( m_Handling );
      closeStatement( m_CreditAmt );
      closeStatement( m_LinesByReason );
      closeStatement( m_AmtByReason );
      closeStatement( m_ReasonDisp );

      m_Cust = null;
      m_CustStatus = null;
      m_Sales = null;
      m_LinesShip = null;
      m_LinesRet = null;
      m_Handling = null;
      m_CreditAmt = null;
      m_LinesByReason = null;
      m_AmtByReason = null;
      m_ReasonDisp = null;
   }
   
   /**
    * Generic ResultSet close procedure.  Logs any exception.
    * @param data the result set to be closed
    */
   public void closeResultSet(ResultSet data)
   {
      try {
         data.close();
      }
      catch ( Exception ex ) {
      }
   }

   /**
    * Generic PreparedStatment close procedure.  Logs any exception.
    * @param stmt the PreparedStatement to be closed
    */
   public void closeStatement(Statement stmt)
   {
      try {
         if ( stmt != null )
            stmt.close();
      }
      catch ( Exception ex ) {
      }
   }

   /**
    * Creates the report.
    * @see com.emerywaterhouse.rpt.server.Report#createReport()
    */
   public boolean createReport()
   {
      boolean created = false;
      m_Status = RptServer.RUNNING;
      
      try {         
         m_OraConn = m_RptProc.getOraConn();
         
         //
         //In order to use prepared statements with a dynamic list of Facility facilities, 
         //we must know the facilities before preparation. 
         loadFacilities();
         this.setFacilityCondition(buildFacilityCondition());
         
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
    * Loads list of whse facility ids and names.
    * 
    * @throws SQLException
    */
   private void loadFacilities() throws SQLException
   {
      ResultSet rs = null;
      Statement stm = null;
      StringBuffer tmp = new StringBuffer();
      
      try {
         if (m_WhseId == null || m_WhseId.equals("")){
            tmp.append("select fas_facility_id, name from warehouse");
         }
         else{
            tmp.append("select fas_facility_id, name from warehouse where fas_facility_id = '");
            tmp.append(m_WhseId);
            tmp.append("' ");
            //
            //Notice order by is critical as it must match order by in data collection statements
            //for data to line up with column headers.
            tmp.append("order by name'");
         }
            
         this.getFacilityList().clear();
         stm = m_OraConn.createStatement();
         rs = stm.executeQuery(tmp.toString());
         
         while ( rs.next() ) 
            this.getFacilityList().add(new Facility(rs.getString("name"), rs.getString("fas_facility_id")));
      }
      finally {
         closeRSet(rs);
         closeStatement(stm);
         rs = null;
      }
   }   
   
   /**
    * Gets a facilty, by name, from m_FacilityList (a convenience method).
    * 
    * @param name - String facility name
    * @throws Exception
    */
   private Facility getFacilityByName(String name) 
   {
      Facility facility = null;
      
      if (name == null || name.equals(""))
         return facility;
      
      for (Facility f : this.getFacilityList()){
         if (f.getName().equals(name))
            return f;
      }
      
      return facility;
   }    
 
   /**
    * Reads the report parameters (begin date & end date) and prepares
    * the database queries.
    * 
    * @return true if the statements are prepared false if not.
    */
   public boolean prepareStatements()
   {
      StringBuffer sql = new StringBuffer(256);
      boolean isPrepared = false;

      if ( m_OraConn != null ) {         
         try {            
            sql.append("select distinct cust_nbr ");
            sql.append("from inv_hdr ");
            sql.append("where sale_type = 'WAREHOUSE' and ");
            sql.append("      invoice_date >= to_date('" + m_BegDate + "', 'mm/dd/yyyy') and ");
            sql.append("      invoice_date <= to_date('" + m_EndDate + "', 'mm/dd/yyyy') and ");
            sql.append("      warehouse in " + this.getFacilityCondition());
            sql.append(" order by cust_nbr");
      
            m_Cust = m_OraConn.prepareStatement(sql.toString());
      
            sql.setLength(0);
            sql.append("select status, sales_rep, can_substitute, name ");
            sql.append("from cust_view ");
            sql.append("where customer_id = ?");
            m_CustStatus = m_OraConn.prepareStatement(sql.toString());
      
            sql.setLength(0);
            sql.append("select /*+rule*/ warehouse, sum(dollars_shipped) sales from inv_hdr ");
            sql.append("where cust_nbr = ? and ");
            sql.append("      sale_type = 'WAREHOUSE' and ");
            sql.append("      invoice_date >= to_date('" + m_BegDate + "', 'mm/dd/yyyy') and ");
            sql.append("      invoice_date <= to_date('" + m_EndDate + "', 'mm/dd/yyyy') and ");
            sql.append("      warehouse in " + this.getFacilityCondition());
            sql.append(" group by warehouse ");
            m_Sales = m_OraConn.prepareStatement(sql.toString());
      
            sql.setLength(0);
            sql.append("select /*+rule*/ warehouse, sum(lines_shipped) linesship from inv_hdr " );
            sql.append("where cust_nbr = ? and ");
            sql.append("      sale_type = 'WAREHOUSE' and ");
            sql.append("      tran_type = 'SALE' and ");
            sql.append("      invoice_date >= to_date('" + m_BegDate + "', 'mm/dd/yyyy') and ");
            sql.append("      invoice_date <= to_date('" + m_EndDate + "', 'mm/dd/yyyy') and ");
            sql.append("      warehouse in " + this.getFacilityCondition());    
            sql.append(" group by warehouse ");
            m_LinesShip = m_OraConn.prepareStatement(sql.toString());
      
            sql.setLength(0);
            sql.append("select /*+rule*/ warehouse, sum(lines_shipped) linesret from inv_hdr ");
            sql.append("where cust_nbr = ? and ");
            sql.append("      sale_type = 'WAREHOUSE' and ");
            sql.append("      tran_type in ('CREDIT','RETURN') and ");
            sql.append("      invoice_date >= to_date('" + m_BegDate + "', 'mm/dd/yyyy') and ");
            sql.append("      invoice_date <= to_date('" + m_EndDate + "', 'mm/dd/yyyy') and ");
            sql.append("      warehouse in " + this.getFacilityCondition());   
            sql.append(" group by warehouse ");
            m_LinesRet = m_OraConn.prepareStatement(sql.toString());
      
            sql.setLength(0);
            sql.append("select /*+rule*/ inv_hdr.warehouse, inv_hdr.warehouse, sum(adder_amount) handling ");
            sql.append("from inv_hdr, inv_adder ");
            sql.append("where inv_hdr.cust_nbr = ? and ");
            sql.append("      inv_hdr.sale_type in ('WAREHOUSE', 'ACE DIRECT') and ");
            sql.append("      inv_hdr.tran_type in ('CREDIT','RETURN') and ");
            sql.append("      inv_hdr.invoice_date >= to_date('" + m_BegDate + "', 'mm/dd/yyyy') and ");
            sql.append("      inv_hdr.invoice_date <= to_date('" + m_EndDate + "', 'mm/dd/yyyy') and ");
            sql.append("      inv_hdr.warehouse in " + this.getFacilityCondition() + " and ");            
            sql.append("      inv_adder.inv_hdr_id = inv_hdr.inv_hdr_id and ");
            sql.append("      adder_descr = 'HANDLING'");
            sql.append("group by inv_hdr.warehouse ");
            m_Handling = m_OraConn.prepareStatement(sql.toString());
      
            sql.setLength(0);
            sql.append("select /*+rule*/ inv_hdr.warehouse, sum(dollars_shipped) creditamt ");
            sql.append("from inv_hdr ");
            sql.append("where inv_hdr.cust_nbr = ? and ");
            sql.append("      inv_hdr.sale_type in ('WAREHOUSE', 'ACE DIRECT') and ");
            sql.append("      inv_hdr.tran_type in ('CREDIT','RETURN') and ");
            sql.append("      inv_hdr.invoice_date >= to_date('" + m_BegDate + "', 'mm/dd/yyyy') and ");
            sql.append("      inv_hdr.invoice_date <= to_date('" + m_EndDate + "', 'mm/dd/yyyy') and ");
            sql.append("      inv_hdr.warehouse in " + this.getFacilityCondition());       
            sql.append(" group by inv_hdr.warehouse ");
            m_CreditAmt = m_OraConn.prepareStatement(sql.toString());
      
            sql.setLength(0);
            sql.append("select /*+rule*/ count(*) lines, warehouse, return_reason_cd, return_disposition_cd ");
            sql.append("from inv_dtl ");
            sql.append("where inv_dtl.cust_nbr = ? and ");
            sql.append("      inv_dtl.sale_type = 'WAREHOUSE' and ");
            sql.append("      inv_dtl.tran_type in ('CREDIT','RETURN') and ");
            sql.append("      inv_dtl.invoice_date >= to_date('" + m_BegDate + "', 'mm/dd/yyyy') and ");
            sql.append("      inv_dtl.invoice_date <= to_date('" + m_EndDate + "', 'mm/dd/yyyy') and ");
            sql.append("      inv_dtl.warehouse in " + this.getFacilityCondition());            
            sql.append(" group by warehouse, return_reason_cd, return_disposition_cd ");
            m_LinesByReason = m_OraConn.prepareStatement(sql.toString());
      
            sql.setLength(0);
            sql.append("select /*+rule*/ sum(ext_sell) amt, warehouse, return_reason_cd, return_disposition_cd ");
            sql.append("from inv_dtl ");
            sql.append("where inv_dtl.cust_nbr = ? and ");
            sql.append("      inv_dtl.sale_type = 'WAREHOUSE' and ");
            sql.append("      inv_dtl.tran_type in ('CREDIT','RETURN') and ");
            sql.append("      inv_dtl.invoice_date >= to_date('" + m_BegDate + "', 'mm/dd/yyyy') and ");
            sql.append("      inv_dtl.invoice_date <= to_date('" + m_EndDate + "', 'mm/dd/yyyy') and ");
            sql.append("      inv_dtl.warehouse in " + this.getFacilityCondition());
            sql.append(" group by warehouse, return_reason_cd, return_disposition_cd ");
            m_AmtByReason = m_OraConn.prepareStatement(sql.toString());
      
            sql.setLength(0);
            sql.append("select distinct warehouse, return_reason_cd, return_disposition_cd ");
            sql.append("from inv_dtl ");
            sql.append("where sale_type = 'WAREHOUSE' and ");
            sql.append("      tran_type in ('CREDIT','RETURN') and ");
            sql.append("      invoice_date >= to_date('" + m_BegDate + "', 'mm/dd/yyyy') and ");
            sql.append("      invoice_date <= to_date('" + m_EndDate + "', 'mm/dd/yyyy') and ");
            sql.append("      warehouse in " + this.getFacilityCondition());            
            sql.append(" order by warehouse, return_reason_cd, return_disposition_cd ");
            m_ReasonDisp = m_OraConn.prepareStatement(sql.toString());
            
            isPrepared = true;
         }
         
         catch ( Exception ex ) {
            log.fatal("exception:", ex);
         }
         
         finally {
            sql = null;
         }
      }
      
      return isPrepared;
   }
   

   /**
    * @return ArrayList<Facility> - the list of facility for current report.
    */
   public List<Facility> getFacilityList() {
      return m_FacilityList;
   }

   /**
    * @param facilityList - ArrayList<Facility>, the list of facility for current report.
    */
   public void setFacilityList(ArrayList<Facility> facilityList) {
      m_FacilityList = facilityList;
   }   

   /**
    * param(0) = begin date
    * param(1) = end date
    * 
    * Note - The report file name gets set here.
    * 
    * @see com.emerywaterhouse.rpt.server.Report#setParams(java.util.ArrayList)
    */
   public void setParams(ArrayList<Param> params)
   {
      StringBuffer tmp = new StringBuffer();
      
      try {
         
         // 
         //processes user parameters from EIS       
         for (Param p : params) {
            
            if ( p.name.equals("begdate"))
               m_BegDate = p.value.trim();
            
            if ( p.name.equals("enddate"))
               m_EndDate = p.value.trim();
            
            if ( p.name.equals("dc"))
               m_WhseId = p.value.trim();
         }
         
         tmp.append("returns");
         tmp.append(m_BegDate.substring(0,2));
         tmp.append(m_BegDate.substring(3,5));
         tmp.append(m_BegDate.substring(6));
         tmp.append("-");
         tmp.append(m_EndDate.substring(0,2));
         tmp.append(m_EndDate.substring(3,5));
         tmp.append(m_EndDate.substring(6));
         tmp.append(".dat");
         
         m_FileNames.add(tmp.toString());
      }
      catch(Exception e){
         log.fatal("ItemReturnAnalysis.setParams ", e);
         m_ErrMsg.append("The report had the following Error(s) in: \r\n");
         m_ErrMsg.append(e.getClass().getName() + "\r\n" + e.getMessage());
      }
      
      finally {
         tmp = null;
      }
   }
   
   

   /**
    * @return String - facility sql condition fragment.
    */
   public String getFacilityCondition() {
      return m_FacilityCondition;
   }

   /**
    * @param facilityCondition - String facility condition fragment.
    */
   public void setFacilityCondition(String facilityCondition) {
      m_FacilityCondition = facilityCondition;
   }

   /**
    * @return List<ReasonDispCode> - list of reason disp codes objects. 
    */
   public List<ReasonDispCode> getReasonDispCodes() {
      return m_ReasonDispCodes;
   }
   
   /**
    * @param reason 
    * @param disp 
    * @param f 
    * @return ReasonDispCode - reason disp code object from current list. 
    */
   public ReasonDispCode getReasonDispCode(String reason, String disp, Facility f) {
      ReasonDispCode rdc = null;
      
      if (reason == null || disp == null || f == null)
         return rdc;
      //
      //Return reference to specified code in list if it exists
      for (ReasonDispCode code : this.getReasonDispCodes()){
         if ( reason.equals(code.getReason()) && disp.equals(code.getDisp()) && f.equals(code.getFacility()) ){
            return code;
         }
            
      }
      return rdc;
   }   

   /**
    * @param reasonDispCodes - List<ReasonDispCode>, list of reason disp codes objects.
    */
   public void setReasonDispCodes(List<ReasonDispCode> reasonDispCodes) {
      m_ReasonDispCodes = reasonDispCodes;
   }

}
